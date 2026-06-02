#!/usr/bin/env python3
"""
库存金额跟踪系统 (Inventory Tracking System)
============================================
功能：
1. 从上月底Excel文件读取期初库存余额（支持 .xlsx / .xlsb）
2. 连接SQL Server查询当月MTD物料事务（使用 pymssql，无需 ODBC 驱动）
3. 按TransType+RefType分类：Received / Consumed / Other Transaction
4. 按Item汇总Qty和AMT（金额直接使用 TotalPosted）
5. 计算期末余额 = 期初 + Received + Consumed + Other（按源数据符号）
6. 导出Excel报表（Project Summary + Summary + Detail 三页，Project Summary 在前）
7. 支持多站点（310/330/410）批量运行
8. 通过 SMTP 发送邮件（支持内网 Relay，无需 Outlook）
9. 通过 .env 外部化配置（DB、SMTP、邮件收件人、Infor API）

作者: WorkBuddy for NAI Group
日期: 2026-05-25 | 更新: 2026-05-31 (移除 config.ini，统一用 .env 配置)
"""

import os
import sys
import time
import glob
import warnings
import argparse
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple
from pathlib import Path

import json
import urllib.request
import urllib.error
import urllib.parse
import subprocess

import pymssql
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter


# ──────────────────────────────────────────────────────────────
# TransType + RefType → Category 映射规则
# ──────────────────────────────────────────────────────────────
CATEGORY_MAP: Dict[Tuple[str, str], str] = {
    # Received (增加库存)
    ("R", "P"): "Received",   # 采购收料
    ("C", "J"): "Received",   # 工单完工
    ("F", "J"): "Received",   # 工单完工
    ("W", "J"): "Received",   # 工单退料（退回仓库）
    ("W", "R"): "Received",   # RMA 退货（退回仓库）

    # Consumed (减少库存)
    ("I", "J"): "Consumed",   # 工单发料
    ("S", "O"): "Consumed",   # 客户订单发货
    ("W", "P"): "Consumed",   # 采购退料（退回供应商）

    # Other Transaction (库存调整)
    ("A", "I"): "Other",      # 库存调整
    ("M", "I"): "Other",      # 库存调整
    ("G", "I"): "Other",      # 库存调整
    ("H", "I"): "Other",      # 库存调整
    ("N", "J"): "Ignored",    # 工单工序转移 — 不计入任何MTD列
}

TRANS_DESCRIPTIONS: Dict[Tuple[str, str], str] = {
    ("R", "P"): "PO Receipt",
    ("C", "J"): "Job Complete",
    ("F", "J"): "Job Finish",
    ("W", "J"): "Job Return to Whse",
    ("W", "R"): "RMA Return",
    ("I", "J"): "Job Issue",
    ("S", "O"): "Customer Shipment",
    ("W", "P"): "PO Return to Vendor",
    ("A", "I"): "Inventory Adjustment",
    ("M", "I"): "Inventory Adjustment",
    ("G", "I"): "Inventory Adjustment",
    ("H", "I"): "Inventory Adjustment",
    ("N", "J"): "Job Transfer (Ignored)",
}

SITE_NAMES = {"310": "Plant1", "330": "Plant2", "410": "PNG"}
# Site 330 的原始货币是 CNY，需要按汇率转为 USD
SITE_CURRENCY = {"310": ("USD", 1.0), "330": ("CNY->USD", 1 / 6.838784), "410": ("USD", 1.0)}


@dataclass
class ItemBalance:
    item: str
    description: str = ""
    project_code: str = ""
    prev_balance: float = 0.0
    prev_qty: float = 0.0
    recv_qty: float = 0.0
    recv_amt: float = 0.0
    cons_qty: float = 0.0
    cons_amt: float = 0.0
    other_qty: float = 0.0
    other_amt: float = 0.0

    @property
    def balance(self) -> float:
        return self.prev_balance + self.recv_amt + self.cons_amt + self.other_amt

    @property
    def net_qty(self) -> float:
        return self.prev_qty + self.recv_qty + self.cons_qty + self.other_qty


# ──────────────────────────────────────────────────────────────
# 样式常量
# ──────────────────────────────────────────────────────────────
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14)
BALANCE_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
TOTAL_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)


class InventoryTracker:
    def __init__(
        self,
        server: Optional[str] = None,
        database: Optional[str] = None,
        username: Optional[str] = None,
        password: Optional[str] = None,
        port: int = 1433,
        site_ref: str = "310",
        prev_balance_file: Optional[str] = None,
        output_file: Optional[str] = None,
        as_of_date: Optional[str] = None,
    ):
        # DB 参数统一从调用参数或 .env 读取，禁止硬编码默认账号密码
        self.server = server or os.environ.get("SQL_SERVER_HOST", "")
        self.database = database or os.environ.get("SQL_SERVER_DATABASE", "")
        self.username = username or os.environ.get("SQL_SERVER_USERNAME", "")
        self.password = password or os.environ.get("SQL_SERVER_PASSWORD", "")
        self.port = port
        self.site_ref = site_ref
        self.prev_balance_file = prev_balance_file
        if as_of_date:
            try:
                self.report_date = datetime.strptime(as_of_date, "%Y-%m-%d")
            except ValueError as e:
                raise ValueError("as_of_date 格式错误，应为 YYYY-MM-DD") from e
        else:
            self.report_date = datetime.today()

        # Currency: site 330 is CNY, convert to USD
        curr_info = SITE_CURRENCY.get(site_ref, ("USD", 1.0))
        self.currency_label = curr_info[0]
        self.fx_rate = curr_info[1]
        self.usd_symbol = "$"

        site_label = SITE_NAMES.get(site_ref, site_ref)
        today = self.report_date
        self.period_str = today.strftime("%Y-%m")
        default_name = f"Inventory_Balance_{site_label}_{today.strftime('%Y%m')}.xlsx"
        self.output_file = output_file or default_name

    def _get_connection(self):
        """使用 pymssql 建立数据库连接（无需 ODBC 驱动）"""
        missing = []
        if not self.server:
            missing.append("SQL_SERVER_HOST")
        if not self.database:
            missing.append("SQL_SERVER_DATABASE")
        if not self.username:
            missing.append("SQL_SERVER_USERNAME")
        if not self.password:
            missing.append("SQL_SERVER_PASSWORD")
        if missing:
            raise RuntimeError(f"缺少数据库配置: {', '.join(missing)}（请检查 .env）")

        # SQL Server 命名实例格式：host\instance，pymssql 需拆分
        host = self.server
        instance = None
        if "\\" in host:
            host, instance = host.split("\\", 1)

        kwargs = dict(
            server=host,
            user=self.username,
            password=self.password,
            database=self.database,
            port=self.port,
            tds_version="7.4",
            login_timeout=15,
            conn_properties="SET TEXTSIZE 65536",
        )
        if instance:
            # pymssql 通过 server="host\\instance" 格式支持命名实例
            kwargs["server"] = f"{host}\\{instance}"
            del kwargs["port"]  # 命名实例时不指定端口，由 SQL Browser 解析

        return pymssql.connect(**kwargs)

    # ──────────────────────────────────────────────────────────
    # 1. 读取期初库存（Excel）
    # ──────────────────────────────────────────────────────────
    def load_previous_balance(self) -> pd.DataFrame:
        """
        读取上月末余额。优先从数据库获取，fallback 到 Excel 文件。
        数据库方式：从 SLTotalInventory 按 BalanceDate=上月月末日期读取 RM/FG 汇总。
        Excel 方式：支持三种列名格式（Prev_Balance / Unitscost / Unitcost）。
        """
        # ── 优先从数据库读取 ──
        if not self.prev_balance_file:
            return self._load_prev_from_db()

        if not os.path.exists(self.prev_balance_file):
            print(f"  ⚠️  期初余额文件未找到：{self.prev_balance_file}")
            print("     将尝试从数据库读取期初余额...")
            return self._load_prev_from_db()

        filepath = self.prev_balance_file
        def _normalize_cols(cols):
            return [str(c).strip() for c in cols]

        def _tok(v):
            return str(v).strip().lower().replace(" ", "").replace("_", "")

        def _canonicalize_columns(df_in: pd.DataFrame) -> pd.DataFrame:
            rename_map = {}
            for c in df_in.columns:
                t = _tok(c)
                if t == "item":
                    rename_map[c] = "Item"
                elif t == "prevqty":
                    rename_map[c] = "Prev_Qty"
                elif t == "prevbalance":
                    rename_map[c] = "Prev_Balance"
                elif t == "per":
                    rename_map[c] = "Per"
                elif t == "unitscost":
                    rename_map[c] = "Unitscost"
                elif t == "unitcost":
                    rename_map[c] = "Unitcost"
                elif t == "description":
                    rename_map[c] = "Description"
                elif t == "uegdldescription":
                    rename_map[c] = "ue_GDL_Description"
            return df_in.rename(columns=rename_map)

        def _has_supported_cols(cols):
            cset = set(_normalize_cols(cols))
            return (
                ("Item" in cset and "Prev_Qty" in cset and "Prev_Balance" in cset)
                or ("Item" in cset and "Per" in cset and "Unitscost" in cset)
                or ("Item" in cset and "Per" in cset and "Unitcost" in cset)
            )

        # 先按默认表头读取；若失败则在所有sheet扫描前60行定位真实表头
        xls = pd.ExcelFile(filepath)
        df = None
        for sheet in xls.sheet_names:
            candidate = pd.read_excel(filepath, sheet_name=sheet)
            candidate.columns = _normalize_cols(candidate.columns)
            candidate = _canonicalize_columns(candidate)
            if _has_supported_cols(candidate.columns):
                df = candidate
                break

            raw = pd.read_excel(filepath, sheet_name=sheet, header=None, nrows=80)
            header_row = None
            scan_rows = min(60, len(raw))
            for ridx in range(scan_rows):
                row_vals = [_tok(v) for v in raw.iloc[ridx].tolist() if str(v).strip()]
                has_item = "item" in row_vals
                has_std = ("prevqty" in row_vals and "prevbalance" in row_vals)
                has_unit = ("per" in row_vals and "unitcost" in row_vals)
                has_units = ("per" in row_vals and "unitscost" in row_vals)
                if has_item and (has_std or has_unit or has_units):
                    header_row = ridx
                    break

            if header_row is not None:
                candidate = pd.read_excel(filepath, sheet_name=sheet, skiprows=header_row)
                candidate.columns = _normalize_cols(candidate.columns)
                candidate = _canonicalize_columns(candidate)
                if _has_supported_cols(candidate.columns):
                    df = candidate
                    break

        if df is None:
            # 用第一个sheet做错误展示，便于快速定位输入文件问题
            df = pd.read_excel(filepath, sheet_name=xls.sheet_names[0])
            df.columns = _normalize_cols(df.columns)
            df = _canonicalize_columns(df)

        # 智能识别列名并计算期初金额
        if "Prev_Qty" in df.columns and "Prev_Balance" in df.columns:
            # 标准格式
            qty_col, amt_col = "Prev_Qty", "Prev_Balance"
            calc_extended = False
        elif "Per" in df.columns and "Unitscost" in df.columns:
            # Infor API 格式：Unitscost 是扩展金额（= Units × Unitcost，由 API 直接提供）
            # Prev_Balance 直接使用 Unitscost，无需计算
            qty_col, amt_col = "Per", "Unitscost"
            calc_extended = False
        elif "Per" in df.columns and "Unitcost" in df.columns:
            # 新Infor格式：Unitcost 是单价，需要 Per × Unitcost
            qty_col = "Per"
            calc_extended = True
            print("  ℹ️  检测到 Unitcost（单价），将计算 Per × Unitcost 作为期初金额")
        else:
            raise ValueError(
                f"无法识别期初余额列名。现有列：{list(df.columns)}\n"
                f"需要包含 (Item + Prev_Qty + Prev_Balance) 或 (Item + Per + Unitcost/Unitscost)"
            )

        df["Item"] = df["Item"].astype(str).str.strip().str.upper()
        df = df[df["Item"] != "NAN"].copy()
        df["Prev_Qty"] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)

        if calc_extended:
            unit_cost = pd.to_numeric(df["Unitcost"], errors="coerce").fillna(0)
            df["Prev_Balance"] = df["Prev_Qty"] * unit_cost
        else:
            df["Prev_Balance"] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0)

        # 尝试保留 Description
        desc_col = None
        for c in ["Description", "ue_GDL_Description"]:
            if c in df.columns:
                desc_col = c
                break
        if desc_col:
            df["Description"] = df[desc_col].astype(str).str.strip()
        else:
            df["Description"] = ""

        df = df[["Item", "Prev_Qty", "Prev_Balance", "Description"]].copy()

        # Currency conversion for non-USD sites
        if self.fx_rate != 1.0:
            df["Prev_Balance"] = df["Prev_Balance"] * self.fx_rate
            print(f"  💱 汇率转换 {self.currency_label} ÷ 6.838784 = USD")

        total_bal = df["Prev_Balance"].sum()
        print(f"  ✅ 已加载 {len(df)} 个Item，期初总金额: ${total_bal:,.2f}")
        return df

    def _load_prev_from_db(self) -> pd.DataFrame:
        """
        从数据库 SLTotalInventory 读取上月末 RM/FG 汇总金额，
        构建与 Excel 方式兼容的 prev_df（按 Item 粒度）。
        如果数据库查询失败或无数据，返回空 DataFrame（期初为 0）。
        """
        prev_date = (self.report_date.replace(day=1) - timedelta(days=1)).strftime("%Y-%m-%d")
        site_label = SITE_NAMES.get(self.site_ref, self.site_ref)
        print(f"  🗄️  从数据库读取期初余额 (Site {self.site_ref}, BalanceDate={prev_date})...")

        try:
            conn = self._get_connection()
            import warnings as _w
            with _w.catch_warnings():
                _w.simplefilter("ignore")
                query = f"""
                SELECT
                    i.[Item],
                    i.[Source],
                    ISNULL(i.[Per], 0) AS Per,
                    ISNULL(i.[Unitscost], 0) AS Unitscost
                FROM dbo.SLTotalInventory i
                WHERE i.[SiteRef] = '{self.site_ref}'
                  AND i.[BalanceDate] = '{prev_date}'
                  AND ISNULL(i.[Unitscost], 0) <> 0
                ORDER BY i.[Item]
                """
                df = pd.read_sql(query, conn)
            conn.close()
        except Exception as e:
            print(f"  ⚠️  数据库期初查询失败: {e}")
            return pd.DataFrame(columns=["Item", "Prev_Qty", "Prev_Balance", "Description"])

        if df.empty:
            print(f"  ⚠️  未找到期初数据 (Site {self.site_ref}, {prev_date})，期初余额设为 0")
            return pd.DataFrame(columns=["Item", "Prev_Qty", "Prev_Balance", "Description"])

        df["Item"] = df["Item"].astype(str).str.strip().str.upper()
        df = df[df["Item"] != "NAN"].copy()
        df["Prev_Qty"] = pd.to_numeric(df["Per"], errors="coerce").fillna(0)
        df["Prev_Balance"] = pd.to_numeric(df["Unitscost"], errors="coerce").fillna(0)

        # Currency conversion for non-USD sites
        if self.fx_rate != 1.0:
            df["Prev_Balance"] = df["Prev_Balance"] * self.fx_rate
            print(f"  💱 汇率转换 {self.currency_label} ÷ 6.838784 = USD")

        # Description 不可用于 SLTotalInventory，直接设空
        df["Description"] = ""

        df = df[["Item", "Prev_Qty", "Prev_Balance", "Description"]].copy()

        total_bal = df["Prev_Balance"].sum()
        print(f"  ✅ 已加载 {len(df)} 个Item，期初总金额: ${total_bal:,.2f}")
        return df

    # ──────────────────────────────────────────────────────────
    def fetch_mtd_transactions(self) -> pd.DataFrame:
        today = self.report_date
        prev_month_end = today.replace(day=1) - timedelta(days=1)

        query = f"""
        SELECT
            m.[SiteRef], m.[TransNum], m.[TransDate], m.[TransType], m.[RefType], m.[Backflush],
            m.[Whse], m.[Loc], m.[Lot], m.[Wc], m.[RefNum], m.[RefLineSuf], m.[RefRelease],
            m.[Item], m.[ue_GDL_Description],
            m.[Qty], m.[TotalPosted],
            m.[RowPointer], m.[RecordDate], m.[DocumentNumber],
            dbo.GET_Customer_Formate_ProjectCode(
                CAST(m.[SiteRef] AS int), m.[Item]
            ) AS ProjectCode
        FROM [csi_datawarehouse].[dbo].[SLMatltrans] m
        WHERE m.[SiteRef] = '{self.site_ref}'
          AND m.[TransDate] > '{prev_month_end.strftime('%Y-%m-%d')}'
          AND m.[TransDate] <= '{today.strftime('%Y-%m-%d')}'
        ORDER BY m.[TransDate], m.[TransNum]
        """

        print(f"  📊 查询 {self.database} Site {self.site_ref} 所有物料 (P+M) ({prev_month_end.strftime('%Y-%m-%d')} < TransDate <= {today.strftime('%Y-%m-%d')})")

        try:
            conn = self._get_connection()
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                df = pd.read_sql(query, conn)
            conn.close()
        except Exception as e:
            print(f"  ❌ 数据库连接失败：{e}")
            sys.exit(1)

        df["Item"] = df["Item"].astype(str).str.strip().str.upper()
        df = df[df["Item"] != "NAN"].copy()
        df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(0)
        df["TotalAmt"] = pd.to_numeric(df["TotalPosted"], errors="coerce").fillna(0)

        # Currency conversion for non-USD sites
        if self.fx_rate != 1.0:
            df["TotalAmt"] = df["TotalAmt"] * self.fx_rate

        print(f"  ✅ 查询到 {len(df):,} 条事务记录，{df['Item'].nunique():,} 个Item")
        return df

    # ──────────────────────────────────────────────────────────
    # 3. 分类
    # ──────────────────────────────────────────────────────────
    def classify(self, row: pd.Series) -> str:
        key = (str(row.get("TransType", "")).strip().upper(),
               str(row.get("RefType", "")).strip().upper())
        # 严格口径：Other 仅来自 CATEGORY_MAP 显式定义；未映射事务不计入任何 MTD 列
        return CATEGORY_MAP.get(key, "Ignored")

    def _desc(self, tt: str, rt: str) -> str:
        return TRANS_DESCRIPTIONS.get((str(tt).strip().upper(), str(rt).strip().upper()),
                                      f"{tt}/{rt}")

    # ──────────────────────────────────────────────────────────
    # 4. 按 Item 汇总
    # ──────────────────────────────────────────────────────────
    def calculate_balances(self, prev_df: pd.DataFrame, trans_df: pd.DataFrame):
        trans_df["Category"] = trans_df.apply(self.classify, axis=1)
        trans_df["TransDesc"] = trans_df.apply(
            lambda r: self._desc(r["TransType"], r["RefType"]), axis=1
        )
        ignored_count = int((trans_df["Category"] == "Ignored").sum())
        if ignored_count:
            print(f"  ℹ️  Ignored 事务: {ignored_count:,} 条（未映射或 N/J，不计入 MTD）")

        # Detail
        detail_cols = [
            "Item", "TransDate", "TransNum", "TransType", "RefType",
            "Category", "TransDesc", "Qty", "TotalAmt",
            "Whse", "Loc", "Lot", "RefNum", "RefLineSuf",
        ]
        available = [c for c in detail_cols if c in trans_df.columns]
        detail_df = trans_df[available].copy()
        detail_df["TransDate"] = pd.to_datetime(detail_df["TransDate"]).dt.strftime("%Y-%m-%d")

        # Build description + ProjectCode lookup from prev_df
        desc_map = {}
        if "Description" in prev_df.columns:
            for _, row in prev_df.iterrows():
                desc_map[str(row["Item"])] = row["Description"]

        # Build ProjectCode lookup from transaction data (via GET_Customer_Formate_ProjectCode)
        project_code_map: Dict[str, str] = {}
        if "ProjectCode" in trans_df.columns:
            for _, row in trans_df.iterrows():
                pc = str(row.get("ProjectCode", "")).strip()
                if pc and pc != "NAN" and pc != "NONE":
                    project_code_map[str(row["Item"])] = pc

        # 查询 SLItems 获取 PMTCode（P=采购物料/RM，M=制造件/FG）
        pmtcode_map: Dict[str, str] = {}
        try:
            conn = pymssql.connect(
                server=self.server, database=self.database,
                user=self.username, password=self.password, port=self.port,
            )
            cursor = conn.cursor(as_dict=True)
            cursor.execute("""
                SELECT Item, PMTCode
                FROM [csi_datawarehouse].[dbo].[SLItems]
                WHERE SiteRef = %s
            """, (self.site_ref,))
            for row in cursor.fetchall():
                item_key = str(row["Item"]).strip().upper()
                pmtcode_map[item_key] = str(row["PMTCode"]).strip() if row["PMTCode"] else ""
            conn.close()
        except Exception as e:
            print(f"  ⚠️  查询 PMTCode 失败：{e}，按 'P' 处理")

        # Summary aggregation
        items: Dict[str, ItemBalance] = {}
        for _, row in prev_df.iterrows():
            item = str(row["Item"])
            items[item] = ItemBalance(
                item=item,
                prev_balance=row["Prev_Balance"],
                prev_qty=row["Prev_Qty"],
                description=desc_map.get(item, ""),
                project_code=project_code_map.get(item, ""),
            )

        for _, row in trans_df.iterrows():
            item = str(row["Item"])
            cat = row["Category"]
            qty, amt = row["Qty"], row["TotalAmt"]
            if item not in items:
                items[item] = ItemBalance(item=item, project_code=project_code_map.get(item, ""))
            ib = items[item]
            if cat == "Received":
                ib.recv_qty += qty; ib.recv_amt += amt
            elif cat == "Consumed":
                ib.cons_qty += qty; ib.cons_amt += amt
            elif cat == "Other":
                ib.other_qty += qty; ib.other_amt += amt
            # "Ignored" 类别（如 N/J 工单转移）直接跳过，不计入任何 MTD 列

        summary_data = []
        for item, ib in items.items():
            pmt = pmtcode_map.get(item, "")
            summary_data.append({
                "ProjectCode": ib.project_code,
                "Item": item,
                "Description": ib.description,
                "PMTCode": pmt,
                "Prev_Qty": round(ib.prev_qty, 4),
                "Prev_Balance": round(ib.prev_balance, 2),
                "Received_Qty": round(ib.recv_qty, 4),
                "Received_AMT": round(ib.recv_amt, 2),
                "Consumed_Qty": round(ib.cons_qty, 4),
                "Consumed_AMT": round(ib.cons_amt, 2),
                "Other_Qty": round(ib.other_qty, 4),
                "Other_AMT": round(ib.other_amt, 2),
                "Balance_Qty": round(ib.net_qty, 4),
                "Balance_AMT": round(ib.balance, 2),
            })

        summary_df = pd.DataFrame(summary_data)
        # 按期末余额从大到小排序
        summary_df = summary_df.sort_values("Balance_AMT", ascending=False).reset_index(drop=True)
        return summary_df, detail_df

    # ──────────────────────────────────────────────────────────
    # 5. 导出 Excel
    # ──────────────────────────────────────────────────────────
    def export_excel(self, summary_df: pd.DataFrame, detail_df: pd.DataFrame) -> None:
        output_path = Path(self.output_file)
        today = self.report_date
        site_label = SITE_NAMES.get(self.site_ref, self.site_ref)
        prev_month_end = today.replace(day=1) - timedelta(days=1)
        prev_label = f"{prev_month_end.month}/{prev_month_end.day}"

        # ── 先用 pandas 写入原始数据（快速，无样式）──
        print(f"  📝 导出 ({len(summary_df):,} summary + {len(detail_df):,} detail rows)...")

        # 创建空白 workbook 并手动写入，确保完全控制行位置
        wb = Workbook()

        # ═══════════════════════════════════════════════════════════
        # Sheet 1: Project Code Summary (按项目代码汇总) ← 放第一个
        # ═══════════════════════════════════════════════════════════
        ws_proj = wb.active
        ws_proj.title = "Project Summary"

        # 按 ProjectCode 分组汇总（不含 Items 计数）
        proj_group = summary_df.groupby("ProjectCode", dropna=False).agg(
            Prev_Balance=("Prev_Balance", "sum"),
            Received_AMT=("Received_AMT", "sum"),
            Consumed_AMT=("Consumed_AMT", "sum"),
            Other_AMT=("Other_AMT", "sum"),
            Balance_AMT=("Balance_AMT", "sum"),
        ).reset_index()
        proj_group = proj_group.sort_values("Balance_AMT", ascending=False).reset_index(drop=True)
        proj_group["ProjectCode"] = proj_group["ProjectCode"].fillna("")

        # Row 1: 标题
        ws_proj.merge_cells("A1:H1")
        ws_proj["A1"] = f"按项目代码汇总 - Site {self.site_ref} ({site_label})"
        ws_proj["A1"].font = TITLE_FONT
        ws_proj["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws_proj.row_dimensions[1].height = 28

        # Row 2: 表头
        proj_headers = ["Project Code", f"{prev_label} Balance",
                        "MTD Received", "MTD Consumption", "MTD Other Transaction",
                        "MTD Daily Balance", "公式", ""]
        for col_idx, col_name in enumerate(proj_headers, 1):
            c = ws_proj.cell(row=2, column=col_idx, value=col_name)
            c.fill = HEADER_FILL
            c.font = HEADER_FONT
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = THIN_BORDER
        ws_proj.merge_cells("G2:H2")
        ws_proj.row_dimensions[2].height = 25

        # Row 3+: 数据行
        for row_data in proj_group.itertuples(index=False):
            ws_proj.append(list(row_data))

        # 为每行添加公式说明
        data_start = 3
        data_end = ws_proj.max_row
        for r in range(data_start, data_end + 1):
            ws_proj.cell(row=r, column=7, value=f"=B{r}+C{r}+D{r}+E{r}")
            ws_proj.cell(row=r, column=7).number_format = "#,##0.00"
            ws_proj.cell(row=r, column=8, value=f"{prev_label}+Recv+Cons+Other")
            ws_proj.cell(row=r, column=8).font = Font(size=9, color="666666")
            for col_idx in [2, 3, 4, 5, 6]:
                ws_proj.cell(row=r, column=col_idx).number_format = "#,##0.00"

        # Total row
        total_r = data_end + 1
        ws_proj.cell(row=total_r, column=1, value="TOTAL")
        ws_proj.cell(row=total_r, column=1).font = Font(bold=True)
        ws_proj.cell(row=total_r, column=1).fill = TOTAL_FILL
        ws_proj.cell(row=total_r, column=1).border = THIN_BORDER
        proj_totals = {
            2: proj_group["Prev_Balance"].sum(),
            3: proj_group["Received_AMT"].sum(),
            4: proj_group["Consumed_AMT"].sum(),
            5: proj_group["Other_AMT"].sum(),
            6: proj_group["Balance_AMT"].sum(),
        }
        for col_idx, total_val in proj_totals.items():
            c = ws_proj.cell(row=total_r, column=col_idx)
            c.value = round(total_val, 2)
            c.font = Font(bold=True)
            c.fill = TOTAL_FILL
            c.border = THIN_BORDER
            c.number_format = "#,##0.00"
        t_prev = proj_totals[2]; t_recv = proj_totals[3]; t_cons = proj_totals[4]; t_other = proj_totals[5]
        ws_proj.cell(row=total_r, column=7, value=round(t_prev + t_recv - t_cons - t_other, 2))
        ws_proj.cell(row=total_r, column=7).font = Font(bold=True)
        ws_proj.cell(row=total_r, column=7).fill = TOTAL_FILL
        ws_proj.cell(row=total_r, column=7).number_format = "#,##0.00"

        proj_widths = [18, 16, 16, 16, 14, 16, 20, 20]
        for i, w in enumerate(proj_widths, 1):
            ws_proj.column_dimensions[get_column_letter(i)].width = w
        ws_proj.freeze_panes = "A3"

        # ═══════════════════════════════════════════════════════════
        # Sheet 2: Summary (Item 明细)
        # ═══════════════════════════════════════════════════════════
        ws = wb.create_sheet("Summary")

        ws.merge_cells("A1:M1")
        ws["A1"] = f"库存金额跟踪报表 - Site {self.site_ref} ({site_label}) - 所有物料 (P+M)"
        ws["A1"].font = TITLE_FONT
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("A2:M2")
        ws["A2"] = (f"报告周期：{today.strftime('%Y年%m月')} 1日 - {today.strftime('%m月%d日')}   |   "
                 f"数据截至：{today.strftime('%Y-%m-%d')}   |   "
                     f"含采购物料 (P) + 制造件 (M)")
        ws["A2"].alignment = Alignment(horizontal="center")
        ws["A2"].font = Font(size=10, italic=True, color="666666")

        headers = ["Project Code", "Item", "Description", f"{prev_label} Qty", f"{prev_label} Balance",
                    "MTD Received Qty", "MTD Received", "MTD Consumption Qty", "MTD Consumption",
                    "MTD Other Qty", "MTD Other Transaction", "MTD Daily Balance Qty", "MTD Daily Balance"]
        ws.append(headers)
        for col_num in range(1, len(headers) + 1):
            c = ws.cell(row=3, column=col_num)
            c.fill = HEADER_FILL
            c.font = HEADER_FONT
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = THIN_BORDER

        for row_data in summary_df.drop(columns=["PMTCode"]).itertuples(index=False):
            ws.append(list(row_data))

        total_row = ws.max_row + 1
        ws.cell(row=total_row, column=1, value="TOTAL")
        ws.cell(row=total_row, column=1).font = Font(bold=True)
        ws.cell(row=total_row, column=1).fill = TOTAL_FILL
        ws.cell(row=total_row, column=1).border = THIN_BORDER
        col_to_field = {
            4:  ("Prev_Qty",     "#,##0.0000"),
            5:  ("Prev_Balance", "#,##0.00"),
            6:  ("Received_Qty", "#,##0.0000"),
            7:  ("Received_AMT", "#,##0.00"),
            8:  ("Consumed_Qty", "#,##0.0000"),
            9:  ("Consumed_AMT", "#,##0.00"),
            10: ("Other_Qty",    "#,##0.0000"),
            11: ("Other_AMT",    "#,##0.00"),
            12: ("Balance_Qty",  "#,##0.0000"),
            13: ("Balance_AMT",  "#,##0.00"),
        }
        for col_idx, (field, fmt) in col_to_field.items():
            total_val = summary_df[field].sum() if field in summary_df.columns else 0
            c = ws.cell(row=total_row, column=col_idx)
            c.value = round(total_val, 4 if "Qty" in field else 2)
            c.font = Font(bold=True)
            c.fill = TOTAL_FILL
            c.border = THIN_BORDER
            c.number_format = fmt

        widths = [16, 20, 40, 12, 15, 14, 15, 14, 15, 12, 14, 12, 15]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

        ws.freeze_panes = "A4"
        ws.auto_filter.ref = f"A3:{get_column_letter(len(headers))}{total_row - 1}"
        ws.row_dimensions[1].height = 28
        ws.row_dimensions[2].height = 20
        ws.row_dimensions[3].height = 30

        # ═══════════════════════════════════════════════════════════
        # Sheet 3: Detail
        # ═══════════════════════════════════════════════════════════
        ws_det = wb.create_sheet("Detail")

        det_cols = len(detail_df.columns)
        ws_det.merge_cells(f"A1:{get_column_letter(det_cols)}1")
        ws_det["A1"] = f"物料事务明细 - Site {self.site_ref} ({site_label}) - 所有物料 (P+M)"
        ws_det["A1"].font = TITLE_FONT
        ws_det["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws_det.row_dimensions[1].height = 28

        for col_idx, col_name in enumerate(detail_df.columns, 1):
            c = ws_det.cell(row=2, column=col_idx, value=col_name)
            c.fill = HEADER_FILL
            c.font = HEADER_FONT
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = THIN_BORDER
        ws_det.row_dimensions[2].height = 25

        for row_data in detail_df.itertuples(index=False):
            ws_det.append(list(row_data))

        ws_det.freeze_panes = "A3"
        ws_det.auto_filter.ref = f"A2:{get_column_letter(det_cols)}{ws_det.max_row}"

        for col_idx in range(1, det_cols + 1):
            col_letter = get_column_letter(col_idx)
            col_name = detail_df.columns[col_idx - 1]
            if col_name in ("Item", "TransNum", "RefNum"):
                ws_det.column_dimensions[col_letter].width = 18
            elif col_name in ("TransDate", "TransType", "RefType", "Category", "TransDesc", "Whse", "Loc", "Lot"):
                ws_det.column_dimensions[col_letter].width = 14
            elif col_name == "Description":
                ws_det.column_dimensions[col_letter].width = 40
            elif col_name in ("Qty", "TotalAmt"):
                ws_det.column_dimensions[col_letter].width = 14
            else:
                ws_det.column_dimensions[col_letter].width = 12

        wb.save(output_path)
        print(f"  📄 报表已导出：{output_path.absolute()}")
        print(f"     Summary: {len(summary_df):,} Items (P+M)")
        print(f"     Detail: {len(detail_df):,} records")
        print(f"     Project Summary: {len(proj_group):,} Project Codes")

    # ──────────────────────────────────────────────────────────
    # 主流程
    # ──────────────────────────────────────────────────────────
    def run(self):
        site_label = SITE_NAMES.get(self.site_ref, self.site_ref)
        print(f"\n{'='*60}")
        print(f"  库存跟踪 - Site {self.site_ref} ({site_label})")
        print(f"{'='*60}")

        print("\n📁 Step 1: 加载期初库存余额...")
        prev_df = self.load_previous_balance()

        print("\n🗄️  Step 2: 查询当月MTD物料事务...")
        trans_df = self.fetch_mtd_transactions()

        print("\n📊 Step 3: 分类汇总计算...")
        summary_df, detail_df = self.calculate_balances(prev_df, trans_df)

        print("\n📤 Step 4: 导出Excel报表...")
        self.export_excel(summary_df, detail_df)

        # Quick summary
        prev_total = summary_df["Prev_Balance"].sum()
        recv_total = summary_df["Received_AMT"].sum()
        cons_total = summary_df["Consumed_AMT"].sum()
        other_total = summary_df["Other_AMT"].sum()
        bal_total = prev_total + recv_total + cons_total + other_total

        print(f"\n  ┌─────────────────────────────────────────┐")
        print(f"  │  期初余额:   ${prev_total:>15,.2f}       │")
        print(f"  │  + Received: ${recv_total:>15,.2f}       │")
        print(f"  │  + Consumed: ${cons_total:>15,.2f}       │")
        print(f"  │  + Other:    ${other_total:>15,.2f}       │")
        print(f"  │  = Balance:  ${bal_total:>15,.2f}       │")
        print(f"  └─────────────────────────────────────────┘")
        print(f"\n✅ Site {self.site_ref} ({site_label}) 完成！")
        return summary_df, detail_df


# ──────────────────────────────────────────────────────────────
# 多站点批量运行
# ──────────────────────────────────────────────────────────────
def run_all_sites(
    server: str, database: str, username: str, password: str, port: int = 1433,
    sites: Optional[List[str]] = None,
    as_of_date: Optional[str] = None,
):
    """批量运行所有站点（期初余额从数据库读取，不再依赖 Excel 文件）"""
    if sites is None:
        sites = ["310", "330", "410"]

    report_date = datetime.strptime(as_of_date, "%Y-%m-%d") if as_of_date else datetime.today()
    all_summaries = {}
    for site in sites:
        tracker = InventoryTracker(
            server=server, database=database,
            username=username, password=password, port=port,
            site_ref=site,
            as_of_date=as_of_date,
        )
        summary_df, detail_df = tracker.run()
        summary_df["Site"] = site
        all_summaries[site] = summary_df

    if not all_summaries:
        print("❌ 无可用数据，退出")
        return None

    # 生成合并汇总
    combined = pd.concat(all_summaries.values(), ignore_index=True)
    print(f"\n{'='*60}")
    print(f"  合并汇总 ({len(all_summaries)} 个站点)")
    print(f"{'='*60}")

    # ── 按 PMTCode(P=RM / M=FG) 分类汇总（供邮件表格使用）──
    def _agg_site(df: pd.DataFrame, pmt_filter: str) -> Dict[str, float]:
        """按 PMTCode 过滤后汇总关键字段"""
        sub = df[df["PMTCode"] == pmt_filter]
        return {
            "prev": sub["Prev_Balance"].sum(),
            "recv": sub["Received_AMT"].sum(),
            "cons": sub["Consumed_AMT"].sum(),
            "other": sub["Other_AMT"].sum(),
            "bal": sub["Balance_AMT"].sum(),
        }

    site_breakdown: Dict[str, Dict[str, Dict[str, float]]] = {}
    for site in all_summaries:
        df_site = all_summaries[site]
        site_breakdown[site] = {
            "RM": _agg_site(df_site, "P"),
            "FG": _agg_site(df_site, "M"),
        }
        print(f"\n  Site {site} ({SITE_NAMES.get(site, site)}):")
        rm = site_breakdown[site]["RM"]
        fg = site_breakdown[site]["FG"]
        print(f"    RM  (P): 期初=${rm['prev']:,.2f} +Recv=${rm['recv']:,.2f} +Cons=${rm['cons']:,.2f} +Other=${rm['other']:,.2f} =${rm['bal']:,.2f}")
        print(f"    FG  (M): 期初=${fg['prev']:,.2f} +Recv=${fg['recv']:,.2f} +Cons=${fg['cons']:,.2f} +Other=${fg['other']:,.2f} =${fg['bal']:,.2f}")

    # ── Daily Balance 以数据库快照为准（SLTotalInventory, BalanceDate=as_of_date）──
    inv_daily_bal = fetch_inventory_daily_balances(
        sites=list(all_summaries.keys()),
        as_of_date=report_date.strftime("%Y-%m-%d"),
        db_config={
            "server": server,
            "database": database,
            "username": username,
            "password": password,
            "port": port,
        },
    )
    for site in all_summaries:
        snap = inv_daily_bal.get(site, {})
        site_breakdown[site]["RM"]["bal"] = snap.get("rm")
        site_breakdown[site]["FG"]["bal"] = snap.get("fg")

    # 保留旧版 group_totals（兼容 Excel 等）
    group_totals = combined.groupby("Site").agg(
        Items=("Item", "count"),
        Prev_Balance=("Prev_Balance", "sum"),
        Received_AMT=("Received_AMT", "sum"),
        Consumed_AMT=("Consumed_AMT", "sum"),
        Other_AMT=("Other_AMT", "sum"),
        Balance_AMT=("Balance_AMT", "sum"),
    ).reset_index()

    grand_prev = combined["Prev_Balance"].sum()
    grand_recv = combined["Received_AMT"].sum()
    grand_cons = combined["Consumed_AMT"].sum()
    grand_other = combined["Other_AMT"].sum()
    grand_bal = grand_prev + grand_recv + grand_cons + grand_other
    print(f"\n  🏢 ALL SITES Grand Total (USD):")
    print(f"    期初: ${grand_prev:,.2f}  +Recv: ${grand_recv:,.2f}  "
          f"+Cons: ${grand_cons:,.2f}  +Other: ${grand_other:,.2f}  "
          f"= ${grand_bal:,.2f}")

    # ── 获取 WIP 数据 ──
    wip_totals = fetch_wip_totals(
        sites=list(all_summaries.keys()),
        as_of_date=report_date.strftime("%Y-%m-%d"),
        db_config={
            "server": server,
            "database": database,
            "username": username,
            "password": password,
            "port": port,
        },
    )
    grand_wip = sum(v for v in wip_totals.values() if v is not None)
    print(f"\n  📦 WIP Grand Total (USD): ${grand_wip:,.2f}")

    return {
        "combined": combined,
        "group_totals": group_totals,
        "grand_prev": grand_prev, "grand_recv": grand_recv,
        "grand_cons": grand_cons, "grand_other": grand_other,
        "grand_bal": grand_bal,
        "wip_totals": wip_totals, "grand_wip": grand_wip,
        "site_breakdown": site_breakdown,
        "report_date": report_date.strftime("%Y-%m-%d"),
    }


# ──────────────────────────────────────────────────────────────
# 从数据库动态读取上期库存余额
# ──────────────────────────────────────────────────────────────
def fetch_prev_balance_from_db(server, username, password, database, port,
                                site_ref: str, target_date: str) -> dict:
    """
    从 SLTotalInventory 表读取指定日期、站点的 RM/FG 汇总金额。

    参数:
        target_date: 日期字符串，如 "2026-04-30"
    返回:
        {"rm": float, "fg": float}
        RM  = SUM(Unitscost) WHERE Source = 'Purchased'
        FG  = SUM(Unitscost) WHERE Source = 'Manufactured'
    """
    result = {"rm": 0.0, "fg": 0.0}
    try:
        conn = _db_connect(server, username, password, database, port)
        cur = conn.cursor()
        cur.execute("""
            SELECT Source, ISNULL(SUM(Unitscost), 0) AS TotalAmt
            FROM dbo.SLTotalInventory
            WHERE SiteRef = %s AND BalanceDate = %s
            GROUP BY Source
        """, (site_ref, target_date))
        for row in cur.fetchall():
            src = str(row[0]).strip().upper()
            amt = float(row[1]) if row[1] else 0.0
            if src == "PURCHASED":
                result["rm"] = amt
            elif src == "MANUFACTURED":
                result["fg"] = amt
        conn.close()
    except Exception as e:
        print(f"  ⚠️  查询上期库存失败 (Site {site_ref}, {target_date}): {e}")
    return result


def fetch_prev_wip_from_db(server, username, password, database, port,
                           site_ref: str, target_date: str) -> float:
    """
    从 SLTotalWIPValueByAcountReport 表读取指定日期、站点的 WIP 汇总金额。

    参数:
        target_date: 日期字符串，如 "2026-05-31"
    返回:
        WIP 总金额 (float)
    """
    try:
        conn = _db_connect(server, username, password, database, port)
        cur = conn.cursor()
        cur.execute("""
            SELECT ISNULL(SUM(AcctTot), 0) AS WipTotal
            FROM dbo.SLTotalWIPValueByAcountReport
            WHERE SiteRef = %s AND BalanceDate = %s
        """, (site_ref, target_date))
        row = cur.fetchone()
        conn.close()
        wip_total = float(row[0]) if row and row[0] else 0.0
        # Site 330 原始口径为 CNY，邮件展示统一 USD
        if site_ref == "330":
            wip_total = wip_total * SITE_CURRENCY["330"][1]
        return wip_total
    except Exception as e:
        print(f"  ⚠️  查询上期 WIP 失败 (Site {site_ref}, {target_date}): {e}")
        return 0.0


# ──────────────────────────────────────────────────────────────
# 发送汇总邮件（SMTP 版本）
# ──────────────────────────────────────────────────────────────
def send_summary_email(
    result: dict,
    to_addr: str = "jason.pang@nai-group.com;shirley.ni@nai-group.com;devin.hua@nai-group.com;chn_planners@nai-group.com;chn_buyer@nai-group.com",
    cc_addr: str = "sky.li@nai-group.com;frank.liu@nai-group.com;shirley.ni@nai-group.com",
    smtp_host: str = "localhost",
    smtp_port: int = 25,
    smtp_user: str = "",
    smtp_password: str = "",
    smtp_tls: bool = False,
    from_addr: str = "inventory-report@nai-group.com",
    db_config: dict | None = None,
    report_date: Optional[str] = None,
):
    """生成固定格式的 HTML 邮件并通过 SMTP 发送

    邮件格式：每个站点一个独立表格
    ┌──────────────────────────────────────────────────────────────────────────┐
    │ Site 310 (Plant1)                                                        │
    ├──────────────┬─────────────┬──────────────┬──────────────┬───────────────┤
    │              │5/30 Balance │MTD Received  │MTD Consumed  │MTD Daily Bal  │
    ├──────────────┼─────────────┼──────────────┼──────────────┼───────────────┤
    │ RM           │ $xxx        │ $xxx         │ $xxx         │ $xxx          │
    │ FG/Semi FG   │ $xxx        │ $xxx         │ $xxx         │ $xxx          │
    │ WIP          │             │ 0            │ 0            │ $xxx          │
    │ Total        │ $xxx        │ $xxx         │ $xxx         │ $xxx          │
    └──────────────┴─────────────┴──────────────┴──────────────┴───────────────┘
    """
    site_breakdown = result.get("site_breakdown", {})
    wip_totals = result.get("wip_totals", {})

    def fmt_num(v):
        if v is None:
            return ""
        if v < 0:
            return f"-${abs(v):,.2f}"
        return f"${v:,.2f}"

    if report_date:
        try:
            today = datetime.strptime(report_date, "%Y-%m-%d")
        except ValueError as e:
            raise ValueError("report_date 格式错误，应为 YYYY-MM-DD") from e
    else:
        today = datetime.today()
    month_label = today.strftime("%B %Y")
    date_label = today.strftime("%B %d %Y")
    prev_month = (today.replace(day=1) - timedelta(days=1))
    prev_eom_label = prev_month.strftime("%m/%d")

    # ── 上月月末日期（用于上期 WIP，如 5/31）──
    prev_date_str = prev_month.strftime("%Y-%m-%d")

    # ── 动态从数据库读取上期 WIP（用于上个月 Balance 列）──
    PREV_MONTH_WIP = {}  # {site: float}
    all_sites = ["310", "330", "410"]

    if db_config:
        print(f"  📊 从数据库读取上期 WIP ({prev_eom_label} = {prev_date_str})...")
        for site in all_sites:
            wip_val = fetch_prev_wip_from_db(
                db_config["server"], db_config["username"], db_config["password"],
                db_config["database"], db_config.get("port", 1433),
                site, prev_date_str,
            )
            PREV_MONTH_WIP[site] = wip_val
            print(f"    Site {site}: WIP={wip_val:,.2f}")
    else:
        # 无数据库配置时使用硬编码 fallback（兼容旧调用方式）
        print("  ⚠️  未提供 db_config，使用硬编码上期 WIP")
        PREV_MONTH_WIP = {
            "310": 200903.16,
            "330": 543.63,
            "410": 157920.87,
        }

    tables_html = ""

    for site in all_sites:
        if site not in site_breakdown:
            continue

        sd = site_breakdown[site]
        rm = sd.get("RM", {})
        fg = sd.get("FG", {})
        rm_bal = rm.get("bal")
        fg_bal = fg.get("bal")
        wip_bal = wip_totals.get(site)
        prev_wip = PREV_MONTH_WIP.get(site, 0.0)
        # 计算 Total 行
        t_prev = rm.get("prev", 0.0) + fg.get("prev", 0.0) + prev_wip
        t_recv = rm.get("recv", 0.0) + fg.get("recv", 0.0)
        t_cons = rm.get("cons", 0.0) + fg.get("cons", 0.0)
        t_other = rm.get("other", 0.0) + fg.get("other", 0.0)
        t_bal = (rm_bal + fg_bal + wip_bal) if (rm_bal is not None and fg_bal is not None and wip_bal is not None) else None

        # MTD Variances = Daily Balance - prev Balance - (Received + Consumed + Other)
        rm_var = (rm_bal - rm.get("prev", 0.0) - (rm.get("recv", 0.0) + rm.get("cons", 0.0) + rm.get("other", 0.0))) if rm_bal is not None else None
        fg_var = (fg_bal - fg.get("prev", 0.0) - (fg.get("recv", 0.0) + fg.get("cons", 0.0) + fg.get("other", 0.0))) if fg_bal is not None else None
        wip_var = (wip_bal - prev_wip) if wip_bal is not None else None  # WIP 没有事务列
        t_var = (t_bal - t_prev - (t_recv + t_cons + t_other)) if t_bal is not None else None

        label = f"{site} ({SITE_NAMES.get(site, site)})"
        site_color = {"310": "#1F4E79", "330": "#1F4E79", "410": "#1F4E79"}[site]

        tables_html += f"""
<p style="margin-top:18px"><b>{label} — {date_label}</b></p>
<table border="1" cellpadding="5" cellspacing="0"
  style="border-collapse:collapse;font-size:10pt;text-align:right;width:100%;max-width:1050px">
<tr style="background-color:{site_color};color:white;text-align:center">
  <th style="width:12%;text-align:left">&nbsp;</th>
    <th style="width:14%">{prev_eom_label} Balance</th>
  <th style="width:12%">MTD Received</th>
  <th style="width:12%">MTD Consumed</th>
  <th style="width:12%">MTD Other Transaction</th>
  <th style="width:12%">MTD Variances</th>
    <th style="width:14%">MTD Daily Balance</th>
</tr>
<tr>
  <td style="text-align:left;font-weight:bold">RM</td>
  <td>{fmt_num(rm.get('prev', 0))}</td>
  <td>{fmt_num(rm.get('recv', 0))}</td>
  <td>{fmt_num(rm.get('cons', 0))}</td>
  <td>{fmt_num(rm.get('other', 0))}</td>
  <td>{fmt_num(rm_var)}</td>
    <td><b>{fmt_num(rm_bal)}</b></td>
</tr>
<tr>
  <td style="text-align:left;font-weight:bold">FG/Semi FG</td>
  <td>{fmt_num(fg.get('prev', 0))}</td>
  <td>{fmt_num(fg.get('recv', 0))}</td>
  <td>{fmt_num(fg.get('cons', 0))}</td>
  <td>{fmt_num(fg.get('other', 0))}</td>
  <td>{fmt_num(fg_var)}</td>
    <td><b>{fmt_num(fg_bal)}</b></td>
</tr>
<tr style="background-color:#FFF2CC">
  <td style="text-align:left;font-weight:bold">WIP</td>
  <td>{fmt_num(prev_wip)}</td>
  <td>&nbsp;</td>
  <td>&nbsp;</td>
  <td>&nbsp;</td>
  <td>{fmt_num(wip_var)}</td>
  <td><b>{fmt_num(wip_bal)}</b></td>
</tr>
<tr style="background-color:#D6E4F0;font-weight:bold">
  <td style="text-align:left">Total</td>
  <td>{fmt_num(t_prev)}</td>
  <td>{fmt_num(t_recv)}</td>
  <td>{fmt_num(t_cons)}</td>
  <td>{fmt_num(t_other)}</td>
  <td>{fmt_num(t_var)}</td>
  <td><b>{fmt_num(t_bal)}</b></td>
</tr>
</table>
"""

    # ── Asia Total 汇总表 ──
    asia_rm_prev = sum(site_breakdown.get(s, {}).get("RM", {}).get("prev", 0) for s in all_sites)
    asia_rm_recv = sum(site_breakdown.get(s, {}).get("RM", {}).get("recv", 0) for s in all_sites)
    asia_rm_cons = sum(site_breakdown.get(s, {}).get("RM", {}).get("cons", 0) for s in all_sites)
    asia_rm_other = sum(site_breakdown.get(s, {}).get("RM", {}).get("other", 0) for s in all_sites)
    asia_rm_bal_vals = [site_breakdown.get(s, {}).get("RM", {}).get("bal") for s in all_sites]
    asia_rm_bal = sum(v for v in asia_rm_bal_vals if v is not None) if all(v is not None for v in asia_rm_bal_vals) else None

    asia_fg_prev = sum(site_breakdown.get(s, {}).get("FG", {}).get("prev", 0) for s in all_sites)
    asia_fg_recv = sum(site_breakdown.get(s, {}).get("FG", {}).get("recv", 0) for s in all_sites)
    asia_fg_cons = sum(site_breakdown.get(s, {}).get("FG", {}).get("cons", 0) for s in all_sites)
    asia_fg_other = sum(site_breakdown.get(s, {}).get("FG", {}).get("other", 0) for s in all_sites)
    asia_fg_bal_vals = [site_breakdown.get(s, {}).get("FG", {}).get("bal") for s in all_sites]
    asia_fg_bal = sum(v for v in asia_fg_bal_vals if v is not None) if all(v is not None for v in asia_fg_bal_vals) else None

    asia_wip_vals = [wip_totals.get(s) for s in all_sites]
    asia_wip_bal = sum(v for v in asia_wip_vals if v is not None) if any(v is not None for v in asia_wip_vals) else None

    asia_prev_wip = sum(PREV_MONTH_WIP.get(s, 0.0) for s in all_sites)
    asia_t_prev = asia_rm_prev + asia_fg_prev + asia_prev_wip
    asia_t_recv = asia_rm_recv + asia_fg_recv
    asia_t_cons = asia_rm_cons + asia_fg_cons
    asia_t_other = asia_rm_other + asia_fg_other
    asia_t_bal = (asia_rm_bal + asia_fg_bal + asia_wip_bal) if (asia_rm_bal is not None and asia_fg_bal is not None and asia_wip_bal is not None) else None

    # Asia Total MTD Variances
    asia_rm_var = (asia_rm_bal - asia_rm_prev - (asia_rm_recv + asia_rm_cons + asia_rm_other)) if asia_rm_bal is not None else None
    asia_fg_var = (asia_fg_bal - asia_fg_prev - (asia_fg_recv + asia_fg_cons + asia_fg_other)) if asia_fg_bal is not None else None
    asia_wip_var = (asia_wip_bal - asia_prev_wip) if asia_wip_bal is not None else None
    asia_t_var = (asia_t_bal - asia_t_prev - (asia_t_recv + asia_t_cons + asia_t_other)) if asia_t_bal is not None else None

    tables_html += f"""
<p style="margin-top:24px"><b>Asia Total — {date_label}</b></p>
<table border="1" cellpadding="5" cellspacing="0"
  style="border-collapse:collapse;font-size:10pt;text-align:right;width:100%;max-width:1050px">
<tr style="background-color:#4472C4;color:white;text-align:center">
  <th style="width:12%;text-align:left">&nbsp;</th>
    <th style="width:14%">{prev_eom_label} Balance</th>
  <th style="width:12%">MTD Received</th>
  <th style="width:12%">MTD Consumed</th>
  <th style="width:12%">MTD Other Transaction</th>
  <th style="width:12%">MTD Variances</th>
  <th style="width:14%">MTD Daily Balance</th>
</tr>
<tr>
  <td style="text-align:left;font-weight:bold">RM</td>
  <td>{fmt_num(asia_rm_prev)}</td>
  <td>{fmt_num(asia_rm_recv)}</td>
  <td>{fmt_num(asia_rm_cons)}</td>
  <td>{fmt_num(asia_rm_other)}</td>
  <td>{fmt_num(asia_rm_var)}</td>
  <td><b>{fmt_num(asia_rm_bal)}</b></td>
</tr>
<tr>
  <td style="text-align:left;font-weight:bold">FG/Semi FG</td>
  <td>{fmt_num(asia_fg_prev)}</td>
  <td>{fmt_num(asia_fg_recv)}</td>
  <td>{fmt_num(asia_fg_cons)}</td>
  <td>{fmt_num(asia_fg_other)}</td>
  <td>{fmt_num(asia_fg_var)}</td>
  <td><b>{fmt_num(asia_fg_bal)}</b></td>
</tr>
<tr style="background-color:#FFF2CC">
  <td style="text-align:left;font-weight:bold">WIP</td>
  <td>{fmt_num(asia_prev_wip)}</td>
  <td>&nbsp;</td>
  <td>&nbsp;</td>
  <td>&nbsp;</td>
  <td>{fmt_num(asia_wip_var)}</td>
  <td><b>{fmt_num(asia_wip_bal)}</b></td>
</tr>
<tr style="background-color:#4472C4;color:white;font-weight:bold">
  <td style="text-align:left">Total</td>
  <td>{fmt_num(asia_t_prev)}</td>
  <td>{fmt_num(asia_t_recv)}</td>
  <td>{fmt_num(asia_t_cons)}</td>
  <td>{fmt_num(asia_t_other)}</td>
  <td>{fmt_num(asia_t_var)}</td>
  <td><b>{fmt_num(asia_t_bal)}</b></td>
</tr>
</table>
"""

    html_body = f"""<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt">
<p>Below is the {month_label} Inventory Valuation Tracking Report.
<br>All amounts in USD. Site 330 CNY amounts converted at rate 6.838784.
<br>Daily Balance values (RM/FG/WIP) are snapshot-based from DB by BalanceDate. If snapshot missing: historical dates stay blank, today may auto-sync.</p>
{tables_html}
</div>"""

    # ── 构建 MIME 邮件 ──
    msg = MIMEMultipart()
    msg["From"] = from_addr
    msg["To"] = to_addr
    if cc_addr:
        msg["Cc"] = cc_addr
    msg["Subject"] = f"Inventory Valuation Tracking Daily Report - {date_label}"
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    # 附件
    group_totals = result.get("group_totals", pd.DataFrame())
    for site in group_totals["Site"]:
        fname = f"Inventory_Balance_{SITE_NAMES.get(site, site)}_{today.strftime('%Y%m')}.xlsx"
        fpath = fname
        if os.path.exists(fpath):
            try:
                with open(fpath, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f'attachment; filename="{fname}"')
                msg.attach(part)
                print(f"  📎 Attached: {fname}")
            except Exception as e:
                print(f"  ⚠️ Failed to attach {fname}: {e}")
        else:
            print(f"  ⚠️ File not found: {fpath}")

    # 收件人列表（To + Cc 合并）
    all_recipients = [a.strip() for a in to_addr.split(";") if a.strip()]
    if cc_addr:
        all_recipients += [a.strip() for a in cc_addr.split(";") if a.strip()]
    all_recipients = list(set(all_recipients))  # 去重

    # ── 发送 ──
    try:
        if smtp_tls:
            server = smtplib.SMTP(smtp_host, smtp_port, timeout=30)
            server.ehlo()
            server.starttls()
            server.ehlo()
        else:
            server = smtplib.SMTP(smtp_host, smtp_port, timeout=30)

        if smtp_user and smtp_password:
            server.login(smtp_user, smtp_password)

        server.sendmail(from_addr, all_recipients, msg.as_string())
        server.quit()
        print(f"  ✅ 邮件已通过 SMTP ({smtp_host}:{smtp_port}) 发送")
        print(f"     To : {to_addr}")
        if cc_addr:
            print(f"     Cc : {cc_addr}")
    except Exception as e:
        print(f"  ❌ 邮件发送失败：{e}")
        raise


# ──────────────────────────────────────────────────────────────
# 期初库存生成 — 从 Infor CSI API 获取库存快照
# ──────────────────────────────────────────────────────────────

# Infor CSI API 站点参数配置
# clmParam 格式：M,PM,V,ABC,0,T,,,,,,,0,0,{site}
# M=制造, PM=PMTCode(采购件P+制造件M), V=Valuation, ABC=ABC分类, T=Include all, {site}=站点
INFOR_API_BASE = "https://mingle-ionapi.inforcloudsuite.com"
INFOR_TENANT = "NAIGROUP_PRD"
INFOR_IDO = "SLItemCostingReport"
INFOR_REPORT_PROC = "Rpt_ItemCostingSp"
INFOR_PROPERTIES = "Seq,RptSeq,Item,Itemdesc,Units,Unitcost,Unitscost,Pmtcode,Prodcode"

# 每个站点的 clmParam 及 MongooseConfig header
SITE_API_CONFIG = {
    "310": {
        "clmParam": "M,PM,V,ABC,0,T,,,,,,,0,0,310",
        "mongoose_config": "NAIGROUP_PRD_310",
    },
    "330": {
        "clmParam": "M,PM,V,ABC,0,T,,,,,,,0,0,330",
        "mongoose_config": "NAIGROUP_PRD_330",
    },
    "410": {
        "clmParam": "M,PM,V,ABC,0,T,,,,,,,0,0,410",
        "mongoose_config": "NAIGROUP_PRD_410",
    },
}


# ──────────────────────────────────────────────────────────────
# WIP (Work-In-Process) API 配置
# IDO:  SLTotalWIPValuebyAccountReport
# Proc: Rpt_TotalWIPValuebyAccountSp
# 关键字段：AcctTot（各科目合计），IsDetail=1 为明细行
# ──────────────────────────────────────────────────────────────
INFOR_WIP_IDO   = "SLTotalWIPValuebyAccountReport"
INFOR_WIP_PROC  = "Rpt_TotalWIPValuebyAccountSp"
INFOR_WIP_PROPS = (
    "JobAcct,AcctUnit1,AcctUnit2,AcctUnit3,AcctUnit4,"
    "JobWipLbrTotal,JobWipMatlTotal,JobWipFovhdTotal,JobWipVovhdTotal,JobWipOutTotal,"
    "AcctTot,Des,IsDetail,GLTotal,Diff"
)
# clmParam 格式：,,,,,,,,0000,9999,RS,0,0,0,0,0,,,OI,{site},0
# 第 20 个字段（0-based index 19）= SiteRef
SITE_WIP_CONFIG = {
    "310": {
        "clmParam": ",,,,,,,,0000,9999,RS,0,0,0,0,0,,,OI,310,0",
        "mongoose_config": "NAIGROUP_PRD_310",
    },
    "330": {
        "clmParam": ",,,,,,,,0000,9999,RS,0,0,0,0,0,,,OI,330,0",
        "mongoose_config": "NAIGROUP_PRD_330",
    },
    "410": {
        "clmParam": ",,,,,,,,0000,9999,RS,0,0,0,0,0,,,OI,410,0",
        "mongoose_config": "NAIGROUP_PRD_410",
    },
}


# ── Infor OAuth2 Token 缓存（模块级，进程生命周期内有效）────────
_infor_token_cache: str | None = None
_infor_token_expires_at: float = 0.0
_daily_snapshot_synced_dates: set[str] = set()


def _read_infor_config() -> dict:
    """
    从环境变量读取 OAuth2 配置。
    优先级：环境变量 > 代码内默认值
    """
    def _env(env_key, default=""):
        return os.environ.get(env_key, "").strip() or default

    return {
        "token_url":    _env("INFOR_TOKEN_URL",
                             f"https://mingle-sso.inforcloudsuite.com:443/{INFOR_TENANT}/as/token.oauth2"),
        "auth_basic":   _env("INFOR_AUTH_BASIC"),
        "username":     _env("INFOR_USERNAME"),
        "password":     _env("INFOR_PASSWORD"),
        "bearer_token": _env("INFOR_BEARER_TOKEN"),
    }


def _fetch_oauth_token(config: dict) -> tuple[str, int]:
    """
    发起 OAuth2 token 请求，返回 (access_token, expires_in)。
    使用 urllib（不依赖 httpx），兼容 Docker 内的精简环境。
    """
    token_url = config["token_url"]
    auth_basic = config["auth_basic"]
    username = config["username"]
    password = config["password"]

    body = urllib.parse.urlencode({
        "grant_type": "password",
        "username": username,
        "password": password,
    }).encode("utf-8")

    req = urllib.request.Request(
        token_url,
        data=body,
        headers={
            "Authorization": f"Basic {auth_basic}",
            "Content-Type": "application/x-www-form-urlencoded",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            raw = resp.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(
            f"❌ OAuth2 Token 请求失败 (HTTP {e.code})\n"
            f"   URL: {token_url}\n"
            f"   响应: {err_body[:500]}"
        ) from e
    except urllib.error.URLError as e:
        raise RuntimeError(
            f"❌ Token 端点网络连接失败: {e.reason}\n"
            f"   URL: {token_url}\n"
            f"   请检查 VPN 或网络连接"
        ) from e

    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        raise RuntimeError(
            f"❌ Token 端点返回非 JSON: {e}\n"
            f"   响应: {raw[:500]}"
        ) from e

    token = data.get("access_token")
    if not token:
        error = data.get("error", "unknown")
        desc = data.get("error_description", "")
        raise RuntimeError(
            f"❌ Token 响应中无 access_token\n"
            f"   error: {error}\n"
            f"   description: {desc}"
        )

    expires_in = int(data.get("expires_in", 3600))
    return token, expires_in


def _load_infor_token(force_refresh: bool = False) -> str:
    """
    获取 Infor CSI API 的有效 access_token。

    策略：
      1. 如果有缓存 token 且未过期（提前 60s），直接返回
      2. 读取 OAuth2 配置（Basic Auth + password grant）
      3. POST token endpoint 获取新 token
      4. 对 400/429/502/503/504 瞬时错误自动重试 3 次
      5. 如果 OAuth2 配置缺失，降级使用手动 Bearer Token

    force_refresh=True 时忽略缓存，强制重新获取（用于 401 重试场景）。
    """
    global _infor_token_cache, _infor_token_expires_at

    # ── 检查缓存 ──
    if not force_refresh and _infor_token_cache:
        now = time.time()
        if now < _infor_token_expires_at - 60:
            return _infor_token_cache

    # ── 读取配置 ──
    config = _read_infor_config()

    # ── 模式 A：OAuth2 password grant（推荐，自动刷新）──
    if config["auth_basic"] and config["username"] and config["password"]:
        print(f"  🔐 OAuth2 获取 Token: {config['token_url']}")
        last_err = None
        for attempt in range(1, 4):
            try:
                token, expires_in = _fetch_oauth_token(config)
                _infor_token_cache = token
                _infor_token_expires_at = time.time() + expires_in
                print(f"  ✅ Token 获取成功，有效期 {expires_in}s")
                return token
            except urllib.error.HTTPError as e:
                last_err = e
                if e.code in (400, 429, 502, 503, 504):
                    wait = 2 ** (attempt - 1)
                    print(f"  ⚠️  Token 请求第 {attempt} 次失败 (HTTP {e.code})，{wait}s 后重试...")
                    time.sleep(wait)
                else:
                    raise
            except RuntimeError as e:
                last_err = e
                if attempt < 3:
                    wait = 2 ** (attempt - 1)
                    print(f"  ⚠️  Token 请求第 {attempt} 次异常：{e}")
                    print(f"  ⏳ {wait}s 后重试...")
                    time.sleep(wait)
                else:
                    break

        raise RuntimeError(f"❌ OAuth2 Token 连续获取失败（3次重试）：{last_err}") from last_err

    # ── 模式 B：手动 Bearer Token（降级，需要手动刷新）──
    if config["bearer_token"]:
        _infor_token_cache = config["bearer_token"]
        _infor_token_expires_at = time.time() + 86400  # 手动 token 假设 24h 有效
        print("  🔑 使用手动 Bearer Token（注意：过期需手动更新）")
        return _infor_token_cache

    raise RuntimeError(
        "❌ 未找到 Infor API 认证凭据！\n"
        "   请在 .env 文件中配置 OAuth2 凭据：\n"
        "   ── 推荐：OAuth2 password grant（自动刷新，无需手动维护 Token）──\n"
        "   INFOR_AUTH_BASIC=<Base64(client_id:client_secret)>\n"
        "   INFOR_USERNAME=<infor_username>\n"
        "   INFOR_PASSWORD=<infor_password>\n"
        "   ── 备选：手动 Bearer Token（需手动更新，不推荐）──\n"
        "   INFOR_BEARER_TOKEN=<token>\n"
        "   ── 获取 OAuth2 凭据 ──\n"
        "   1. client_id / client_secret：联系 Infor 管理员或从 csi_datawarehouse .env 获取\n"
        "   2. INFOR_AUTH_BASIC = Base64(client_id:client_secret)\n"
        "      Linux: echo -n 'client_id:client_secret' | base64\n"
        "      Python: base64.b64encode(b'client_id:client_secret').decode()\n"
    )



def _fetch_infor_site(site_ref: str, token: str, token_expired_retry: bool = False) -> pd.DataFrame:
    """
    调用 Infor CSI IDO API 获取指定站点的库存成本报告（Purchased Material）。
    返回 DataFrame，列：Item, Itemdesc, Prodcode, Units, Unitcost

    token_expired_retry=True 时表示已在重试中，不再重试 401。
    """
    site_cfg = SITE_API_CONFIG[site_ref]
    clm_param = urllib.parse.quote(site_cfg["clmParam"], safe="")
    url = (
        f"{INFOR_API_BASE}/{INFOR_TENANT}/CSI/IDORequestService/ido/load/{INFOR_IDO}"
        f"?clm={INFOR_REPORT_PROC}"
        f"&clmParam={clm_param}"
        f"&readonly=true"
        f"&properties={INFOR_PROPERTIES}"
    )

    headers = {
        "Authorization": f"Bearer {token}",
        "X-Infor-MongooseConfig": site_cfg["mongoose_config"],
        "Accept": "application/json",
        "Content-Type": "application/json",
    }

    print(f"  🌐 调用 Infor API: Site {site_ref} ({site_cfg['mongoose_config']})...")

    req = urllib.request.Request(url, headers=headers, method="GET")
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            raw = resp.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        if e.code == 401 and not token_expired_retry:
            print(f"  🔄 Site {site_ref}: Token 过期 (401)，强制刷新...")
            new_token = _load_infor_token(force_refresh=True)
            return _fetch_infor_site(site_ref, new_token, token_expired_retry=True)
        elif e.code == 401:
            raise RuntimeError(
                f"❌ Infor API 认证失败 (401) - Site {site_ref}\n"
                f"   Token 刷新后仍然无效，请检查 OAuth2 凭据"
            ) from e
        elif e.code == 502:
            raise RuntimeError(
                f"❌ Infor API 502 Bad Gateway - Site {site_ref}\n"
                f"   可能原因：VPN 未连接、Infor CloudSuite 服务暂时不可用\n"
                f"   响应：{body[:300]}"
            ) from e
        else:
            raise RuntimeError(
                f"❌ Infor API HTTP {e.code} - Site {site_ref}\n"
                f"   响应：{body[:300]}"
            ) from e
    except urllib.error.URLError as e:
        raise RuntimeError(
            f"❌ 网络连接失败 - Site {site_ref}: {e.reason}\n"
            f"   请检查 VPN 或网络连接"
        ) from e

    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        raise RuntimeError(
            f"❌ Infor API 返回非 JSON 内容 - Site {site_ref}: {e}\n"
            f"   原始响应前 500 字符：{raw[:500]}"
        ) from e

    rows = []
    props = INFOR_PROPERTIES.split(",")

    try:
        item_list = data.get("Items", {})
        if isinstance(item_list, dict):
            item_list = item_list.get("Items", [])
        if not isinstance(item_list, list):
            raise ValueError(f"响应结构异常，Items 不是列表: {type(item_list)}")

        for record in item_list:
            row = {}
            for prop in props:
                row[prop] = record.get(prop, "")
            rows.append(row)

    except Exception as e:
        raise RuntimeError(
            f"❌ Infor API 响应解析失败 - Site {site_ref}: {e}\n"
            f"   响应结构：{str(data)[:500]}"
        ) from e

    df = pd.DataFrame(rows, columns=props)
    print(f"  ✅ Site {site_ref}: API 返回 {len(df)} 条记录")

    if df.empty:
        print(f"  ⚠️  Site {site_ref}: 未返回数据")
        return pd.DataFrame(columns=["Item", "Description", "Per", "Unitcost", "Unitscost", "ProductCode", "Source"])

    # 列名标准化
    df = df.rename(columns={
        "Itemdesc":  "Description",
        "Units":     "Per",
        "Prodcode":  "ProductCode",
        "Pmtcode":   "Source",
    })

    df["Item"]      = df["Item"].astype(str).str.strip().str.upper()
    df["Description"] = df["Description"].astype(str).str.strip()
    df["Per"]       = pd.to_numeric(df["Per"],       errors="coerce").round(8).fillna(0)
    df["Unitcost"]  = pd.to_numeric(df["Unitcost"],  errors="coerce").round(8).fillna(0)
    df["Unitscost"] = pd.to_numeric(df["Unitscost"], errors="coerce").round(8).fillna(0)
    if "ProductCode" in df.columns:
        df["ProductCode"] = df["ProductCode"].astype(str).str.strip()
    if "Source" in df.columns:
        df["Source"] = df["Source"].astype(str).str.strip()
        df["Source"] = df["Source"].map(
            lambda v: {"P": "Purchased", "M": "Manufactured"}.get(v, v)
        )

    df = df[df["Unitscost"] != 0].copy()
    return df[["Item", "Description", "Per", "Unitcost", "Unitscost", "ProductCode", "Source"]]


def _fetch_wip_site(site_ref: str, token: str, token_expired_retry: bool = False) -> tuple:
    """
    调用 Infor CSI IDO API 获取指定站点的 WIP 数据。

    返回 (total, records)：
      total   = float，AcctTot 汇总行合计（原始货币，330 为 CNY）
      records = list[dict]，每条明细记录的 JobAcct, Des, AcctTot 等字段
    调用方负责货币换算和入库。

    token_expired_retry=True 时表示已在重试中，不再重试 401。
    """
    site_cfg = SITE_WIP_CONFIG[site_ref]
    clm_param = urllib.parse.quote(site_cfg["clmParam"], safe="")
    url = (
        f"{INFOR_API_BASE}/{INFOR_TENANT}/CSI/IDORequestService/ido/load/{INFOR_WIP_IDO}"
        f"?clm={INFOR_WIP_PROC}"
        f"&properties={urllib.parse.quote(INFOR_WIP_PROPS, safe='')}"
        f"&clmParam={clm_param}"
        f"&readonly=true"
    )

    headers = {
        "Authorization": f"Bearer {token}",
        "X-Infor-MongooseConfig": site_cfg["mongoose_config"],
        "Accept": "application/json",
        "Content-Type": "application/json",
    }

    print(f"  🌐 WIP API: Site {site_ref} ({site_cfg['mongoose_config']})...")

    req = urllib.request.Request(url, headers=headers, method="GET")
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            raw = resp.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        if e.code == 401 and not token_expired_retry:
            print(f"  🔄 WIP Site {site_ref}: Token 过期 (401)，强制刷新...")
            new_token = _load_infor_token(force_refresh=True)
            return _fetch_wip_site(site_ref, new_token, token_expired_retry=True)
        elif e.code == 401:
            raise RuntimeError(
                f"❌ WIP API 认证失败 (401) - Site {site_ref}\n"
                f"   Token 刷新后仍然无效，请检查 OAuth2 凭据"
            ) from e
        elif e.code == 502:
            raise RuntimeError(
                f"❌ WIP API 502 Bad Gateway - Site {site_ref}\n"
                f"   响应：{body[:300]}"
            ) from e
        else:
            raise RuntimeError(
                f"❌ WIP API HTTP {e.code} - Site {site_ref}\n"
                f"   响应：{body[:300]}"
            ) from e
    except urllib.error.URLError as e:
        raise RuntimeError(
            f"❌ WIP 网络连接失败 - Site {site_ref}: {e.reason}"
        ) from e

    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        raise RuntimeError(
            f"❌ WIP API 返回非 JSON - Site {site_ref}: {e}\n"
            f"   响应前 500 字符：{raw[:500]}"
        ) from e

    # 解析 IDO 响应
    item_list = data.get("Items", [])
    if not isinstance(item_list, list):
        raise ValueError(f"响应结构异常，Items 不是列表: {type(item_list)}")

    if not item_list:
        print(f"  ⚠️  WIP Site {site_ref}: 未返回数据，金额视为 0")
        return (0.0, [])

    # 分离汇总行和明细行
    total = 0.0
    summary_records = []  # 汇总行（IsDetail==0），用于写入 SLTotalWIPValueByAcountReport
    for record in item_list:
        is_detail = str(record.get("IsDetail", "")).strip()
        acct_tot_val = record.get("AcctTot")
        try:
            acct_tot = float(acct_tot_val) if acct_tot_val not in ("", None) else 0.0
        except (ValueError, TypeError):
            acct_tot = 0.0

        if is_detail in ("", "0", "false", "False"):
            total += acct_tot
            summary_records.append({
                "JobAcct": str(record.get("JobAcct", "")).strip() or None,
                "Des": str(record.get("Des", "")).strip() or None,
                "AcctTot": acct_tot,
            })

    print(f"  ✅ WIP Site {site_ref}: AcctTot 合计 = {total:,.2f} ({len(summary_records)} 汇总行)")
    return (total, summary_records)


def _fetch_wip_totals_from_db(
    sites: List[str],
    balance_date: str,
    server: str,
    username: str,
    password: str,
    database: str,
    port: int = 1433,
) -> Dict[str, Optional[float]]:
    """从 SLTotalWIPValueByAcountReport 读取指定日期 WIP（USD）。缺失站点返回 None。"""
    result: Dict[str, Optional[float]] = {s: None for s in sites}
    try:
        conn = _db_connect(server, username, password, database, port)
        cur = conn.cursor()
        cur.execute(
            """
            SELECT SiteRef, ISNULL(SUM(AcctTot), 0) AS WipTotal
            FROM dbo.SLTotalWIPValueByAcountReport
            WHERE BalanceDate = %s AND SiteRef IN (%s, %s, %s)
            GROUP BY SiteRef
            """,
            (balance_date, "310", "330", "410"),
        )
        rows = cur.fetchall()
        conn.close()

        for site_ref, wip_total in rows:
            site = str(site_ref).strip()
            if site not in result:
                continue
            val = float(wip_total) if wip_total is not None else 0.0
            _, fx = SITE_CURRENCY.get(site, ("USD", 1.0))
            result[site] = round(val * fx, 2)
    except Exception as e:
        print(f"  ⚠️  从数据库读取 WIP 失败 (BalanceDate={balance_date})：{e}")

    return result


def _fetch_inventory_balances_from_db(
    sites: List[str],
    balance_date: str,
    server: str,
    username: str,
    password: str,
    database: str,
    port: int = 1433,
) -> Dict[str, Dict[str, Optional[float]]]:
    """从 SLTotalInventory 读取指定日期 RM/FG Daily Balance（USD）。缺失站点返回 None。"""
    result: Dict[str, Dict[str, Optional[float]]] = {
        s: {"rm": None, "fg": None} for s in sites
    }
    try:
        conn = _db_connect(server, username, password, database, port)
        cur = conn.cursor()
        cur.execute(
            """
            SELECT SiteRef, Source, ISNULL(SUM(Unitscost), 0) AS TotalAmt
            FROM dbo.SLTotalInventory
            WHERE BalanceDate = %s AND SiteRef IN (%s, %s, %s)
            GROUP BY SiteRef, Source
            """,
            (balance_date, "310", "330", "410"),
        )
        rows = cur.fetchall()
        conn.close()

        seen_sites: set[str] = set()
        for site_ref, source, total_amt in rows:
            site = str(site_ref).strip()
            if site not in result:
                continue
            if site not in seen_sites:
                result[site] = {"rm": 0.0, "fg": 0.0}
                seen_sites.add(site)

            val = float(total_amt) if total_amt is not None else 0.0
            _, fx = SITE_CURRENCY.get(site, ("USD", 1.0))
            usd_val = round(val * fx, 2)
            src = str(source).strip().upper()
            if src == "PURCHASED":
                result[site]["rm"] = usd_val
            elif src == "MANUFACTURED":
                result[site]["fg"] = usd_val
    except Exception as e:
        print(f"  ⚠️  从数据库读取 Total Inventory 失败 (BalanceDate={balance_date})：{e}")

    return result


def _run_daily_snapshot_once(balance_date: str) -> bool:
    """按日期触发一次 daily_inventory_snapshot.py，同日期重复调用会跳过。"""
    if balance_date in _daily_snapshot_synced_dates:
        return True

    try:
        snapshot_script = Path(__file__).with_name("daily_inventory_snapshot.py")
        cmd = [sys.executable, str(snapshot_script), "--balance-date", balance_date]
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
        if proc.returncode == 0:
            _daily_snapshot_synced_dates.add(balance_date)
            print("  ✅ daily_inventory_snapshot.py 同步完成")
            return True
        print("  ⚠️  daily_inventory_snapshot.py 同步失败")
        if proc.stdout:
            print(proc.stdout[-1000:])
        if proc.stderr:
            print(proc.stderr[-1000:])
        return False
    except Exception as e:
        print(f"  ⚠️  触发 daily_inventory_snapshot.py 失败：{e}")
        return False


def fetch_inventory_daily_balances(
    sites: Optional[List[str]] = None,
    as_of_date: Optional[str] = None,
    db_config: Optional[dict] = None,
) -> Dict[str, Dict[str, Optional[float]]]:
    """
    获取 RM/FG 的 MTD Daily Balance（USD）。
    规则：
    - 先读 SLTotalInventory(BalanceDate=as_of_date)
    - 若 as_of_date < 今天 且站点缺失：保留空白(None)
    - 若 as_of_date = 今天 且站点缺失：自动触发 daily_inventory_snapshot.py 后重查
    """
    if sites is None:
        sites = ["310", "330", "410"]
    result = {s: {"rm": None, "fg": None} for s in sites}
    if not db_config:
        return result

    report_dt = datetime.strptime(as_of_date, "%Y-%m-%d") if as_of_date else datetime.today()
    balance_date = report_dt.strftime("%Y-%m-%d")
    today_str = datetime.today().strftime("%Y-%m-%d")

    print(f"\n📦 读取 Total Inventory 快照 (BalanceDate={balance_date})...")
    inv_db = _fetch_inventory_balances_from_db(
        sites=sites,
        balance_date=balance_date,
        server=db_config["server"],
        username=db_config["username"],
        password=db_config["password"],
        database=db_config["database"],
        port=db_config.get("port", 1433),
    )

    missing_sites = [s for s in sites if inv_db.get(s, {}).get("rm") is None and inv_db.get(s, {}).get("fg") is None]
    if not missing_sites:
        print("  ✅ Total Inventory 快照齐全")
        return inv_db

    if balance_date < today_str:
        print(f"  ⚠️  Total Inventory 快照缺失站点: {missing_sites}（历史日期，保持空白）")
        return inv_db

    if balance_date == today_str:
        print(f"  🔄 当天 Total Inventory 快照缺失站点: {missing_sites}，尝试自动同步...")
        _run_daily_snapshot_once(balance_date)
        inv_db_retry = _fetch_inventory_balances_from_db(
            sites=sites,
            balance_date=balance_date,
            server=db_config["server"],
            username=db_config["username"],
            password=db_config["password"],
            database=db_config["database"],
            port=db_config.get("port", 1433),
        )
        remain_missing = [s for s in sites if inv_db_retry.get(s, {}).get("rm") is None and inv_db_retry.get(s, {}).get("fg") is None]
        if remain_missing:
            print(f"  ⚠️  同步后 Total Inventory 仍缺失站点: {remain_missing}（保持空白）")
        return inv_db_retry

    return inv_db


def fetch_wip_totals(
    sites: Optional[List[str]] = None,
    as_of_date: Optional[str] = None,
    db_config: Optional[dict] = None,
) -> Dict[str, Optional[float]]:
    """
    获取各站点 WIP 总金额（USD）。
    - 310 / 410：API 返回 USD，直接使用
    - 330：API 返回 CNY，按 ÷6.838784 换算 USD

    返回 {site_ref: wip_usd_amount or None}
    - 有值：该站点 WIP（USD）
    - None：该日期无快照（历史日期保留空白；当天同步失败时也保留空白）
    """
    if sites is None:
        sites = ["310", "330", "410"]

    report_dt = datetime.strptime(as_of_date, "%Y-%m-%d") if as_of_date else datetime.today()
    balance_date = report_dt.strftime("%Y-%m-%d")
    today_str = datetime.today().strftime("%Y-%m-%d")

    # 1) 先按 as-of-date 读数据库快照
    if db_config:
        print(f"\n📦 读取 WIP 快照 (BalanceDate={balance_date})...")
        wip_db = _fetch_wip_totals_from_db(
            sites=sites,
            balance_date=balance_date,
            server=db_config["server"],
            username=db_config["username"],
            password=db_config["password"],
            database=db_config["database"],
            port=db_config.get("port", 1433),
        )
        missing_sites = [s for s in sites if wip_db.get(s) is None]
        if not missing_sites:
            grand = sum(v for v in wip_db.values() if v is not None)
            print(f"  ✅ WIP 快照齐全，Grand Total (USD): ${grand:,.2f}")
            return wip_db

        # 2) 历史日期缺失：保留空白，不回落 API
        if balance_date < today_str:
            print(f"  ⚠️  WIP 快照缺失站点: {missing_sites}（历史日期，保持空白）")
            grand = sum(v for v in wip_db.values() if v is not None)
            print(f"  📊 WIP Grand Total (USD, available only): ${grand:,.2f}")
            return wip_db

        # 3) 当天缺失：尝试调用 daily_inventory_snapshot.py 同步后再读一次
        if balance_date == today_str:
            print(f"  🔄 当天 WIP 快照缺失站点: {missing_sites}，尝试自动同步 daily_inventory_snapshot.py ...")
            _run_daily_snapshot_once(balance_date)
            print("  🔁 重新读取 WIP 快照...")

            wip_db_retry = _fetch_wip_totals_from_db(
                sites=sites,
                balance_date=balance_date,
                server=db_config["server"],
                username=db_config["username"],
                password=db_config["password"],
                database=db_config["database"],
                port=db_config.get("port", 1433),
            )
            remain_missing = [s for s in sites if wip_db_retry.get(s) is None]
            if remain_missing:
                print(f"  ⚠️  同步后仍缺失站点: {remain_missing}（保持空白）")
            grand = sum(v for v in wip_db_retry.values() if v is not None)
            print(f"  📊 WIP Grand Total (USD, available only): ${grand:,.2f}")
            return wip_db_retry

    # 无 db_config 时保留原有 API 逻辑（向后兼容）
    print("\n📦 获取 WIP 数据...")
    try:
        token = _load_infor_token()
    except RuntimeError as e:
        print(f"  ❌ WIP Token 获取失败：{e}")
        return {s: None for s in sites}

    wip_totals: Dict[str, Optional[float]] = {}
    for site_ref in sites:
        try:
            raw_total, _ = _fetch_wip_site(site_ref, token)
            # 330 是 CNY，换算 USD
            _, fx = SITE_CURRENCY.get(site_ref, ("USD", 1.0))
            usd_total = raw_total * fx
            if fx != 1.0:
                print(f"  💱 WIP Site {site_ref}: CNY {raw_total:,.2f} → USD {usd_total:,.2f}")
            wip_totals[site_ref] = round(usd_total, 2)
        except RuntimeError as e:
            print(f"  ❌ WIP Site {site_ref} 获取失败：{e}\n     金额留空")
            wip_totals[site_ref] = None

    grand_wip = sum(v for v in wip_totals.values() if v is not None)
    print(f"  📊 WIP Grand Total (USD): ${grand_wip:,.2f}")
    return wip_totals


def _db_connect(server, username, password, database, port=1433):
    """独立的数据库连接函数（供 WIP/库存同步使用）"""
    host = server
    instance = None
    if "\\" in host:
        host, instance = host.split("\\", 1)

    kwargs = dict(
        server=f"{host}\\{instance}" if instance else host,
        user=username,
        password=password,
        database=database,
        tds_version="7.4",
        login_timeout=15,
        conn_properties="SET TEXTSIZE 65536",
    )
    if not instance:
        kwargs["port"] = port

    return pymssql.connect(**kwargs)


def _enrich_with_slitems(df: pd.DataFrame, site_ref: str,
                          server, username, password, database, port=1433) -> pd.DataFrame:
    """
    从 SLItems 表按 Item + SiteRef 关联补充项目信息。
    补充字段：ProductCode, Sourcing（PMTCode）
    若数据库不可用，则以空值填充（不阻断主流程）。
    """
    if df.empty:
        return df

    items = df["Item"].dropna().unique().tolist()
    if not items:
        return df

    try:
        conn = _db_connect(server, username, password, database, port)
        placeholders = ",".join(["%s"] * len(items))
        query = f"""
            SELECT UPPER(LTRIM(RTRIM(item))) AS Item,
                   ProductCode,
                   PMTCode AS Sourcing
            FROM [csi_datawarehouse].[dbo].[SLItems]
            WHERE SiteRef = %s
              AND UPPER(LTRIM(RTRIM(item))) IN ({placeholders})
        """
        params = [site_ref] + items
        ref_df = pd.read_sql(query, conn, params=params)
        conn.close()

        ref_df["Item"] = ref_df["Item"].astype(str).str.strip().str.upper()
        ref_df["ProductCode"] = ref_df["ProductCode"].astype(str).str.strip()
        ref_df["Sourcing"] = ref_df["Sourcing"].astype(str).str.strip()
        ref_df = ref_df.drop_duplicates(subset="Item")

        df = df.merge(ref_df[["Item", "ProductCode", "Sourcing"]], on="Item", how="left")
        df["ProductCode"] = df["ProductCode"].fillna("")
        df["Sourcing"] = df["Sourcing"].fillna("Unknown")
        print(f"  🔗 Site {site_ref}: SLItems 关联成功，{len(ref_df)} 条记录匹配")
    except Exception as e:
        print(f"  ⚠️  Site {site_ref}: SLItems 关联失败（{e}），ProductCode/Sourcing 将为空")
        df["ProductCode"] = ""
        df["Sourcing"] = ""

    return df



def save_inventory_to_db(df: pd.DataFrame, site_ref: str,
                          server, username, password, database, port=1433,
                          balance_date: Optional[str] = None) -> int:
    """
    将 API 获取的库存数据写入 SLTotalInventory 表。
    同一 BalanceDate + SiteRef 只保留最新一次数据（先删后插）。

    df 列：Item, Description, Per, Unitcost, Unitscost, [ProductCode], [Source]
    返回插入行数。
    """
    if df.empty:
        return 0

    balance_date = balance_date or datetime.now().strftime("%Y-%m-%d")
    create_date = datetime.now()

    col_product = "ProductCode" if "ProductCode" in df.columns else None
    col_source = "Source" if "Source" in df.columns else None

    rows = []
    for _, row in df.iterrows():
        rows.append((
            str(row["Item"]).strip(),
            str(row.get("Description", "")).strip() or None,
            float(row["Per"]) if pd.notna(row["Per"]) else None,
            float(row["Unitcost"]) if pd.notna(row["Unitcost"]) else None,
            float(row["Unitscost"]) if pd.notna(row["Unitscost"]) else None,
            str(row[col_product]).strip() if col_product and pd.notna(row.get(col_product)) else None,
            str(row[col_source]).strip() if col_source and pd.notna(row.get(col_source)) else None,
            site_ref,
            balance_date,
            create_date,
        ))

    try:
        conn = _db_connect(server, username, password, database, port)
        cur = conn.cursor()
        cur.execute(
            "DELETE FROM dbo.SLTotalInventory WHERE SiteRef = %s AND BalanceDate = %s",
            (site_ref, balance_date),
        )
        deleted = cur.rowcount
        cur.executemany(
            "INSERT INTO dbo.SLTotalInventory "
            "(Item, [Description], Per, Unitcost, Unitscost, ProductCode, Source, SiteRef, BalanceDate, CreateDate) "
            "VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)",
            rows,
        )
        conn.commit()
        cur.close()
        conn.close()
        msg = f"  💾 Site {site_ref}: {len(rows)} 行写入 SLTotalInventory (BalanceDate={balance_date})"
        if deleted:
            msg += f" [替换 {deleted} 行旧数据]"
        print(msg)
        return len(rows)
    except Exception as e:
        print(f"  ⚠️  Site {site_ref}: 写入数据库失败 ({e})，不影响主流程")
        return 0


def save_wip_to_db(site_ref: str, wip_raw_amount: float, wip_records: list,
                    server, username, password, database, port=1433,
                    balance_date: Optional[str] = None) -> bool:
    """
    将单站点 WIP 数据写入 SLTotalWIPValueByAcountReport 表。
    同一 BalanceDate + SiteRef 只保留最新一次数据（先删后插）。

    wip_raw_amount: API 返回的原始合计金额（用于日志）
    wip_records: list[dict]，每条含 JobAcct, Des, AcctTot（汇总行明细）
    返回 True 表示写入成功。
    """
    balance_date = balance_date or datetime.now().strftime("%Y-%m-%d")
    create_date = datetime.now()

    try:
        conn = _db_connect(server, username, password, database, port)
        cur = conn.cursor()

        # 先删除同一天同站点的旧数据
        cur.execute(
            "DELETE FROM dbo.SLTotalWIPValueByAcountReport WHERE SiteRef = %s AND BalanceDate = %s",
            (site_ref, balance_date),
        )

        if wip_records:
            rows = []
            for rec in wip_records:
                rows.append((
                    site_ref,
                    balance_date,
                    rec.get("JobAcct"),
                    rec.get("Des"),
                    round(rec["AcctTot"], 8),
                    create_date,
                ))
            cur.executemany(
                "INSERT INTO dbo.SLTotalWIPValueByAcountReport "
                "(SiteRef, BalanceDate, JobAcct, Des, AcctTot, CreateDate) "
                "VALUES (%s, %s, %s, %s, %s, %s)",
                rows,
            )
            print(f"  💾 WIP Site {site_ref}: {len(rows)} 条汇总行写入 SLTotalWIPValueByAcountReport (BalanceDate={balance_date})")
        else:
            # 即使没有明细行，也插入一条合计记录（JobAcct/Des 为 NULL）
            cur.execute(
                "INSERT INTO dbo.SLTotalWIPValueByAcountReport "
                "(SiteRef, BalanceDate, JobAcct, Des, AcctTot, CreateDate) "
                "VALUES (%s, %s, NULL, NULL, %s, %s)",
                (site_ref, balance_date, round(wip_raw_amount, 8), create_date),
            )
            print(f"  💾 WIP Site {site_ref}: 合计 {wip_raw_amount:,.2f} 写入 SLTotalWIPValueByAcountReport (BalanceDate={balance_date})")

        conn.commit()
        cur.close()
        conn.close()
        return True
    except Exception as e:
        print(f"  ⚠️  WIP Site {site_ref}: 写入数据库失败 ({e})，不影响主流程")
        return False




# ──────────────────────────────────────────────────────────────
# 守护进程 — 每天 00:00 同步 + 每天 09:00 跑报表
# ──────────────────────────────────────────────────────────────
DAEMON_HOUR_SNAPSHOT = 0
DAEMON_MINUTE_SNAPSHOT = 0
DAEMON_HOUR_REPORT = 9
DAEMON_MINUTE_REPORT = 0

def _next_schedule(now):
    """
    计算下一次执行时间和任务类型。
    规则：
      - 每天 00:00：同步 daily_inventory_snapshot.py（Inventory + WIP + SLMatltrans）
      - 每天 09:00：运行库存跟踪报表
    返回 (target_datetime, task_type)  task_type = 'snapshot' | 'report'
    """
    def _make(hour, minute):
        return now.replace(hour=hour, minute=minute, second=0, microsecond=0)

    today_snapshot = _make(DAEMON_HOUR_SNAPSHOT, DAEMON_MINUTE_SNAPSHOT)
    today_report = _make(DAEMON_HOUR_REPORT, DAEMON_MINUTE_REPORT)

    candidates = []
    # 每天 00:00 同步 snapshot
    if now < today_snapshot:
        candidates.append((today_snapshot, "snapshot"))
    # 每天 09:00 跑报表
    if now < today_report:
        candidates.append((today_report, "report"))

    if candidates:
        return min(candidates, key=lambda x: x[0])

    # 今天的任务都过了，看明天
    tomorrow = (now + timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    return (tomorrow.replace(hour=DAEMON_HOUR_SNAPSHOT, minute=DAEMON_MINUTE_SNAPSHOT), "snapshot")


def schedule_loop(run_func, args):
    """
    双调度守护进程：
      - 每天 00:00：同步 daily_inventory_snapshot.py（Inventory + WIP + SLMatltrans）
      - 每天 09:00：运行库存跟踪报表 + 发邮件
    Docker 停止时收到 SIGTERM 自然退出。
    """

    print(f"⏰ Daemon mode started:")
    print(f"   每天   {DAEMON_HOUR_SNAPSHOT:02d}:{DAEMON_MINUTE_SNAPSHOT:02d} → 同步 Inventory + WIP + SLMatltrans")
    print(f"   每天   {DAEMON_HOUR_REPORT:02d}:{DAEMON_MINUTE_REPORT:02d} → 运行库存跟踪报表")

    while True:
        target, task_type = _next_schedule(datetime.now())
        wait_seconds = (target - datetime.now()).total_seconds()

        task_labels = {"snapshot": "同步数据快照", "report": "运行库存报表"}
        label = task_labels.get(task_type, task_type)
        print(f"⏳ Next: {target.strftime('%Y-%m-%d %H:%M:%S')} [{label}] ({wait_seconds/3600:.1f}h)")

        try:
            time.sleep(max(wait_seconds, 0))
        except KeyboardInterrupt:
            print("\n🛑 Daemon stopped.")
            break

        try:
            if task_type == "snapshot":
                print(f"\n{'='*60}")
                print(f"  🔄 定时任务：同步 daily_inventory_snapshot.py")
                print(f"{'='*60}")
                snapshot_script = Path(__file__).with_name("daily_inventory_snapshot.py")
                balance_date = datetime.now().strftime("%Y-%m-%d")
                cmd = [sys.executable, str(snapshot_script), "--balance-date", balance_date]
                print(f"  📋 执行: {' '.join(cmd)}")
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
                if result.stdout:
                    print(result.stdout)
                if result.returncode != 0:
                    print(f"  ⚠️  snapshot 退出码 {result.returncode}")
                    if result.stderr:
                        print(f"  STDERR: {result.stderr[:500]}")
                else:
                    print(f"  ✅ daily_inventory_snapshot.py 同步完成")
            elif task_type == "report":
                run_func(args)
        except Exception as e:
            print(f"❌ Scheduled run failed: {e}")
            import traceback
            traceback.print_exc()
            time.sleep(300)


def run_once(args):
    """执行一次完整的库存跟踪+邮件流程"""
    if args.all_sites:
        result = run_all_sites(
            server=args.server, database=args.database,
            username=args.username, password=args.password, port=args.db_port,
            as_of_date=args.as_of_date or None,
        )
        if result and not args.no_email:
            print("\n📧 发送汇总邮件...")
            send_summary_email(
                result,
                to_addr=args.email_to,
                cc_addr=args.email_cc,
                smtp_host=args.smtp_host,
                smtp_port=args.smtp_port,
                smtp_user=args.smtp_user,
                smtp_password=args.smtp_password,
                smtp_tls=args.smtp_tls,
                from_addr=args.email_from,
                db_config={
                    "server": args.server,
                    "database": args.database,
                    "username": args.username,
                    "password": args.password,
                    "port": args.db_port,
                },
                report_date=args.as_of_date or result.get("report_date"),
            )
        return result
    else:
        tracker = InventoryTracker(
            server=args.server, database=args.database,
            username=args.username, password=args.password, port=args.db_port,
            site_ref=args.site,
            output_file=args.output,
            as_of_date=args.as_of_date or None,
        )
        return tracker.run()


# ──────────────────────────────────────────────────────────────
# CLI
# ──────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="库存金额跟踪系统（pymssql + SMTP）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  # 单次批量运行
  python inventory_tracking.py --all-sites

  # 守护模式：每天 00:00 同步 + 09:00 报表（Docker 推荐）
  python inventory_tracking.py --all-sites --daemon

  # 指定日期运行
  python inventory_tracking.py --all-sites --as-of-date 2026-06-02

  # 单站点运行
  python inventory_tracking.py --site 310

  # 不发送邮件
  python inventory_tracking.py --all-sites --no-email
        """,
    )
    # 数据库
    parser.add_argument("-s", "--server",   default=os.environ.get("SQL_SERVER_HOST", ""))
    parser.add_argument("-d", "--database", default=os.environ.get("SQL_SERVER_DATABASE", ""))
    parser.add_argument("-u", "--username", default=os.environ.get("SQL_SERVER_USERNAME", ""))
    parser.add_argument("-p", "--password", default=os.environ.get("SQL_SERVER_PASSWORD", ""))
    parser.add_argument("--db-port", type=int, default=int(os.environ.get("SQL_SERVER_PORT", 1433)))
    # 模式
    parser.add_argument("--site", default="310", help="单站点模式 (默认310)")
    parser.add_argument("--all-sites", action="store_true", help="批量运行310/330/410")
    parser.add_argument("--daemon", action="store_true", help="守护模式：每天 00:00 同步 + 09:00 跑报表")
    parser.add_argument("-o", "--output", help="输出文件名 (单站点模式)")
    parser.add_argument("--as-of-date", default="", help="模拟报表日期，格式 YYYY-MM-DD")
    # 邮件
    parser.add_argument("--no-email", action="store_true", help="不发送邮件")
    parser.add_argument("--email-to",
        default=os.environ.get("MAIL_TO", ""))
    parser.add_argument("--email-cc",
        default=os.environ.get("MAIL_CC", ""))
    parser.add_argument("--email-from",
        default=os.environ.get("MAIL_FROM", "suzinventoryvaluationdailyreport@nai-group.com"))
    # SMTP（优先级：.env → 默认值）
    parser.add_argument("--smtp-host",     default=os.environ.get("SMTP_HOST", "localhost"))
    parser.add_argument("--smtp-port",     type=int, default=int(os.environ.get("SMTP_PORT", 25)))
    parser.add_argument("--smtp-user",     default=os.environ.get("SMTP_USER", ""))
    parser.add_argument("--smtp-password", default=os.environ.get("SMTP_PASSWORD", ""))
    parser.add_argument("--smtp-tls",      action="store_true",
        default=os.environ.get("SMTP_TLS", "false").lower() == "true")

    args = parser.parse_args()

    if args.as_of_date:
        try:
            datetime.strptime(args.as_of_date, "%Y-%m-%d")
        except ValueError:
            print("❌ --as-of-date 格式错误，应为 YYYY-MM-DD")
            sys.exit(1)

    missing_db = []
    if not args.server:
        missing_db.append("SQL_SERVER_HOST")
    if not args.database:
        missing_db.append("SQL_SERVER_DATABASE")
    if not args.username:
        missing_db.append("SQL_SERVER_USERNAME")
    if not args.password:
        missing_db.append("SQL_SERVER_PASSWORD")
    if missing_db:
        print(f"❌ 缺少数据库配置: {', '.join(missing_db)}")
        print("   请在 .env 中配置后重试。")
        sys.exit(1)

    if args.daemon:
        if args.as_of_date:
            print("❌ 守护模式不支持固定 --as-of-date，请去掉该参数。")
            sys.exit(1)
        # 守护模式：双调度（每月1日 00:15 期初 + 每天 09:00 报表）
        schedule_loop(lambda a: run_once(a), args)
    else:
        run_once(args)


if __name__ == "__main__":
    # ── 加载 .env 到环境变量 ──
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except ImportError:
        pass
    main()
