#!/usr/bin/env python3
"""
库存金额跟踪系统 (Inventory Tracking System)
============================================
功能：
1. 从上月底Excel文件读取期初库存余额（支持 .xlsx / .xlsb）
2. 连接SQL Server查询当月MTD物料事务（使用 pymssql，无需 ODBC 驱动）
3. 按TransType+RefType分类：Received / Consumed / Other Transaction
4. 按Item汇总Qty和AMT（金额直接使用 TotalPosted）
5. 计算期末余额 = 期初 + Received - Consumed - Other
6. 导出Excel报表（Project Summary + Summary + Detail 三页，Project Summary 在前）
7. 支持多站点（310/330/410）批量运行
8. 通过 SMTP 发送邮件（支持内网 Relay，无需 Outlook）
9. 通过 config.ini 外部化配置（DB、SMTP、邮件收件人）

作者: WorkBuddy for NAI Group
日期: 2026-05-25 | 更新: 2026-05-27 (pymssql + SMTP + config.ini)
"""

import os
import sys
import time
import glob
import warnings
import argparse
import configparser
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple
from pathlib import Path

import pymssql
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter


# ──────────────────────────────────────────────────────────────
# 读取 config.ini（如存在）
# ──────────────────────────────────────────────────────────────
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.ini")

def load_config() -> configparser.ConfigParser:
    cfg = configparser.ConfigParser()
    if os.path.exists(CONFIG_FILE):
        cfg.read(CONFIG_FILE, encoding="utf-8")
    return cfg


# ──────────────────────────────────────────────────────────────
# TransType + RefType → Category 映射规则
# ──────────────────────────────────────────────────────────────
CATEGORY_MAP: Dict[Tuple[str, str], str] = {
    # Received (入库)
    ("R", "P"): "Received",   # PO Receipt
    ("W", "P"): "Received",   # PO Withdraw

    # Consumed (消耗)
    ("I", "J"): "Consumed",   # Job Issue / WIP Change
    ("W", "J"): "Consumed",   # Job Withdrawal / Return
    ("S", "O"): "Consumed",   # Order Ship

    # Other Transaction (其他)
    ("A", "I"): "Other",      # Adjustment
    ("M", "I"): "Other",      # Stock Move
    ("G", "I"): "Other",      # Misc Receipt
    ("H", "I"): "Other",      # Misc Issue
    ("C", "J"): "Other",      # Job Complete
    ("F", "J"): "Other",      # Job Finish
    ("N", "J"): "Other",      # Job Labor / Next Operation
    ("W", "R"): "Other",      # RMA Withdraw
}

TRANS_DESCRIPTIONS: Dict[Tuple[str, str], str] = {
    ("R", "P"): "PO Receipt",
    ("W", "P"): "PO Withdraw",
    ("I", "J"): "Job Issue / WIP Change",
    ("W", "J"): "Job Withdrawal / Return",
    ("S", "O"): "Order Ship",
    ("A", "I"): "Adjustment",
    ("M", "I"): "Stock Move",
    ("G", "I"): "Misc Receipt",
    ("H", "I"): "Misc Issue",
    ("C", "J"): "Job Complete",
    ("F", "J"): "Job Finish",
    ("N", "J"): "Job Labor / Next Operation",
    ("W", "R"): "RMA Withdraw",
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
        return self.prev_balance + self.recv_amt - self.cons_amt - self.other_amt

    @property
    def net_qty(self) -> float:
        return self.prev_qty + self.recv_qty - self.cons_qty - self.other_qty


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
        server: str = r"SUZVPRINT01\CUSTOMSSYS",
        database: str = "csi_datawarehouse",
        username: str = "sa",
        password: str = "xxVcDW9ED24YWX",
        port: int = 1433,
        site_ref: str = "310",
        prev_balance_file: Optional[str] = None,
        output_file: Optional[str] = None,
    ):
        self.server = server
        self.database = database
        self.username = username
        self.password = password
        self.port = port
        self.site_ref = site_ref
        self.prev_balance_file = prev_balance_file

        # Currency: site 330 is CNY, convert to USD
        curr_info = SITE_CURRENCY.get(site_ref, ("USD", 1.0))
        self.currency_label = curr_info[0]
        self.fx_rate = curr_info[1]
        self.usd_symbol = "$"

        site_label = SITE_NAMES.get(site_ref, site_ref)
        today = datetime.today()
        self.period_str = today.strftime("%Y-%m")
        default_name = f"Inventory_Balance_{site_label}_{today.strftime('%Y%m')}.xlsx"
        self.output_file = output_file or default_name

    def _get_connection(self):
        """使用 pymssql 建立数据库连接（无需 ODBC 驱动）"""
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
        从库存Excel读取上月末余额。
        支持两种列名格式：
          A) 标准格式：Item, Prev_Qty, Prev_Balance
          B) Infor导出格式：Item, Per, Unitcost（单价）→ Prev_Balance = Per × Unitcost
          C) 旧Infor格式：Item, Per, Unitscost（扩展成本）→ 直接使用
        """
        if not self.prev_balance_file or not os.path.exists(self.prev_balance_file):
            print(f"  ⚠️  期初余额文件未找到：{self.prev_balance_file}")
            print("     期初余额将设为 0")
            return pd.DataFrame(columns=["Item", "Prev_Qty", "Prev_Balance", "Description"])

        filepath = self.prev_balance_file
        # 自动检测是否需要跳过第一行（site 410 等文件第一行是标题）
        skip = 0
        if filepath.lower().endswith(".xlsb"):
            skip = 1
        else:
            # 读取前2行，检查列名是否全为 "Unnamed"（说明第一行是标题行，不是列名）
            test_df = pd.read_excel(filepath, nrows=2)
            if all(c.startswith("Unnamed") for c in test_df.columns):
                skip = 1

        df = pd.read_excel(filepath, skiprows=skip)
        df.columns = [c.strip() for c in df.columns]

        # 智能识别列名并计算期初金额
        if "Prev_Qty" in df.columns and "Prev_Balance" in df.columns:
            # 标准格式
            qty_col, amt_col = "Prev_Qty", "Prev_Balance"
            calc_extended = False
        elif "Per" in df.columns and "Unitscost" in df.columns:
            # 旧Infor格式：Unitscost 是扩展成本
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
                f"需要包含 (Item + Prev_Qty + Prev_Balance) 或 (Item + Per + Unitcost)"
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

    # ──────────────────────────────────────────────────────────
    # 2. 查询数据库（MTD事务）
    # ──────────────────────────────────────────────────────────
    def fetch_mtd_transactions(self) -> pd.DataFrame:
        today = datetime.today()
        month_start = today.replace(day=1)

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
        INNER JOIN [csi_datawarehouse].[dbo].[SLItems] i
            ON m.[Item] = i.[Item] AND m.[SiteRef] = i.[SiteRef]
        WHERE m.[SiteRef] = '{self.site_ref}'
          AND m.[TransDate] >= '{month_start.strftime('%Y-%m-%d')}'
          AND i.[PMTCode] = 'P'
        ORDER BY m.[TransDate], m.[TransNum]
        """

        print(f"  📊 查询 {self.database} Site {self.site_ref} 采购物料 (PMTCode=P) ({month_start.strftime('%Y-%m-%d')} ~)")

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
        df["TotalAmt"] = pd.to_numeric(df["TotalPosted"], errors="coerce").fillna(0).abs()

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
        return CATEGORY_MAP.get(key, "Other")

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
            else:
                ib.other_qty += qty; ib.other_amt += amt

        summary_data = []
        for item, ib in items.items():
            summary_data.append({
                "ProjectCode": ib.project_code,
                "Item": item,
                "Description": ib.description,
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
        today = datetime.today()
        site_label = SITE_NAMES.get(self.site_ref, self.site_ref)

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
        proj_headers = ["Project Code", "4/30 Balance",
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
            ws_proj.cell(row=r, column=7, value=f"=B{r}+C{r}-D{r}-E{r}")
            ws_proj.cell(row=r, column=7).number_format = "#,##0.00"
            ws_proj.cell(row=r, column=8, value="4/30+Recv-Cons-Other")
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
        ws["A1"] = f"库存金额跟踪报表 - Site {self.site_ref} ({site_label}) - 采购物料 (PMTCode=P)"
        ws["A1"].font = TITLE_FONT
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("A2:M2")
        ws["A2"] = (f"报告周期：{today.strftime('%Y年%m月')} 1日 - {today.strftime('%m月%d日')}   |   "
                     f"数据截至：{today.strftime('%Y-%m-%d %H:%M')}   |   "
                     f"仅含采购物料")
        ws["A2"].alignment = Alignment(horizontal="center")
        ws["A2"].font = Font(size=10, italic=True, color="666666")

        headers = ["Project Code", "Item", "Description", "4/30 Qty", "4/30 Balance",
                    "MTD Received Qty", "MTD Received", "MTD Consumption Qty", "MTD Consumption",
                    "MTD Other Qty", "MTD Other Transaction", "MTD Daily Balance Qty", "MTD Daily Balance"]
        ws.append(headers)
        for col_num in range(1, len(headers) + 1):
            c = ws.cell(row=3, column=col_num)
            c.fill = HEADER_FILL
            c.font = HEADER_FONT
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = THIN_BORDER

        for row_data in summary_df.itertuples(index=False):
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
        ws_det["A1"] = f"物料事务明细 - Site {self.site_ref} ({site_label}) - 采购物料 (PMTCode=P)"
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
        print(f"     Summary: {len(summary_df):,} Items (PMTCode=P)")
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
        bal_total = prev_total + recv_total - cons_total - other_total

        print(f"\n  ┌─────────────────────────────────────────┐")
        print(f"  │  期初余额:   ${prev_total:>15,.2f}       │")
        print(f"  │  + Received: ${recv_total:>15,.2f}       │")
        print(f"  │  - Consumed: ${cons_total:>15,.2f}       │")
        print(f"  │  - Other:    ${other_total:>15,.2f}       │")
        print(f"  │  = Balance:  ${bal_total:>15,.2f}       │")
        print(f"  └─────────────────────────────────────────┘")
        print(f"\n✅ Site {self.site_ref} ({site_label}) 完成！")
        return summary_df, detail_df


# ──────────────────────────────────────────────────────────────
# 多站点批量运行
# ──────────────────────────────────────────────────────────────
def run_all_sites(
    server: str, database: str, username: str, password: str, port: int = 1433,
    prev_dir: str = "Previous Balance", output_dir: Optional[str] = None,
    sites: Optional[List[str]] = None,
):
    """批量运行所有站点"""
    if sites is None:
        sites = ["310", "330", "410"]

    # 自动匹配期初文件到站点（优先匹配新命名 site XXX.xlsx）
    prev_files = {"310": None, "330": None, "410": None}
    prev_path = Path(prev_dir)
    # 第一轮：精确匹配 "site XXX.xlsx"
    for site in ["310", "330", "410"]:
        exact = prev_path / f"site {site}.xlsx"
        if exact.exists():
            prev_files[site] = str(exact)
    # 第二轮：模糊匹配旧命名 (Plant1/Plant2/PNG)
    if not all(prev_files.values()):
        for f in prev_path.iterdir():
            fname = f.name.upper()
            if not fname.endswith((".XLSX", ".XLSB")):
                continue
            if "PLANT1" in fname and not prev_files["310"]:
                prev_files["310"] = str(f)
            elif "PLANT2" in fname and not prev_files["330"]:
                prev_files["330"] = str(f)
            elif "PNG" in fname and not prev_files["410"]:
                prev_files["410"] = str(f)

    out_dir = Path(output_dir) if output_dir else Path(".")
    out_dir.mkdir(parents=True, exist_ok=True)

    all_summaries = {}
    for site in sites:
        pf = prev_files.get(site)
        if not pf or not os.path.exists(pf):
            print(f"\n⚠️  Site {site}: 未找到期初文件，跳过")
            continue

        tracker = InventoryTracker(
            server=server, database=database,
            username=username, password=password, port=port,
            site_ref=site, prev_balance_file=pf,
            output_file=str(out_dir / f"Inventory_Balance_{SITE_NAMES.get(site, site)}_{datetime.today().strftime('%Y%m')}.xlsx"),
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

    group_totals = combined.groupby("Site").agg(
        Items=("Item", "count"),
        Prev_Balance=("Prev_Balance", "sum"),
        Received_AMT=("Received_AMT", "sum"),
        Consumed_AMT=("Consumed_AMT", "sum"),
        Other_AMT=("Other_AMT", "sum"),
        Balance_AMT=("Balance_AMT", "sum"),
    ).reset_index()

    for _, row in group_totals.iterrows():
        s = row["Site"]
        print(f"\n  Site {s} ({SITE_NAMES.get(s, s)}):")
        print(f"    Items: {row['Items']:,}")
        print(f"    期初: ${row['Prev_Balance']:,.2f}  +Recv: ${row['Received_AMT']:,.2f}  "
              f"-Cons: ${row['Consumed_AMT']:,.2f}  -Other: ${row['Other_AMT']:,.2f}  "
              f"= ${row['Balance_AMT']:,.2f}")

    grand_prev = combined["Prev_Balance"].sum()
    grand_recv = combined["Received_AMT"].sum()
    grand_cons = combined["Consumed_AMT"].sum()
    grand_other = combined["Other_AMT"].sum()
    grand_bal = grand_prev + grand_recv - grand_cons - grand_other
    print(f"\n  🏢 ALL SITES Grand Total (USD):")
    print(f"    期初: ${grand_prev:,.2f}  +Recv: ${grand_recv:,.2f}  "
          f"-Cons: ${grand_cons:,.2f}  -Other: ${grand_other:,.2f}  "
          f"= ${grand_bal:,.2f}")

    return {
        "combined": combined,
        "group_totals": group_totals,
        "grand_prev": grand_prev, "grand_recv": grand_recv,
        "grand_cons": grand_cons, "grand_other": grand_other,
        "grand_bal": grand_bal, "out_dir": out_dir,
    }


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
):
    """生成固定格式的 HTML 邮件并通过 SMTP 发送"""
    gt = result["group_totals"]
    combined = result["combined"]

    def fmt_num(v):
        return f"${abs(v):,.2f}"

    def proj_count(site):
        site_df = combined[combined["Site"] == site]
        pc = site_df["ProjectCode"].dropna()
        pc = pc[pc.str.strip() != ""]
        return len(pc.unique())

    # 构建表格行
    rows_html = ""
    for _, row in gt.iterrows():
        s = row["Site"]
        label = f"{s} ({SITE_NAMES.get(s, s)})"
        pc_count = proj_count(s)
        rows_html += (
            f"<tr>"
            f"<td style='text-align:left'>{label}</td>"
            f"<td>{int(row['Items']):,}</td>"
            f"<td>{pc_count}</td>"
            f"<td>{fmt_num(row['Prev_Balance'])}</td>"
            f"<td>{fmt_num(row['Received_AMT'])}</td>"
            f"<td>{fmt_num(row['Consumed_AMT'])}</td>"
            f"<td>{fmt_num(row['Other_AMT'])}</td>"
            f"<td><b>{fmt_num(row['Balance_AMT'])}</b></td>"
            f"</tr>\n"
        )

    total_items = int(gt["Items"].sum())
    total_pc = sum(proj_count(s) for s in gt["Site"])
    rows_html += (
        f"<tr style='background-color:#D6E4F0;font-weight:bold'>"
        f"<td style='text-align:left'>Grand Total</td>"
        f"<td>{total_items:,}</td><td>-</td>"
        f"<td>{fmt_num(result['grand_prev'])}</td>"
        f"<td>{fmt_num(result['grand_recv'])}</td>"
        f"<td>{fmt_num(result['grand_cons'])}</td>"
        f"<td>{fmt_num(result['grand_other'])}</td>"
        f"<td><b>{fmt_num(result['grand_bal'])}</b></td>"
        f"</tr>\n"
    )

    today = datetime.today()
    month_label = today.strftime("%B %Y")
    date_label = today.strftime("%B %d %Y")
    prev_month = (today.replace(day=1) - timedelta(days=1))
    prev_eom_label = prev_month.strftime("%m/%d")

    html_body = f"""<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt">
<p>Below is the {month_label} Inventory Valuation Tracking Report (Purchased Materials, PMTCode=P). All amounts in USD. Site 330 CNY amounts converted at rate 6.838784.</p>

<table border="1" cellpadding="5" cellspacing="0"
  style="border-collapse:collapse;font-size:10pt;text-align:right">
<tr style="background-color:#4472C4;color:white;text-align:center">
  <th>Site</th><th>Items</th><th>Projects</th>
  <th>{prev_eom_label} Balance</th><th>MTD Received</th>
  <th>MTD Consumption</th><th>MTD Other Transaction</th>
  <th>MTD Daily Balance</th>
</tr>
{rows_html}
</table>
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
    for site in gt["Site"]:
        fname = f"Inventory_Balance_{SITE_NAMES.get(site, site)}_{today.strftime('%Y%m')}.xlsx"
        fpath = str(result["out_dir"] / fname)
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
# 守护进程 — 每日 9:00 自动执行
# ──────────────────────────────────────────────────────────────
DAEMON_HOUR = 9
DAEMON_MINUTE = 0

def schedule_loop(run_func, args):
    """无限循环，每天 9:00 执行 run_func。Docker 停止时收到 SIGTERM 自然退出。"""
    print(f"⏰ Daemon mode: will run daily at {DAEMON_HOUR:02d}:{DAEMON_MINUTE:02d} (Asia/Shanghai)")
    while True:
        now = datetime.now()
        target = now.replace(hour=DAEMON_HOUR, minute=DAEMON_MINUTE, second=0, microsecond=0)
        if now >= target:
            target += timedelta(days=1)
        wait_seconds = (target - now).total_seconds()
        print(f"⏳ Next run at {target.strftime('%Y-%m-%d %H:%M:%S')} ({wait_seconds/3600:.1f}h)")
        try:
            time.sleep(wait_seconds)
        except KeyboardInterrupt:
            print("\n🛑 Daemon stopped.")
            break
        try:
            run_func(args)
        except Exception as e:
            print(f"❌ Scheduled run failed: {e}")
            import traceback
            traceback.print_exc()
            # 出错后等待 5 分钟再重试，避免死循环刷日志
            time.sleep(300)


def run_once(args):
    """执行一次完整的库存跟踪+邮件流程"""
    if args.all_sites:
        result = run_all_sites(
            server=args.server, database=args.database,
            username=args.username, password=args.password, port=args.db_port,
            prev_dir=args.prev_dir, output_dir=args.output_dir,
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
            )
        return result
    else:
        if not args.prev_file:
            prev_dir = Path(args.prev_dir)
            site_label = SITE_NAMES.get(args.site, args.site)
            for f in prev_dir.iterdir():
                if site_label.upper() in f.name.upper() or args.site in f.name:
                    args.prev_file = str(f)
                    break
            if not args.prev_file:
                print(f"❌ 未找到 Site {args.site} 的期初余额文件。请用 -f 指定。")
                sys.exit(1)

        tracker = InventoryTracker(
            server=args.server, database=args.database,
            username=args.username, password=args.password, port=args.db_port,
            site_ref=args.site, prev_balance_file=args.prev_file,
            output_file=args.output,
        )
        return tracker.run()


# ──────────────────────────────────────────────────────────────
# CLI
# ──────────────────────────────────────────────────────────────
def main():
    # 读取 config.ini 作为默认值
    cfg = load_config()
    db_sec = cfg["database"] if "database" in cfg else {}
    mail_sec = cfg["email"] if "email" in cfg else {}
    smtp_sec = cfg["smtp"] if "smtp" in cfg else {}

    parser = argparse.ArgumentParser(
        description="库存金额跟踪系统（pymssql + SMTP）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  # 单次批量运行
  python inventory_tracking.py --all-sites

  # 守护模式：每天 09:00 自动执行，启动后不立即跑（Docker 推荐）
  python inventory_tracking.py --all-sites --daemon

  # 单站点运行
  python inventory_tracking.py --site 310 -f "Previous Balance/site 310.xlsx"

  # 不发送邮件
  python inventory_tracking.py --all-sites --no-email
        """,
    )
    # 数据库
    parser.add_argument("-s", "--server",   default=os.environ.get("SQL_SERVER_HOST") or db_sec.get("server", r"SUZVPRINT01\CUSTOMSSYS"))
    parser.add_argument("-d", "--database", default=os.environ.get("SQL_SERVER_DATABASE") or db_sec.get("database", "csi_datawarehouse"))
    parser.add_argument("-u", "--username", default=os.environ.get("SQL_SERVER_USERNAME") or db_sec.get("username", "sa"))
    parser.add_argument("-p", "--password", default=os.environ.get("SQL_SERVER_PASSWORD") or db_sec.get("password", "xxVcDW9ED24YWX"))
    parser.add_argument("--db-port", type=int, default=int(os.environ.get("SQL_SERVER_PORT") or db_sec.get("port", 1433)))
    # 模式
    parser.add_argument("--site", default="310", help="单站点模式 (默认310)")
    parser.add_argument("-f", "--prev-file", help="单站点期初余额Excel路径")
    parser.add_argument("--all-sites", action="store_true", help="批量运行310/330/410")
    parser.add_argument("--daemon", action="store_true", help="守护模式：每天 09:00 自动执行")
    parser.add_argument("--prev-dir", default=db_sec.get("prev_dir", "Previous Balance"))
    parser.add_argument("--output-dir", default=db_sec.get("output_dir", None))
    parser.add_argument("-o", "--output", help="输出文件名 (单站点模式)")
    # 邮件
    parser.add_argument("--no-email", action="store_true", help="不发送邮件")
    parser.add_argument("--email-to",
        default=os.environ.get("MAIL_TO") or mail_sec.get("to", "jason.pang@nai-group.com;shirley.ni@nai-group.com;devin.hua@nai-group.com;chn_planners@nai-group.com;chn_buyer@nai-group.com"))
    parser.add_argument("--email-cc",
        default=os.environ.get("MAIL_CC") or mail_sec.get("cc", "sky.li@nai-group.com;frank.liu@nai-group.com;shirley.ni@nai-group.com"))
    parser.add_argument("--email-from",
        default=os.environ.get("MAIL_FROM") or mail_sec.get("from_addr", "suzinventoryvaluationdailyreport@nai-group.com"))
    # SMTP（优先级：.env → config.ini → 默认值）
    parser.add_argument("--smtp-host",
        default=os.environ.get("SMTP_HOST") or smtp_sec.get("host", "localhost"))
    parser.add_argument("--smtp-port",     type=int,
        default=int(os.environ.get("SMTP_PORT") or smtp_sec.get("port", 25)))
    parser.add_argument("--smtp-user",
        default=os.environ.get("SMTP_USER") or smtp_sec.get("user", ""))
    parser.add_argument("--smtp-password",
        default=os.environ.get("SMTP_PASSWORD") or smtp_sec.get("password", ""))
    parser.add_argument("--smtp-tls",      action="store_true",
        default=(os.environ.get("SMTP_TLS") or smtp_sec.get("tls", "false")).lower() == "true")

    args = parser.parse_args()

    if args.daemon:
        # 守护模式：不立即执行，等到 09:00 再跑
        schedule_loop(lambda a: run_once(a), args)
    else:
        run_once(args)


if __name__ == "__main__":
    main()
