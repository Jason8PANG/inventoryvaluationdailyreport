"""
SLMatltrans 物料交易表同步（IDO: SLMatltrans）
多站点增量同步，每个站点分别抓取，SiteRef 存短码（310/330/410）

独立运行版本 — 不依赖 csi_datawarehouse 包。
复用 inventory_tracking.py 的 OAuth2 Token 获取和 IDO API 调用模式。

运行方式：
  python slmatltrans.py
"""

import os
import sys
import time
import json
import urllib.request
import urllib.error
import urllib.parse
from datetime import date, timedelta
from dotenv import load_dotenv

sys.stdout.reconfigure(encoding="utf-8")
load_dotenv()

# ── 复用 inventory_tracking.py 的常量和函数 ──
from inventory_tracking import (
    _load_infor_token,
    INFOR_API_BASE,
    INFOR_TENANT,
    _db_connect,
)

SITES = ["310", "330", "410"]

# 站点 → MongooseConfig header
SITE_MONGOOSE = {
    "310": "NAIGROUP_PRD_310",
    "330": "NAIGROUP_PRD_330",
    "410": "NAIGROUP_PRD_410",
}

# 增量同步回补天数（首次无数据时）
FULL_SYNC_DAYS = 30

# IDO 属性列表
MATLTRANS_PROPERTIES = (
    "TransNum,TransDate,TransType,RefType,Backflush,"
    "Whse,Loc,Lot,Wc,RefNum,RefLineSuf,RefRelease,"
    "Item,ue_GDL_Manufacturer,ue_GDL_ManufacturerItem,"
    "ue_GDL_LotSupplierLot,ue_GDL_Uf_Customer,ue_GDL_Description,"
    "Qty,MatlCost,LbrCost,FovhdCost,VovhdCost,OutCost,"
    "DocumentNum,"
    "DerMatltranCost,DerMatlTranViewTotalPosted,MatlTranViewTypeDesc,"
    "RowPointer,CreateDate,RecordDate"
)


def _to_decimal(val, default=None):
    try:
        return float(val)
    except (ValueError, TypeError):
        return default


def _fetch_ido_data(site_ref: str, token: str, filter_str: str,
                     token_expired_retry: bool = False) -> list:
    """
    调用 Infor CSI IDO API 加载 SLMatltrans 数据。
    使用 ido/load 端点 + filter 参数进行增量查询。
    """
    mongoose = SITE_MONGOOSE[site_ref]
    props_encoded = urllib.parse.quote(MATLTRANS_PROPERTIES, safe="")
    filter_encoded = urllib.parse.quote(filter_str, safe="")

    url = (
        f"{INFOR_API_BASE}/{INFOR_TENANT}/CSI/IDORequestService/ido/load/SLMatltrans"
        f"?properties={props_encoded}"
        f"&filter={filter_encoded}"
    )

    headers = {
        "Authorization": f"Bearer {token}",
        "X-Infor-MongooseConfig": mongoose,
        "Accept": "application/json",
        "Content-Type": "application/json",
    }

    req = urllib.request.Request(url, headers=headers, method="GET")
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            raw = resp.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        if e.code == 401 and not token_expired_retry:
            print(f"    🔄 SLMatltrans Site {site_ref}: Token 过期 (401)，强制刷新...")
            new_token = _load_infor_token(force_refresh=True)
            return _fetch_ido_data(site_ref, new_token, filter_str, token_expired_retry=True)
        else:
            raise RuntimeError(
                f"SLMatltrans API HTTP {e.code} - Site {site_ref}\n"
                f"   响应: {body[:300]}"
            ) from e
    except urllib.error.URLError as e:
        raise RuntimeError(
            f"SLMatltrans 网络连接失败 - Site {site_ref}: {e.reason}"
        ) from e

    data = json.loads(raw)
    # API 返回 {"Items": [...]} 或 {"SLMatltrans": {"items": [...]}}
    item_list = data.get("Items", data.get("SLMatltrans", {}))
    if isinstance(item_list, dict):
        item_list = item_list.get("items", item_list.get("Items", []))
    if not isinstance(item_list, list):
        return []

    return item_list


def _get_last_record_date(cursor, site_short: str) -> str | None:
    """查询 SLMatltrans 表指定站点的最大 RecordDate"""
    cursor.execute(
        "SELECT MAX(RecordDate) FROM dbo.SLMatltrans WHERE SiteRef = %s",
        (site_short,),
    )
    result = cursor.fetchone()[0]
    if result is None:
        return None
    if isinstance(result, str):
        return result[:19]
    return f"{result.year:04d}-{result.month:02d}-{result.day:02d} {result.hour:02d}:{result.minute:02d}:{result.second:02d}"


def _sync_site(site_ref: str, token: str,
               server, username, password, database, port):
    """同步单个站点的 SLMatltrans 增量数据"""
    print(f"  🔄 SLMatltrans 同步 Site {site_ref}...")
    start = time.time()

    try:
        # ── 1. 确定增量锚点 ──
        conn = _db_connect(server, username, password, database, port)
        cur = conn.cursor()
        last_record_date = _get_last_record_date(cur, site_ref)

        if last_record_date:
            date_from = last_record_date
            print(f"    增量锚点: RecordDate > {last_record_date}")
        else:
            date_from = (date.today() - timedelta(days=FULL_SYNC_DAYS)).isoformat()
            print(f"    首次同步，回补 {FULL_SYNC_DAYS} 天（从 {date_from}）")

        # ── 2. 调用 IDO API ──
        filter_str = f"RecordDate > '{date_from}'"
        rows = _fetch_ido_data(site_ref, token, filter_str)

        if not rows:
            print(f"    本次无新增数据")
            conn.close()
            return 0

        print(f"    拉取 {len(rows)} 条，开始 Upsert...")

        # ── 3. Upsert（MERGE by RowPointer）──
        count = 0
        for row in rows:
            cur.execute("""
                MERGE dbo.SLMatltrans AS t
                USING (SELECT
                    %s AS SiteRef,
                    %s AS TransNum,
                    %s AS TransDate,
                    %s AS TransType,
                    %s AS RefType,
                    %s AS Backflush,
                    %s AS Whse,
                    %s AS Loc,
                    %s AS Lot,
                    %s AS Wc,
                    %s AS RefNum,
                    %s AS RefLineSuf,
                    %s AS RefRelease,
                    %s AS Item,
                    %s AS ue_GDL_Manufacturer,
                    %s AS ue_GDL_ManufacturerItem,
                    %s AS ue_GDL_LotSupplierLot,
                    %s AS ue_GDL_Uf_Customer,
                    %s AS ue_GDL_Description,
                    %s AS Qty,
                    %s AS MatlCost,
                    %s AS LbrCost,
                    %s AS FovhdCost,
                    %s AS VovhdCost,
                    %s AS OutCost,
                    %s AS DocumentNumber,
                    %s AS UnitCost,
                    %s AS TotalPosted,
                    %s AS TransactionDescription,
                    %s AS RowPointer,
                    %s AS CreateDate,
                    %s AS RecordDate
                ) AS s
                ON  t.RowPointer = s.RowPointer

                WHEN MATCHED THEN
                    UPDATE SET
                        t.TransDate                = s.TransDate,
                        t.TransType                = s.TransType,
                        t.RefType                  = s.RefType,
                        t.Backflush                = s.Backflush,
                        t.Whse                     = s.Whse,
                        t.Loc                      = s.Loc,
                        t.Lot                      = s.Lot,
                        t.Wc                       = s.Wc,
                        t.RefNum                   = s.RefNum,
                        t.RefLineSuf               = s.RefLineSuf,
                        t.RefRelease               = s.RefRelease,
                        t.Item                     = s.Item,
                        t.ue_GDL_Manufacturer      = s.ue_GDL_Manufacturer,
                        t.ue_GDL_ManufacturerItem  = s.ue_GDL_ManufacturerItem,
                        t.ue_GDL_LotSupplierLot    = s.ue_GDL_LotSupplierLot,
                        t.ue_GDL_Uf_Customer       = s.ue_GDL_Uf_Customer,
                        t.ue_GDL_Description       = s.ue_GDL_Description,
                        t.Qty                      = s.Qty,
                        t.MatlCost                 = s.MatlCost,
                        t.LbrCost                  = s.LbrCost,
                        t.FovhdCost                = s.FovhdCost,
                        t.VovhdCost                = s.VovhdCost,
                        t.OutCost                  = s.OutCost,
                        t.DocumentNumber           = s.DocumentNumber,
                        t.UnitCost                 = s.UnitCost,
                        t.TotalPosted              = s.TotalPosted,
                        t.TransactionDescription   = s.TransactionDescription,
                        t.RowPointer               = s.RowPointer,
                        t.CreateDate               = s.CreateDate,
                        t.RecordDate               = s.RecordDate
                WHEN NOT MATCHED THEN
                    INSERT (
                        SiteRef, TransNum, TransDate, TransType, RefType, Backflush,
                        Whse, Loc, Lot, Wc, RefNum, RefLineSuf, RefRelease,
                        Item, ue_GDL_Manufacturer, ue_GDL_ManufacturerItem,
                        ue_GDL_LotSupplierLot, ue_GDL_Uf_Customer, ue_GDL_Description,
                        Qty, MatlCost, LbrCost, FovhdCost, VovhdCost, OutCost,
                        DocumentNumber, UnitCost, TotalPosted, TransactionDescription,
                        RowPointer, CreateDate, RecordDate
                    ) VALUES (
                        s.SiteRef, s.TransNum, s.TransDate, s.TransType, s.RefType, s.Backflush,
                        s.Whse, s.Loc, s.Lot, s.Wc, s.RefNum, s.RefLineSuf, s.RefRelease,
                        s.Item, s.ue_GDL_Manufacturer, s.ue_GDL_ManufacturerItem,
                        s.ue_GDL_LotSupplierLot, s.ue_GDL_Uf_Customer, s.ue_GDL_Description,
                        s.Qty, s.MatlCost, s.LbrCost, s.FovhdCost, s.VovhdCost, s.OutCost,
                        s.DocumentNumber, s.UnitCost, s.TotalPosted, s.TransactionDescription,
                        s.RowPointer, s.CreateDate, s.RecordDate
                    );
            """, (
                site_ref,
                row.get("TransNum"),
                row.get("TransDate"),
                row.get("TransType"),
                row.get("RefType"),
                row.get("Backflush"),
                row.get("Whse"),
                row.get("Loc"),
                row.get("Lot"),
                row.get("Wc"),
                row.get("RefNum"),
                row.get("RefLineSuf"),
                row.get("RefRelease"),
                row.get("Item"),
                row.get("ue_GDL_Manufacturer"),
                row.get("ue_GDL_ManufacturerItem"),
                row.get("ue_GDL_LotSupplierLot"),
                row.get("ue_GDL_Uf_Customer"),
                row.get("ue_GDL_Description"),
                _to_decimal(row.get("Qty")),
                _to_decimal(row.get("MatlCost")),
                _to_decimal(row.get("LbrCost")),
                _to_decimal(row.get("FovhdCost")),
                _to_decimal(row.get("VovhdCost")),
                _to_decimal(row.get("OutCost")),
                row.get("DocumentNum"),
                _to_decimal(row.get("DerMatltranCost")),
                _to_decimal(row.get("DerMatlTranViewTotalPosted")),
                row.get("MatlTranViewTypeDesc"),
                row.get("RowPointer"),
                row.get("CreateDate"),
                row.get("RecordDate"),
            ))
            count += 1

        conn.commit()
        conn.close()

        elapsed = round(time.time() - start, 2)
        print(f"    ✅ Site {site_ref} 完成，Upsert {count} 条，耗时 {elapsed}s")
        return count

    except Exception as e:
        print(f"    ❌ Site {site_ref} 同步失败: {e}")
        try:
            conn.close()
        except Exception:
            pass
        return 0


def sync_slmatltrans(server, username, password, database, port):
    """
    同步所有站点的 SLMatltrans 数据。
    复用已获取的 Infor Token（或重新获取）。

    返回: dict {site_ref: upsert_count}
    """
    print(f"\n  ── SLMatltrans 同步开始 ──")

    # 获取 Token（复用缓存）
    try:
        token = _load_infor_token()
    except RuntimeError as e:
        print(f"  ❌ Token 获取失败：{e}")
        return {}

    results = {}
    total = 0
    start = time.time()

    for site_ref in SITES:
        count = _sync_site(site_ref, token, server, username, password, database, port)
        results[site_ref] = count
        total += count

    elapsed = round(time.time() - start, 2)
    print(f"\n  SLMatltrans 同步完成 | 总计 {total} 条 | 耗时 {elapsed}s")

    return results


if __name__ == "__main__":
    server   = os.getenv("SQL_SERVER_HOST")
    port     = int(os.getenv("SQL_SERVER_PORT", "1433"))
    database = os.getenv("SQL_SERVER_DATABASE")
    username = os.getenv("SQL_SERVER_USERNAME")
    password = os.getenv("SQL_SERVER_PASSWORD")

    if not server:
        print("  ❌ 缺少数据库配置（SQL_SERVER_HOST），请检查 .env")
        sys.exit(1)

    sync_slmatltrans(server, username, password, database, port)
