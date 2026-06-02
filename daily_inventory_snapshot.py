#!/usr/bin/env python3
"""
每日库存快照脚本
================
从 Infor CSI API 抓取三站点的：
  1. SLItemCostingReport → SLTotalInventory 表（库存明细）
  2. SLTotalWIPValuebyAccountReport → SLTotalWIPValueByAcountReport 表（WIP 按科目汇总）
  3. SLMatltrans → SLMatltrans 表（物料交易记录，增量同步）

直接写入数据库。不生成 Excel，不发送邮件。
同一 BalanceDate + SiteRef 多次运行只保留最新数据（先删后插）。

运行方式：
  python daily_inventory_snapshot.py

依赖：inventory_tracking.py 中的 _load_infor_token, _fetch_infor_site,
      _fetch_wip_site, save_inventory_to_db, save_wip_to_db
      slmatltrans.py 中的 sync_slmatltrans
"""

import os
import sys
import argparse
from datetime import datetime, timedelta
from dotenv import load_dotenv

sys.stdout.reconfigure(encoding="utf-8")
load_dotenv()

# 复用 inventory_tracking.py 中的函数
from inventory_tracking import (
    _load_infor_token,
    _fetch_infor_site,
    _fetch_wip_site,
    save_inventory_to_db,
    save_wip_to_db,
    SITE_CURRENCY,
)
from slmatltrans import sync_slmatltrans

SITES = ["310", "330", "410"]


def main():
    parser = argparse.ArgumentParser(description="每日库存快照（API 入库）")
    parser.add_argument(
        "--balance-date",
        default="",
        help="手动指定 BalanceDate（YYYY-MM-DD），优先级最高",
    )
    parser.add_argument(
        "--date-offset-days",
        type=int,
        default=0,
        help="在自动计算的业务日期基础上再回退天数，默认 0",
    )
    parser.add_argument(
        "--midnight-grace-minutes",
        type=int,
        default=5,
        help="午夜后宽限分钟数（0点后这段时间内仍记前一天），默认 5",
    )
    args = parser.parse_args()

    run_time = datetime.now()
    # 业务日期计算：
    # 1) 指定 --balance-date 时直接使用
    # 2) 否则在午夜宽限期内记前一天（解决 23:59:59 任务跨到 00:00:00 启动的问题）
    if args.balance_date:
        try:
            business_dt = datetime.strptime(args.balance_date, "%Y-%m-%d")
        except ValueError:
            print("  ❌ --balance-date 格式错误，请使用 YYYY-MM-DD")
            sys.exit(1)
        strategy = "manual"
    else:
        if run_time.hour == 0 and run_time.minute < max(args.midnight_grace_minutes, 0):
            business_dt = run_time - timedelta(days=1)
            strategy = "midnight-grace"
        else:
            business_dt = run_time
            strategy = "runtime-date"

    balance_date = (business_dt - timedelta(days=args.date_offset_days)).strftime("%Y-%m-%d")
    print(f"{'='*60}")
    print(f"  每日库存快照 | BalanceDate = {balance_date}")
    print(f"  RunTime      = {run_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  DateStrategy = {strategy}")
    print(f"  OffsetDays   = {args.date_offset_days}")
    print(f"  GraceMinutes = {args.midnight_grace_minutes}")
    print(f"{'='*60}")

    # ── 1. 获取 Infor API Token ──
    try:
        token = _load_infor_token()
    except RuntimeError as e:
        print(f"  ❌ Token 获取失败：{e}")
        sys.exit(1)

    # ── 2. DB 连接参数 ──
    server   = os.getenv("SQL_SERVER_HOST")
    port     = int(os.getenv("SQL_SERVER_PORT", "1433"))
    database = os.getenv("SQL_SERVER_DATABASE")
    username = os.getenv("SQL_SERVER_USERNAME")
    password = os.getenv("SQL_SERVER_PASSWORD")

    if not server:
        print("  ❌ 缺少数据库配置（SQL_SERVER_HOST），请检查 .env")
        sys.exit(1)

    # ── 3. 逐站点抓取库存数据并入库 ──
    total_inserted = 0
    failed_inventory = []
    wip_results = {}

    for site_ref in SITES:
        print(f"\n  ── Site {site_ref} ──")

        # 3a. 库存明细（SLItemCostingReport）
        try:
            df = _fetch_infor_site(site_ref, token)
        except RuntimeError as e:
            print(f"  ❌ 库存 API 调用失败：{e}")
            failed_inventory.append(site_ref)
        else:
            if df.empty:
                print(f"  ⚠️  库存返回空数据")
                failed_inventory.append(site_ref)
            else:
                print(f"  📡 库存 API 返回 {len(df)} 行")
                inserted = save_inventory_to_db(
                    df, site_ref, server, username, password, database, port,
                    balance_date=balance_date,
                )
                total_inserted += inserted

        # 3b. WIP 按科目汇总（SLTotalWIPValuebyAccountReport → SLTotalWIPValueByAcountReport）
        try:
            raw_wip, wip_records = _fetch_wip_site(site_ref, token)
            save_wip_to_db(
                site_ref, raw_wip, wip_records, server, username, password, database, port,
                balance_date=balance_date,
            )
            currency = "CNY" if site_ref == "330" else "USD"
            _, fx = SITE_CURRENCY.get(site_ref, ("USD", 1.0))
            wip_results[site_ref] = (raw_wip, currency, round(raw_wip * fx, 2))
        except RuntimeError as e:
            print(f"  ❌ WIP API 调用失败：{e}")
            wip_results[site_ref] = (0.0, "N/A", 0.0)

    # ── 4. 汇总 ──
    print(f"\n{'='*60}")
    print(f"  库存 | BalanceDate={balance_date} | 插入 {total_inserted} 行")
    if failed_inventory:
        print(f"  库存失败站点：{failed_inventory}")

    print(f"\n  WIP 汇总：")
    grand_wip = 0.0
    for s in SITES:
        raw, curr, usd = wip_results.get(s, (0, "N/A", 0))
        print(f"    Site {s}: {raw:,.2f} {curr}  →  USD {usd:,.2f}")
        grand_wip += usd
    print(f"    Grand Total (USD): ${grand_wip:,.2f}")

    # ── 5. 同步 SLMatltrans（物料交易记录）──
    print(f"\n  ── SLMatltrans 同步 ──")
    try:
        matltrans_results = sync_slmatltrans(server, username, password, database, port)
        if matltrans_results:
            for s in SITES:
                cnt = matltrans_results.get(s, 0)
                print(f"    Site {s}: {cnt} 条 Upsert")
    except Exception as e:
        print(f"  ❌ SLMatltrans 同步异常: {e}")

    print(f"\n{'='*60}")


if __name__ == "__main__":
    main()
