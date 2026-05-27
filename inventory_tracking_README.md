# 库存金额跟踪系统 (Inventory Tracking System)

## 功能概述

自动从数据库查询当月 MTD 物料事务，按分类汇总，生成带期初余额 + 发生额 + 期末余额的 Excel 报表。

## 计算公式

```
期末金额 Balance = 期初金额(Previous Balance)
                 + Received AMT
                 - Consumed AMT
                 - Other Transaction AMT

期末数量 = 期初数量 + Received Qty - Consumed Qty - Other Transaction Qty
```

## 分类规则 (TransType + RefType)

| Category | TransType | RefType | Description |
|----------|-----------|---------|-------------|
| **Received** | R | P | PO Receipt |
| **Received** | W | P | PO Withdraw |
| **Consumed** | I | J | Job Issue / WIP Change |
| **Consumed** | W | J | Job Withdrawal / Return |
| **Consumed** | S | O | Order Ship |
| **Other Transaction** | A | I | Adjustment |
| **Other Transaction** | M | I | Stock Move |
| **Other Transaction** | G | I | Misc Receipt |
| **Other Transaction** | H | I | Misc Issue |
| **Other Transaction** | C | J | Job Complete |
| **Other Transaction** | F | J | Job Finish |
| **Other Transaction** | N | J | Job Labor / Next Operation |
| **Other Transaction** | W | R | RMA Withdraw |

## 依赖安装

```bash
pip install pandas pyodbc openpyxl
```

> **注意**：连接 SQL Server 需要安装 [ODBC Driver for SQL Server](https://docs.microsoft.com/zh-cn/sql/connect/odbc/download-odbc-driver-for-sql-server)

## 使用方法

### 1. 准备期初余额 Excel

创建文件 `prev_balance.xlsx`，格式如下：

| Item | Prev_Qty | Prev_Balance |
|------|----------|--------------|
| ITM-001 | 100 | 5000.00 |
| ITM-002 | 50 | 2500.00 |

### 2. 运行脚本

```bash
# 基础用法
python inventory_tracking.py -f prev_balance.xlsx

# 指定 SQL Server
python inventory_tracking.py -s 10.0.6.134 -f prev_balance.xlsx

# 指定站点（如 330）
python inventory_tracking.py --site 330 -f prev_balance.xlsx

# 完整参数
python inventory_tracking.py \
    -s 10.0.6.134 \
    -d csi_datawarehouse \
    --site 310 \
    -f prev_balance.xlsx \
    -o May_Inventory_Report.xlsx
```

### 3. 输出文件

默认生成 `Inventory_Balance_YYYYMM.xlsx`，包含两个 Sheet：

- **Summary**：按 Item 汇总的期初/发生/期末数据
- **Detail**：每条物料事务的明细记录

## 报表字段说明

### Summary Sheet

| 字段 | 说明 |
|------|------|
| Item | 料号 |
| 期初数量 | 上月末库存数量 |
| 期初金额 | 上月末库存金额 |
| Received Qty | 本月入库数量 |
| Received AMT | 本月入库金额（成本 = MatlCost+LbrCost+FovhdCost+VovhdCost+OutCost） |
| Consumed Qty | 本月消耗数量 |
| Consumed AMT | 本月消耗金额 |
| Other Qty | 其他事务数量 |
| Other AMT | 其他事务金额 |
| 期末数量 | 期初数量 + Received Qty - Consumed Qty - Other Qty |
| 期末金额 (Balance) | 期初金额 + Received AMT - Consumed AMT - Other AMT |

## 文件清单

| 文件 | 说明 |
|------|------|
| `inventory_tracking.py` | 主脚本 |
| `inventory_tracking_README.md` | 使用说明 |
| `prev_balance_sample.xlsx` | 期初余额 Excel 模板 |

## 修改日志

| 日期 | 版本 | 修改内容 |
|------|------|----------|
| 2026-05-25 | v1.0 | 初始版本，支持按 Item 汇总 |
