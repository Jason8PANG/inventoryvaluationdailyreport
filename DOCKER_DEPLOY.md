# Inventory Valuation Tracking — Docker 部署说明

## 目录结构

```
naiitemvaluationtracking/
├── .env                     # ⭐ 主配置文件（DB、SMTP、收件人）
├── inventory_tracking.py    # 主脚本（pymssql + smtplib）
├── config.ini               # 后备配置文件（.env 未设置时生效）
├── requirements.txt         # Python 依赖
├── Dockerfile               # Docker 镜像构建
├── docker-compose.yml       # 容器编排
├── Previous Balance/        # 期初余额 Excel（挂载进容器）
│   ├── site 310.xlsx
│   ├── site 330.xlsx
│   └── site 410.xlsx
└── reports/                 # 报表输出目录（自动创建）
```

---

## 快速部署步骤

### 1. 上传文件到 CentOS 服务器

```bash
scp -r ./naiitemvaluationtracking user@centos-server:/opt/nai-inventory
```

### 2. 配置 .env 文件

所有敏感配置集中在 `.env`，按需修改：

```bash
# 数据库连接
SQL_SERVER_HOST=SUZVPRINT01\CUSTOMSSYS
SQL_SERVER_PORT=1433
SQL_SERVER_DATABASE=csi_datawarehouse
SQL_SERVER_USERNAME=datasync
SQL_SERVER_PASSWORD=R3p0rts

# 邮件收件人（多个用 ; 分隔）
MAIL_TO=jason.pang@nai-group.com;shirley.ni@nai-group.com;...
MAIL_CC=sky.li@nai-group.com;frank.liu@nai-group.com;...

# SMTP 服务器
SMTP_HOST=mail.smtp2go.com
SMTP_PORT=2525
SMTP_USER=Suzhou
SMTP_PASSWORD=xxx

# 发件人
MAIL_FROM=suzinventoryvaluationdailyreport@nai-group.com
```

> **优先级**: `.env` → `config.ini` → 脚本硬编码默认值。无需改代码。

### 3. 构建并启动（守护模式）

```bash
cd /opt/nai-inventory

# 创建外部网络（仅首次）
docker network create public-net 2>/dev/null

# 构建镜像
docker compose build

# 启动守护进程（每天 09:00 自动执行）
docker compose up -d

# 查看日志确认启动正常
docker compose logs -f
```

启动后**不会立即执行**，等待到下一个 09:00 才运行。

```bash
# 常用运维命令
docker compose logs --tail 50    # 最近 50 行日志
docker compose restart           # 重启
docker compose down              # 停止（调度随之停止）
docker compose up -d             # 重新启动
```

### 4. 定时执行机制

调度**内置于 Docker 容器**中，无需宿主机 cron：

| 行为 | 说明 |
|---|---|
| 启动时 | 不执行，等待到 09:00 |
| 每日 09:00 | 自动运行全站点库存跟踪 + 邮件发送 |
| `docker compose down` | 容器停止 → 调度自动停止 |
| `docker compose up -d` | 重新启动，等到下一个 09:00 |

> 手动调试：`docker compose exec app python inventory_tracking.py --all-sites`
> 如果某天执行失败，5 分钟后自动重试一次。

---

## 配置说明

### config.ini 完整说明

| 节 | 键 | 说明 |
|---|---|---|
| `[database]` | `server` | SQL Server 主机名或 IP，支持 `HOST\INSTANCE` 格式（后备） |
| | `database` | 数据库名（后备） |
| | `username` / `password` | 数据库账户（后备） |
| | `prev_dir` | 期初文件目录（容器内路径） |
| `[email]` | `to` | 收件人（后备，优先 .env MAIL_TO） |
| | `cc` | 抄送（后备，优先 .env MAIL_CC） |
| | `from_addr` | 发件人地址（后备，优先 .env MAIL_FROM） |
| `[smtp]` | `host` | SMTP 服务器（后备，优先 .env SMTP_HOST） |
| | `port` | 端口（后备，优先 .env SMTP_PORT） |
| | `user` / `password` | SMTP 认证（后备，优先 .env SMTP_USER/PASSWORD） |
| | `tls` | 是否 STARTTLS（后备，优先 .env SMTP_TLS） |

### 命令行参数（可覆盖 config.ini）

```bash
# 自定义收件人
python inventory_tracking.py --all-sites \
  --email-to "a@nai-group.com;b@nai-group.com" \
  --email-cc "c@nai-group.com"

# 自定义 SMTP
python inventory_tracking.py --all-sites \
  --smtp-host 10.0.1.100 --smtp-port 25

# 不发邮件（只生成 Excel）
python inventory_tracking.py --all-sites --no-email

# 守护模式（每天 09:00 自动执行）
python inventory_tracking.py --all-sites --daemon
```

---

## 与 Windows 版本的差异

| 项目 | Windows 版 | Docker/Linux 版 |
|---|---|---|
| 数据库驱动 | `pyodbc` + ODBC Driver 17 | `pymssql`（纯 Python，无需系统驱动） |
| 邮件发送 | `win32com` Outlook COM | `smtplib` SMTP（标准库） |
| 收件人配置 | 硬编码在脚本 | `config.ini` 外部化 |
| 定时任务 | WorkBuddy Automation | Docker 内置 sleep loop（守护模式） |
| 期初文件 | 本地目录 | Docker volume 挂载 |
