# 商品自动上传机器人（Playwright）

把人工按网站模板备好的 xlsx，自动提交到站点批量上传表单（2026-08-15 起站点直接收 xlsx，不再转 CSV）。

## 项目结构
```
upload_products_cn/
├── upload.py              # 主程序：登录 → 上传 xlsx → 轮询任务中心 → 抓取成败明细
├── upload.bat             # 一键启动器（拖放/双击/命令行；含固定账号，仅本机）
├── header_template.txt    # 基准表头（33 列，表头严格校验用；--save-template-header 可刷新）
├── clean_old_runs.py      # 清理运行期产物（按天清旧的失败 CSV / 日志）
├── README.md
├── legacy/                # 旧工具归档（calibrate.py 及 calibrate/ 目录，2026-07-31）
├── failed/                # 运行产物：失败行 CSV（失败行_*.csv）；调试快照按需生成
├── session.json           # 登录态 cookies（含 token，机密，已被 .gitignore 忽略）
├── last_run.json          # 最近一次运行结果（权威记录）
└── upload.log             # 运行日志
```
> `session.json` / `last_run.json` / `upload.log` / `failed/` 均不进版本库（见 `.gitignore`）。

## 设计要点
- **数据源**：人工备好的 xlsx 模板，程序只做格式/编码转换，不解析业务字段。
- **机制**：网页表单 + Playwright；需登录，自动复用 cookies 会话、过期才重登。
- **编码**：站点直接收 xlsx（openpyxl 原样读/写）。早期曾转 UTF-8 CSV，2026-08-15 起站点改版不再接受 CSV，故改为直接提交原始 xlsx。
- **触发**：手动 CLI，预留调度接口。
- **安全阀**：基础预检（可读/非空/必需列齐全）+ `--dry-run` 试跑。
- **失败处理**：解析“导入完成”成功弹窗与“请检查导入数据”校验失败弹窗；失败行的**错误原因**（如「第 N 行：找不到1级商品分类」）会被抽取并写入 `failed/失败行_*.csv`（含原始行号 + 整行原数据）；运行记 `upload.log`；有失败非零退出；**不自动重试**。

## 安装
```bash
pip install playwright openpyxl
# 本机已装系统 Chrome，upload.py 用 channel="chrome" 复用，无需下载内核。
# 如需改用 Playwright 自带 chromium：playwright install chromium，再把 upload.py 里 launch 的 channel 改为 None。
```

## 运行
```bash
# 真实上传
python upload.py --file "商品导入模板_CN 0729--本 - 副本.xlsx"

# 只转换+预览，不提交（推荐先跑一次确认转换无误）
python upload.py --file 商品.xlsx --dry-run

# 首次/调试：手动登录一次（会话 cookies 自动保存，之后可全自动）
python upload.py --file 商品.xlsx --headless=false

# 保留临时转换的 CSV（排查用）
python upload.py --file 商品.xlsx --keep-csv

# 站点模板变更时，把正确模板的表头刷新为新的基准（不提交）
python upload.py --file 正确模板.xlsx --save-template-header
```

## 表头严格校验（默认开启，2026-08-01 新增）
上传前会**逐列**比对「基准表头」，列数 / 列名 / 顺序任一不符即**拒绝上传**（在打开浏览器之前就拦下，不会提交到站点）。

- 基准表头来自 `header_template.txt`（自正确模板抓取，33 列）。
- 校验未通过会打印精确差异，例如：
  ```
  表头校验未通过，已拒绝上传。差异如下：
    - 第 2 列不符：期望「货号」，实际「货号XXX」
    - 第 33 列不符：期望「供应商名称」，实际「多余列」
  ```
  同时把完整报告写到 `failed/header_check_<时间>.txt`（含实际表头逐列）。
- **建议先在 `--dry-run` 下跑一遍**确认表头无误，再正式提交。
- 站点模板改版后：用上面 `--save-template-header` 刷新基准即可，无需改代码。

## 一键启动器（upload.bat）
凭据通过环境变量 `SITE_USER` / `SITE_PASS` 提供（见下节），**无需敲路径**，三种用法：
1. **拖放**：把 xlsx 直接拖到 `upload.bat` 图标上 → 自动上传。
2. **双击**：直接双击 `upload.bat` → 弹出文件选择框挑模板 → 自动上传。
3. **命令行**（PowerShell 要加 `.\`）：
   ```bat
   .\upload.bat "D:\某个目录\商品导入模板_CN.xlsx"
   .\upload.bat "路径" dry          :: 试跑（只转换不提交）
   ```
- 不传参数 → 自动弹文件选择器；第二个参数写 `dry` 即走 `--dry-run`。
- 启动器内部用隔离 venv 的 python 跑 `upload.py`，与手动命令等价。结尾有 `pause`，窗口停留便于看结果。
- **便携性**：`upload.bat` 优先用「自身旁边的 `upload.py`」，找不到才回退到 `<clone 目录>\upload_products_cn\upload.py`。本文件夹是完整自包含工作副本，直接在此运行即可。
- 凭据只从环境变量读取，**不要把账号密码写进 `upload.bat` 或任何会进版本库的文件**。

## 登录凭据（推荐环境变量，勿写死代码）
```bash
set SITE_USER=你的账号
set SITE_PASS=你的密码
```
- 有凭据 → 自动登录（已校准：`#login_mobilePhone` / `#login_password` / `form button[type=submit]`）。
- 无凭据且 `--headless=false` → 浏览器打开后手动登录，登录成功后脚本继续并保存会话。
- 之后 headless 默认 True，复用已保存 cookies（`session.json`），平时全自动。

## 选择器（已用 calibrate.py 实测校准，2026-07-31）
站点：`https://admin-cn.example.com/goods/import-products`（React + antd refine SPA）

| 用途 | 选择器 | 说明 |
|------|--------|------|
| 文件选择框 | `input[type=file]` | antd Upload dragger 的隐藏 input（`display:none`），`set_input_files` 可作用 |
| 提交按钮 | `button.ant-btn-primary:has-text('批量导入')` | ⚠️ 卡片标题也是“批量导入商品”，必须限定为 `button` 元素，否则会误点标题 |
| 登录用户框 | `#login_mobilePhone` | |
| 登录密码框 | `#login_password` | |
| 登录提交 | `form button[type=submit]` | 按钮文字为“登 录”（antd 带空格），用表单提交按钮最稳 |
| 文件列表项 | `.ant-upload-list-item` | 选完文件后出现的上传列表项，用于确认文件已载入 |

### 关键时序（站点明文提示）
页面提示：**“选择一个文件，等待出现 totalRow 数量提示之后再操作。”**
→ `upload.py` 在 `set_input_files` 后：先等 `.ant-upload-list-item` 出现（文件已载入），
再 `wait_for_timeout(5000)` 留出前端解析 CSV（totalRow 数量提示）的时间，然后才点击“批量导入”。

## 结果解析（任务中心异步 + 接口 errorDetail，2026-08-26 重写）
提交后 `upload.py` 会：
1. 提交回执 `{"code":200,"data":<task_id>}` 中取 `task_id`；
2. 进「任务中心」轮询该 `task_id` 的列表接口 `status` 字段（3=已完成成功 / 1=失败 / 0,2=进行中），秒级判定终态；
3. 终态为失败时，定位该任务行（`data-row-key`==task_id，精确，避免同名多次上传撞错行），点「失败明细」；
4. **优先拦截任务详情接口响应**（`POST /api` 的 `data.errorDetail` JSON 数组：`[{barcode,code,detail:"第N行:原因"},...]`）——它比偶发为「暂无数据」的弹窗可靠，弹窗仅作兜底；
5. 失败原因结构化写入 `failed/失败行_*.csv`：**原始行号 + 货号 + 条码 + 错误信息 + 整行原数据**（无「第N行」前缀的错误用货号/条码反查行号）；
6. 调试快照（`failed/debug_task_*.html` / `task_*.png`）仍按需生成，便于人工核对。

## 校准脚本（已归档）
`calibrate.py` 及 `calibrate/` 目录已移至 `legacy/`（2026-07-31 旧工具，只读、不提交）。如需重校准选择器：
```bash
SITE_USER=xxx SITE_PASS=yyy python legacy/calibrate.py
```

## 输出
- `upload.log`：每次运行一行摘要（时间/文件/成功失败数/异常）。
- `last_run.json`：**权威结果记录**（时间/成功数/失败数/是否完成/退出码/调试路径/summary）。由于本机 Chrome 关闭时 shell 可能被掐断、进程退出码不可靠，判定成功与否以 `last_run.json` 的 `ok` 字段为准（而非进程退出码）。
- `failed/失败行_*.csv`：失败行清单，**含原始行号、货号、条码、错误信息、整行原数据**，便于直接定位修改（仅当有失败时）。
- `failed/debug_task_*.html` / `failed/task_*.png`：最近一次的结果页 HTML / 截图（调试用，自动被 `.gitignore` 忽略）。
- `session.json`：保存的登录会话 cookies（含 token，机密，勿外传、勿进版本库）。

## 目录治理
- `clean_old_runs.py` 按天清理运行产物（默认清 7 天前的失败 CSV、旧的 `upload.log`）：
  ```bash
  python clean_old_runs.py              # 清 7 天前
  python clean_old_runs.py --days 30    # 清 30 天前
  python clean_old_runs.py --dry        # 只预览不删
  ```
- `session.json` / `last_run.json` / `upload.log` / `failed/` 已在 `.gitignore` 忽略；`upload.bat` 含明文账号，勿外传。

## 结果弹窗格式（已实测，2026-08-26 更新）
导入改为**异步任务**：提交后到「任务中心」查看，任务「已完成」后点「成功明细」/「失败明细」按钮看明细。弹窗摘要格式：
```
导入成功：共11条，成功11条（新增0条，修改11条），失败0条
```
- `upload.py` 现按此摘要行精确计数（共/成功/新增/修改/失败），并写入 `last_run.json` 的 `summary` 字段。
- `成功明细` 弹窗会显示该摘要；`失败明细` 弹窗除摘要外，若含「第 N 行：xxx」逐行错误也会抽取写入 `failed/失败行_*.csv`。
- 解析两种结果弹窗（旧版「导入完成 / 总行数 / 更新商品 / 新增商品」已不再出现，旧解析逻辑保留为兼容但不再命中）：
  - 成功：`导入成功 / 共 N 条 / 成功 N 条（新增 X / 修改 Y）/ 失败 Z 条`。
  - 失败：`失败明细` 按钮 + 逐行错误。
