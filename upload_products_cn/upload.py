#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
商品自动上传 - Playwright 上传机器人
====================================
把人工按网站模板备好的 xlsx，自动转换为 UTF-8 CSV 并提交到站点批量上传表单。

设计来源：grill-me 设计访谈（2026-07-31）
  - 源：人工备好的 xlsx 模板（不解析业务字段，只做格式/编码转换）
  - 机制：网页表单 + Playwright；需登录，自动复用 cookies 会话、过期才重登
  - 编码：纯 UTF-8（无 BOM）——Mac 可传已佐证站点接受无 BOM
  - 触发：手动 CLI，预留调度接口
  - 安全阀：基础预检 + --dry-run 试跑
  - 失败：解析结果；失败行写 csv + 原因；upload.log；有失败非零退出；不自动重试

依赖：
  pip install playwright openpyxl
  （本环境用系统已装 Chrome，channel="chrome"；如需自带 chromium：
   playwright install chromium，再把 launch 的 channel 改为 None）
登录凭据（推荐用环境变量，勿写死）：
  SITE_USER=xxx   SITE_PASS=yyy

用法：
  python upload.py --file 商品.xlsx
  python upload.py --file 商品.xlsx --dry-run          # 只转换+预览，不提交
  python upload.py --file 商品.xlsx --headless=false   # 手动登录/调试
  python upload.py --file 商品.xlsx --keep-csv         # 保留临时转换的 CSV
  python upload.py --file 商品.xlsx --save-template-header  # 刷新基准表头(站点模板变更时)

表头校验（严格模式，默认开启）：上传前会逐列比对“基准表头”(header_template.txt)，
列数/列名/顺序任一不符即拒绝上传并打印差异，不会提交到站点。基准表头可用
--save-template-header 从正确模板刷新。
"""

import argparse
import csv
import datetime as dt
import json
import os
import re
import shutil
import sys
import time
import traceback
from pathlib import Path

# 多站点配置（站点差异集中在此模块，upload.py 不写死站点逻辑）
try:
    import station_config as SC
except Exception:
    SC = None

try:
    import openpyxl
except ImportError:
    sys.exit("缺少依赖 openpyxl，请先: pip install openpyxl")

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
except ImportError:
    sys.exit("缺少依赖 playwright，请先: pip install playwright")

import socket
import urllib.parse


def _proxy_port_open(host: str, port: int, timeout: float = 2.0) -> bool:
    """探测本地代理端口是否可用。"""
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except Exception:
        return False


def build_browser_args():
    """决定 Chromium 代理参数（自动探测，避免硬编码直连/代理任一失败）：
    - USE_PROXY=on/1   → 用系统代理（依赖本地代理客户端运行）
    - USE_PROXY=off/0  → 强制直连 --no-proxy-server
    - USE_PROXY=auto(默认) → 探测到本地代理端口开放则用系统代理，否则直连
    """
    mode = os.environ.get("USE_PROXY", "auto").strip().lower()
    if mode in ("off", "0"):
        return ["--no-proxy-server"]
    if mode in ("on", "1"):
        return []  # 使用系统/环境代理
    hp = os.environ.get("HTTPS_PROXY") or os.environ.get("HTTP_PROXY") or ""
    port = 7897
    if hp:
        try:
            p = urllib.parse.urlparse(hp).port
            if p:
                port = p
        except Exception:
            pass
    if _proxy_port_open("127.0.0.1", port):
        log(f"[proxy] 检测到本地代理端口 {port} 开放，使用系统代理")
        return []
    log(f"[proxy] 本地代理端口 {port} 未开放，使用直连")
    return ["--no-proxy-server"]


# ---------------------------------------------------------------------------
# 多站点辅助
# ---------------------------------------------------------------------------
def task_center_url_of(upload_url: str) -> str:
    """由上传地址推导任务中心地址（替换路径段为 /task-center）。

    https://admin-cn.example.com/goods/import-products
      -> https://admin-cn.example.com/task-center
    CN 下推导结果与原 TASK_CENTER_URL 完全一致，故不影响现有 CN 行为。
    """
    try:
        scheme, rest = upload_url.split("//", 1)
        host = rest.split("/", 1)[0]
        return f"{scheme}//{host}/task-center"
    except Exception:
        return TASK_CENTER_URL


def backup_after_upload(xlsx_path: Path, station: str) -> str | None:
    """上传成功后把提交文件备份到 已导入表格/<站>/（覆盖同名，幂等安全）。

    返回备份路径或 None（未配置站点 / 目录不可建）。失败仅告警不阻断主流程。
    """
    if SC is None:
        return None
    if station not in SC.BACKUP_DIR:
        return None
    try:
        bdir = SC.BACKUP_DIR[station]
        os.makedirs(bdir, exist_ok=True)
        dst = os.path.join(bdir, xlsx_path.name)
        shutil.copy2(str(xlsx_path), dst)
        log(f"已备份上传表到: {dst}")
        return dst
    except Exception as e:
        log(f"[警告]  上传后备份失败（不阻断）: {e}")
        return None


# ---------------------------------------------------------------------------
# 配置区（按实际站点调整，已用 calibrate.py 实测校准）
# ---------------------------------------------------------------------------
UPLOAD_URL = "https://admin-cn.example.com/goods/import-products"
# 任务中心（2026-08-15 系统更新：导入改为异步任务，结果需到任务中心查看）
TASK_CENTER_URL = "https://admin-cn.example.com/task-center"

# 2026-08-15 系统更新：后台导入接口不再接受 CSV，只收 xlsx（与 GR 一致）。
# 故 CN 改为直接提交原始 xlsx，不再做 xlsx→CSV 转换。
SUBMIT_XLSX = True

# 会话 / 日志 / 失败行落盘位置（与脚本同目录）
BASE_DIR = Path(__file__).resolve().parent
SESSION_FILE = BASE_DIR / "session.json"          # 保存的 cookies
LOG_FILE = BASE_DIR / "upload.log"                # 运行日志
FAILED_DIR = BASE_DIR / "failed"                  # 失败行 csv / 调试 HTML 输出目录

# 编码：纯 UTF-8 无 BOM（站点接受，Mac 已佐证；Windows Excel 默认 GBK 会乱码）
ENCODING = "utf-8"

# 预检：xlsx 必须包含的“核心标识列”（按模板表头；缺失则报错不提交）
REQUIRED_HEADERS = ["品名", "货号", "产品分类"]

# 表头严格校验（严格模式，2026-08-01 新增）：与正确模板逐列（列数+列名+顺序）
# 完全一致才放行。基准表头来自 header_template.txt（可由 --save-template-header 刷新），
# 文件缺失时退回内置 CANONICAL_HEADER。
HEADER_TEMPLATE_FILE = BASE_DIR / "header_template.txt"
CANONICAL_HEADER = [
    "品名", "货号", "条形码", "产品分类", "二级分类", "三级分类", "四级分类",
    "五级分类", "六级分类", "批发价格", "采购单价", "单箱重量(kg)", "单箱体积(m³)",
    "单箱长度(cm)", "单箱宽度(cm)", "单箱高度(cm)", "中包方式", "中包每包可装个数",
    "单箱每箱装箱个数", "每托可装个数", "每集装箱可装个数", "销量", "最低起订量",
    "币种", "是否可以打折(Y/N)", "品牌", "标签", "产品详情", "产品属性", "商品认证",
    "原产地", "供应商编码", "供应商名称",
]

# ---- 选择器（已 calibrate.py 实测校准，2026-07-31）----
SEL_FILE_INPUT = "input[type=file]"              # 文件选择框（antd dragger 隐藏 input）
# 上传页“提交”按钮：class=ant-btn ant-btn-primary，文字“批量导入”
# 注意：卡片标题也是“批量导入商品”，故必须限定为 button 元素
SEL_SUBMIT_BTN = "button.ant-btn-primary:has-text('批量导入')"
SEL_SUBMIT_FALLBACK = "button:has-text('批量导入')"
# 登录页（实测可用）
SEL_LOGIN_USER = "#login_mobilePhone"
SEL_LOGIN_PASS = "#login_password"
SEL_LOGIN_SUBMIT = "form button[type=submit]"
PASSWORD_INPUT = "input[type=password]"
# 文件载入后的上传列表项（antd Upload 渲染）
SEL_UPLOAD_ITEM = ".ant-upload-list-item"

# 结果等待：提交后页面出现以下任一文字即视为结果已出
# 注意：包含“请检查导入数据”这种前置校验失败弹窗
RESULT_WAIT_TEXT_RE = r"成功|失败|导入结果|上传结果|处理完成|完成|请检查导入数据|错误信息"

# 结果解析正则（站点结果弹窗实测格式，如 “导入完成 / 总行数: 2 / 更新商品: 2”）
RE_DONE = re.compile(r"导入完成", re.I)                 # 结果弹窗标题
RE_ERROR_PROMPT = re.compile(r"请检查导入数据", re.I)   # 前置校验失败弹窗标题
RE_TOTAL = re.compile(r"总行数[:：]\s*(\d+)", re.I)
RE_UPDATED = re.compile(r"更新商品[:：]\s*(\d+)", re.I)
RE_ADDED = re.compile(r"新增商品[:：]\s*(\d+)", re.I)
RE_FAIL_COUNT = re.compile(r"失败(?:商品|行|记录)?[:：]\s*(\d+)", re.I)
# 错误行文字（两种实测格式）：
#   旧（前置校验）：第 2 行：找不到1级商品分类「家具卫浴」
#   新（任务中心失败明细，截图 导入.xlsx）：
#     执行异常：…请修正后重新导入：第 41 行，第 18 列 [中包含可包含个数]：值 [无] 格式错误，无法转换为数字…
# 行号取「第 N 行」；group(2) 为错误描述（含「第 M 列 [字段]：值 [X]…」，便于定位）。
# 非贪婪 + 前瞻到下一个「第 N 行」或文末：支持同一弹窗多条错误的逐条拆分。
RE_ERROR_ROW = re.compile(
    r"第\s*(\d+)\s*行"                      # 行号
    r"(?:\s*，?\s*第\s*\d+\s*列)?"          # 可选「，第 M 列」
    r"\s*[:：]?\s*(.*?)"                    # 非贪婪：错误描述
    r"(?=第\s*\d+\s*行|下载明细|关闭|$)", # 边界：下一错误行 / 弹窗按钮 / 文末
    re.I | re.S,
)
# 兜底：兼容 “成功 N 条 / 失败 M 条” 写法
RE_SUCCESS = re.compile(r"成功\s*(\d+)\s*条?", re.I)
RE_FAIL = re.compile(r"失败\s*(\d+)\s*条?", re.I)
# 结果弹出层（antd Modal / message）
SEL_RESULT_MODAL = ".ant-modal, .ant-message"
# 失败明细表常见容器（antd Modal 中的 table）
SEL_ERROR_TABLE = ".ant-modal-confirm-content table, .ant-modal-content table"

# 任务中心「成功明细」弹窗摘要行格式（2026-08-26 实测）：
#   导入成功：共11条，成功11条（新增0条，修改11条），失败0条
# 失败明细弹窗同样含同类摘要。逐条正则更稳，避免顺序/括号干扰。
RE_SUM_TOTAL = re.compile(r"共\s*(\d+)\s*条")
RE_SUM_SUCC = re.compile(r"成功\s*(\d+)\s*条")
RE_SUM_ADD = re.compile(r"新增\s*(\d+)\s*条")
RE_SUM_MOD = re.compile(r"修改\s*(\d+)\s*条")
RE_SUM_FAIL = re.compile(r"失败\s*(\d+)\s*条")


def parse_task_summary(text):
    """从任务中心明细弹窗文本抽取摘要计数。

    返回 (共, 成功, 新增, 修改, 失败) 五元组，或 None（无摘要行）。
    弹窗示例：导入成功：共11条，成功11条（新增0条，修改11条），失败0条
    """
    if not text:
        return None
    m_total = RE_SUM_TOTAL.search(text)
    m_succ = RE_SUM_SUCC.search(text)
    m_fail = RE_SUM_FAIL.search(text)
    if not (m_total and m_succ and m_fail):
        return None  # 不是标准的「共/成功/失败」摘要
    m_add = RE_SUM_ADD.search(text)
    m_mod = RE_SUM_MOD.search(text)
    return (
        int(m_total.group(1)),
        int(m_succ.group(1)),
        int(m_add.group(1)) if m_add else 0,
        int(m_mod.group(1)) if m_mod else 0,
        int(m_fail.group(1)),
    )


# ---------------------------------------------------------------------------
# 预检 + 转换
# ---------------------------------------------------------------------------
def load_canonical_header():
    """读取基准表头：优先 header_template.txt，缺失退回内置 CANONICAL_HEADER。"""
    if HEADER_TEMPLATE_FILE.exists():
        try:
            lines = [ln.strip() for ln in HEADER_TEMPLATE_FILE.read_text(encoding="utf-8").splitlines()]
            lines = [ln for ln in lines if ln]
            if lines:
                return lines
        except Exception:
            pass
    return list(CANONICAL_HEADER)


def validate_header(header, canonical=None):
    """严格表头校验：列数 + 列名 + 顺序 必须完全一致。
    返回 (ok, diffs:list[str])；ok=False 时 diffs 给出每处差异。"""
    if canonical is None:
        canonical = load_canonical_header()
    diffs = []
    if len(header) != len(canonical):
        diffs.append(f"列数不符：期望 {len(canonical)} 列，实际 {len(header)} 列")
    n = min(len(header), len(canonical))
    for i in range(n):
        if header[i] != canonical[i]:
            diffs.append(f"第 {i+1} 列不符：期望「{canonical[i]}」，实际「{header[i]}」")
    if len(header) > len(canonical):
        for i in range(n, len(header)):
            diffs.append(f"第 {i+1} 列多余：实际「{header[i]}」（基准模板无此列）")
    if len(header) < len(canonical):
        for i in range(n, len(canonical)):
            diffs.append(f"缺少第 {i+1} 列：期望「{canonical[i]}」")
    return (len(diffs) == 0), diffs


def preflight(xlsx_path: Path, canonical=None):
    """基础预检：可读 / 非空 / 核心列齐全 / 表头与基准模板严格一致。不通过抛异常。

    canonical: 指定基准表头（多站点时用 station_config.TEMPLATE_HEADER[station]）。
    为 None 时退回 CN 惯例（load_canonical_header：header_template.txt 或内置 CANONICAL_HEADER）。
    """
    if not xlsx_path.exists():
        raise FileNotFoundError(f"文件不存在: {xlsx_path}")
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = ws.iter_rows(values_only=True)
    try:
        header = next(rows)
    except StopIteration:
        raise ValueError("文件为空（无表头行）")
    header = [str(h).strip() for h in header if h is not None]
    if not header:
        raise ValueError("表头行为空")
    # 仅统计“非全空”的数据行：xlsx 常有大量空尾行，若按物理行数计会严重虚高，
    # 并导致结果里“成功行数”被夸大到上千。失败行映射仍基于物理行号（见 write_failed_csv），不受影响。
    data_rows_all = list(rows)
    data_count = sum(1 for r in data_rows_all
                     if any(c is not None and str(c).strip() != "" for c in r))
    if data_count == 0:
        raise ValueError("文件无任何非空数据行")

    # 表头严格校验（2026-08-01 新增，默认开启）：与正确模板逐列（列数+列名+顺序）
    # 完全一致才放行。放在最前，保证差异报告精确到具体列。
    ok, diffs = validate_header(header, canonical)
    if not ok:
        report = ("表头校验未通过，已拒绝上传。差异如下：\n"
                  + "\n".join(f"  - {d}" for d in diffs))
        try:
            FAILED_DIR.mkdir(exist_ok=True)
            stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
            rep = FAILED_DIR / f"header_check_{stamp}.txt"
            rep.write_text(
                report + "\n\n实际表头（%d 列）：\n" % len(header)
                + "\n".join(f"{i+1}. {h}" for i, h in enumerate(header)),
                encoding="utf-8")
            report += f"\n（详细报告见: {rep}）"
        except Exception:
            pass
        raise ValueError(report)

    # 核心列兜底（严格校验已覆盖，此处仅作极端兜底）
    missing = [h for h in REQUIRED_HEADERS if h not in header]
    if missing:
        raise ValueError(f"缺少必需列: {missing}")

    wb.close()
    return header, data_count


def convert_xlsx_to_csv(xlsx_path: Path, csv_path: Path, encoding=ENCODING):
    """xlsx -> UTF-8 CSV（无 BOM），保持列顺序与空单元格。"""
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    with open(csv_path, "w", encoding=encoding, newline="", errors="replace") as f:
        writer = csv.writer(f)
        for row in ws.iter_rows(values_only=True):
            # None -> 空字符串，保持列对齐
            writer.writerow(["" if c is None else c for c in row])
    wb.close()


def read_xlsx_data(xlsx_path: Path):
    """读取 xlsx 全部数据行（不含表头），用于失败时回贴原始内容到失败行 CSV。"""
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return []
    return [["" if c is None else c for c in r] for r in rows[1:]]


# ---------------------------------------------------------------------------
# 日志
# ---------------------------------------------------------------------------
def log(line: str):
    ts = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    text = f"[{ts}] {line}"
    print(text)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(text + "\n")
    except Exception:
        pass


# ---------------------------------------------------------------------------
# 会话
# ---------------------------------------------------------------------------
def load_session(context):
    if SESSION_FILE.exists():
        try:
            data = json.loads(SESSION_FILE.read_text(encoding="utf-8"))
            cookies = data.get("cookies") if isinstance(data, dict) else data
            if isinstance(cookies, list) and cookies:
                context.add_cookies(cookies)
                log("已载入保存的会话 cookies")
            else:
                log("会话文件无有效 cookies，将重新登录")
        except Exception as e:
            log(f"载入会话失败（将重新登录）: {e}")


def save_session(context):
    try:
        cookies = context.cookies()
        if not cookies:
            # 没有 cookies 时不要覆盖已有会话（否则会把有效会话冲成空）
            log("无 cookies 可保存（跳过，保留已有会话文件）")
            return
        sess = {"cookies": cookies}
        SESSION_FILE.write_text(json.dumps(sess, ensure_ascii=False), encoding="utf-8")
        log("已保存会话 cookies")
    except Exception as e:
        log(f"保存会话失败: {e}")


# ---------------------------------------------------------------------------
# 登录（按需）
# ---------------------------------------------------------------------------
def ensure_logged_in(page, headless: bool):
    """若检测到登录页，则登录。优先用环境变量凭据；否则等待手动登录。"""
    if not page.query_selector(PASSWORD_INPUT):
        return  # 看起来已登录（无密码框）
    user = os.environ.get("SITE_USER")
    pwd = os.environ.get("SITE_PASS")
    if user and pwd:
        log("检测到登录页，使用环境变量凭据自动登录")
        user_box = page.query_selector(SEL_LOGIN_USER) or page.query_selector("input[type=text]")
        if user_box:
            user_box.fill(user)
        page.fill(SEL_LOGIN_PASS, pwd)
        btn = page.locator(SEL_LOGIN_SUBMIT)
        if btn.count() == 0:
            btn = page.locator("button").last
        btn.first.click()
        # 等真正离开 login 页再继续
        try:
            page.wait_for_function("() => !location.href.includes('login')", timeout=20000)
        except PWTimeout:
            log("警告：登录后未检测到离开 login 页，继续尝试…")
        page.wait_for_timeout(2000)
    else:
        if headless:
            raise RuntimeError("检测到登录页但无凭据，且为 headless 模式。"
                               "请改用 --headless=false 手动登录一次，或用 SITE_USER/SITE_PASS 环境变量。")
        log("检测到登录页：请在打开的浏览器中手动登录，登录成功后脚本继续…")
        for _ in range(120):  # 最多等 2 分钟
            if not page.query_selector(PASSWORD_INPUT) or page.query_selector(SEL_FILE_INPUT):
                break
            page.wait_for_timeout(1000)
        else:
            raise RuntimeError("手动登录超时（2 分钟）")


def click_submit(page):
    """点击上传页的“批量导入”提交按钮（限定 button，避开卡片标题）。"""
    for sel in (SEL_SUBMIT_BTN, SEL_SUBMIT_FALLBACK):
        loc = page.locator(sel)
        if loc.count() > 0:
            loc.first.click()
            return sel
    # 兜底：带“导入/提交/上传”文字的任意按钮
    loc = page.locator("button", has_text=re.compile(r"导入|提交|上传"))
    if loc.count() > 0:
        loc.first.click()
        return "button[导入/提交/上传]"
    raise RuntimeError("未找到提交按钮（批量导入）")


# ---------------------------------------------------------------------------
# 结果解析
# ---------------------------------------------------------------------------
def parse_result(page, api_capture: list):
    """读取结果弹窗文字，提取成功/失败条数，并尽力抽取失败行明细与错误原因。

    站点实测两种弹窗：
      1) 成功：导入完成 / 总行数: N / 更新商品: N（或 新增商品: N）
      2) 前置校验失败：请检查导入数据 / 表格“错误信息”列（第 N 行：xxx）

    返回：(n_succ, n_fail, failed_rows, dbg, shot, completed, error_messages)
      error_messages: 错误明细字符串列表（如“第 2 行：找不到1级商品分类「家具卫浴」”）
    """
    try:
        page.wait_for_function(
            f"document.body && /{RESULT_WAIT_TEXT_RE}/.test(document.body.innerText)",
            timeout=60000,
        )
    except PWTimeout:
        log("警告：提交后未在 60s 内检测到结果文字，按当前页面解析")

    body_text = page.inner_text("body") or ""

    is_error_prompt = bool(RE_ERROR_PROMPT.search(body_text))
    completed = bool(RE_DONE.search(body_text))
    total = RE_TOTAL.search(body_text)
    updated = RE_UPDATED.search(body_text)
    added = RE_ADDED.search(body_text)
    fail_m = RE_FAIL_COUNT.search(body_text)

    # 成功数：优先 更新+新增；否则退化为 总行数 / “成功 N 条”
    n_succ = None
    if updated or added:
        n_succ = (int(updated.group(1)) if updated else 0) + \
                 (int(added.group(1)) if added else 0)
    elif total:
        n_succ = int(total.group(1))
    else:
        s = RE_SUCCESS.search(body_text)
        if s:
            n_succ = int(s.group(1))

    # 失败数 + 错误明细
    error_messages = []
    failed_rows = []

    if is_error_prompt:
        # 从弹窗 table 抽取错误信息列（antd Modal 内的 table）
        try:
            err_tbls = page.query_selector_all(SEL_ERROR_TABLE)
            for tbl in err_tbls:
                t = tbl.inner_text() or ""
                lines = [ln.strip() for ln in t.splitlines() if ln.strip()]
                for line in lines:
                    if re.match(r"^错误信息$", line):
                        continue  # 跳过表头
                    if line:
                        error_messages.append(line)
                        failed_rows.append([line])
        except Exception as e:
            log(f"错误弹窗表格抽取异常（已忽略）: {e}")
        n_fail = len(error_messages)
    else:
        # 成功/其它弹窗：优先 “失败: M”；否则 “失败 M 条”；导入完成且无失败文字→视为 0
        if fail_m:
            n_fail = int(fail_m.group(1))
        else:
            f = RE_FAIL.search(body_text)
            if f:
                n_fail = int(f.group(1))
            elif completed:
                n_fail = 0
            else:
                n_fail = None
        # 尽力抽取失败行：遍历页面上的 table，找含“失败/原因/错误”的表
        try:
            tables = page.query_selector_all("table")
            for tbl in tables:
                t = tbl.inner_text() or ""
                if re.search(r"失败|原因|错误|不通过|无效|行号|第.*行", t):
                    lines = [ln for ln in t.splitlines() if ln.strip()]
                    for line in lines[1:]:  # 跳过表头
                        cells = [c.strip() for c in line.split("\t")]
                        if any(cells):
                            failed_rows.append(cells)
                            if len(cells) == 1:
                                error_messages.append(cells[0])
        except Exception as e:
            log(f"失败行表格抽取异常（已忽略）: {e}")

    log(f"结果解析：导入完成={completed} 校验失败弹窗={is_error_prompt} "
        f"总行数={total.group(1) if total else '?'} "
        f"更新={updated.group(1) if updated else 0} 新增={added.group(1) if added else 0} "
        f"失败={n_fail}")
    if error_messages:
        log("错误明细：")
        for m in error_messages:
            log(f"  - {m}")

    # 始终保存原始结果 HTML + 截图，便于核对/调选择器
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    FAILED_DIR.mkdir(exist_ok=True)
    dbg = FAILED_DIR / f"debug_result_{stamp}.html"
    try:
        dbg.write_text(page.content(), encoding="utf-8")
    except Exception:
        dbg = None
    try:
        shot = FAILED_DIR / f"result_{stamp}.png"
        page.screenshot(path=str(shot), full_page=True)
    except Exception:
        shot = None

    # 留存接口响应（若截获到）
    if api_capture:
        cap_path = FAILED_DIR / f"api_capture_{stamp}.json"
        try:
            cap_path.write_text(json.dumps(api_capture, ensure_ascii=False, indent=2),
                                encoding="utf-8")
        except Exception:
            pass

    return n_succ, n_fail, failed_rows, dbg, shot, completed, error_messages


def write_failed_csv(error_messages, header, data_rows, reason_note: str,
                     error_detail=None):
    """把错误明细写出为 CSV：原始行号 + 货号/条码 + 错误信息 + 该行的原始数据。

    error_messages : 字符串列表，如 “第 2 行：单箱宽度不能为空…”（弹窗兜底用）
    error_detail   : list[dict]{barcode,code,detail}（接口结构化，优先），
                     detail 形如 “第 N 行：原因”
    两种来源合并去重后写盘，含原始行数据便于直接定位修改。
    """
    rows_out = []  # (row_num, code, barcode, err_text)
    seen = set()
    def _add(row_num, code, barcode, err_text):
        # 去重以错误原因为准：同一错误原因（无论行号是否已知）只写一行，
        # 避免接口结构化版（有行号）与弹窗兜底版（无行号）重复写出。
        key = err_text
        if key in seen:
            return
        seen.add(key)
        rows_out.append((row_num, code or "", barcode or "", err_text or ""))
    # 0) 反查表：货号/条码 -> 数据行索引（detail 无『第N行』时用）
    code_idx = None
    barcode_idx = None
    if header:
        for i, h in enumerate(header):
            hl = (h or "").strip()
            if hl == "货号":
                code_idx = i
            elif hl == "条形码":
                barcode_idx = i
    # 1) 结构化接口数据（优先，含货号/条码）
    if error_detail:
        for it in error_detail:
            if not isinstance(it, dict):
                continue
            detail = it.get("detail") or ""
            m = RE_ERROR_ROW.search(detail)
            row_num = int(m.group(1)) if m else None
            err_text = m.group(2).strip() if m else detail
            # 无行号时，用货号/条码在 data_rows 反查（站点亦有『货号 XXX 已存在』式错误）
            if row_num is None and data_rows:
                code = str(it.get("code") or "").strip()
                barcode = str(it.get("barcode") or "").strip()
                for idx, row in enumerate(data_rows):
                    if code_idx is not None and code and str(row[code_idx]).strip() == code:
                        row_num = idx + 1; break
                    if barcode_idx is not None and barcode and str(row[barcode_idx]).strip() == barcode:
                        row_num = idx + 1; break
            _add(row_num, it.get("code"), it.get("barcode"), err_text)
    # 2) 弹窗文本兜底
    if error_messages:
        for msg in error_messages:
            m = RE_ERROR_ROW.search(msg)
            row_num = int(m.group(1)) if m else None
            err_text = m.group(2).strip() if m else msg
            _add(row_num, None, None, err_text)
    if not rows_out:
        return None
    FAILED_DIR.mkdir(exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    out = FAILED_DIR / f"失败行_{stamp}.csv"
    with open(out, "w", encoding=ENCODING, newline="") as f:
        w = csv.writer(f)
        w.writerow(["原始行号", "货号", "条码", "错误信息"] + (list(header) if header else []))
        for row_num, code, barcode, err_text in rows_out:
            data_idx = (row_num - 1) if row_num else None
            raw = data_rows[data_idx] if (data_rows and data_idx is not None
                                          and 0 <= data_idx < len(data_rows)) else []
            w.writerow([row_num if row_num else "", code, barcode, err_text] + list(raw))
        w.writerow([])
        w.writerow(["# 备注:", reason_note])
    return out


def record_outcome(n_succ, n_fail, completed, dbg, shot,
                  summary=None):
    """把本次结果落到 last_run.json——Chrome 关闭时 shell 可能被掐断，
    导致进程退出码不可靠；此文件作为权威结果记录，供人或自动化读取。
    summary: 任务中心明细弹窗的 (共, 成功, 新增, 修改, 失败) 五元组（2026-08-26 新格式）。"""
    code = 0 if (completed and not (n_fail and n_fail > 0)) else 1
    out = {
        "ts": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "success": n_succ,
        "fail": n_fail,
        "completed": completed,
        "ok": code == 0,
        "exit_code": code,
        "debug_html": str(dbg) if dbg else None,
        "screenshot": str(shot) if shot else None,
    }
    if summary:
        out["summary"] = {
            "total": summary[0], "success": summary[1],
            "added": summary[2], "modified": summary[3], "fail": summary[4],
        }
    try:
        (BASE_DIR / "last_run.json").write_text(
            json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass
        return out["ok"]


# ---------------------------------------------------------------------------
# 任务中心（2026-08-15 系统更新：导入改为异步任务，结果到任务中心查看）
# ---------------------------------------------------------------------------
TASK_POLL_TIMEOUT = 900  # 秒：等待任务到达终态的最长总时长（实测慢任务约 8 分钟才落终态）


def wait_submit_popup(page):
    """提交后等待站点弹出的“导入任务已提交，请到任务中心查看进度”回执。
    站点现已改异步导入：提交成功不再同步弹“导入完成”，而是弹此回执。"""
    try:
        page.wait_for_function(
            "document.body && /导入任务已提交|任务中心查看进度|请到任务中心/.test(document.body.innerText)",
            timeout=60000,
        )
        log("检测到提交回执：导入任务已提交（异步任务已建立）")
        return True
    except PWTimeout:
        log("警告：提交后未检测到“导入任务已提交”回执，尝试直接进任务中心轮询")
        return False


def poll_task_center(page, file_name, task_id=None, country_code="CN",
                    timeout=TASK_POLL_TIMEOUT, headless=True,
                    task_center_url=TASK_CENTER_URL):
    """进入任务中心轮询本次提交的任务，直到到达终态（已完成 / 失败）。

    判定以列表接口返回的 status 字段为准（比 DOM 文本稳健，SPA 列表加载抖动大）：
      status=0 待处理/队列中；status=2 进行中；status=3 已完成(成功)；status=1 失败。
    匹配优先用提交回执里的精确 task_id（data:196），回退才用文件名模糊匹配。
    返回 (status, has_error, task_id, info)：
      status    : 'completed' | 'timeout'
      has_error : 任务是否为失败任务（status==1 或 hasErrorDetail）
      task_id   : 匹配到的任务 id
      info      : 命中的列表项 dict（含 fileName/status/hasErrorDetail 等）
    """
    # 站点任务状态枚举（2026-08-26 实测：列表接口 data.list[].status）
    ST_DONE_OK = 3      # 已完成（成功）
    ST_FAILED = 1       # 失败
    ST_PROCESSING = (0, 2)  # 待处理/队列中/进行中

    captured = {}  # taskId(str) -> task info dict（来自列表接口响应）

    def on_response(response):
        if "api" in response.url and response.request.method == "POST":
            try:
                body = response.text()
            except Exception:
                return
            if '"fileName"' in body and '"status"' in body:
                try:
                    data = json.loads(body)
                except Exception:
                    return
                d = data.get("data") if isinstance(data.get("data"), dict) else data
                if not isinstance(d, dict):
                    return
                for item in (d.get("list") or []):
                    fn = item.get("fileName") or ""
                    tid = str(item.get("id"))
                    # 精确 id 命中 或 文件名模糊命中
                    if (task_id and str(task_id) == tid) or (file_name and file_name in str(fn)):
                        captured[tid] = item

    page.on("response", on_response)

    log(f"进入任务中心轮询（file={file_name}，task_id={task_id}）")
    deadline = time.time() + timeout

    def _enter_and_capture():
        """进任务中心并尽量抓取列表；以 captured（接口数据）为准，DOM 仅辅助。"""
        try:
            page.goto(task_center_url, wait_until="load", timeout=30000)
        except Exception as e:
            log(f"进任务中心异常（忽略）: {e}")
        if page.query_selector(PASSWORD_INPUT):
            ensure_logged_in(page, headless)
        # 给列表 XHR 留出加载时间（SPA：占位行→真实数据约需数秒）
        page.wait_for_timeout(4000)

    # 首轮：若提交回执已带回 task_id，先用它精确匹配（不依赖列表渲染）
    def _match():
        if task_id and str(task_id) in captured:
            return captured[str(task_id)]
        # 回退：文件名精确→模糊
        for tid, it in captured.items():
            if file_name and file_name == (it.get("fileName") or ""):
                return it
        for tid, it in captured.items():
            if file_name and file_name in (it.get("fileName") or ""):
                return it
        # DOM 兜底（列表接口未捕获时）
        try:
            rows = [r for r in page.query_selector_all(".ant-table-tbody tr")
                    if "measure-row" not in (r.get_attribute("class") or "")
                    and "暂无数据" not in (r.inner_text() or "")]
            for r in rows:
                if file_name and file_name in (r.inner_text() or ""):
                    return {"_dom": True}
        except Exception:
            pass
        return None

    _enter_and_capture()
    while time.time() < deadline:
        it = _match()
        if it is None:
            log("任务中心尚未出现本次任务（列表未刷新），等待后重试…")
        else:
            st = it.get("status")
            has_err = bool(it.get("hasErrorDetail")) or st == ST_FAILED
            if st in ST_PROCESSING or st is None:
                log(f"任务处理中（status={st}），继续等待…")
            else:
                if st == ST_DONE_OK:
                    log("任务已到终态：已完成（成功）")
                    return ("completed", False, str(it.get("id") or task_id), it)
                if st == ST_FAILED:
                    log("任务已到终态：失败")
                    return ("completed", True, str(it.get("id") or task_id), it)
                # 未知 status 但已非处理中：保守按完成计
                log(f"任务已到终态（未知 status={st}），保守按完成计")
                return ("completed", has_err, str(it.get("id") or task_id), it)
        page.wait_for_timeout(5000)
        _enter_and_capture()
    return ("timeout", False, str(task_id) if task_id else None, None)


# 失败明细弹窗内容异步填充、偶发空（“暂无数据”），需重试多次抓取真实错误
FAIL_DETAIL_RETRY = 4
FAIL_DETAIL_GAP = 4  # 秒


def _modal_has_real_detail(text):
    """弹窗文本是否含真实错误（而非『暂无数据』占位）。"""
    if not text:
        return False
    return "暂无数据" not in text


def fetch_task_error_detail_via_click(page, task_id, click_btn, retries=3):
    """拦截前端点『失败明细』时自发发出的任务详情接口响应，取结构化 errorDetail。

    点按钮前端会用正确签名发 POST /api（body 含 taskId），响应 data.errorDetail
    是 JSON 数组：[{barcode, code, detail:"第 N 行：原因"}, ...]，比弹窗可靠
    （弹窗偶发『暂无数据』但接口始终有数据）。
    返回 list[dict] 或 None。
    """
    captured = []
    def on_resp(r):
        try:
            if r.url.endswith("/api") and r.request.method == "POST":
                b = r.text()
                if '"errorDetail"' in b or 'errorDetail":' in b:
                    captured.append(b)
        except Exception:
            pass
    page.on("response", on_resp)
    try:
        for attempt in range(1, retries + 1):
            try:
                click_btn.click()
            except Exception:
                pass
            # 等待详情接口返回
            for _ in range(20):
                page.wait_for_timeout(500)
                if captured:
                    break
            if captured:
                break
            # 没捕获到：尝试关闭弹窗再点
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
            page.wait_for_timeout(1000)
        if not captured:
            return None
        for body in reversed(captured):
            try:
                data = json.loads(body)
            except Exception:
                continue
            d = data.get("data") if isinstance(data, dict) else None
            ed = d.get("errorDetail") if isinstance(d, dict) else None
            if not ed:
                continue
            items = json.loads(ed) if isinstance(ed, str) else ed
            if isinstance(items, list) and items:
                return items
        return None
    finally:
        page.remove_listener("response", on_resp)

def open_task_detail_modal(page, file_name, task_id=None, task_center_url=TASK_CENTER_URL):
    """在任务列表定位本次任务行，读取失败/成功明细。

    失败原因优先用【任务详情接口】的结构化 errorDetail（data.errorDetail JSON 数组），
    它比弹窗可靠——弹窗偶发为『暂无数据』，但接口始终返回完整错误。
    弹窗文本仅作为兜底/可视化。
    定位优先级：task_id（data-row-key，精确，避免同名多次上传撞错行）→ 文件名模糊。
    返回 (modal_text, kind, error_detail)：
      modal_text  : 弹窗文本（成功明细多为 None，失败明细可能为空）
      kind        : '失败明细' | '成功明细' | None
      error_detail: list[dict]{barcode,code,detail} 或 None（结构化失败原因）
    """
    def _find_target():
        on_tc = task_center_url in (page.url or "")
        if not on_tc:
            _enter_tc()
        try:
            page.wait_for_selector(".ant-table-tbody tr", state="attached", timeout=20000)
        except PWTimeout:
            return None
        page.wait_for_timeout(2500)
        rows = page.query_selector_all(".ant-table-tbody tr")
        rows = [r for r in rows
                if "measure-row" not in (r.get_attribute("class") or "")
                and "暂无数据" not in (r.inner_text() or "")]
        if task_id:
            for r in rows:
                if (r.get_attribute("data-row-key") or "") == str(task_id):
                    return r
        for r in rows:
            t = r.inner_text() or ""
            if file_name and file_name in t:
                return r
        stem = file_name.rsplit(".", 1)[0] if file_name else ""
        for r in rows:
            t = r.inner_text() or ""
            if stem and stem in t:
                return r
        return None

    def _enter_tc():
        try:
            page.goto(task_center_url, wait_until="load", timeout=30000)
        except Exception as e:
            log(f"进任务中心异常（忽略）: {e}")
        if page.query_selector(PASSWORD_INPUT):
            ensure_logged_in(page, True)
        page.wait_for_timeout(3000)

    target = _find_target()
    if not target:
        _enter_tc()
        target = _find_target()
    if not target:
        log(f"任务中心列表未找到匹配 [{file_name}] 的行")
        return (None, None, None)

    fail_btn = target.query_selector("button:has-text('失败明细')")
    ok_btn = target.query_selector("button:has-text('成功明细')")
    kind = "失败明细" if fail_btn else ("成功明细" if ok_btn else None)
    if not (fail_btn or ok_btn):
        log("匹配行内未找到「明细」按钮")
        return (None, None, None)

    error_detail = None
    modal_text = None
    # 失败明细：优先拦截前端点『失败明细』时自发请求的详情接口响应，
    # 取结构化 errorDetail（data.errorDetail JSON 数组），比偶发为空的弹窗可靠。
    if kind == "失败明细" and fail_btn:
        error_detail = fetch_task_error_detail_via_click(page, task_id, fail_btn)
        if error_detail:
            log(f"已从任务详情接口取到失败明细 {len(error_detail)} 条")
    # 成功明细多为触发下载（不弹窗），直接返回
    if kind == "成功明细":
        return (None, "成功明细", error_detail)
    # 弹窗文本兜底（成功明细无弹窗；失败明细接口为空时再试弹窗）
    if not error_detail:
        best_text = None
        for attempt in range(1, FAIL_DETAIL_RETRY + 1):
            log(f"点击「失败明细」（第 {attempt}/{FAIL_DETAIL_RETRY} 次）")
            try:
                (fail_btn or ok_btn).click()
                page.wait_for_selector(".ant-modal-content", state="visible", timeout=15000)
                page.wait_for_timeout(1500)
                modal_text = page.inner_text(".ant-modal-content")
            except Exception:
                modal_text = None
            if _modal_has_real_detail(modal_text):
                best_text = modal_text
                break
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
            page.wait_for_timeout(FAIL_DETAIL_GAP * 1000)
            if attempt < FAIL_DETAIL_RETRY:
                _enter_tc()
                target = _find_target()
                if not target:
                    break
                fail_btn = target.query_selector("button:has-text('失败明细')")
        if best_text:
            modal_text = best_text
        elif not modal_text:
            log("失败明细弹窗为『暂无数据』，已尝试接口与弹窗均未取得明细，请人工核对任务中心")
    return (modal_text, "失败明细", error_detail)


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def run(file_arg: str, dry_run: bool, headless: bool, keep_csv: bool, station="cn"):
    xlsx_path = Path(file_arg).expanduser().resolve()
    log(f"开始处理: {xlsx_path}（站点: {SC.STATION_NAMES.get(station, station) if SC else station}）")

    # 多站点：上传地址 + 任务中心地址
    upload_url = (SC.UPLOAD_URLS.get(station, UPLOAD_URL) if SC else UPLOAD_URL)
    task_center_url = task_center_url_of(upload_url)

    # 1) 预检（多站点时按站点模板表头严格校验；CN 维持原基准表头逻辑）
    canonical = None
    if SC is not None and station != "cn":
        canonical = SC.TEMPLATE_HEADER.get(station)
    header, data_count = preflight(xlsx_path, canonical=canonical)
    log(f"预检通过：表头 {len(header)} 列，数据行 {data_count}")

    # 2) 转换（仅 CSV 模式需要；2026-08-15 起 CN 直接传 xlsx，不再转换）
    csv_path = None
    if not SUBMIT_XLSX:
        csv_path = BASE_DIR / f"_tmp_upload_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        convert_xlsx_to_csv(xlsx_path, csv_path, ENCODING)
        log(f"已转换为 UTF-8 CSV: {csv_path}")
    else:
        log("SUBMIT_XLSX=True：直接提交原始 xlsx（站点 2026-08-15 起不再接受 CSV）")

    # 3) dry-run：只预览
    if dry_run:
        if SUBMIT_XLSX:
            data_rows = read_xlsx_data(xlsx_path)
            nz = sum(1 for r in data_rows if any(c is not None and str(c).strip() != "" for c in r))
            log("=== DRY-RUN 预览（不提交）===")
            log(f"将提交 xlsx: {xlsx_path}")
            log(f"表头 {len(header)} 列，非空数据行 {nz}")
            for i, row in enumerate(data_rows[:3], 1):
                log(f"  样例行{i}: {row[:4]}")
        else:
            with open(csv_path, encoding=ENCODING, newline="") as f:
                head = [next(f) for _ in range(2)]
            log("=== DRY-RUN 预览（不提交）===")
            log("".join(head).strip())
            log(f"临时 CSV 路径（保留）: {csv_path}")
        print("\n[DRY-RUN] 未提交。去掉 --dry-run 再运行即真实上传。")
        return 0

    # 4) Playwright 提交（实际读取原始数据行，供失败行回贴）
    data_rows = read_xlsx_data(xlsx_path)
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=headless,
            channel="chrome",
            args=build_browser_args(),  # 自动探测：代理开着用代理，没开直连
        )
        context = browser.new_context(accept_downloads=True)
        load_session(context)
        page = context.new_page()

        # 截获提交后的接口响应，辅助解析结果
        api_capture = []
        def on_response(response):
            if "api" in response.url and response.request.method == "POST":
                try:
                    body = response.text()
                except Exception:
                    return
                if re.search(r"成功|失败|导入|totalRow|row|条", body):
                    api_capture.append({"url": response.url, "body": body[:3000]})
        page.on("response", on_response)

        try:
            log(f"打开上传页: {upload_url}")
            # 后台常有持续心跳/统计请求，networkidle 易超时；改用 load+元素等待
            page.goto(upload_url, wait_until="load", timeout=30000)
            # 无 session 时站点为客户端重定向（import-products → login?to=...），
            # 必须等页面稳定（出现密码框或文件框）再判断登录态，否则会误判为已登录。
            try:
                page.wait_for_function(
                    "() => document.querySelector('input[type=password]') "
                    "|| document.querySelector('input[type=file]')",
                    timeout=30000,
                )
            except PWTimeout:
                log("警告：打开上传页后既未出现登录框也未出现文件框，尝试继续")
            ensure_logged_in(page, headless)

            # 登录后确保在上传页（带鉴权）
            if page.query_selector(PASSWORD_INPUT) or "import-products" not in page.url:
                log("重新跳转上传页（带鉴权）")
                page.goto(upload_url, wait_until="load", timeout=30000)
                page.wait_for_timeout(2000)

            log("定位文件输入框并选择文件")
            # antd Upload 的文件 input 是隐藏的(display:none)，set_input_files 可作用，
            # 故只等“已挂载”而非“可见”，否则会因隐藏而超时。
            page.wait_for_selector(SEL_FILE_INPUT, state="attached", timeout=30000)
            if SUBMIT_XLSX:
                page.set_input_files(SEL_FILE_INPUT, str(xlsx_path))
            else:
                page.set_input_files(SEL_FILE_INPUT, str(csv_path))

            # 站点提示：等待出现 totalRow 数量提示之后再操作
            try:
                page.wait_for_selector(SEL_UPLOAD_ITEM, timeout=30000)
                log("文件已载入上传区")
            except PWTimeout:
                log("警告：未检测到上传文件列表项，仍尝试继续")
            # 给前端解析 CSV 留出时间（totalRow 数量提示）
            page.wait_for_timeout(5000)

            log("点击“批量导入”提交")
            used = click_submit(page)
            log(f"已点击提交按钮（选择器: {used}）")

            # 2026-08-15 起：导入为异步任务，提交后回执“请到任务中心查看进度”，
            # 不再同步弹“导入完成”。改为：等回执 → 进任务中心轮询 → 抓明细弹窗。
            wait_submit_popup(page)

            # 同步错误检测：若提交接口直接返回非 200（如重复提交 / 文件非法），
            # 不会生成异步任务，需立即报错退出，避免干等任务中心 10 分钟。
            submit_resp = None
            for cap in api_capture:
                if "import" in cap["url"].lower() or re.search(r"导入任务已提交|保存导入文件失败|文件.*不正确|重复", cap["body"]):
                    submit_resp = cap["body"]
                    break
            if submit_resp:
                if (re.search(r'"code"\s*:\s*(?!200)\d+', submit_resp)
                        or re.search(r"保存导入文件失败|不正确|重复|已存在|不规范|失败", submit_resp)):
                    log(f"[警告]  提交被站点同步拒绝（未生成异步任务）: {submit_resp[:600]}")
                    record_outcome(None, None, False, None, None)
                    save_session(context)
                    browser.close()
                    return 1
                log(f"提交接口响应: {submit_resp[:300]}")
                # 提取任务 id（回执 data 字段，如 {"msg":"导入任务已提交...","code":200,"data":196}）
                try:
                    _sr = json.loads(submit_resp)
                    _d = _sr.get("data")
                    submit_task_id = int(_d) if isinstance(_d, int) else None
                except Exception:
                    submit_task_id = None
                if submit_task_id:
                    log(f"已获取任务 id={submit_task_id}")

            file_stem = xlsx_path.name
            status, has_error, tid, _info = poll_task_center(
                page, file_stem, task_id=submit_task_id, country_code=station.upper(),
                headless=headless, task_center_url=task_center_url)
            completed = (status == "completed")
            n_succ = None
            n_fail = 0
            error_messages = []

            if status == "timeout":
                log("[警告]  任务中心轮询超时，无法确认结果；退出码 1")
                record_outcome(n_succ, n_fail, False, None, None)
                save_session(context)
                browser.close()
                return 1

            # 抓取明细弹窗（失败明细 / 成功明细）
            # open_task_detail_modal 内部负责进入任务中心并等待表格稳定（SPA 异步），
            # 这里不再重复 goto，避免把已稳定的表格再冲掉。
            modal_text, kind, error_detail = open_task_detail_modal(
                page, file_stem, task_id=submit_task_id, task_center_url=task_center_url)
            # 成败以按钮文字(kind)为准；明细弹窗摘要行给出精确计数。
            # 2026-08-26 起格式：导入成功：共N条，成功N条（新增X条，修改Y条），失败Z条
            summary = parse_task_summary(modal_text) if modal_text else None

            if kind == "成功明细":
                if summary:
                    n_succ = summary[1]
                    n_fail = summary[4]
                    log(f"任务成功：共 {summary[0]} 条，成功 {n_succ} 条"
                        f"（新增 {summary[2]}，修改 {summary[3]}），失败 {n_fail} 条")
                else:
                    # 成功明细触发下载而非弹窗，无法读摘要：保守按行数计
                    n_succ = data_count
                    n_fail = 0
                    log(f"任务成功（明细弹窗未读取到摘要）：成功 {n_succ} 行")
            elif kind == "失败明细":
                # 失败任务：优先用【接口结构化 error_detail】（含货号/条码/行号/原因），
                # 弹窗文本仅作为兜底。两者合并去重写入失败 CSV。
                error_messages = []
                if error_detail:
                    for it in error_detail:
                        if not isinstance(it, dict):
                            continue
                        detail = it.get("detail") or ""
                        m = RE_ERROR_ROW.search(detail)
                        if m:
                            error_messages.append(
                                f"第 {m.group(1)} 行：{m.group(2).strip()} "
                                f"（货号 {it.get('code','')} / 条码 {it.get('barcode','')}）")
                        else:
                            error_messages.append(detail)
                # 弹窗文本兜底（接口无数据时）
                if not error_messages and modal_text:
                    for m in RE_ERROR_ROW.finditer(modal_text):
                        err_text = m.group(2).strip()
                        if err_text:
                            error_messages.append(f"第 {m.group(1)} 行：{err_text}")
                # 去重
                seen = set()
                error_messages = [x for x in error_messages if not (x in seen or seen.add(x))]
                if summary:
                    n_fail = summary[4]
                    n_succ = max(0, summary[0] - summary[4])
                    log(f"失败明细解析：共 {summary[0]} 条，成功 {n_succ} 条，失败 {n_fail} 条")
                elif error_detail:
                    n_fail = len(error_detail)
                    n_succ = max(0, data_count - n_fail) if data_count else n_fail
                    log(f"失败明细解析（接口结构化）：失败 {n_fail} 条 / 共 {data_count} 数据行")
                elif error_messages:
                    n_fail = len(error_messages)
                    n_succ = max(0, data_count - n_fail) if data_count else n_fail
                    log(f"失败明细解析（弹窗文本）：失败 {n_fail} 行 / 共 {data_count} 数据行")
                else:
                    n_fail = -1
                    log("[警告]  任务为失败任务，接口与弹窗均未取得逐行原因，请人工核对任务中心")
                for m in error_messages:
                    log(f"  - {m}")
            else:
                # 弹窗/按钮均未能读取：以 poll_task_center 的权威 status 判定，
                # 不打印疑似失败的措辞。has_error 由接口 hasErrorDetail/status 决定。
                if has_error:
                    n_succ = 0
                    n_fail = max(1, data_count)
                    log(f"[警告]  任务为失败任务（status/hasErrorDetail），且未能读取失败明细弹窗，请人工核对任务中心")
                else:
                    n_succ = data_count
                    n_fail = 0
                    log(f"任务已完成（明细弹窗未读取到，按接口状态计）成功 {n_succ} 行")

            # 调试产物（在 browser.close 可能掐断 shell 之前落盘）
            FAILED_DIR.mkdir(exist_ok=True)
            stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
            try:
                dbg = FAILED_DIR / f"debug_task_{stamp}.html"
                dbg.write_text(page.content(), encoding="utf-8")
            except Exception:
                dbg = None
            try:
                shot = FAILED_DIR / f"task_{stamp}.png"
                page.screenshot(path=str(shot), full_page=True)
            except Exception:
                shot = None

            # 先把权威结果落盘（含新格式摘要）
            record_outcome(n_succ, n_fail, completed, dbg, shot, summary=summary)

            failed_csv = None
            if error_messages or error_detail:
                failed_csv = write_failed_csv(
                    error_messages, header, data_rows,
                    f"站点导入失败（成功{n_succ}/失败{n_fail}）",
                    error_detail=error_detail)
                if failed_csv:
                    log(f"已写出失败行清单（含错误原因/货号/条码）: {failed_csv}")
                else:
                    log("[警告]  站点报告有失败，但未能写出失败行 CSV（见调试HTML）")

            save_session(context)
            browser.close()

            if n_fail and n_fail > 0:
                log(f"[警告]  存在失败行 {n_fail} 条，退出码 1。请查看失败行 CSV 中的错误原因。")
                return 1
            if completed:
                log("[完成]  上传完成（任务中心：已完成）")
                # 仅在“确认上传成功”后备份到 已导入表格/<站>/
                # （站点报失败 / 轮询超时 / 未确认完成 都不备份，避免污染已导入台账）
                backup_after_upload(xlsx_path, station)
                return 0
            log("[警告]  未确认导入完成，退出码 1")
            return 1
        except Exception as e:
            log(f"运行异常: {e}")
            log(traceback.format_exc())
            try:
                save_session(context)
            except Exception:
                pass
            try:
                browser.close()
            except Exception:
                pass
            return 2
        finally:
            # 仅 CSV 模式才清理临时转换文件；SUBMIT_XLSX 模式下 xlsx 是用户原文件，绝不能删
            if (not SUBMIT_XLSX) and (not keep_csv) and csv_path and csv_path.exists():
                try:
                    csv_path.unlink()
                except Exception:
                    pass


def main():
    ap = argparse.ArgumentParser(description="商品自动上传机器人（Playwright）")
    ap.add_argument("--file", required=True, help="人工备好的 xlsx 模板路径")
    ap.add_argument("--dry-run", action="store_true", help="只转换+预览，不真正提交")
    ap.add_argument("--headless", action="store_true", default=True,
                    help="无头模式（默认 True）；调试/手动登录用 --headless=false")
    ap.add_argument("--keep-csv", action="store_true", help="保留临时转换的 CSV")
    ap.add_argument("--save-template-header", action="store_true",
                    help="把当前 --file 的表头存为基准模板(header_template.txt)，不提交")
    ap.add_argument("--station", default="cn", choices=["cn", "es", "gr"],
                    help="上传站点：cn=中国站(默认) / es=西班牙站 / gr=希腊站；"
                         "决定上传地址、表头校验基准、成功后备份目录")
    args = ap.parse_args()

    # 刷新基准表头（站点模板变更时用）：仅存模板，不提交
    if args.save_template_header:
        try:
            hb = openpyxl.load_workbook(args.file, read_only=True, data_only=True)
            hws = hb[hb.sheetnames[0]]
            hdr = [str(x).strip() for x in next(hws.iter_rows(values_only=True)) if x is not None]
            hb.close()
            # 多站点：非 CN 存到独立文件，避免覆盖 CN 基准
            if args.station != "cn":
                tpl = BASE_DIR / f"header_template_{args.station}.txt"
            else:
                tpl = HEADER_TEMPLATE_FILE
            tpl.write_text("\n".join(hdr) + "\n", encoding="utf-8")
            print(f"[OK] 已把 {len(hdr)} 列表头存为基准模板({args.station}): {tpl}")
            sys.exit(0)
        except Exception as e:
            print(f"保存模板表头失败: {e}")
            sys.exit(2)

    try:
        code = run(args.file, args.dry_run, args.headless, args.keep_csv, station=args.station)
        sys.exit(code)
    except Exception as e:
        log(f"致命错误: {e}")
        log(traceback.format_exc())
        sys.exit(2)


if __name__ == "__main__":
    main()
