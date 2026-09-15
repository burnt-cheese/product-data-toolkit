# -*- coding: utf-8 -*-
"""
站点适配转换器 (site_adapt) 
===========================
把 China 形态的「货号匹配结果.xlsx」 (或任意已按模板填好的源文件) 转换成
目标站点 (cn/es/gr) 的上传模板形态, 并应用各站规则: 

  - 固定填值 (税率/币种/品牌/打折/原产地) 
  - 计算规则 (最低起订量 = 中包每包可装个数) 
  - 查表规则 (西班牙 西语品名 ← A006.外文名称, 按货号) 
  - 翻译规则 (希腊 希腊语品名 ← 品名 机器翻译 ZH->EL) 
  - 强制留空 (西班牙 产品详情/产品属性) 

站点选择: 不传 --station 会弹 WinForms 窗口 (station_pick.ps1) 让你选, 默认中国站.

备份: 默认不备份——按用户要求, 仅当 upload.py **上传成功后**才备份到 已导入表格/<站>/.
 (需要生成时就留档可显式加 --backup) 

用法: 
  python site_adapt.py --input 货号匹配结果.xlsx                # 弹窗选站
  python site_adapt.py --station cn --input 货号匹配结果.xlsx
  python site_adapt.py --station es --input 货号匹配结果.xlsx --out 输出
  python site_adapt.py --station gr --input 货号匹配结果.xlsx --pathfile 上次生成.txt
"""

import argparse
import os
import re
import sys
import time
import shutil
import zipfile
import pickle
import xml.etree.ElementTree as ET

# 复用 station_config 里的集中配置
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import station_config as SC

WORKSPACE = SC.WORKSPACE
DEFAULT_A006 = os.path.join(WORKSPACE, "货号匹配器", "数据源", "A006-货号统计（全部）.xlsx")
DEFAULT_OUT = os.path.join(WORKSPACE, "货号匹配器", "输出")

ILLEGAL_B = re.compile(rb"[\x00-\x08\x0b\x0c\x0e-\x1f]")

# 翻译相关: 免费端点不稳定 (偶发 500/限流, 重试退避最坏每 row约 15s) , 
# 故做「持久化缓存 + 连续失败熔断」, 避免大表把流程拖到几十分钟.
TRANS_CACHE_PATH = os.path.join(SC.CACHE_DIR, "trans_el.pkl")
MAX_CONSEC_FAIL = 5                                              # 连续失败这么多次 => 判定被限流
MAX_COOLDOWN_CYCLES = 3                                          # 最多长冷却几次 (之后才放弃) 
COOLDOWN_SECONDS = 90                                            # 每次长冷却秒数 (等限流窗口过去) 
TRANS_SLEEP = float(os.environ.get("TRANS_SLEEP", "0.4"))  # 每次调接口后的礼貌间隔


def col_letter_to_idx(col):
    """列字母 (如 'AB') 转 0-based 索引."""
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def norm_code(c):
    """货号规范化: 去空格, 去 .0 尾巴, 便于 A006 查表匹配."""
    c = str(c).strip()
    if c.endswith(".0"):
        c = c[:-2]
    return c


# ---------------------------------------------------------------------------
# A006 外文名称查表 (货号 -> 外文名称) , 带缓存
# ---------------------------------------------------------------------------
def read_a006_foreign(a006_path, cache_dir):
    """解析 A006, 返回 {货号: 外文名称}.结果按 A006 修改时间缓存."""
    os.makedirs(cache_dir, exist_ok=True)
    cache = os.path.join(cache_dir, "a006_foreign.pkl")
    try:
        mtime = os.path.getmtime(a006_path)
    except Exception:
        mtime = 0
    if os.path.exists(cache):
        try:
            with open(cache, "rb") as f:
                d = pickle.load(f)
            if d.get("mtime") == mtime and "lookup" in d:
                print(f"  [A006] cache hit: {len(d['lookup'])}  entries of foreign names")
                return d["lookup"]
        except Exception:
            pass
    print("  [A006] first parse of foreign names (code->name), large file takes ~10s...")
    z = zipfile.ZipFile(a006_path)
    ss_bytes = ILLEGAL_B.sub(b"", z.read("xl/sharedStrings.xml"))
    ss_bytes = re.sub(rb'\sxmlns="[^"]+"', b"", ss_bytes, count=1)
    ss_root = ET.fromstring(ss_bytes)
    shared = ["".join(t.text or "" for t in si.iter("t")) for si in ss_root.iter("si")]
    ws_bytes = ILLEGAL_B.sub(b"", z.read("xl/worksheets/sheet1.xml"))
    ws_bytes = re.sub(rb'\sxmlns="[^"]+"', b"", ws_bytes, count=1)
    wroot = ET.fromstring(ws_bytes)
    rows = wroot.findall(".//row")
    # 表头列字母
    hdr = {}
    for c in rows[0].findall("c"):
        ref = c.get("r") or ""
        col = "".join(ch for ch in ref if ch.isalpha())
        t = c.get("t"); v = c.find("v")
        val = v.text if v is not None else ""
        if t == "s":
            try:
                val = shared[int(val)]
            except Exception:
                val = ""
        hdr[str(val).strip()] = col
    code_col = hdr.get("货号"); foreign_col = hdr.get("外文名称")
    if not code_col or not foreign_col:
        print("  [A006] 货号/外文名称  columns missing, skip lookup")
        return {}
    ci = col_letter_to_idx(code_col); fi = col_letter_to_idx(foreign_col)
    lookup = {}
    for row in rows[1:]:
        cells = {}
        for c in row.findall("c"):
            ref = c.get("r") or ""
            col = "".join(ch for ch in ref if ch.isalpha())
            idx = col_letter_to_idx(col)
            t = c.get("t"); v = c.find("v"); isn = c.find("is")
            if t == "s" and v is not None:
                try:
                    val = shared[int(v.text)]
                except Exception:
                    val = ""
            elif isn is not None:
                val = "".join(tt.text or "" for tt in isn.iter("t"))
            elif v is not None:
                val = v.text
            else:
                val = ""
            cells[idx] = val
        code = norm_code(cells.get(ci))
        foreign = (cells.get(fi) or "").strip()
        if code:
            lookup[code] = foreign
            if code.isdigit():
                lookup[str(int(code))] = foreign  # 去前导零版本
    with open(cache, "wb") as f:
        pickle.dump({"mtime": mtime, "lookup": lookup}, f)
    print(f"  [A006] foreign-name lookup ready: {len(lookup)} entries")
    return lookup


# ---------------------------------------------------------------------------
# 翻译 (中文 -> 希腊语) , 带重试 + 退避 + 失败降级
# ---------------------------------------------------------------------------
# 免费端点偶尔返回错误页 HTML 而非抛异常, 需显式识别为失败
_ERR_PAT = re.compile(r"error|server error|that'?s an error|translate\.google|<!doctype|<!DOCTYPE", re.I)


def translate_zh_to_el(text, max_retries=4, base_delay=1.5):
    """翻译中文品名为希腊语.任意失败返回 None (由调用方留空+警告) .

    Note: deep-translator 在免费端点 500 时可能把错误页 HTML 当"翻译结果"返回
     (不抛异常) , 故对返回值做错误特征校验, 命中则当作失败重试.
    """
    from deep_translator import GoogleTranslator
    text = (text or "").strip()
    if not text:
        return ""
    last = None
    for attempt in range(1, max_retries + 1):
        try:
            r = GoogleTranslator(source="auto", target="el").translate(text)
        except Exception as e:
            last = e
            time.sleep(base_delay * attempt)
            continue
        if r and not _ERR_PAT.search(str(r)):
            return r
        # 返回值是错误页 -> 视为失败, 重试
        last = ValueError("bad translation response: " + str(r)[:60])
        time.sleep(base_delay * attempt)
    return None


def translate_batch_zh_to_el(texts, max_retries=3, base_delay=2.0):
    """批量翻译 (中文->希腊语等) .texts 等长返回 list[str]; 整体失败返回 None.

    把 N 次逐 row请求合成 N/批 次, 免费端点下整体快 10~20 倍.
    端点偶发 500 会整批返错误页/抛异常, 统一判失败; 单条异常留空不拖垮整批.
    """
    from deep_translator import GoogleTranslator
    clean = [(t or "").strip() for t in texts]
    if not any(clean):
        return clean
    last = None
    for attempt in range(1, max_retries + 1):
        try:
            res = GoogleTranslator(source="auto", target="el").translate_batch(clean)
        except Exception as e:
            last = e
            time.sleep(base_delay * attempt)
            continue
        if isinstance(res, list) and len(res) == len(clean):
            out = []
            for r in res:
                if r and not _ERR_PAT.search(str(r)):
                    out.append(r)
                else:
                    out.append("")   # 单条失败留空
            return out
        last = ValueError("bad batch translation response")
        time.sleep(base_delay * attempt)
    return None


def load_trans_cache(path=TRANS_CACHE_PATH):
    """载入希腊语翻译缓存 {中文品名: 希腊语}, 跨次运 row复用."""
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
    except Exception:
        pass
    if os.path.exists(path):
        try:
            with open(path, "rb") as f:
                d = pickle.load(f)
            if isinstance(d, dict):
                return d
        except Exception:
            pass
    return {}


def save_trans_cache(cache, path=TRANS_CACHE_PATH):
    """保存翻译缓存 (失败不影响主流程) ."""
    try:
        with open(path, "wb") as f:
            pickle.dump(cache, f)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# 站点选择 (GUI) 
# ---------------------------------------------------------------------------
def pick_station_via_gui(default="cn", timeout=600):
    """弹出 station_pick.ps1 站点选择窗, 返回 cn / es / gr.

    任何失败 (ps1 缺失 / PowerShell 异常 / 超时 / 结果非法) 都回落到 default, 
    保证 CLI 与 bat 编排不会被卡住.
    """
    import subprocess
    import tempfile

    ps1 = os.path.join(SC.WORKSPACE, "station_pick.ps1")
    if not os.path.exists(ps1):
        print(f"  [NOTE] not found: {ps1}, falling back to default {SC.STATION_NAMES.get(default, default)}")
        return default

    out = os.path.join(tempfile.gettempdir(), f"station_pick_{os.getpid()}.txt")
    try:
        if os.path.exists(out):
            os.remove(out)
    except Exception:
        pass

    print("  Station picker popup (default China)...")
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps1, out],
            timeout=timeout, check=False,
        )
    except Exception as e:
        print(f"  [NOTE] station picker unavailable ({e}) , falling back to default {SC.STATION_NAMES.get(default, default)}")
        return default

    try:
        with open(out, "r", encoding="utf-8") as f:
            v = (f.read() or "").strip().lower()
    except Exception:
        return default
    if v in SC.TEMPLATE_HEADER:
        print(f"  Selected station: {SC.STATION_NAMES.get(v, v)} ({v}) ")
        return v
    return default


# ---------------------------------------------------------------------------
# 主转换
# ---------------------------------------------------------------------------
def adapt(input_path, station, out_dir, a006_path, backup=False, pathfile=None):
    if station not in SC.TEMPLATE_HEADER:
        raise ValueError(f"未知站点: {station} (可选 cn/es/gr) ")
    cfg_header = SC.TEMPLATE_HEADER[station]
    print(f">>> Station adapt: {SC.STATION_NAMES[station]} ({station}), target header {len(cfg_header)} cols")

    # 读源
    import openpyxl
    wb = openpyxl.load_workbook(input_path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        raise ValueError("源文件无 data rows")
    src_header_raw = [str(h).strip() if h is not None else "" for h in rows[0]]
    src_header = [SC.normalize_header(h) for h in src_header_raw]
    print(f"    Source header {len(src_header)} cols,  data rows {len(rows)-1}")

    # A006 查表 (仅 ES 需要) 
    a006_lookup = {}
    if SC.LOOKUP.get(station):
        if not os.path.exists(a006_path):
            print(f"    [WARN] A006 missing: {a006_path}, Spanish names left blank")
        else:
            a006_lookup = read_a006_foreign(a006_path, SC.CACHE_DIR)

    # 预解析: 目标列 -> 源列索引
    src_idx = {col: SC.resolve_source_col(col, src_header) for col in cfg_header}
    # 计算规则源列
    computed_src = {}
    for tgt, src in SC.COMPUTED.get(station, {}).items():
        computed_src[tgt] = SC.resolve_source_col(src, src_header)
    # 翻译规则源列
    translate_src = {}
    for tgt, (src, sl, tl) in SC.TRANSLATE.get(station, {}).items():
        translate_src[tgt] = (SC.resolve_source_col(src, src_header), sl, tl)
    # 查表规则 (匹配键列 + A006 源列名) 
    lookup_info = {}
    for tgt, (a006_col, key_col) in SC.LOOKUP.get(station, {}).items():
        lookup_info[tgt] = (SC.resolve_source_col(key_col, src_header), a006_col)

    # 翻译缓存 (跨次运 row复用) : 命中即不调接口, 大幅减少请求数
    trans_cache = {}
    if translate_src:
        trans_cache = load_trans_cache()
        print(f"  [TRANS] loaded cache {len(trans_cache)} entries")
    # 批量翻译准备: 先逐 row吃缓存, 未命中收集唯一值, 最后整批翻 (请求数从~N降到~N/批大小) 
    pending = []          # ( row号, 目标列, 源key)  待补译
    seen_unique = set()   # 已登记的唯一源key
    unique_keys = []      # 保持首次出现顺序的唯一源key列表
    n_trans = 0           # 本次新译条数
    n_cache_hit = 0       # 本次缓存命中条数
    consec_fail = 0       # 连续失败计数 (判定限流用) 
    cooldown_used = 0     # 已用掉的长冷却次数
    stopped_translate = False

    out_rows = []
    warnings = []  # ( row号, 原文) 翻译失败
    for r_i, row in enumerate(rows[1:], 1):
        cells = ["" if v is None else v for v in row]
        vals = {}
        for col in cfg_header:
            si = src_idx.get(col)
            vals[col] = cells[si] if (si is not None and si < len(cells)) else ""

        # 1) 计算规则: 最低起订量 = 中包每包可装个数
        for tgt, src in SC.COMPUTED.get(station, {}).items():
            si = computed_src.get(tgt)
            v = cells[si] if (si is not None and si < len(cells)) else ""
            vals[tgt] = v

        # 2) 查表规则: 仅当目标为空才填 (不覆盖已有值) 
        for tgt, (key_idx, a006_col) in lookup_info.items():
            if str(vals.get(tgt, "")).strip():
                continue
            key = cells[key_idx] if (key_idx is not None and key_idx < len(cells)) else ""
            key = norm_code(key)
            foreign = a006_lookup.get(key, "")
            if not foreign and key.isdigit():
                foreign = a006_lookup.get(str(int(key)), "")
            vals[tgt] = foreign

        # 3) 翻译规则: 先吃缓存; 未命中登记, 留待循环结束后批量翻译
        for tgt, (src_i, sl, tl) in translate_src.items():
            if str(vals.get(tgt, "")).strip():
                continue
            src_v = cells[src_i] if (src_i is not None and src_i < len(cells)) else ""
            key = str(src_v).strip()
            if not key:
                continue
            if key in trans_cache:
                vals[tgt] = trans_cache[key]
                n_cache_hit += 1
            else:
                if key not in seen_unique:
                    seen_unique.add(key)
                    unique_keys.append(key)
                pending.append((r_i, tgt, key))

        # 翻译进度改为批次后统一打印 (见下方收尾) 

        # 4) 固定填值 (强制覆盖) 
        for col, v in SC.DEFAULTS.get(station, {}).items():
            vals[col] = v

        # 5) 强制留空
        for col in SC.FORCE_EMPTY.get(station, set()):
            vals[col] = ""

        out_rows.append([vals.get(col, "") for col in cfg_header])

    # 翻译收尾: 批量补译未命中项 (整批请求, 远少于逐 row请求) , 再回填到 out_rows
    if translate_src:
        BATCH = 20                                   # 每批条数 (免费端点单批≤~25 较稳) 
        tgt_idx = {tgt: cfg_header.index(tgt) for tgt in translate_src}
        failed_keys = set()                          # 翻译失败的唯一key (用于告警去重) 
        n_unique = len(unique_keys)
        if n_unique:
            print(f"  [TRANS] cache miss {n_unique}  unique names, 按批 (每批 {BATCH})...",
                  flush=True)
            for ci in range(0, n_unique, BATCH):
                chunk = unique_keys[ci:ci + BATCH]
                label = f"{ci+1}-{min(ci+BATCH, n_unique)}/{n_unique}"
                # 疑似被限流: 长冷却后整批重试
                if consec_fail >= MAX_CONSEC_FAIL and cooldown_used < MAX_COOLDOWN_CYCLES:
                    cooldown_used += 1
                    consec_fail = 0
                    print(f"    [TRANS] consecutive failures {MAX_CONSEC_FAIL} times, likely rate-limited;"
                          f"冷却 {COOLDOWN_SECONDS}s 后继续"
                          f" (第 {cooldown_used}/{MAX_COOLDOWN_CYCLES} 轮) ", flush=True)
                    save_trans_cache(trans_cache)  # 先落盘, 中途中断也不丢
                    time.sleep(COOLDOWN_SECONDS)
                got = translate_batch_zh_to_el(chunk)
                time.sleep(TRANS_SLEEP)
                if got is None:
                    consec_fail += 1
                    if consec_fail >= MAX_CONSEC_FAIL and cooldown_used >= MAX_COOLDOWN_CYCLES:
                        stopped_translate = True
                    for k in chunk:
                        failed_keys.add(k)
                    print(f"    [TRANS] batch {label} 失败 (留空, 稍后重跑可补) ", flush=True)
                else:
                    consec_fail = 0
                    ok = 0
                    for k, tr in zip(chunk, got):
                        if tr:
                            trans_cache[k] = tr
                            ok += 1
                        else:
                            failed_keys.add(k)
                    n_trans += ok
                    print(f"    [TRANS] batch {label} 完成 (new {ok} entries) ", flush=True)
        # 回填到 out_rows
        for (r_i, tgt, key) in pending:
            tr = trans_cache.get(key, "")
            if not tr:
                warnings.append((r_i, key))
            out_rows[r_i - 1][tgt_idx[tgt]] = tr
        save_trans_cache(trans_cache)
        print(f"  [TRANS] new {n_trans}  entries / cache hit {n_cache_hit} entries"
              f" (缓存共 {len(trans_cache)} entries已保存) ")
        if stopped_translate:
            print(f"  [WARN] 免费翻译端点持续失败 (已长冷却 {cooldown_used} rounds),"
                  f"已停止翻译剩余 row.相关单元格留空, 可事后手填; "
                  f"重跑会命中缓存继续补译 (不需重头再来) .")

    # 写输出
    os.makedirs(out_dir, exist_ok=True)
    stamp = time.strftime("%Y%m%d_%H%M%S")
    base = os.path.splitext(os.path.basename(input_path))[0]
    out_name = f"{station}_{base}_{stamp}.xlsx"
    out_path = os.path.join(out_dir, out_name)
    from openpyxl import Workbook
    wb2 = Workbook(); ws2 = wb2.active; ws2.title = "Sheet1"
    ws2.append(cfg_header)
    for r in out_rows:
        ws2.append(r)
    wb2.save(out_path)
    print(f"    Generated: {out_path} ({len(out_rows)}  rows)")

    # 供 bat 串联: 把生成文件的绝对路径写到 --pathfile, bat 用 set /p 取回
    if pathfile:
        try:
            with open(pathfile, "w", encoding="utf-8") as f:
                f.write(os.path.abspath(out_path))
        except Exception as e:
            print(f"  [WARN] 写出 pathfile 失败 (不影响生成) : {e}")

    # 备份到 已导入表格/<站>/ (默认关闭, 改由上传成功后备份) 
    bak_path = None
    if backup:
        bdir = SC.BACKUP_DIR[station]
        os.makedirs(bdir, exist_ok=True)
        bak_path = os.path.join(bdir, out_name)
        shutil.copy2(out_path, bak_path)
        print(f"    Backed up: {bak_path}")

    if warnings:
        print(f"  [WARN] {len(warnings)}  rows Greek translation failed (已留空, 可事后手填) : ")
        for r_i, src_v in warnings[:20]:
            print(f"       row{r_i}: {src_v}")
    return out_path, len(out_rows), warnings


def main():
    ap = argparse.ArgumentParser(description="站点适配转换器 (cn/es/gr) ")
    ap.add_argument("--station", choices=["cn", "es", "gr"],
                    help="目标站点; 不传则弹窗选择 (默认中国站) ")
    ap.add_argument("--input", required=True, help="源 xlsx (China 形态匹配结果, 或已按模板填好的文件) ")
    ap.add_argument("--out", default=DEFAULT_OUT, help="输出目录 (默认 货号匹配器/输出) ")
    ap.add_argument("--a006", default=DEFAULT_A006, help="A006 路径 (西语品名查表用) ")
    ap.add_argument("--backup", action="store_true",
                    help="生成时即备份到 已导入表格/<站>/ (默认不备份, 改由上传成功后备份) ")
    ap.add_argument("--pathfile", help="把生成文件的绝对路径写入该文件, 供 bat 串联读取")
    args = ap.parse_args()

    station = args.station
    if not station:
        station = pick_station_via_gui()

    try:
        out_path, n, warns = adapt(args.input, station,
                                   args.out, args.a006,
                                   backup=args.backup, pathfile=args.pathfile)
        print(f"\nDone: {SC.STATION_NAMES[station]} upload table generated ({n}  rows)")
        if warns:
            print(f"   Note: {len(warns)}  rows translation failed (见上) , 其余正常.")
        sys.exit(0)
    except Exception as e:
        import traceback
        print(f"FATAL: {e}")
        traceback.print_exc()
        sys.exit(2)


if __name__ == "__main__":
    main()
