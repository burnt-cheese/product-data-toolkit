# -*- coding: utf-8 -*-
"""
货号匹配自动化 (可移植项目版) 
================================
项目结构 (整个文件夹可拷到其他电脑) : 
  sku_matcher/
  ├── scripts/generate_matched.py   本脚本
  ├── data_source/A006-货号统计（全部）.xlsx   SAP 导出表 (会更新, 替换同名文件即可) 
  ├── data_source/商品导入模板.xlsx            模板 (仅取 34 列表头布局) 
  ├── input/                             把要匹配的货号表放这里
  ├── output/                             结果文件自动输出到这里
  ├── setup.bat                        其他电脑: 双击一键建 .venv + 装依赖
  └── 商品匹配.bat                     把货号表拖到它上面即自动出表

用法: 
  python generate_matched.py                     # 默认读 input/货号信息.xlsx
  python generate_matched.py "某货号表.xlsx"      # 指定输入文件 (任意路径) 
  python generate_matched.py "xx.xlsx" --sap "A006.xlsx" --tpl "模板.xlsx"  # 覆盖数据源

输入表约定: 
  - 必须有表头, 且某一列表头含「货号」二字 (否则取第 1 列) 
  - 可选: 某列表头含「客户」/customer, 则按客户分多个 sheet
"""
import re, zipfile, sys, os, glob, time, pickle, hashlib
try:
    from lxml import etree as _ET
    _USING_LXML = True
except ImportError:                                  # 兜底用 Python 标准库
    import xml.etree.ElementTree as _ET
    _USING_LXML = False
from pathlib import Path
from decimal import Decimal, ROUND_UP
from openpyxl import load_workbook, Workbook

# ===================== 项目内路径 (自动定位, 无需修改) =====================
BASE    = Path(__file__).resolve().parent.parent          # 项目根
DATA    = BASE / "data_source"
IN_DIR  = BASE / "input"
OUT_DIR = BASE / "output"
OUT_DIR.mkdir(parents=True, exist_ok=True)

SAP  = DATA / "A006-货号统计（全部）.xlsx"
TPL  = DATA / "商品导入模板.xlsx"
OUT  = OUT_DIR / "sku_match_result.xlsx"

# ---------- .xls 老格式(BIFF)自动转换 ----------
# 放在顶层代码 (INPUT 解析) 之前, 因为该代码会在模块加载时立即执行并调用本函数.
# openpyxl 不支持 .xls, 先用 xlrd 读, 再用 openpyxl 写出标准 .xlsx, 
# 之后走与 .xlsx 完全相同的流程.日期->YYYY-MM-DD 字符串, 布尔->bool, 数字原样保留
#  (norm_key 会把 44969.0 归一成 44969, 货号不会因浮点丢失) .
def _convert_xls_to_xlsx(src, dst):
    import xlrd
    from openpyxl import Workbook
    rb = xlrd.open_workbook(str(src))
    wb = Workbook()
    wb.remove(wb.active)
    for sh in rb.sheet_names():
        rs = rb.sheet_by_name(sh)
        ws = wb.create_sheet(title=(sh[:31] or "Sheet1"))
        for r in range(rs.nrows):
            row = []
            for c in range(rs.ncols):
                v = rs.cell_value(r, c)
                ct = rs.cell_type(r, c)
                if ct == xlrd.XL_CELL_DATE:
                    v = xlrd.xldate_as_datetime(v, rb.datemode).strftime("%Y-%m-%d")
                elif ct == xlrd.XL_CELL_BOOLEAN:
                    v = bool(v)
                row.append(v)
            ws.append(row)
    wb.save(str(dst))

# 解析命令行参数: 位置参数=输入货号表; --sap/--tpl 可覆盖数据源
_input_arg = None
_keep_arg = None
_args = sys.argv[1:]
for i, a in enumerate(_args):
    if a == "--sap" and i + 1 < len(_args):
        SAP = Path(_args[i + 1])
    elif a == "--tpl" and i + 1 < len(_args):
        TPL = Path(_args[i + 1])
    elif a == "--keep" and i + 1 < len(_args):
        _keep_arg = _args[i + 1]          # 逗号分隔的列名, 跳过弹窗 (脚本化/批量用) 
    elif not a.startswith("-") and _input_arg is None:
        _input_arg = Path(a)
INPUT = _input_arg if _input_arg is not None else IN_DIR / "货号信息.xlsx"

# 输入文件存在性检查 (拖入即用时给出友好提示) 
if not os.path.exists(INPUT):
    print(f"[ERROR] Input goods-code table not found: {INPUT}")
    print("      用法1: python generate_matched.py \"货号表.xlsx\"")
    print("      Usage2: drag the goods-code file onto 货号匹配-拖入即用.bat")
    sys.exit(1)

# .xls 老格式自动转换: openpyxl 不支持 .xls, 先转成临时 .xlsx 再走正常流程
if INPUT.suffix.lower() == ".xls":
    try:
        import xlrd  # noqa: F401
    except ImportError:
        print("[ERROR] Input is legacy .xls but xlrd is missing.")
        print("      Run in project .venv: python -m pip install xlrd")
        sys.exit(1)
    import tempfile
    _tmpdir = Path(tempfile.mkdtemp(prefix="xls2xlsx_"))
    _converted = _tmpdir / (INPUT.stem + ".xlsx")
    print(f">>> 检测到 .xls 老格式, 自动转换为 .xlsx 以继续: {INPUT.name} -> {_converted.name}")
    _convert_xls_to_xlsx(INPUT, _converted)
    INPUT = _converted
if not os.path.exists(SAP):
    print(f"[ERROR] SAP source not found: {SAP}")
    print("      Put the SAP export (A006-货号统计（全部）.xlsx) into the data_source folder")
    sys.exit(1)
if not os.path.exists(TPL):
    print(f"[ERROR] Product import template not found: {TPL}")
    print("      Put 商品导入模板.xlsx into the data_source folder")
    sys.exit(1)

# 折扣白名单 (供应商名称, 命中则「是否可以打折」=N)
# 【公开版】真实供应商名已替换为占位符 —— 请填入你自己 SAP 里的「业务伙伴名称」原文
WHITE = {"示例供应商A", "示例供应商B", "示例供应商C", "示例供应商D", "示例供应商E"}
WL_CONST = "{" + ";".join(f'"{w}"' for w in WHITE) + "}"

ILLEGAL     = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')     # str 模式 (兼容旧用法) 
ILLEGAL_B   = re.compile(rb'[\x00-\x08\x0b\x0c\x0e-\x1f]')    # bytes 模式 (SAP 26MB+ xlsx 直接 z.read()) 

# ---------- 工具 ----------
def col_letter(ref):
    return re.match(r'([A-Z]+)', ref).group(1)

def norm_key(v):
    if v is None: return ""
    if isinstance(v, bool): return str(v)
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else str(v)
    if isinstance(v, int): return str(v)
    return str(v).strip()

# ---------- 容错读 SAP (自动清洗垂直制表符等非法 XML 字符) ----------
# 缓存策略: SAP 文件 26MB+, 解析需 20+ 秒, 将结果 pickle 到 data_source/.sap_cache/.
# 缓存键 = (文件名, mtime, size), SAP 内容有更新时自动失效.
_CACHE_DIR = DATA / ".sap_cache"
_CACHE_DIR.mkdir(parents=True, exist_ok=True)

def _sap_cache_key(p):
    s = os.stat(p)
    return (Path(p).name, int(s.st_mtime), int(s.st_size))

def _sap_cache_path(p):
    name, mt, sz = _sap_cache_key(p)
    h = hashlib.md5(f"{name}|{mt}|{sz}".encode("utf-8")).hexdigest()[:12]
    return _CACHE_DIR / f"sap_{h}.pkl"

def _load_sap_cache(p):
    cp = _sap_cache_path(p)
    if not cp.exists():
        return None
    try:
        with open(cp, "rb") as f:
            data = pickle.load(f)
        if isinstance(data, dict) and data.get("key") == _sap_cache_key(p):
            return data
    except Exception:
        try: cp.unlink()
        except Exception: pass
    return None

def _save_sap_cache(p, lookup, dup):
    cp = _sap_cache_path(p)
    data = {"key": _sap_cache_key(p), "lookup": lookup, "dup": dup,
            "saved_at": int(time.time()),
            "backend": "lxml" if _USING_LXML else "ElementTree"}
    tmp = cp.with_suffix(".tmp")
    try:
        with open(tmp, "wb") as f:
            pickle.dump(data, f, protocol=pickle.HIGHEST_PROTOCOL)
        os.replace(tmp, cp)
        return True
    except Exception as e:
        print(f"    [WARN] Cache write failed: {e}")
        try: tmp.unlink()
        except Exception: pass
        return False

def read_sap_lookup(path):
    # 1) 缓存命中 -> 直接返回 (~0.5s) 
    cached = _load_sap_cache(path)
    if cached is not None:
        sz_mb = os.path.getsize(_sap_cache_path(path)) / 1024 / 1024
        print(f"    Loaded SAP from cache ({len(cached['lookup'])}  goods-codes, {cached.get('backend','?')}, "
              f"缓存 {sz_mb:.1f}MB, "
              f"建于 {time.strftime('%Y-%m-%d %H:%M', time.localtime(cached.get('saved_at',0)))}) ")
        return cached["lookup"], cached.get("dup", 0)

    # 2) 缓存未命中 -> 解析 (首次约 8~30s, 依赖 lxml 是否可用) 
    t0 = time.time()
    backend = "lxml" if _USING_LXML else "ElementTree"
    print(f"    First parse of SAP ({backend}) , 大文件需要十几秒至几十秒 ...")
    z = zipfile.ZipFile(path)

    # sharedStrings.xml (22MB) 
    t1 = time.time()
    ss_bytes = ILLEGAL_B.sub(b"", z.read("xl/sharedStrings.xml"))
    ss_bytes = re.sub(rb'\sxmlns="[^"]+"', b"", ss_bytes, count=1)
    ss_root = _ET.fromstring(ss_bytes)
    shared = ["".join(t.text or "" for t in si.iter("t"))
              for si in ss_root.iter("si")]
    del ss_bytes, ss_root
    print(f"    sharedStrings parsed ({len(shared)} entries): {time.time()-t1:.1f}s")

    # sheet1.xml (234MB! 瓶颈) 
    t2 = time.time()
    ws_bytes = ILLEGAL_B.sub(b"", z.read("xl/worksheets/sheet1.xml"))
    ws_bytes = re.sub(rb'\sxmlns="[^"]+"', b"", ws_bytes, count=1)
    wroot = _ET.fromstring(ws_bytes)
    del ws_bytes
    print(f"    sheet1 XML parsed: {time.time()-t2:.1f}s (cumulative {time.time()-t0:.1f}s) ")

    # 找所有 row: lxml 去掉 xmlns 后默认 ns 为空, 可用 './/row'; ET 必须 'sheetData/row'
    if _USING_LXML:
        rows = wroot.findall(".//row")
    else:
        rows = wroot.findall("sheetData/row")

    needed = ["货号","条形码","品名","外文名称","首选供应商","业务伙伴名称","装箱量","采购单价",
              "单箱重量","单箱体积","长度","宽度","高度","中包方式","中包数量",
              "产品材质","外箱条码","产品详情","产品分类","二级物料分类","三级物料分类"]

    hdr = {}
    for c in rows[0].findall("c"):
        ref = c.get("r") or ""
        col = "".join(L for L in ref if L.isalpha())
        t = c.get("t"); v = c.find("v")
        val = v.text if v is not None else ""
        if t == "s" and val != "":
            try: val = shared[int(val)]
            except (ValueError, IndexError): pass
        hdr[col] = val

    lookup = {}; dup = 0
    data_rows = rows[1:]
    n = len(data_rows)
    t3 = time.time()
    # 流式进度: 每 10% 打印一次, 避免刷屏
    next_pct = 10
    for i, row in enumerate(data_rows, start=1):
        cells = {}; hv = None
        for c in row.findall("c"):
            ref = c.get("r") or ""
            col = "".join(L for L in ref if L.isalpha())
            t = c.get("t"); v = c.find("v")
            val = v.text if v is not None else ""
            if t == "s" and val != "":
                try: val = shared[int(val)]
                except (ValueError, IndexError): pass
            if col == "B": hv = val
            if col in hdr and hdr[col] in needed:
                cells[hdr[col]] = val
        if hv in (None, ""): continue
        k = norm_key(hv)
        if k in lookup: dup += 1
        lookup[k] = cells

        # 进度: 每 10% / 每 1s 打印一次 (不强刷屏) 
        pct = i * 100 // n if n else 100
        if pct >= next_pct:
            print(f"    Building lookup: {pct:3d}%  ({i}/{n} rows, "
                  f"已用 {time.time()-t3:.1f}s, 累计 {time.time()-t0:.1f}s) ", flush=True)
            while next_pct <= pct: next_pct += 10

    print(f"    Lookup built: {time.time()-t3:.1f}s")

    # 3) 保存缓存
    print(f"    Saving cache ...")
    _save_sap_cache(path, lookup, dup)
    print(f"    SAP parse total: {time.time()-t0:.1f}s  -> 货号 {len(lookup)} 个, 重复 {dup}")
    return lookup, dup

# ---------- 列映射 (输出列 -> SAP字段) ----------
MAP = {
    1:("src","品名"), 3:("src","条形码"), 4:("src","产品分类"),
    5:("src","二级物料分类"), 6:("src","三级物料分类"), 11:("src","采购单价"),
    12:("src","单箱重量"), 13:("src","单箱体积"), 14:("src","长度"),
    15:("src","宽度"), 16:("src","高度"), 17:("src","中包方式"),
    18:("src","中包数量"), 19:("src","装箱量"), 23:("src","装箱量"),
    28:("src","产品详情"), 29:("src","产品材质"), 32:("src","首选供应商"),
    33:("src","业务伙伴名称"),
}
# 保留为真正的 Excel 公式 (打开即实时计算) 
FORMULA_COLS = {
    25: lambda r: f'=IF(ISNUMBER(MATCH(AG{r},{WL_CONST},0)),"N","Y")',          # 是否可以打折
    10: lambda r: f'=IF(Y{r}="Y",ROUNDUP(ROUNDUP(K{r}/0.97,2)/0.88,2),ROUNDUP(K{r}/0.97,2))',  # 批发价格
    26: lambda r: f'=IF(Y{r}="Y","示例品牌",AG{r})',                            # 品牌
}
NUMERIC={"采购单价","单箱重量","单箱体积","长度","宽度","高度","中包数量","装箱量"}
BLANK_COLS={27, 34}          # AA 标签, AH 不填
FIXED={22:100, 24:"CNY"}     # V 销量=100, X 币种=CNY

# ---------- 「保留输入列」: 弹窗选择 + 列名Matched ----------
# 用户输入表里某些列 (如 品名, 单价) 比 A006 准确, 勾选后该列直接取自输入表, 
# 不再去 A006 匹配.仅覆盖「输出模板中已有的对应列」, 避免破坏 34 列导入结构.
def norm_header(s):
    """归一化表头名用于匹配: 去空格, 去括号及单位, 去'单箱/每箱/每包'等前缀."""
    s = str(s).lower()
    s = s.replace(" (", "(").replace(") ", ")")
    s = re.sub(r'\([^)]*\)', '', s)                    # 去掉括号内容, 如 (kg)
    s = re.sub(r'[　\s]', '', s)                       # 去掉所有空白 (含全角空格) 
    s = re.sub(r'(kg|cm|m³|³|m2|²)', '', s)           # 去掉单位
    s = s.replace("单箱","").replace("每箱","").replace("每包","").replace("每托","")
    return s

# 输入表头归一化名 -> 输出模板表头名 (用于输入列名与输出列名不完全一致时) 
_KEEP_SYNONYM = {
    "条形码": "条形码", "条码号": "条形码", "条码": "条形码",
    "采购单价": "采购单价", "单价": "采购单价", "进货价": "采购单价",
    "二级物料分类": "二级分类", "三级物料分类": "三级分类",
    "中包数量": "中包每包可装个数",
    "装箱量": "单箱每箱装箱个数",
    "产品材质": "产品属性",
    "毛重": "单箱重量(kg)", "净重": "单箱重量(kg)",
    "供应商编码": "供应商编码", "首选供应商": "供应商编码",
    "供应商名称": "供应商名称", "业务伙伴名称": "供应商名称",
}
# 数字型输出列 (覆盖时顺便转数值) 
NUMERIC_OUT = {col for col,(kind,field) in MAP.items() if field in NUMERIC}

def match_output_col(input_header, out_headers):
    """把输入列表头映射到输出模板列号; 无对应列返回 None."""
    norm = norm_header(input_header)
    for i, h in enumerate(out_headers, 1):           # 1) 精确归一化匹配
        if h is not None and norm_header(h) == norm:
            return i
    syn = _KEEP_SYNONYM.get(norm)                    # 2) 同义词表
    if syn:
        for i, h in enumerate(out_headers, 1):
            if h is not None and norm_header(h) == norm_header(syn):
                return i
    return None

def to_num(v):
    if v is None or v=="":
        return v
    if isinstance(v,(int,float)):
        return v
    s=str(v).strip().replace(",","")
    try:
        f=float(s)
        return int(f) if f.is_integer() else f
    except ValueError:
        return v

# ---------- 货号列 / 客户列 自动识别 ----------
def _is_code_like(v):
    """判断单元格值是否像'货号/编码': 多为数字或短字母数字, 不含中文."""
    if v is None or v == "":
        return False
    s = str(v).strip()
    if not s:
        return False
    if any('\u4e00' <= ch <= '\u9fff' for ch in s):   # 含中文 -> 不像货号
        return False
    core = re.sub(r'[\s\-_./]', '', s)                 # 去掉常见分隔符
    return bool(core) and bool(re.fullmatch(r'[A-Za-z0-9]+', core))

def _score_huohao_header(hs):
    """给'货号列'表头打分: 0 表示不像货号列 (会排除条码/名称/价格等) ."""
    if hs is None:
        return 0
    s = str(hs).lower()
    if any(k in s for k in ("条码", "bar", "名称", "品名", "描述", "desc",
                            "规格", "颜色", "数量", "价格", "金额", "重量", "体积")):
        return 0
    if "货号" in str(hs): return 100
    if "sku" in s:        return 95
    if "物料编码" in str(hs) or "物料号" in str(hs): return 90
    if any(k in str(hs) for k in ("商品编码","商品编号","产品编码","产品编号",
                                  "货品编码","货品编号","产品货号","商品货号")): return 85
    if "编号" in str(hs): return 70
    if "编码" in str(hs): return 60
    if "型号" in str(hs): return 55
    if "item" in s:       return 50
    if "code" in s:       return 45
    if "no" in s:         return 40
    return 0

def _score_cust_header(hs):
    if hs is None:
        return 0
    s = str(hs).lower()
    if "客户" in str(hs): return 100
    if "customer" in s:   return 95
    if "买手" in str(hs) or "采购员" in str(hs): return 60
    if "国家" in str(hs) or "country" in s:        return 50
    return 0

def _header_row_score(ws, rr):
    """给某一行打分: 越像'货号明细表头行'分数越高.

    兼容订货合同这类'表头不在第 1 rows, 前面有公司抬头, 后面有条款备注'的表格.
    重点奖励含'货号'二字 (+100) , 其余表头关键词 (客户/品名/条码/单价…) 加分.
    """
    score = 0
    for c in range(1, ws.max_column + 1):
        h = ws.cell(rr, c).value
        if h is None:
            continue
        hs = str(h)
        if "货号" in hs:
            score += 100
        elif _score_huohao_header(h) > 0:
            score += 10
        if "客户" in hs or "customer" in hs.lower():
            score += 20
        if any(k in hs for k in ("品名", "条码", "单价", "数量", "金额",
                                 "规格", "装箱", "件数", "型号", "材质")):
            score += 5
    return score

def _detect_cols(ws):
    """返回 (hcol, ccol, notes, has_header, header_row).自动识别货号列与客户列: 
    - 第一步: 在前 30 rows内寻找'含货号表头'的那一行 (兼容合同类等表头不在第 1 rows) ; 
    - 找到后在该行内定位货号列/客户列, 数据从该行下一行读起; 
    - 找不到货号表头行 -> 回退原逻辑 (第 1 rowsHeader / 无表头数据特征推断) , header_row=1.
    notes 始终说明最终选用了哪一列, 便于核对.
    """
    notes = []
    # ---- 第一步: 定位货号表头行 (兼容表头不在第 1 rows) ----
    hdr_row = None
    best = 0
    for rr in range(1, min(ws.max_row, 30) + 1):
        sc = _header_row_score(ws, rr)
        if sc > best:
            best, hdr_row = sc, rr
    if hdr_row is not None and best >= 100:   # 确有含'货号'的表头行
        hcol = ccol = None
        best_h = best_c = 0
        for c in range(1, ws.max_column + 1):
            h = ws.cell(hdr_row, c).value
            if h is None:
                continue
            sc = _score_huohao_header(h)
            if sc > best_h:
                best_h, hcol = sc, c
            sc2 = _score_cust_header(h)
            if sc2 > best_c:
                best_c, ccol = sc2, c
        htxt = ws.cell(hdr_row, hcol).value
        notes.append(f"货号列: 第 {hcol} 列 (第 {hdr_row} rows表头'{htxt}') ")
        if ccol is not None:
            ctxt = ws.cell(hdr_row, ccol).value
            notes.append(f"客户列: 第 {ccol} 列 (表头'{ctxt}') ")
        return hcol, ccol, notes, True, hdr_row

    # ---- 第二步: 回退原逻辑 (第 1 rowsHeader / 无表头推断) ----
    hcol = ccol = None
    best_h = best_c = 0
    for c in range(1, ws.max_column + 1):
        h = ws.cell(1, c).value
        if h is None:
            continue
        sc = _score_huohao_header(h)
        if sc > best_h:
            best_h, hcol = sc, c
        sc2 = _score_cust_header(h)
        if sc2 > best_c:
            best_c, ccol = sc2, c
    if hcol is None:
        if ws.max_column == 1:
            hcol = 1
            notes.append("未找到货号表头且文件仅 1 列 -> 使用第 1 列作为货号列")
        else:
            best_score, pick = -1.0, None
            for c in range(1, ws.max_column + 1):
                sample = [ws.cell(r, c).value for r in range(2, min(ws.max_row, 50) + 1)]
                sample = [x for x in sample if x not in (None, "")]
                if not sample:
                    continue
                code = sum(1 for x in sample if _is_code_like(x))
                intlike = sum(1 for x in sample
                              if _is_code_like(x) and '.' not in str(x))
                # 兼顾「像编码」与「像整数编码(无小数点)」, 避免误选价格列
                score = 0.5 * (code / len(sample)) + 0.5 * (intlike / len(sample))
                if score > best_score:
                    best_score, pick = score, c
            if pick and best_score >= 0.6:
                hcol = pick
                notes.append(f"未找到货号表头, 按数据特征自动选用第 {pick} 列作为货号列"
                             f" (编码特征 {best_score*100:.0f}%) ")
            else:
                hcol = 1
                notes.append("未找到货号表头且无法判定编码列 -> 默认使用第 1 列"
                             " (如不对, 请给货号列加含'货号'的表头) ")
    # 是否有表头行: 找到过表头关键词, 且货号列首行不像真实数据 (含数字的编码) 
    has_header = (best_h > 0 or best_c > 0)
    if hcol is not None:
        first = ws.cell(1, hcol).value
        if _is_code_like(first) and re.search(r'\d', str(first)):
            has_header = False   # 首行就是编码数据 -> 视为无表头, 从第 1 rows读

    # 始终说明最终选用的列, 便于核对
    htxt = ws.cell(1, hcol).value
    if has_header and htxt is not None and str(htxt).strip() != "":
        notes.append(f"货号列: 第 {hcol} 列 (表头'{htxt}') ")
    else:
        notes.append(f"货号列: 第 {hcol} 列 (无表头, 按数据特征推断) ")
    if ccol is not None:
        ctxt = ws.cell(1, ccol).value
        notes.append(f"客户列: 第 {ccol} 列 (表头'{ctxt}') ")
    return hcol, ccol, notes, has_header, 1

# ---------- 读入「含货号的表格」----------
def read_huohao_list(path):
    """读取货号表 -> (out, dup_codes, input_cols).

    out: 列表 [(原始值, 归一化key, 客户, 该行所有列的值dict)].
    严格「一行输入 = 一行输出」: 不做去重, 避免静默丢行.
    自动识别货号列 (支持多列源表 / 无表头推断) , 处理前先剔除无数据行保持整洁.
    input_cols: 主表 (含货号表头的那张) 的「全部列」清单 [(列号, 表头名), ...], 
                供弹窗让用户勾选「要保留的输入列」 (后续会排除货号列) .
    """
    wb = load_workbook(path, data_only=True)
    out=[]; seen=set(); dup_codes=set()
    skipped=0
    input_cols=[]; cols_captured=False
    for ws in wb.worksheets:
        # 跳过完全是空的表 (如 .xls 转换后残留的 Sheet2/Sheet3) , 避免误报「剔除无数据行」
        if ws.max_row <= 1 and ws.max_column <= 1:
            continue
        hcol, ccol, notes, has_header, header_row = _detect_cols(ws)
        for nt in notes:
            print(f"    [COL] {nt}")
        if not cols_captured and hcol is not None:   # 仅从含货号表头的主表取列清单
            input_cols = [(c, ws.cell(header_row, c).value)
                          for c in range(1, ws.max_column + 1) if c != hcol]
            cols_captured = True
        if has_header:
            start = header_row + 1       # 表头行的下一行开始读
        else:
            start = 1                    # 无表头: 从第 1 rows读
        blank_run = 0                    # 连续空货号行计数, 用于截断明细表后的条款/备注
        for r in range(start, ws.max_row + 1):
            v = ws.cell(r, hcol).value
            nk = norm_key(v)
            if nk == "":                 # 货号空/纯空白 -> 无数据行
                skipped += 1
                blank_run += 1
                if blank_run >= 5:      # 连续 5 rows无货号 -> 视为明细表已结束, 停止
                    break
                continue
            blank_run = 0
            cust = ws.cell(r, ccol).value if ccol else None
            extra = {c: ws.cell(r, c).value for c in range(1, ws.max_column + 1)}
            if nk in seen:
                dup_codes.add(nk)   # 该货号出现多次 -> 记为重复项 (仍保留这一行) 
            else:
                seen.add(nk)
            out.append((v, nk, cust, extra))
    if skipped:
        print(f"    [数据整洁] Dropped {skipped} empty rows (blank goods-code); matching {len(out)} rows")
    return out, dup_codes, input_cols

# ---------- 弹窗: 选择要保留的输入列 (不取 A006) ----------
def _ask_keep_columns(input_cols):
    """弹窗让用户勾选'需要原样保留, 不取 A006 的输入列'.

    input_cols: [(列号, 表头名), ...] (已排除货号列) .
    返回勾选的列号列表 (空列表 = 不保留任何列, 全部按 A006 匹配) .
    无法加载 GUI 时降级为「不保留」 (避免在非交互环境卡死) .

    实现说明: 受管 Python 3.13 venv 不带 tkinter, 统一改成调用同目录的
    `keep_cols_dialog.ps1` (PowerShell WinForms) .任何路径异常都降级
    为「不保留」, 绝不丢行——与原 tkinter 实现的容错策略一致.
    """
    if not input_cols:
        return []

    import subprocess, json, os
    here = os.path.dirname(os.path.abspath(__file__))
    ps1 = os.path.join(here, 'keep_cols_dialog.ps1')
    if not os.path.exists(ps1):
        print(f"    [NOTE] Picker script not found {ps1}, skipped keep-columns; all use A006.")
        return []

    # 临时文件放在系统临时目录的子文件夹里 (用 PID+随机数区分并发) 
    import tempfile, secrets
    tmp_dir = os.path.join(tempfile.gettempdir(), 'keep_cols')
    try:
        os.makedirs(tmp_dir, exist_ok=True)
    except Exception:
        tmp_dir = tempfile.gettempdir()
    safe = f"keep_{os.getpid()}_{secrets.token_hex(3)}"
    in_path  = os.path.join(tmp_dir, safe + '.in.json')
    out_path = os.path.join(tmp_dir, safe + '.out.json')

    try:
        with open(in_path, 'w', encoding='utf-8') as f:
            json.dump([[int(c[0]), str(c[1])] for c in input_cols],
                      f, ensure_ascii=False)
    except Exception as e:
        print(f"    [NOTE] Failed to write column list ({e}) , skipped keep-columns; all use A006.")
        return []

    cmd = [
        'powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass',
        '-File', ps1, in_path, out_path,
    ]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True,
                           encoding='utf-8', errors='replace', timeout=600)
    except subprocess.TimeoutExpired:
        print("    [提示] GUI 弹窗超时 (10 分钟未操作) , skipped keep-columns; all use A006.")
        return []
    except Exception as e:
        print(f"    [NOTE] GUI launch failed ({e}) , skipped keep-columns; all use A006.")
        return []

    if r.returncode != 0:
        err = (r.stderr or r.stdout or '').strip().splitlines()[-3:]  # 末三行足够
        print(f"    [NOTE] GUI returned error (exit={r.returncode}) , skipped keep-columns; all use A006.")
        for ln in err:
            print(f"        {ln}")
        return []

    if not os.path.exists(out_path):
        print("    [NOTE] GUI produced no result; all use A006.")
        return []

    try:
        with open(out_path, 'r', encoding='utf-8-sig') as f:
            result = json.load(f)
    except Exception as e:
        print(f"    [NOTE] Failed to parse GUI result ({e}) , 全部按 A006 匹配.")
        return []

    if not isinstance(result, list):
        print("    [NOTE] GUI result malformed; all use A006.")
        return []
    out = []
    for x in result:
        try:
            out.append(int(x))
        except (TypeError, ValueError):
            pass
    return out

# ============================ 主流程 ============================
# 1) 先读模板表头布局 (快速; 弹窗列名匹配要用到) 
print(">>> Reading template header layout (target sheet only, skip the 220k-row side sheet) ...")
# 关键: 模板里"Sheet1"有 224513 rows, openpyxl 默认全量加载会拖到 30s+.
# 用 read_only=True + data_only=False 只取"洪都拉斯孙伯总"的第 1 rows表头, ~4 秒.
wb_t = load_workbook(TPL, data_only=False, read_only=True)
tpl_ws = wb_t["洪都拉斯孙伯总"]
HEADERS = []
for row in tpl_ws.iter_rows(min_row=1, max_row=1, values_only=True):
    HEADERS = list(row)
    break
MAXC = len(HEADERS)
wb_t.close()
# read_only 模式无法访问 column_dimensions; 模板里也没自定义 cols, 
# 输出 xlsx 会自动按内容长度自适应 (或维持 openpyxl 写入时的默认 13) .
print(f"    Header {MAXC} cols (read_only, default widths)")

# 2) 读取输入表 (快速) -> 拿到列清单, 弹窗让你选要保留的列
print(">>> 读取货号文件:", os.path.basename(INPUT))
items, dup_codes, input_cols = read_huohao_list(INPUT)
print(f"    Goods-code rows: {len(items)}")
if dup_codes:
    # 重复项提示: 仅在命令行显示「有几项 + 具体货号」, 结果表中不做任何标记
    print(f"    WARNING: input has {len(dup_codes)} duplicate codes (kept as-is, not flagged in result):")
    print("       " + ", ".join(sorted(dup_codes)))
else:
    print("    No duplicate codes")

# 3) 决定「保留列」: 优先命令行 --keep (脚本化/批量) , 否则弹窗交互
if _keep_arg:
    keep_names = [x.strip() for x in _keep_arg.split(",") if x.strip()]
    name_to_col = {str(h).strip(): c for c, h in input_cols if h}
    keep_input_cols = [name_to_col[n] for n in keep_names if n in name_to_col]
    picked = [n for n in keep_names if n in name_to_col]
    missing = [n for n in keep_names if n not in name_to_col]
    print(f">>> Keep-cols from --keep: {picked}")
    if missing:
        print(f"    [警告] These columns do not exist in input, ignored: {missing}")
else:
    print(">>> Popup: pick input columns to keep as-is (otherwise filled by A006) ...")
    keep_input_cols = _ask_keep_columns(input_cols)
KEEP_MAP = {}   # 输入列号 -> 输出列号
for ci in keep_input_cols:
    hname = dict(input_cols).get(ci)
    oc = match_output_col(hname, HEADERS) if hname is not None else None
    if oc:
        KEEP_MAP[ci] = oc
        print(f"    [保留列] '{hname}'-> 输出第 {oc} 列'{HEADERS[oc-1]}' (from input, not A006)")
    else:
        print(f"    [保留列] '{hname}'在输出模板中无对应列, 已忽略 (keeps 34-col structure)")
if not KEEP_MAP:
    print("    (No keep-cols selected; all fields from A006)")

# 4) 读取并清洗 SAP (慢, 带缓存) 
print(">>> Reading and cleaning SAP file ...")
sap_lookup, dup = read_sap_lookup(SAP)
print(f"    SAP goods-code count: {len(sap_lookup)}, duplicates: {dup}")

# 5) 按客户分组 (无客户列则归入「匹配结果」) + 生成
groups={}
for orig,nk,cust,extra in items:
    key=cust if cust else "匹配结果"
    groups.setdefault(key, []).append((orig,nk,extra))

def _apply_keep(ows, i, extra):
    """把勾选的保留列从输入表原样写入输出 (覆盖 A006 值) ."""
    for in_c, out_c in KEEP_MAP.items():
        v = extra.get(in_c)
        if v in (None, ""):
            continue
        ows.cell(i, out_c, to_num(v) if out_c in NUMERIC_OUT else v)

out_wb=Workbook(); out_wb.remove(out_wb.active)
unmatched=[]
total_match=0; total_unmatch=0

for gname, glist in groups.items():
    title = str(gname)[:31]
    ows=out_wb.create_sheet(title=title)
    # 表头 (read_only 模式拿不到 column_dimensions, 列宽用 Excel 默认即可) 
    for c in range(1, MAXC+1):
        ows.cell(1,c, HEADERS[c-1] if c-1 < len(HEADERS) else None)
    mcnt=0; ucnt=0
    for i,(orig,nk,extra) in enumerate(glist, start=2):
        ows.cell(i,2, orig)                       # B 货号
        rec=sap_lookup.get(nk)
        if rec is None:
            ucnt+=1
            unmatched.append((title, i, orig))
            # 未匹配: 清空除货号外的所有列
            for c in range(1, MAXC+1):
                if c!=2: ows.cell(i,c).value=None
            _apply_keep(ows, i, extra)            # 即便未匹配也保留勾选列
            continue
        mcnt+=1
        # 源字段列 (分类 D/E/F 直接取 SAP 原分类, 不做修正/核对) 
        for col,(kind,field) in MAP.items():
            v=rec.get(field)
            ows.cell(i,col, to_num(v) if field in NUMERIC else v)
        # 固定值
        for col,val in FIXED.items():
            ows.cell(i,col, val)
        # 不填列
        for col in BLANK_COLS:
            ows.cell(i,col).value=None
        # 公式列
        for col in FORMULA_COLS:
            ows.cell(i,col, FORMULA_COLS[col](i))
        _apply_keep(ows, i, extra)                # 覆盖 A006: 勾选列以输入表值为准
    print(f"    [{title}] matched {mcnt}, unmatched {ucnt}")
    total_match+=mcnt; total_unmatch+=ucnt

# 未匹配清单
us=out_wb.create_sheet(title="未匹配货号")
us.append(["客户页/分组","行号","货号","说明"])
for g,r,h in unmatched:
    us.append([g, r, h, "SAP文件中未找到该货号, 请核对"])

final=OUT
try:
    out_wb.save(final)
except PermissionError:
    n=len(glob.glob(str(OUT.with_name(OUT.stem+"_v*.xlsx"))))
    final=OUT.with_name(f"{OUT.stem}_v{n+1}.xlsx")
    out_wb.save(final)

print(f">>> Generated: {final}")
print(f">>> 总计: matched {total_match}, unmatched {total_unmatch}")
if dup_codes:
    print(f">>> Note: input has {len(dup_codes)} duplicate codes, kept as-is in result (unflagged)")
if KEEP_MAP:
    kept = ", ".join(f"'{dict(input_cols).get(c)}'" for c in KEEP_MAP)
    print(f">>> Kept input columns per your selection: {kept}(these cols are NOT from A006, taken directly from input)")
# summary print 已移交给 run_pipeline.py 统一输出 (避免重复打印) 
