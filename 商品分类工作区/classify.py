# -*- coding: utf-8 -*-
"""
classify.py — 商品自动分类主程序（可复用，四层逻辑）
====================================================
按以下优先级分类，每一级结果都强制校验：必须存在于【中国站分类树】。

  ① 货号精确匹配  —— 用货号去主数据库查，有则采用主数据库分类（最权威）
  ② 原表分类优先  —— 原表三级类本身在中国站树内，则直接采用，不被后面层改动大类
  ③ 品名相似补全  —— 货号没有、原表无效时，用核心词去主数据库找"同大类"商品借三级
                     （门槛：核心词>=4、票数>=3、借到一级==原一级，避免跨类误借）
  ④ 字面+关键词  —— 以上都没有，按品名字面 + 原分类，用关键词规则兜底
  ⑤ 无法确定      —— 以上都不行，标"待确认"（不编造任何类目）

输出：在原表右侧追加 产品分类 / 二级分类 / 三级分类 三列。

用法：
  python classify.py                # 自动取 input/ 下最新 .xlsx
  python classify.py 路径/to/表.xlsx  # 指定输入
输出：output/<原文件名>_已分类.xlsx
"""
import openpyxl, os, re, sys, glob, importlib.util, collections

BASE = os.path.dirname(os.path.abspath(__file__))
_rules_path = os.path.join(BASE, 'rules', 'rules.py')
_spec = importlib.util.spec_from_file_location('cn_rules', _rules_path)
rules = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(rules)

_rules_full_path = os.path.join(BASE, 'rules', 'rules_full.py')
_spec2 = importlib.util.spec_from_file_location('cn_rules_full', _rules_full_path)
rules_full = importlib.util.module_from_spec(_spec2)
_spec2.loader.exec_module(rules_full)

DATA = os.path.join(BASE, 'data')
MASTER = os.path.join(BASE, 'master', '主数据库_商品数据.xlsx')
CN_TREE = os.path.join(DATA, '中国站商品分类.xlsx')
INPUT_DIR = os.path.join(BASE, 'input')
OUTPUT_DIR = os.path.join(BASE, 'output')

NEW_HEADERS = ['产品分类', '二级分类', '三级分类']


def norm(s):
    s = '' if s is None else str(s)
    s = s.lower()
    s = re.sub(r'\s+', '', s)
    return re.sub(r'[，,。.;；:：/\\()（）\\-_×x*]', '', s)


def first_sheet(wb):
    """取工作簿里第一个工作表（兼容中文/异名表名，不再写死 'Sheet1'）。
    优先返回第一个可见且有数据的表。"""
    # 优先选第一个至少有 1 行的表，避免空表
    for ws in wb.worksheets:
        if ws.max_row and ws.max_row >= 1:
            return ws
    return wb.worksheets[0]


def load_cn_tree(path):
    chains = set()
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    for r in first_sheet(wb).iter_rows(min_row=2, values_only=True):
        if r[3] and r[4] and r[5]:
            chains.add((r[3], r[4], r[5]))
    wb.close()
    return chains


def build_leaf_index(cn_chains):
    """三级名 -> 该名在标准树里出现的所有 (一级,二级) 集合。
    用于‘主数据库/旧分类的二级名与中国站树不一致’时，按三级名对齐到标准二级。"""
    idx = collections.defaultdict(set)
    for (t, s, l) in cn_chains:
        idx[l].add((t, s))
    return idx


def normalize_to_tree(chain, leaf_index, prefer_sec=None):
    """把 chain 对齐到中国站树的标准 (一级,二级)。
    - 若 chain 不在树里、但三级名在树里存在 → 换用树里的版本（三级名对齐）。
    - 若三级名在树里有多个 (一级,二级)（如 吹风机 既有 小家电 也有 个人护理电器）
      → 优先 prefer_sec（原二级）、其次非“筐底类”二级、其次与原一级同名。
    对任何已分配 chain（非待确认）都建议调用一次，确保落到树里的规范二级。"""
    t, s, l = chain
    CATCHALL = {'小家电', '其它', '其它类', '杂项', '其它用品', '其它系列'}
    if l in leaf_index and len(leaf_index[l]) >= 1:
        cands = sorted(leaf_index[l])   # 排序保证确定性（set 迭代序受 hash 随机化影响）
        if prefer_sec is not None:
            hit = [c for c in cands if c[1] == prefer_sec]
            if hit:
                return (hit[0][0], hit[0][1], l)
        specific = [c for c in cands if c[1] not in CATCHALL]
        pool = specific if specific else cands
        same_top = [c for c in pool if c[0] == t]
        return (same_top[0][0], same_top[0][1], l) if same_top else (pool[0][0], pool[0][1], l)
    return None


def load_master(path):
    """读取主数据库（商品数据.xlsx 干净版）。
    列序: 品名(0) 货号(1) 条形码(2) 产品分类(3) 二级分类(4) 三级分类(5)
    返回 (按货号索引, 按品名索引)。"""
    by_code, by_name = {}, {}
    if not os.path.exists(path):
        print(f"  [警告] 主数据库缺失：{path}，跳过精确匹配")
        return by_code, by_name
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    for sn in wb.sheetnames:
        for r in wb[sn].iter_rows(min_row=2, values_only=True):
            if not (r[0] or r[1]):
                continue
            chain = (r[3], r[4], r[5])
            if r[1] is not None:
                by_code.setdefault(norm(r[1]), chain)
            if r[0] is not None:
                by_name.setdefault(norm(r[0]), chain)
    wb.close()
    return by_code, by_name


def build_name_index(by_name):
    """为主数据库品名建 n-gram 倒排索引： ngram -> set(品名) """
    index = collections.defaultdict(set)
    for nm in by_name:
        L = len(nm)
        for n in (5, 4, 3, 2):
            if L < n:
                continue
            seen = set()
            for i in range(L - n + 1):
                g = nm[i:i + n]
                if g in seen:
                    continue
                seen.add(g)
                index[g].add(nm)
    return index


def similarity_match(nm, by_name, name_index, orig_top=None):
    """② 品名包含匹配：在主数据库里找品名【包含】本商品核心词的商品，
    借鉴其分类。要求：
      - 核心词长度 >= 3
      - 该主数据库品名与待分类品名的整体相似度(LCS比例) >= 0.5
      - 同类票数 >= 2（多数投票），或单条且核心词>=4
    orig_top: 原表一级类。若提供，则"与原一级相同"的候选票数门槛降到 >=1
    （同大类下补全三级风险低），"跨一级"的候选仍需 >=2。
    返回 (chain, 共享词长度, 票数, 参考品名) 或 None。
    确定性：所有平局按 (一级,二级,三级) 字典序打破，同一输入结果恒定。"""
    if not nm:
        return None
    L = len(nm)
    for n in (5, 4, 3):
        if L < n:
            continue
        # 收集候选主数据库品名（按核心词）
        cand_names = set()
        for i in range(L - n + 1):
            g = nm[i:i + n]
            cand_names |= name_index.get(g, set())
        if not cand_names:
            continue
        # 对每个候选算整体相似度，过滤低相似，再按分类投票
        vote = collections.Counter()
        for mname in cand_names:
            lcs = _lcs_ratio(nm, mname)
            if lcs < 0.5:
                continue
            vote[by_name[mname]] += 1
        if vote:
            # 优先"与原一级相同"的候选(门槛>=1)，否则取票数>=2的
            same = [c for c, v in vote.items() if (orig_top is None or c[0] == orig_top)]
            if same:
                # 平局按 (票数, 分类) 字典序取最大，保证确定性
                chain = max(same, key=lambda c: (vote[c], c))
                votes = vote[chain]
            else:
                chain, votes = max(vote.items(), key=lambda kv: (kv[1], kv[0]))
            # 记录一条参考品名用于审计（排序保证确定性）
            ref = next(m for m in sorted(cand_names) if by_name[m] == chain)
            return (chain, n, votes, ref)
    return None


def _lcs_ratio(a, b):
    """最长公共子序列长度 / max(len) ，衡量整体相似度。"""
    la, lb = len(a), len(b)
    if la == 0 or lb == 0:
        return 0.0
    # 滚动 DP
    dp = [0] * (lb + 1)
    for i in range(1, la + 1):
        prev = 0
        for j in range(1, lb + 1):
            tmp = dp[j]
            if a[i - 1] == b[j - 1]:
                dp[j] = prev + 1
            else:
                dp[j] = max(dp[j], dp[j - 1])
            prev = tmp
    return dp[lb] / max(la, lb)


# 强字面关键词：品名含这些词时，字面已经足够明确，应交给③精确归，不交给②借
STRONG_TOKENS = {'线','毛线','绒线','发箍','抓夹','发夹','发圈','皮筋','蝴蝶结',
                 '耳环','耳钉','项链','手链','戒指','胸针','发带','发绳',
                 '枕芯','枕头','吹风','吹风机',
                 '茶水杯'}  # 茶水杯：字面明确为饮具，避免被"200ml"等容量词借到别的品类

def build_valid_rules(cn_chains):
    toy = [(tok, s, l) for (tok, s, l) in rules.TOY_RULES if ('玩具系列', s, l) in cn_chains]
    hair = [(tok, s, l) for (tok, s, l) in rules.HAIR_CHILD_RULES if ('发饰系列', s, l) in cn_chains]
    baby = [(tok, s, l) for (tok, s, l) in rules.BABY_RULES if ('母婴系列', s, l) in cn_chains]
    dropped = [e for e in rules.TOY_RULES if ('玩具系列', e[1], e[2]) not in cn_chains]
    if dropped:
        print("  [提示] 玩具规则中被中国站树过滤(不存在)的:", dropped)
    full_valid = {}
    for scope, table in rules_full.SCOPE.items():
        kept = [e for e in table if (e[1], e[2], e[3]) in cn_chains]
        if kept:
            full_valid[scope] = kept
    return toy, hair, baby, full_valid


def in_cn(top, sec, leaf, cn_chains):
    return leaf != '待确认' and (top, sec, leaf) in cn_chains


def template_rules(nm, toy_rules, hair_rules, baby_rules):
    for tok, sec, leaf in toy_rules:
        if tok in nm:
            return ('玩具系列', sec, leaf)
    for tok, sec, leaf in hair_rules:
        if tok in nm:
            return ('发饰系列', sec, leaf)
    for tok, sec, leaf in baby_rules:
        if tok in nm:
            return ('母婴系列', sec, leaf)
    return None


# ──────────────────────────────────────────────────────────────
# 表头自动识别（不固定：精确别名优先，关键词模糊兜底）
#   - 每个目标列：先用精确别名命中（已 norm 归一化，忽略大小写/空格/标点）
#   - 未命中再用「关键词」模糊匹配（含任一词即命中，取最左列），避免
#     “商品名 / 货物编码 / SKU码 / 一级类目”等表头认不到而丢列
#   - 同一列只会被一个目标占用，按优先级（name>code>bar>otop>osec>oleaf）排他
# 新增列（产品分类/二级分类/三级分类）的识别规则见 NEW_HEADERS 处理段。
# ──────────────────────────────────────────────────────────────
EXACT = {
    'name': ['品名', 'name', '商品名称', '商品名', '产品名称', '货物名称', '标题', 'title'],
    'code': ['货号', 'code', 'sku', '编码', '货物编码', '商品编码', '货品编码', '款号', '型号', 'item', 'productcode', '货号sku'],
    'bar':  ['条形码', 'bar', 'barcode', '条码', '国条', '国际条码', 'ean', 'upc'],
    'otop': ['原产品分类', '产品分类', '一级分类', '一级类目', '类目', '品类', '产品类目', '原一级分类', '原分类'],
    'osec': ['原二级分类', '二级分类', '二级类目', '原二级', '二级类'],
    'oleaf':['原三级分类', '三级分类', '三级类目', '原三级', '三级类', '最末级分类'],
    'o4':  ['四级分类', '四级类目', '四级类'],
}
# 模糊关键词：表头只要「包含」其中任一词即视为该列（不要求整列相等）
KEYWORDS = {
    'name': ['品名', '商品名', '产品名', '货物名', '名称', 'title', 'name'],
    'code': ['货号', '编码', '款号', '型号', 'sku', 'code'],
    'bar':  ['条码', 'bar', 'barcode', 'ean', 'upc'],
    'otop': ['一级', '类目', '品类', '产品分类'],
    'osec': ['二级'],
    'oleaf':['三级'],
    'o4':  ['四级'],
}
# 目标列优先级：靠前的先占，避免“产品分类”同时命中 otop 与 oleaf 的模糊词
PRIORITY = ['name', 'code', 'bar', 'otop', 'osec', 'oleaf', 'o4']

# 二级分类别名归一：原表/外部数据用的叫法 → 标准树里的规范二级。
# 解决“整链不在标准树”的误判（如 原表‘塑料制品’ = 标准树‘塑料系列’）。
SEC_ALIAS = {
    '塑料制品': '塑料系列',
    '塑料': '塑料系列',
}


def locate_columns(ws):
    raw = [c.value for c in ws[1]]
    header = [('' if h is None else str(h)) for h in raw]
    norm_hdr = [norm(h) for h in header]

    used = set()          # 已被占用的列下标
    result = {}

    # 第 1 轮：精确别名（整列 norm 相等）
    for key in PRIORITY:
        found = None
        for i, nh in enumerate(norm_hdr):
            if i in used:
                continue
            if nh and nh in EXACT[key]:
                found = i
                break
        result[key] = found
        if found is not None:
            used.add(found)

    # 第 2 轮：关键词模糊（表头「包含」任一关键词；多列命中取最左）
    for key in PRIORITY:
        if result[key] is not None:
            continue
        for i, nm in enumerate(norm_hdr):
            if i in used or not nm:
                continue
            if any(k in nm for k in KEYWORDS[key]):
                result[key] = i
                used.add(i)
                break

    return result


def detect_headers(ws):
    """打印表头识别报告，让用户一眼确认程序认对了哪些列。"""
    raw = [c.value for c in ws[1]]
    cols = locate_columns(ws)
    label = {'name': '品名', 'code': '货号', 'bar': '条形码',
             'otop': '原一级分类', 'osec': '原二级分类', 'oleaf': '原三级分类', 'o4': '四级分类'}
    print("表头自动识别:")
    for i, h in enumerate(raw):
        tag = [label[k] for k in PRIORITY if cols[k] == i]
        print(f"  [{i:>2}] {repr(h):<16} -> {('、'.join(tag)) if tag else '（未使用）'}")
    miss = [label[k] for k in PRIORITY if cols[k] is None]
    if miss:
        print(f"  [提示] 未识别到列：{miss}（将跳过对应匹配层）")
    return cols


def classify(src_path, out_path):
    print(f"\n读取商品表: {src_path}")
    cn_chains = load_cn_tree(CN_TREE)
    print(f"中国站分类树: {len(cn_chains)} 个三级链")
    leaf_index = build_leaf_index(cn_chains)
    master_code, master_name = load_master(MASTER)
    print(f"主数据库: 货号{len(master_code)} / 品名{len(master_name)}")
    name_index = build_name_index(master_name)
    toy_rules, hair_rules, baby_rules, full_rules = build_valid_rules(cn_chains)

    wb = openpyxl.load_workbook(src_path)
    ws = first_sheet(wb)
    cols = detect_headers(ws)

    maxcol = ws.max_column
    header = [c.value for c in ws[1]]

    # 原表是否已自带【最终】产品分类/二级分类/三级分类 三列（都识别到）？
    #   · 三列都识别到，且都不是“原”开头的参考列（原产品分类/原二级/原三级等）
    #     → 视为“自带最终分类”，严格优先保留 + 追加「系统复核」列。
    #   · 带“原”字的是旧参考列 → 不算自带，仍按主库/品名/关键词重新分类（追加三列）。
    raw_at = lambda col: (header[col] if col is not None else '')
    HAS_FINAL = (cols['otop'] is not None and cols['osec'] is not None
                 and cols['oleaf'] is not None
                 and not str(raw_at(cols['otop'])).startswith('原')
                 and not str(raw_at(cols['osec'])).startswith('原')
                 and not str(raw_at(cols['oleaf'])).startswith('原'))
    if HAS_FINAL:
        # 严格优先：原表三列原样保留，绝不改写；只在末尾追加“系统复核”列
        #   记录程序意见：有效（整链在标准树）/ 待确认（不在标准树整链）
        ci_review = maxcol + 1
        ws.cell(row=1, column=ci_review, value='系统复核')
        ci_top = ci_sec = ci_leaf = None
    elif all(h in header for h in NEW_HEADERS):
        # 表头已有这三列（但通常不全等原表，仍按覆盖列处理）
        base = header.index('产品分类') + 1
        ci_top, ci_sec, ci_leaf = base, base + 1, base + 2
    else:
        ci_top, ci_sec, ci_leaf = maxcol + 1, maxcol + 2, maxcol + 3
        for off, h in enumerate(NEW_HEADERS):
            ws.cell(row=1, column=maxcol + 1 + off, value=h)

    c1 = c2 = c3 = c_unc = 0
    n = 0
    recommend_rows = []   # 收集【程序推荐的一到三级分类】(品名,货号,一级,二级,三级)
    for row in ws.iter_rows(min_row=2):
        n += 1
        name = row[cols['name']].value if cols['name'] is not None else None
        code = row[cols['code']].value if cols['code'] is not None else None
        otop = row[cols['otop']].value if cols['otop'] is not None else None
        nm = norm(name) if name else ''

        assigned = None
        reason = ''
        # ═══ 始终跑完整四层，得到【程序推荐分类】assigned ═══
        # 不论原表是纯净“产品分类”还是“原产品分类”参考列，推荐值都用同一套逻辑算。
        # 主表是否写回 assigned、还是保留原值另加“系统复核”，由下方 HAS_FINAL 写回段决定。
        # ① 货号精确匹配
        if code is not None and norm(code) in master_code:
            m = master_code[norm(code)]
            if in_cn(*m, cn_chains):
                assigned = m; reason = '货号精确'; c1 += 1
        # ② 原表参考列优先（原列有效则直接采用，不被③/④改动大类）
        if assigned is None and cols['oleaf'] is not None:
            ot = row[cols['otop']].value
            os_ = row[cols['osec']].value
            ol = row[cols['oleaf']].value
            if in_cn(ot, os_, ol, cn_chains):
                assigned = (ot, os_, ol); reason = '原分类有效'; c3 += 1
        # ③ 品名包含匹配（主数据库相似品名借鉴分类；主数据库已是正确源）
        if assigned is None and not any(t in nm for t in STRONG_TOKENS):
            ot = row[cols['otop']].value if cols['otop'] is not None else None
            sim = similarity_match(nm, master_name, name_index, orig_top=ot)
            if sim:
                chain, glen, votes, ref = sim

                # 领域护栏：借到的分类必须与“原表一级”在同一大领域，否则拒绝。
                #   玩具域: 玩具系列/儿童玩具/母婴系列/儿童用品/玩具
                #   饰美域: 饰品系列/发饰系列/美妆用品/美妆系列
                #   家居域: 其余（家居百货/厨房/家电/电子/节日/箱包…）
                # 作用：防止“儿童玩具”被借成“厨房用品”这类明显跨界的错配；
                #       同领域内跨一级（如 电子电器→家电数码）允许。
                def _grp(top):
                    if top in ('玩具系列', '儿童玩具', '母婴系列', '儿童用品', '玩具'):
                        return '玩具'
                    if top in ('饰品系列', '发饰系列', '美妆用品', '美妆系列'):
                        return '饰美'
                    return '家居'

                same_grp = (ot is None) or (_grp(chain[0]) == _grp(ot))
                if in_cn(*chain, cn_chains) and glen >= 3 and votes >= 1 and same_grp:
                    assigned = chain; reason = f'品名相似(g{glen},v{votes})'; c2 += 1
        # ④ 字面+关键词理解（原表无效或缺失时，按字面规则兜底）
        if assigned is None:
            rr = None
            # 跨类强字面（吹风机等，字面已足够明确，直接精确归，不被③歧义借错）
            # 灯具/电工强字面：品名含这些词即足够明确，无需原一级，直接精确归。
            # （插头/插座/灯泡/灯杯 皆为“跨类强字面”，与 吹风/茶水杯 同性质）
            # 顺序：具体词在前，避免被更泛的词误命中。
            GLOBAL = [('灯泡', '家电数码', '灯具', '灯泡'),
                      ('球泡灯', '家电数码', '灯具', '灯泡'),
                      ('灯杯', '家电数码', '灯具', '灯杯'),
                      ('杯灯', '家电数码', '灯具', '灯杯'),
                      ('插头', '家电数码', '开关配件', '插头'),
                      ('插座', '家电数码', '开关配件', '插座'),
                      ('吹风', '家电数码', '个人护理电器', '吹风机'),
                      ('吹风机', '家电数码', '个人护理电器', '吹风机'),
                      ('茶水杯', '厨房用品', '饮具系列', '陶瓷杯')]
            for tok, top, sec, leaf in GLOBAL:
                if tok in nm and in_cn(top, sec, leaf, cn_chains):
                    rr = (top, sec, leaf); break
            if rr is None:
                if otop in ('儿童玩具', '玩具'):
                    rr = template_rules(nm, toy_rules, hair_rules, baby_rules)
                elif otop in full_rules:
                    for tok, top, sec, leaf in full_rules[otop]:
                        if tok in nm:
                            rr = (top, sec, leaf); break
            if rr and in_cn(*rr, cn_chains):
                assigned = rr; reason = '关键词规则'; c3 += 1
            # 就近归并
            if assigned is None and cols['oleaf'] is not None:
                key = (row[cols['otop']].value, row[cols['osec']].value, row[cols['oleaf']].value)
                if key in rules.HOME_FIX:
                    assigned = rules.HOME_FIX[key]; reason = '就近归并'; c3 += 1
        # ⑤ 无法确定
        if assigned is None:
            assigned = ('待确认', '待确认', '待确认'); c_unc += 1

        if not HAS_FINAL:
            # 三级名对齐/规范：把 (一级,二级) 对齐到树里的规范二级
            # （解决“主数据库/旧分类的二级名与中国站树不一致”，以及同三级名多二级时优先非筐底类，
            #  如 小家电/吹风机 → 个人护理电器/吹风机）。
            # 关键：若 chain 本身已在树里（in_cn），保留原二级，不做对齐——
            #   否则“规则明确指定的二级”（如 茶水杯→饮具系列/陶瓷杯）会被按字典序
            #   误对齐到同三级名下的另一个二级（如 陶瓷制品/陶瓷杯）。
            # 注：HAS_FINAL（原表自带最终分类）模式下不做对齐——原表整链不在标准树
            #     就判待确认，绝不把“三级名在树里”的链对齐成有效链来冒充“有效”。
            if assigned[0] != '待确认':
                if not in_cn(*assigned, cn_chains):
                    nrm = normalize_to_tree(assigned, leaf_index, prefer_sec=row[cols['osec']].value if cols['osec'] is not None else None)
                    if nrm:
                        assigned = nrm

        if HAS_FINAL:
            # 严格优先：原表三列原样保留（不写 assigned），仅把程序意见写进“系统复核”列。
            #   系统复核基于【原表值】是否在标准树整链，而非 assigned（推荐值）。
            ot = row[cols['otop']].value
            os_ = row[cols['osec']].value
            ol = row[cols['oleaf']].value
            # 二级别名归一（如 原表‘塑料制品’ → 标准树‘塑料系列’），
            # 避免整链因二级叫法差异被误判为“不在标准树”。
            os_norm = SEC_ALIAS.get(os_, os_)
            orig_ok = bool(ot and os_ and ol and in_cn(ot, os_norm, ol, cn_chains))
            verdict = '有效' if orig_ok else '待确认'
            row[ci_review - 1].value = verdict
        else:
            orig_ok = False
            row[ci_top - 1].value = assigned[0]
            row[ci_sec - 1].value = assigned[1]
            row[ci_leaf - 1].value = assigned[2]

        # 注：四级分类等辅助列不再注入“待确认”——待审核状态统一由「系统复核」列标记，
        #     原表自带列原样保留，避免重复标记。

        # 收集【程序推荐的一到三级分类】用于新建 sheet。
        # 规则：原表分类已通过校验(orig_ok)的商品不需要再推荐 → 不写入推荐 sheet。
        if not orig_ok:
            recommend_rows.append((name, code, assigned[0], assigned[1], assigned[2]))

    wb.save(out_path)

    # ── 新建 sheet：程序推荐的一到三级分类 ──
    # 列：品名 / 货号 / 推荐产品分类 / 推荐二级分类 / 推荐三级分类
    # 与原表主数据分开，方便你对比原分类与程序推荐值、按需采用。
    wb = openpyxl.load_workbook(out_path)
    rec_ws = wb.create_sheet(title='程序推荐分类')
    rec_ws.append(['品名', '货号', '推荐产品分类', '推荐二级分类', '推荐三级分类'])
    for (rn, rc, rt, rs, rl) in recommend_rows:
        rec_ws.append([rn, rc, rt, rs, rl])
    wb.save(out_path)
    print(f"  新建 sheet「程序推荐分类」: {len(recommend_rows)} 行")

    # 校验
    wb2 = openpyxl.load_workbook(out_path, read_only=True, data_only=True)
    ok = unc = bad = 0
    if HAS_FINAL:
        # 原表自带三列：校验“系统复核”列
        for r in first_sheet(wb2).iter_rows(min_row=2, values_only=True):
            v = r[ci_review - 1]
            if v == '待确认':
                unc += 1
            else:
                ok += 1
    else:
        for r in first_sheet(wb2).iter_rows(min_row=2, values_only=True):
            leaf = r[ci_leaf - 1]
            if leaf == '待确认':
                unc += 1
            else:
                ok += 1
                if (r[ci_top - 1], r[ci_sec - 1], leaf) not in cn_chains:
                    bad += 1
                    print("  INVALID:", (r[ci_top - 1], r[ci_sec - 1], leaf))
    wb2.close()
    print(f"\n写出: {out_path}")
    if HAS_FINAL:
        print("  模式: 原表自带分类（严格优先）→ 原列原样保留，仅追加「系统复核」列")
        print(f"  系统复核: 有效={ok} | 待确认={unc}")
    else:
        print(f"  ①货号精确={c1} | ②品名相似={c2} | ③字面/原分类={c3} | ④待确认={c_unc}")
        print(f"  校验: 已归类(中国站真实链)={ok} | 待确认={unc} | 无效链={bad}")


def find_input():
    files = [f for f in glob.glob(os.path.join(INPUT_DIR, '*.xlsx')) if not os.path.basename(f).startswith('~$')]
    return max(files, key=os.path.getmtime) if files else None


if __name__ == '__main__':
    src = sys.argv[1] if len(sys.argv) > 1 else find_input()
    if not src or not os.path.exists(src):
        print("未找到待分类商品表。请把 .xlsx 放到 input/ 目录，或传入路径。")
        sys.exit(1)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    base = os.path.splitext(os.path.basename(src))[0]
    out = os.path.join(OUTPUT_DIR, base + '_已分类.xlsx')
    classify(src, out)
