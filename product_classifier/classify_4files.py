# -*- coding: utf-8 -*-
"""
希腊站 4 张上传表批量分类（按中国站分类树，策略：能精确归类则归类；
无对应叶子的欧洲零售类(汽车养护汽车养护/纸品日化纸品日化)及无性别内衣袜子，
按用户要求"近似归到上级大类"收口，并在报告显著标注近似项供复核）。
文件: 汽车养护(12) 小家电(74) 内衣(163) 纸品日化(64)
"""
import openpyxl, os
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))          # product_classifier/
ROOT = os.path.dirname(HERE)                               # 仓库根
BASE = os.path.join(ROOT, "input")                         # 待分类的希腊站上传表
OUTDIR = os.path.join(HERE, "output")
TREE = os.path.join(HERE, "data", "中国站商品分类.xlsx")

twb = openpyxl.load_workbook(TREE, data_only=True)
tws = twb.active
valid = set()
for r in range(1, tws.max_row + 1):
    a, b, c = tws.cell(r, 4).value, tws.cell(r, 5).value, tws.cell(r, 6).value
    if a and b and c:
        valid.add((str(a), str(b), str(c)))


# ===== 小家电：精确 =====
def cls_small_appliance(n):
    if '收音机' in n: return (None, ('家电数码','电子产品','收音机'))
    if '吹风机' in n: return (None, ('家电数码','个人护理电器','吹风机'))
    if '直发器' in n: return (None, ('家电数码','个人护理电器','夹板'))
    if '卷发棒' in n: return (None, ('家电数码','个人护理电器','卷发棒'))
    if '理发剪' in n: return (None, ('家电数码','个人护理电器','理发剪'))
    if '修剪器' in n: return (None, ('家电数码','个人护理电器','剃毛器'))
    return ('待确认', n)   # VGR 无品类信息

# ===== 内衣 =====
def cls_underwear(n):
    if '文胸' in n: return (None, ('服装服饰','内衣裤','文胸'))
    if '裹胸' in n: return (None, ('服装服饰','内衣裤','文胸'))
    if '连裤袜' in n: return (None, ('服装服饰','袜子','丝袜'))
    if '丁字裤' in n: return (None, ('服装服饰','内衣裤','女士内裤'))
    if '女式内裤' in n: return (None, ('服装服饰','内衣裤','女士内裤'))
    if '男式平角裤' in n: return (None, ('服装服饰','内衣裤','男士内裤'))
    if '男式三角裤' in n: return (None, ('服装服饰','内衣裤','男士内裤'))
    if n == '袜子' or n.startswith('袜子'):
        return ('近似:袜子无性别,按女士内衣系列归女士袜', ('服装服饰','袜子','女士袜'))
    return ('待确认', n)

# ===== 汽车养护：汽车养护，全部近似到汽车清洗护理大类 =====
# (润滑油/防冻液/玻璃水/洗车香波等均无专属叶，统一近似汽车清洁洗护)
def cls_carcare(n):
    return (('近似:汽车养护化学品,无对应叶,归汽车清洗护理大类'),
            ('汽车用品','汽车清洁','汽车清洗护理'))

# ===== 纸品日化：纸品近似餐巾纸叶；洗涤日化近似护理用品类 =====
def cls_paper_daily(n):
    # 纸品
    if '卫生卷纸' in n or '厨房卷纸' in n:
        return (('近似:生活用纸无专属叶,归一次性餐巾纸大类'), ('厨房用品','一次性用品','餐巾纸'))
    if '餐巾纸' in n: return (None, ('厨房用品','一次性用品','餐巾纸'))
    if '面巾纸' in n: return (None, ('厨房用品','一次性用品','餐巾纸'))
    if '手帕纸' in n: return (None, ('厨房用品','一次性用品','餐巾纸'))
    if '湿巾' in n: return (None, ('家居百货','护理用品','湿巾'))
    # 洗涤/清洁日化 -> 护理用品/清洁剂（近似，作为家用清洁剂收口）
    return (('近似:家用洗涤/清洁日化,无专属叶,归护理用品-清洁剂大类'),
            ('家居百货','护理用品','清洁剂'))

jobs = [
    ("希腊站上传_汽车养护.xlsx",      cls_carcare),
    ("希腊站上传_小家电.xlsx",    cls_small_appliance),
    ("希腊站上传_内衣.xlsx",  cls_underwear),
    ("希腊站上传_纸品日化.xlsx",    cls_paper_daily),
]

summary = {}
for fname, fn in jobs:
    src = os.path.join(BASE, fname)
    wb = openpyxl.load_workbook(src)
    ws = wb.active
    for mr in list(ws.merged_cells.ranges):
        if mr.min_col <= 5 and mr.max_col >= 4:
            ws.unmerge_cells(str(mr))
    exact = approx = uncls = 0
    approx_rows, uncls_rows = [], []
    for r in range(2, ws.max_row + 1):
        n = ws.cell(r, 1).value
        if not n or not str(n).strip():
            ws.cell(r,5).value=ws.cell(r,6).value=ws.cell(r,7).value=None
            continue
        n = str(n)
        res = fn(n)
        flag, chain = res
        # chain 若是待确认元组（待确认,n）处理
        if isinstance(chain, str) and chain == '待确认':
            ws.cell(r,5).value=ws.cell(r,6).value=ws.cell(r,7).value=None
            uncls += 1; uncls_rows.append((r, n)); continue
        if chain not in valid:
            ws.cell(r,5).value=ws.cell(r,6).value=ws.cell(r,7).value=None
            uncls += 1; uncls_rows.append((r, n, f"非法链{chain}")); continue
        ws.cell(r,5).value, ws.cell(r,6).value, ws.cell(r,7).value = chain
        if flag is None:
            exact += 1
        else:
            approx += 1; approx_rows.append((r, n, flag, ' / '.join(chain)))
    tag = fname.replace('希腊站上传_','').replace('.xlsx','')
    out = os.path.join(OUTDIR, fname.replace('.xlsx','_classified.xlsx'))
    wb.save(out)
    summary[tag]=(exact,approx,uncls)
    print(f"\n{'='*70}\n[{tag}] {fname}")
    print(f"  精确 {exact} | 近似 {approx} | 待确认 {uncls}")
    if approx_rows:
        print("  近似明细:")
        for r,n,flag,c in approx_rows:
            print(f"    {r:>3} | {n[:34]:<36} ~ {c}")
    if uncls_rows:
        print("  待确认明细:")
        for row in uncls_rows:
            print("    ", row)
    print("  输出:", out)

print("\n\n===== 汇总 =====")
for tag,(e,a,u) in summary.items():
    print(f"  {tag:<8} 精确{e:>3}  近似{a:>3}  待确认{u:>3}")
