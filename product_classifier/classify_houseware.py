# -*- coding: utf-8 -*-
"""
家居用品 希腊站上传表：纠正错误中文品名 + 按中国站分类树分类。
输入：input/希腊站上传_家居用品.xlsx
输出：output/希腊站上传_家居用品_classified.xlsx
"""
import openpyxl, re, os

HERE = os.path.dirname(os.path.abspath(__file__))          # product_classifier/
ROOT = os.path.dirname(HERE)                               # 仓库根
SRC = os.path.join(ROOT, "input", "希腊站上传_家居用品.xlsx")
TREE = os.path.join(HERE, "data", "中国站商品分类.xlsx")
OUT = os.path.join(HERE, "output", "希腊站上传_家居用品_classified.xlsx")

# ---------- 1. 载入中国站分类树，构建合法 3 级链路集合 ----------
twb = openpyxl.load_workbook(TREE, data_only=True)
tws = twb.active
valid = set()
for r in range(1, tws.max_row + 1):
    a, b, c = tws.cell(r, 4).value, tws.cell(r, 5).value, tws.cell(r, 6).value
    if a and b and c:
        valid.add((str(a), str(b), str(c)))

# ---------- 2. 品名纠正表（row -> 正确中文品名） ----------
CORRECTIONS = {
    2:  '感应杯 18oz',
    3:  '感应杯 12oz (400ML)',
    4:  '铜杯 4号 300ML',
    11: '陶瓷浓缩咖啡杯 90cc 碟套装',
    35: '窗帘 纱帘 140X260厘米',
    39: '铜杯 3号 250ML',
    72: '窗帘 黑色遮光 金属边 140X260厘米',
}

# ---------- 3. 分类规则（基于“纠正后”品名，按关键词判定 3 级链路） ----------
def classify(name):
    n = name
    # 杯类
    if '感应杯' in n or '铜杯' in n:
        return ('厨房用品', '饮具系列', '金属杯')
    if '马克杯' in n or '浓缩咖啡杯' in n:
        return ('厨房用品', '饮具系列', '陶瓷杯')
    # 钟表
    if '挂钟' in n:
        return ('家居百货', '钟表', '挂钟')
    # 花 / 植物
    if '花束' in n:
        return ('家居装饰', '仿真植物', '把束')
    if '花朵' in n:
        return ('家居装饰', '仿真植物', '花木')
    if '人造植物' in n:
        return ('家居装饰', '仿真植物', '花木')
    if '盆栽' in n or '花盆' in n:
        return ('家居装饰', '家居饰品', '花盆')
    # 盘碟
    if '汤碟' in n:
        return ('厨房用品', '餐具系列', '汤盘')
    if '长方形盘' in n:
        return ('厨房用品', '厨房工具', '方盘')
    if '餐盘' in n or '意式面碟' in n or '甜点盘' in n:
        return ('厨房用品', '餐具', '餐盘')
    # 床品
    if '羽绒被套' in n:
        return ('家居百货', '家纺系列', '床套')
    if '床垫保护套' in n:
        return ('家居百货', '家纺系列', '床笠')
    if '靠垫' in n:
        return ('家居百货', '家纺系列', '抱枕套')
    if '床单' in n:
        if '带胶条' in n:
            return ('家居百货', '家纺系列', '床笠')
        return ('家居百货', '家居纺织', '床上用品')
    # 卫浴 / 窗帘
    if '毛巾' in n:
        return ('家居卫浴', '洗浴用品', '毛巾')
    if '窗帘杆' in n:
        return ('家居百货', '窗帘及配件', '窗帘杆')
    if '窗帘' in n:
        return ('家居百货', '窗帘及配件', '窗帘')
    # 取暖
    if '壁炉' in n:
        return ('家电数码', '家用电器', '取暖器')
    return None

# ---------- 4. 处理主表 ----------
wb = openpyxl.load_workbook(SRC)
ws = wb.active

# 模板中 D:E（条形码 + 产品分类）按行合并，写产品分类列会命中 MergedCell，先解除合并。
# 解除后数值仍保留在 D（左上角），E 释放为可写单元格。
for mr in list(ws.merged_cells.ranges):
    if mr.min_col <= 5 and mr.max_col >= 4:
        ws.unmerge_cells(str(mr))

corr_log = []      # (row, old, new)
class_log = []     # (row, name, chain) or (row, name, '待确认')
invalid = []

for r in range(2, ws.max_row + 1):
    old = ws.cell(r, 1).value
    new = CORRECTIONS.get(r, old)
    if new != old:
        ws.cell(r, 1).value = new
        corr_log.append((r, old, new))

    chain = classify(new if new is not None else '')
    if chain is None:
        class_log.append((r, new, '待确认'))
        continue
    if chain not in valid:
        invalid.append((r, new, chain))
        class_log.append((r, new, '链路不存在(待确认)'))
        continue
    ws.cell(r, 5).value, ws.cell(r, 6).value, ws.cell(r, 7).value = chain
    class_log.append((r, new, ' / '.join(chain)))

# ---------- 5. 校验 & 落盘 ----------
os.makedirs(os.path.dirname(OUT), exist_ok=True)
wb.save(OUT)

print("=== 品名纠正 (%d 处) ===" % len(corr_log))
for r, o, nw in corr_log:
    print(f"  行{r:>3}: {o!r}  ->  {nw!r}")
print("\n=== 非法链路 (应为 0) ===", invalid if invalid else "无")
print("\n=== 分类分布 ===")
from collections import Counter
cnt = Counter()
for r, n, c in class_log:
    cnt[c] += 1
for c, k in sorted(cnt.items()):
    print(f"  {k:>3}  {c}")

print("\n=== 汇总 ===")
print("总行数:", ws.max_row - 1)
print("已分类:", sum(1 for r, n, c in class_log if c not in ('待确认', '链路不存在(待确认)')))
print("待确认:", sum(1 for r, n, c in class_log if '待确认' in c))
print("输出:", OUT)
