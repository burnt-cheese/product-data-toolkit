# -*- coding: utf-8 -*-
"""把 master/商品数据.xlsx 的三个 sheet 合并成一个干净的、dimension 正确的
主数据库，供 classify.py 用 read_only 秒读。货物号同时建 货号/条形码 索引。
输出: master/主数据库_商品数据.xlsx  (列: 品名,货号,条形码,产品分类,二级分类,三级分类)
"""
import openpyxl, os, re, time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, 'master', '商品数据.xlsx')
OUT = os.path.join(ROOT, 'master', '主数据库_商品数据.xlsx')

def norm(s):
    s = '' if s is None else str(s)
    s = s.lower().strip()
    return re.sub(r'[，,。.;；:：/\\()（）\-_×x*]', '', s)

t0 = time.time()
print('加载源文件(普通模式，约需2-3分钟)...')
wb = openpyxl.load_workbook(SRC, read_only=False, data_only=True)
print(f'  加载耗时 {round(time.time()-t0,1)}s')

out = openpyxl.Workbook()
ws = out.active
ws.title = '主数据库'
ws.append(['品名', '货号', '条形码', '产品分类', '二级分类', '三级分类'])

code_idx = {}
name_idx = {}
dup = 0
total = 0
for sn in wb.sheetnames:
    s = wb[sn]
    for r in s.iter_rows(min_row=2, values_only=True):
        name, code, bar = r[0], r[1], r[2]
        top, sec, leaf = r[4], r[5], r[6]
        if not name and not code:
            continue
        total += 1
        ws.append([name, code, bar, top, sec, leaf])
        nc = norm(code)
        if nc:
            code_idx.setdefault(nc, (top, sec, leaf))
        nn = norm(name)
        if nn:
            name_idx.setdefault(nn, (top, sec, leaf))
    print(f'  sheet {sn} 处理完')
wb.close()

out.save(OUT)
print(f'写出 {OUT}')
print(f'总行数={total} 货号索引={len(code_idx)} 品名索引={len(name_idx)}')

# 校验 dimension 是否正确
wb2 = openpyxl.load_workbook(OUT, read_only=True)
print('  read_only 重新读: max_row=', wb2['主数据库'].max_row)
wb2.close()
