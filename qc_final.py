#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""终稿全面质检：完整性/枚举合法性/标签合法性/格式/一致性/重复。零token。"""
import json, csv, glob, re, collections, openpyxl

import os
BASE = os.environ.get('MHB_TAG_BASE', os.path.dirname(os.path.abspath(__file__)))
wb = openpyxl.load_workbook(f'{BASE}/full/终稿_闵行生物医药企业标注.xlsx', read_only=True)
ws = wb['终稿']
rows = list(ws.iter_rows(min_row=2, values_only=True))
# sheet1（原表+标注）行数校验
if '原表+标注' in wb.sheetnames:
    n0 = wb['原表+标注'].max_row - 1
    if n0 != 2703: print(f'注意：原表+标注行数 {n0} ≠ 2703')
HDR = [c.value for c in next(wb['终稿'].iter_rows(min_row=1, max_row=1))]
ix = {h: i for i, h in enumerate(HDR)}
def g(r, c): return r[ix[c]] if r[ix[c]] is not None else ''

issues = []

# 1) 完整性
ids = [g(r, '序号') for r in rows]
if len(ids) != 2698: issues.append(f'行数 {len(ids)} ≠ 2698')
dups = [k for k, v in collections.Counter(ids).items() if v > 1]
if dups: issues.append(f'序号重复: {dups}')
src_ids = set()
for tsv in glob.glob(f'{BASE}/full/batches/batch_f*.tsv'):
    for i, l in enumerate(open(tsv, encoding='utf-8')):
        if i: src_ids.add(int(l.split('\t')[0]))
miss = src_ids - set(ids)
if miss: issues.append(f'源表序号缺失 {len(miss)}: {sorted(miss)[:10]}')

# 2) 枚举合法性
ENUMS = {
    '相关度': {'R0','R1','R2','R3','R4',''}, '标签置信度': {'高','中','低','登记',''},
    '母公司置信度': {'确定','疑似','无',''}, '知名度': {'知名','普通','未知'},
    '主体系': {'drug','device','service','ai','enabling','digital','周边配套',''},
}
for col, valid in ENUMS.items():
    bad = collections.Counter(g(r, col) for r in rows if g(r, col) not in valid)
    if bad: issues.append(f'{col}非法值: {dict(bad)}')

# 3) 标签合法性（二级必须在词表或补充清单）
valid_l2 = set()
with open(f'{BASE}/技术标签词表.csv', encoding='utf-8-sig') as f:
    for r in csv.DictReader(f): valid_l2.add(r['二级'])
valid_l2 |= {'医药流通与供应链','医药供应链上游配套','材料应用','食品应用','农牧应用','其他应用','复杂制剂','药物递送技术'}
bad_l2 = collections.Counter()
for r in rows:
    for c in ('二级标签主','二级标签副'):
        v = g(r, c)
        if v and v not in valid_l2: bad_l2[v] += 1
if bad_l2: issues.append(f'二级标签不在词表/补充清单: {dict(bad_l2)}')

# 三级合法性
valid_l3 = set()
with open(f'{BASE}/技术标签词表.csv', encoding='utf-8-sig') as f:
    for r in csv.DictReader(f): valid_l3.add(r['三级'])
valid_l3 |= {'制药装备及部件','医用及工业气体','洁净工程','制药公用工程','上游耗材配套','批发零售','冷链物流','进出口贸易','CSO','细胞培养肉','生物基材料','脂质体','微球','纳米粒','复杂注射剂','长效缓释','LNP递送','纳米递送','外泌体递送','靶向递送','递送载体'}
bad_l3 = collections.Counter(g(r, '三级标签') for r in rows if g(r, '三级标签') and g(r, '三级标签') not in valid_l3)
if bad_l3: issues.append(f'三级标签不在词表/补充清单 ({sum(bad_l3.values())}家): {dict(bad_l3.most_common(15))}')

# 4) 地位词闸门复核
n_status = sum(1 for r in rows if g(r, '行业地位'))
n_status_noev = sum(1 for r in rows if g(r, '行业地位') and 'http' not in g(r, '证据摘要'))
print(f'地位词非空 {n_status} 家，其中无证据URL {n_status_noev} 家')
if n_status_noev: issues.append(f'地位词无证据: {n_status_noev} 家')

# 5) 格式检查
fmt_bad = [g(r, '序号') for r in rows if g(r, '一句话简介') and re.match(r'^(该公司|该企业)', g(r, '一句话简介'))]
if fmt_bad: issues.append(f'简介以"该公司/该企业"开头: {fmt_bad[:10]}')
long_intro = [g(r, '序号') for r in rows if len(g(r, '一句话简介')) > 100]
if long_intro: issues.append(f'简介超100字: {len(long_intro)} 家 {long_intro[:8]}')

# 6) 一致性：R3/R4 不应挂领域体系标签
confl = [(g(r,'序号'), g(r,'公司名称')[:14], g(r,'相关度'), g(r,'主体系'), g(r,'二级标签主')) for r in rows
         if g(r,'相关度') in ('R3','R4') and g(r,'主体系') in ('drug','device','ai','digital')]
if confl: issues.append(f'R3/R4却挂核心体系标签 {len(confl)} 家: {confl[:8]}')

# 7) 简介重复（不同公司同一句简介=复制粘贴事故）
intro_map = collections.defaultdict(list)
for r in rows:
    s = g(r, '一句话简介')
    if s and len(s) > 10 and '暂无公开经营信息' not in s and '未搜索' not in s:
        intro_map[s].append(g(r, '序号'))
dup_intro = {k: v for k, v in intro_map.items() if len(v) > 1}
if dup_intro: issues.append(f'简介重复: {list(dup_intro.items())[:5]}')

# 8) 兜底句统计
fb = sum(1 for r in rows if '暂无公开经营信息' in g(r, '一句话简介'))
fb_vague = sum(1 for r in rows if '登记行业信息笼统' in g(r, '一句话简介'))
print(f'兜底句: {fb} 家（其中行业笼统 {fb_vague}）')

# 9) 无公开信息但有内容字段（违反铁律）
viol = [(g(r,'序号'), g(r,'公司名称')[:12]) for r in rows
        if g(r,'信息状态') == '无公开信息' and g(r,'主体系') and g(r,'标签依据') not in ('登记信息',)]
if viol: issues.append(f'无信息却有标签(非登记依据): {viol[:8]}')

print('\n=== 问题清单 ===')
if issues:
    for i, x in enumerate(issues, 1): print(f'{i}. {x}')
else:
    print('全部通过，零问题')
