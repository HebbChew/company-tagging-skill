#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""步骤0：全量切批。源表P1 → 去重/过滤/聚类/分批 → full/batches/batch_fXX.tsv + 清洗报告"""
import openpyxl, re, os, json, collections

SRC = os.environ.get('MHB_SOURCE_XLSX', '')  # 源表路径，必填环境变量
import os
BASE = os.environ.get('MHB_TAG_BASE', os.path.dirname(os.path.abspath(__file__)))
OUT = f'{BASE}/full/batches'
os.makedirs(OUT, exist_ok=True)
BATCH_SIZE = 130

wb = openpyxl.load_workbook(SRC, read_only=True)
ws = wb['P1强相关企业']
rows = list(ws.iter_rows(min_row=2, values_only=True))
# 列: 0序号 1名称 2保留优先度 3调整标记 4高新 5专精特新 6规上 7登记状态 8信用代码 9成立日期 10注册资本 11区县 12地址 ...18行业小类 19规模 22营收 24科创分
def wan(s):
    m = re.search(r'([\d.]+)\s*万', str(s or ''))
    return float(m.group(1)) if m else 0
def rev_yi(s):
    m = re.match(r'([\d.]+)(亿|万)?', str(s or ''))
    if not m: return 0
    v = float(m.group(1))
    return v if m.group(2) == '亿' else v/10000

recs = []
for r in rows:
    recs.append({
        '序号': r[0], '名称': str(r[1]).strip(), '信用代码': str(r[8] or '').strip(),
        '登记状态': str(r[7] or '').strip(), '地址': str(r[12] or '').strip(),
        '行业小类': str(r[18] or '').strip(), '高新': r[4] or '', '专精特新': r[5] or '',
        '规上': r[6] or '', '规模': str(r[19] or '').strip(), '营收亿': round(rev_yi(r[22]), 2),
        '注册资本万': int(wan(r[10])), '科创分': str(r[24] or '').strip(),
    })

# 1) 登记状态过滤
live, dead = [], []
for r in recs:
    (live if r['登记状态'] in ('存续', '在业', '在营', '开业', '正常') else dead).append(r)

# 2) 信用代码去重（有代码的按代码；无代码按名称）
seen, dedup, dups = {}, [], []
for r in live:
    key = r['信用代码'] or ('NAME:' + r['名称'])
    if key in seen:
        dups.append((r['序号'], r['名称'], seen[key]['序号']))
    else:
        seen[key] = r
        dedup.append(r)

# 3) 集团/品牌聚类：商号前缀（2/3/4字）≥2家共享即并群，union-find 合併；通用词前缀不算
def core(name):
    s = re.sub(r'[（(].*?[）)]', '', name)          # 去括号内容
    s = re.sub(r'^(上海市|上海|中国)', '', s)         # 去地区前缀
    s = re.sub(r'(股份|有限责任|有限)?公司$', '', s)   # 去组织形式
    return s.strip()
GENERIC = {'生物','医药','医疗','科技','健康','生命','智能','智慧','医学','制药','基因','细胞','上海','中国'}
parent = {r['序号']: r['序号'] for r in dedup}
def find(x):
    while parent[x] != x:
        parent[x] = parent[parent[x]]
        x = parent[x]
    return x
def union(a, b):
    ra, rb = find(a), find(b)
    if ra != rb:
        parent[rb] = ra
for L in (2, 3, 4):
    pref = collections.defaultdict(list)
    for r in dedup:
        c = core(r['名称'])
        if len(c) >= L:
            pref[c[:L]].append(r['序号'])
    for p, ids in pref.items():
        if len(ids) >= 2 and p not in GENERIC:
            for i in ids[1:]:
                union(ids[0], i)
cluster_of = {r['序号']: find(r['序号']) for r in dedup if any(find(r['序号']) == find(o['序号']) and o is not r for o in dedup)}
# 上面效率差，改为按根分组后只保留≥2家的群
roots = collections.defaultdict(list)
for r in dedup:
    roots[find(r['序号'])].append(r['序号'])
cluster_of = {}
for root, ids in roots.items():
    if len(ids) >= 2:
        for i in ids:
            cluster_of[i] = root
n_clustered = len(cluster_of)
n_clusters = len(set(cluster_of.values()))

# 4) 分批：聚类整体不拆，贪心装批
groups = collections.defaultdict(list)
for r in dedup:
    groups[cluster_of.get(r['序号'], f"s{r['序号']}")].append(r)
batches, cur, cur_n = [], [], 0
for g, members in sorted(groups.items(), key=lambda kv: -len(kv[1])):
    if cur_n + len(members) > BATCH_SIZE and cur:
        batches.append(cur)
        cur, cur_n = [], 0
    cur += members
    cur_n += len(members)
if cur:
    batches.append(cur)
# 批内按序号排序，批号重排
for b in batches:
    b.sort(key=lambda r: r['序号'])

for i, b in enumerate(batches, 1):
    with open(f'{OUT}/batch_f{i:02d}.tsv', 'w', encoding='utf-8') as f:
        f.write('序号\t公司名称\t信用代码\t地址\t行业小类\t高新\t专精特新\t规上\t规模\t营收亿\t注册资本万\t科创分\n')
        for r in b:
            f.write('\t'.join([str(r['序号']), r['名称'], r['信用代码'], r['地址'], r['行业小类'],
                               str(r['高新']), str(r['专精特新']), str(r['规上']), r['规模'],
                               str(r['营收亿']), str(r['注册资本万']), r['科创分']]) + '\n')

report = {
    '源表行数': len(recs), '非存续单列': len(dead), '重复合并': len(dups),
    '进入主流程': len(dedup), '聚类企业数': n_clustered, '聚类簇数': n_clusters,
    '批次数': len(batches), '批均': round(len(dedup)/len(batches), 1),
}
with open(f'{BASE}/full/清洗报告.json', 'w', encoding='utf-8') as f:
    json.dump({'汇总': report,
               '重复主体': [{'序号': d[0], '名称': d[1], '保留序号': d[2]} for d in dups],
               '非存续': [{'序号': r['序号'], '名称': r['名称'], '状态': r['登记状态']} for r in dead]},
              f, ensure_ascii=False, indent=1)
print(json.dumps(report, ensure_ascii=False))
print('批次大小分布:', sorted(len(b) for b in batches))
print('重复样例:', dups[:5])
print('非存续样例:', [(r['序号'], r['名称'][:14], r['登记状态']) for r in dead[:5]])
EOF_MARKER = None
