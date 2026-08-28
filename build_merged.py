#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""构建 Pass A merged 数据集：Pass A 初判 + 搜索层终稿回填（含复找），并打 _已核验 标记。
输出: cache/passA_intro/batch_XXX_v3merged.jsonl"""
import json, glob

import os
BASE = os.environ.get('MHB_TAG_BASE', os.path.dirname(os.path.abspath(__file__)))

# 搜索结果汇总：先常规层，后复找层（复找=更深搜索，覆盖空记录）
import re as _re
_OURS = _re.compile(r'^(full_S2_part\d+|full_S1_part\d+|full_S0A_A\d+|full_S0B_B\d+|full_S0C_C\d+-\d+|full_S0C_C3a|refind_R\d+|audit_fix|reverify_3家|full_S0A_组|full_S0B_组|full_S0C_批|full_S1_组|full_S2_组|full_S0W_W)')
def _our_search_files():
    import glob as _g
    return [p for p in _g.glob(f'{BASE}/cache/search/*.jsonl') if _OURS.match(p.split('/')[-1])]

S = {}
for p in sorted(_our_search_files()):
    for l in open(p, encoding='utf-8'):
        try: r = json.loads(l)
        except: continue
        S[r['序号']] = r
for p in sorted(glob.glob(f'{BASE}/cache/search/refind_*.jsonl') + glob.glob(f'{BASE}/cache/search/audit_*.jsonl')):
    for l in open(p, encoding='utf-8'):
        try: r = json.loads(l)
        except: continue
        S[r['序号']] = r

for tsv in sorted(glob.glob(f'{BASE}/full/batches/batch_f*.tsv')):
    b = tsv.split('/')[-1].replace('.tsv', '')
    out = []
    for l in open(f'{BASE}/cache/passA_intro/{b}_v3.jsonl', encoding='utf-8'):
        a = json.loads(l)
        s = S.get(a['序号'])
        if s:
            a['主营业务'] = s.get('主营业务_终稿') or a.get('主营业务', '')
            a['技术关键词'] = s.get('技术关键词_终稿') or a.get('技术关键词', [])
            a['一句话简介'] = s.get('一句话简介_终稿') or a.get('一句话简介', '')
            a['行业地位'] = s.get('行业地位_终稿') or a.get('行业地位', '')
            a['相关度'] = s.get('相关度') or a.get('相关度', '')
            a['_已核验'] = True          # 有搜索记录=内容经过搜索层核验
        else:
            a['_已核验'] = False
        out.append(a)
    with open(f'{BASE}/cache/passA_intro/{b}_v3merged.jsonl', 'w', encoding='utf-8') as f:
        for a in out:
            f.write(json.dumps(a, ensure_ascii=False) + '\n')
print(f'merged 完成，搜索覆盖 {len(S)} 家')
