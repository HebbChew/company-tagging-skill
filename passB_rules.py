#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Pass B v3: 规则打标（零token）。
输入: cache/passA_intro/{batch}.jsonl（merged 含 _已核验 标记）+ 技术标签词表.csv + method/通用词映射.csv + 批次TSV
匹配优先级: 技术关键词字段(bonus+2) > 词表三级/别名(w3) > 词表二级(w2) > 通用词映射(w1-2)
三级只来自三级名/别名命中（二级名/无三级通用词不下钻）；
多体系时选 prose 有印证的；ASCII 关键词≥3字符整词匹配。

置信度分档 v3（2026-08-27 用户裁定）:
  高   = 词表命中（二级/三级/别名）且内容经搜索核验(_已核验)
  中   = 词表命中但未核验 或 通用词映射命中且已核验
  低   = 通用词映射未核验 / prose兜底 / 流通·周边规则
  登记 = 行业小类映射（批发零售类未经实证）
  空   = 残渣（无标签）
输出: cache/passB_rules/{batch}.jsonl
"""
import json, csv, re, sys, collections, os

import os
BASE = os.environ.get('MHB_TAG_BASE', os.path.dirname(os.path.abspath(__file__)))
BATCH = sys.argv[1] if len(sys.argv) > 1 else 'batch_001_v2'
TSV = BATCH.split('_v')[0]

STOP = {'其他', '服务', '技术', '研发', '生产', '销售', '制造', '无', ''}
kw_index = []   # (kw, 体系, 二级, 三级, 权重, 级别, 来源)  级别2=二级名 3=三级/别名; 来源 dict|map
with open(f'{BASE}/技术标签词表.csv', encoding='utf-8-sig') as f:
    for r in csv.DictReader(f):
        sys_, l2, l3 = r['体系'], r['二级'], r['三级']
        kws = set()
        if l3 and l3 not in STOP: kws.add((l3, 3, 3))
        if l2 and l2 not in STOP: kws.add((l2, 2, 2))
        for a in re.split(r'[;；、,，]', r['别名'] or ''):
            a = a.strip()
            if len(a) >= 2 and a not in STOP: kws.add((a, 3, 3))
        for kw, w, lvl in kws:
            kw_index.append((kw, sys_, l2, l3, w, lvl, 'dict'))
with open(f'{BASE}/method/通用词映射.csv', encoding='utf-8-sig') as f:
    for r in csv.DictReader(f):
        l3 = (r.get('三级') or '').strip()
        kw_index.append((r['关键词'], r['体系'], r['二级'], l3, 2 if l3 else 1, 3 if l3 else 2, 'map'))
# v3.1: 补充二级/三级纳入匹配索引（source='sup'）——模型技术关键词常用补充清单标准名，
# 不纳入则正确关键词匹配不到、被通用词带偏（利和/毕科案例）
SUP_L2 = {
    '医药流通与供应链': ('service', ''), '医药供应链上游配套': ('service', ''), '复杂制剂': ('drug', ''), '药物递送技术': ('drug', ''),
    '材料应用': ('enabling', ''), '食品应用': ('enabling', ''), '农牧应用': ('enabling', ''), '其他应用': ('enabling', ''),
}
SUP_L3 = {'制药装备及部件': ('service', '医药供应链上游配套'), '医用及工业气体': ('service', '医药供应链上游配套'),
          '洁净工程': ('service', '医药供应链上游配套'), '制药公用工程': ('service', '医药供应链上游配套'),
          '上游耗材配套': ('service', '医药供应链上游配套'),
          '批发零售': ('service', '医药流通与供应链'), '冷链物流': ('service', '医药流通与供应链'),
          '进出口贸易': ('service', '医药流通与供应链'), 'CSO': ('service', '医药流通与供应链'),
          '细胞培养肉': ('enabling', '食品应用'), '生物基材料': ('enabling', '材料应用'),
          '脂质体': ('drug', '复杂制剂'), '微球': ('drug', '复杂制剂'), '纳米粒': ('drug', '复杂制剂'),
          '复杂注射剂': ('drug', '复杂制剂'), '长效缓释': ('drug', '复杂制剂'),
          'LNP递送': ('drug', '药物递送技术'), '纳米递送': ('drug', '药物递送技术'), '外泌体递送': ('drug', '药物递送技术'),
          '靶向递送': ('drug', '药物递送技术'), '递送载体': ('drug', '药物递送技术')}
for kw, (sys_, l2) in SUP_L2.items():
    kw_index.append((kw, sys_, kw, '', 2, 2, 'sup'))
for kw, (sys_, l2) in SUP_L3.items():
    kw_index.append((kw, sys_, l2, kw, 3, 3, 'sup'))
kw_index.sort(key=lambda x: (-x[4], -x[5], -len(x[0])))

ASCII_RE = re.compile(r'^[A-Za-z0-9+\-()./]+$')
def kw_in(kw, text):
    if ASCII_RE.match(kw):
        if len(kw) < 3: return False
        return re.search(r'(?<![A-Za-z0-9])' + re.escape(kw) + r'(?![A-Za-z0-9])', text) is not None
    return kw in text

CIRC = ('批发', '零售', '经销', '流通', '贸易')
PERI = ('物业', '停车', '百货', '餐饮', '酒店')
HY小类 = {
    '医疗用品及器材批发': ('service', '医药流通与供应链'),
    '医疗用品及器材零售': ('service', '医药流通与供应链'),
    '西药批发': ('service', '医药流通与供应链'),
    '中药批发': ('service', '医药流通与供应链'),
    '卫生材料及医药用品制造': ('device', '医用材料与耗材'),
}

rows = [json.loads(l) for l in open(f'{BASE}/cache/passA_intro/{BATCH}.jsonl', encoding='utf-8')]
hy = {}
_tsv = f'{BASE}/{TSV}.tsv'
if not os.path.exists(_tsv): _tsv = f'{BASE}/full/batches/{TSV}.tsv'
with open(_tsv, encoding='utf-8') as f:
    header = f.readline().rstrip('\n').split('\t')
    col_hy = header.index('行业小类') if '行业小类' in header else 2
    for line in f:
        p = line.rstrip('\n').split('\t')
        hy[int(p[0])] = p[col_hy]

def match(text, bonus=0):
    hits, seen = [], set()
    if not text: return hits
    for kw, sys_, l2, l3, w, lvl, src in kw_index:
        if kw_in(kw, text):
            key = (sys_, l2, kw)
            if key not in seen:
                seen.add(key)
                hits.append((sys_, l2, l3 if lvl == 3 else '', kw, w + bonus, src))
        if len(hits) >= 8: break
    return hits

out = []
for r in rows:
    verified = bool(r.get('_已核验'))
    kws_field = '、'.join(r.get('技术关键词') or [])
    prose = (r.get('主营业务') or '') + '。' + (r.get('一句话简介') or '')
    hits = match(kws_field, bonus=2)
    hits.sort(key=lambda h: -h[4])
    dedup, seen2 = [], set()
    for h in hits:
        if (h[0], h[1]) not in seen2:
            seen2.add((h[0], h[1])); dedup.append(h)
    hits = dedup[:6]
    prose_hits = match(prose)

    体系计 = collections.Counter(h[0] for h in hits)
    主体系 = 二级主 = 二级副 = 三级 = ''
    conf, resid, basis = '', True, ''
    if hits:
        主体系 = max(体系计, key=lambda s: (体系计[s], -[h[0] for h in hits].index(s)))
        if len(体系计) > 1:   # 多体系时选 prose 有印证的
            prose_sys = {h[0] for h in prose_hits}
            corro = [s for s in 体系计 if s in prose_sys]
            if corro and 主体系 not in corro:
                主体系 = max(corro, key=lambda s: 体系计[s])
        主hits = [h for h in hits if h[0] == 主体系]
        # v3.1b: 同体系内并列时，优先 prose 印证的二级，再带三级的，再权重
        # （东富龙案例：模型噪声词"CDMO"与正确补充二级"供应链上游配套"同权重，prose 含"制药装备"佐证后者）
        prose_l2 = {h[1] for h in prose_hits if h[1]}
        主hits.sort(key=lambda h: (-(1 if h[1] in prose_l2 else 0), -(1 if h[2] else 0), -h[4]))
        二级主, 三级 = 主hits[0][1] or '', 主hits[0][2] or ''
        副hits = [h for h in hits if h[0] != 主体系]
        二级副 = 副hits[0][1] if 副hits else ''
        src0 = 主hits[0][5]          # dict | map
        if src0 == 'dict':
            conf = '高' if verified else '中'
        else:
            conf = '中' if verified else '低'
        resid = False
        basis = '关键词命中'
        # v3.1 角色纠偏：业务角色优先于产品词
        # （甘庆案例：用PCR做科研服务≠IVD产品企业；利和案例：器械批发商≠器械企业）
        if re.search(r'经销|代理|批发|零售|分销|总代', prose) and 二级主 != '医药流通与供应链' \
                and not re.search(r'生产|制造|研发|在研|管线', prose):
            主体系, 二级主, 二级副, 三级 = 'service', '医药流通与供应链', '', ''
            conf = '中' if verified else '低'
            basis = '角色规则-流通'
        elif re.search(r'科研服务|技术服务|研究与服务|实验外包|检测服务|咨询服务', prose) and 主体系 in ('drug', 'device'):
            主体系 = 'service'
            if 二级主 not in ('CRO-发现与前期研究', 'CRO-临床前与GLP', 'CRO-临床试验与运营', 'CRO-注册法规与申报',
                              'CDMO-小分子与化学工艺', 'CDMO-生物药与生物工艺', 'CDMO-细胞与基因治疗CGT', 'CMC/质量与技术转移'):
                二级主, 三级 = 'CRO-发现与前期研究', ''
            conf = '中' if verified else '低'
            basis = '角色规则-服务'
    elif prose_hits:
        p计 = collections.Counter(h[0] for h in prose_hits)
        主体系 = max(p计, key=lambda s: p计[s])
        conf = '低'
        resid = False
        hits = prose_hits[:2]
        basis = 'prose兜底'
    elif any(k in prose for k in CIRC):
        主体系, 二级主, conf, resid = 'service', '医药流通与供应链', '低', False
        basis = '流通规则'
    elif any(k in prose for k in PERI):
        主体系, conf = '周边配套', '低'
        basis = '周边规则'
    else:
        x = hy.get(r['序号'], '')
        if x in HY小类:
            主体系, 二级主 = HY小类[x][0], HY小类[x][1]
            conf = '登记' if ('批发' in x or '零售' in x) else '低'
            basis = '登记信息'
    out.append({
        '序号': r['序号'], '公司名称': r['公司名称'],
        '命中关键词': ';'.join(h[3] for h in hits[:4]),
        '主体系': 主体系, '二级标签主': 二级主, '二级标签副': 二级副, '三级标签': 三级,
        '规则置信度': conf, '标签依据': basis, '残渣': 'Y' if resid else 'N',
    })

with open(f'{BASE}/cache/passB_rules/{BATCH}.jsonl', 'w', encoding='utf-8') as f:
    for o in out:
        f.write(json.dumps(o, ensure_ascii=False) + '\n')

c1 = collections.Counter(o['主体系'] or '未命中' for o in out)
c2 = collections.Counter(o['规则置信度'] or '空' for o in out)
print('体系分布:', dict(c1))
print('置信度:', dict(c2))
print('残渣数:', sum(1 for o in out if o['残渣'] == 'Y'))
