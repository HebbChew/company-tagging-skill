#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""机构类型直标（零搜索）：有明确机构类型的企业不核验直接打体系级标签。
规则：
  医院/门诊部/诊所/基层医疗/其他卫生/医疗管理/健康管理 → service / 医疗服务（补充）
  药房/药店 → service / 医药供应链与流通（补充）
  一律无二级、置信度=低、依据="机构类型直标"
应用范围：未搜索企业 + 已搜索但无公开信息且无标签的企业（不覆盖已有更强标签）。
同时把未搜索且无机构类型的标为"未核验（待补搜）"。
用法: MHB_TAG_BASE=<目录> MHB_SOURCE_XLSX=<源表> MHB_FINAL_XLSX=<终稿> python3 apply_机构类型.py
"""
import os, openpyxl

BASE = os.environ.get('MHB_TAG_BASE', os.path.dirname(os.path.abspath(__file__)))
SRC = os.environ['MHB_SOURCE_XLSX']
FIN = os.environ.get('MHB_FINAL_XLSX', f'{BASE}/full/终稿_企业标注.xlsx')

swb = openpyxl.load_workbook(SRC, read_only=True)
sws = wb0 = swb['P1强相关企业']
hdr = [c for c in next(sws.iter_rows(min_row=1, max_row=1, values_only=True))]
ix = {h: i for i, h in enumerate(hdr)}
orgtype = {int(r[ix['序号']]): str(r[ix['机构类型']] or '').strip() for r in sws.iter_rows(min_row=2, values_only=True)} \
    if '机构类型' in ix else {}

def tag_of(ot):
    if ot in ('医院', '门诊部', '诊所', '基层医疗', '其他卫生', '医疗管理', '健康管理'):
        return ('service', '医疗服务（补充）', '机构类型直标')
    if ot in ('药房', '药店'):
        return ('service', '医药供应链与流通（补充）', '机构类型直标')
    return None

wb = openpyxl.load_workbook(FIN)
direct, applied, wait = 0, 0, 0
ws = wb['终稿']
HDR = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
jx = {h: i for i, h in enumerate(HDR)}
col = {h: jx[h] + 1 for h in HDR}
for r in ws.iter_rows(min_row=2):
    sid = int(r[col['序号'] - 1].value)
    ot = orgtype.get(sid, '')
    searched = str(r[col['搜索层级'] - 1].value or '') not in ('', '未搜索')
    info = str(r[col['信息状态'] - 1].value or '')
    has_tag = bool(str(r[col['主体系'] - 1].value or ''))
    if not searched:
        if ot and tag_of(ot):
            sys_, l1, basis = tag_of(ot)
            r[col['主体系'] - 1].value = sys_
            r[col['一级'] - 1].value = l1
            r[col['标签置信度'] - 1].value = '低'
            r[col['标签依据'] - 1].value = basis
            r[col['信息状态'] - 1].value = '未核验（机构类型直标）'
            direct += 1
        else:
            r[col['信息状态'] - 1].value = '未核验（待补搜）'
            wait += 1
    elif info == '无公开信息' and ot and not has_tag and tag_of(ot):
        sys_, l1, basis = tag_of(ot)
        r[col['主体系'] - 1].value = sys_
        r[col['一级'] - 1].value = l1
        r[col['标签置信度'] - 1].value = '低'
        r[col['标签依据'] - 1].value = basis
        applied += 1

# sheet1（无标签依据列）
if '原表+标注' in wb.sheetnames:
    ws0 = wb['原表+标注']
    H0 = [c.value for c in next(ws0.iter_rows(min_row=1, max_row=1))]
    k0 = {h: i for i, h in enumerate(H0)}
    c0 = {h: k0[h] + 1 for h in H0}
    for r in ws0.iter_rows(min_row=2):
        sid = int(r[c0['序号'] - 1].value)
        ot = orgtype.get(sid, '')
        searched = str(r[c0['搜索层级'] - 1].value or '') not in ('', '未搜索')
        if not searched:
            if ot and tag_of(ot):
                sys_, l1, basis = tag_of(ot)
                r[c0['主体系'] - 1].value = sys_
                r[c0['一级'] - 1].value = l1
                r[c0['标签置信度'] - 1].value = '低'
                r[c0['信息状态'] - 1].value = '未核验（机构类型直标）'
            else:
                r[c0['信息状态'] - 1].value = '未核验（待补搜）'
wb.save(FIN)
print(f'机构类型直标: 未核验直标 {direct} 家, 无信息补标 {applied} 家, 待补搜 {wait} 家')
