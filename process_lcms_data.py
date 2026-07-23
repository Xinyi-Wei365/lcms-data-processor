#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LC-MS/MS 数据处理智能体 v3
从 MassHunter 导出的原始数据生成最终计算浓度表格
自动识别化合物名称，无需预设列表
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import re
import io

# ============================================================
# 样式常量
# ============================================================
FONT_COLOR = '384350'
HEADER_BG = 'E1E3E5'
BLUE_FONT = '1F4E78'
BLUE_BG = 'D9EAF7'
GOLD_FONT = '7F6000'
GOLD_BG = 'FFF2CC'
RED_FONT = 'C00000'
YELLOW_BG = 'FFFF00'

# ============================================================
# 工具函数
# ============================================================
def safe_float(v):
    if v is None: return None
    try: return float(v)
    except: return None

def round6(v):
    if v is None: return None
    return round(float(v), 6)

def round_int(v):
    if v is None: return None
    return round(float(v))

# ============================================================
# 样式工厂
# ============================================================
def make_styles():
    base = Font(name='SimSun', size=11, color=FONT_COLOR)
    bold = Font(name='SimSun', size=11, color=FONT_COLOR, bold=True)
    red_bold = Font(name='SimSun', size=11, color=RED_FONT, bold=True)
    blue = Font(name='SimSun', size=11, color=BLUE_FONT)
    gold = Font(name='SimSun', size=11, color=GOLD_FONT)

    hdr_fill  = PatternFill(start_color=HEADER_BG, end_color=HEADER_BG, fill_type='solid')
    blue_fill = PatternFill(start_color=BLUE_BG, end_color=BLUE_BG, fill_type='solid')
    gold_fill = PatternFill(start_color=GOLD_BG, end_color=GOLD_BG, fill_type='solid')
    yell_fill = PatternFill(start_color=YELLOW_BG, end_color=YELLOW_BG, fill_type='solid')
    no_fill   = PatternFill(fill_type=None)

    ca = Alignment(horizontal='center', vertical='center')
    la = Alignment(horizontal='left', vertical='center')

    return {
        'hdr':  {'font': base, 'fill': hdr_fill,  'alignment': ca},
        'hdrL': {'font': base, 'fill': hdr_fill,  'alignment': la},
        'cmpd': {'font': base, 'fill': no_fill,   'alignment': la},
        'data': {'font': base, 'fill': no_fill,   'alignment': ca},
        'rec':  {'font': red_bold, 'fill': no_fill, 'alignment': ca},
        'stat': {'font': blue, 'fill': blue_fill, 'alignment': ca},
        'gold': {'font': gold, 'fill': gold_fill, 'alignment': ca},
        'yell': {'font': base, 'fill': yell_fill, 'alignment': ca},
        'yellL':{'font': bold, 'fill': yell_fill, 'alignment': la},
        'bold': {'font': bold, 'fill': no_fill,   'alignment': ca},
    }

def sty(cell, s):
    for k in ('font','fill','alignment'):
        if s.get(k): setattr(cell, k, s[k])


# ============================================================
# 化合物名称智能排序
# ============================================================
def classify_compounds(compounds):
    """
    自动将化合物分类并排序
    顺序: BAC -> DDAC -> ATMAC -> BAC代谢物 -> IS内标 -> SS替代物
    返回: (target_compounds, is_compounds, ss_compounds)
    """
    bac, ddac, atmac, metabolites, is_list, ss_list, others = [], [], [], [], [], [], []

    for c in compounds:
        cn = str(c).strip()
        if not cn:
            continue

        # IS 进样内标：含 COOH, +COOH, [C13], -d6, -d3, OH-d3
        if any(kw in cn for kw in ['[C13]', 'COOH', '-d6', '-d3', 'OH-d3']):
            is_list.append(cn)
        # SS 替代物：d7, d9 开头
        elif re.match(r'd\d+-', cn):
            ss_list.append(cn)
        # DDAC 系列
        elif 'DDAC' in cn.upper():
            ddac.append(cn)
        # ATMAC 系列
        elif 'ATMAC' in cn.upper():
            atmac.append(cn)
        # BAC 系列
        elif 'BAC' in cn.upper():
            bac.append(cn)
        else:
            others.append(cn)

    # 各系列按碳链长度排序
    def sort_key(name):
        nums = re.findall(r'C(\d+)', name)
        return int(nums[0]) if nums else 0

    bac.sort(key=sort_key)
    ddac.sort(key=sort_key)
    atmac.sort(key=sort_key)
    metabolites.sort(key=sort_key)
    is_list.sort(key=sort_key)
    ss_list.sort(key=sort_key)

    # 从BAC组中分离代谢产物
    pure_bac = []
    meta = []
    for c in bac:
        if '+O' in c or '+2O' in c:
            meta.append(c)
        else:
            pure_bac.append(c)
    metabolites = meta + metabolites

    target = pure_bac + ddac + atmac + metabolites + others
    return target, is_list, ss_list


# ============================================================
# 原始数据读取
# ============================================================
def read_raw(filepath_or_bytes):
    """读取原始数据，支持文件路径或 bytes"""
    if isinstance(filepath_or_bytes, bytes):
        wb = openpyxl.load_workbook(io.BytesIO(filepath_or_bytes), data_only=True)
    else:
        wb = openpyxl.load_workbook(filepath_or_bytes, data_only=True)
    ws = wb['Sheet1']

    blanks, mss, samps = [], [], []
    for col in range(4, ws.max_column + 1):
        hdr = str(ws.cell(row=1, column=col).value or '')
        cl = get_column_letter(col)
        if 'BLANK' in hdr.upper():
            blanks.append((col, cl, hdr))
        elif 'MS' in hdr.upper():
            mss.append((col, cl, hdr))
        elif re.search(r'\dPPB', hdr.upper()):
            continue
        else:
            samps.append((col, cl, hdr))

    # 读取化合物名称
    compounds_raw = []
    for row in range(3, ws.max_row + 1):
        nm = str(ws.cell(row=row, column=2).value or '').strip()
        if nm:
            compounds_raw.append(nm)

    # 读取数据
    data = {}
    for row in range(3, ws.max_row + 1):
        nm = str(ws.cell(row=row, column=2).value or '').strip()
        if not nm: continue
        data[nm] = {}
        for col in range(4, ws.max_column + 1):
            data[nm][get_column_letter(col)] = ws.cell(row=row, column=col).value
    wb.close()

    # 自动分类
    target, is_c, ss_c = classify_compounds(compounds_raw)
    all_c = target + is_c + ss_c

    return data, blanks, mss, samps, target, is_c, ss_c, all_c


# ============================================================
# Sheet 1: 基质加标浓度
# ============================================================
def build_sheet1(wb, raw_data, ms_cols, all_compounds, S, spike):
    ws = wb.create_sheet('Matrix spike  基质加标浓度')
    n_ms = len(ms_cols)

    ms_start, mid1 = 2, 2 + n_ms
    rec_lbl, rec_start = mid1 + 1, mid1 + 2
    mid2 = rec_start + n_ms
    stat_col, sd_col, se_col = mid2 + 1, mid2 + 2, mid2 + 3
    last_col = se_col

    ws.cell(row=1, column=1, value='化合物方法'); sty(ws.cell(row=1,column=1), S['hdrL'])
    for i in range(n_ms):
        ws.cell(row=1, column=ms_start+i, value=f'matrix spike_{i+1}'); sty(ws.cell(row=1,column=ms_start+i), S['hdr'])
    ws.cell(row=1, column=mid1, value=None)
    ws.cell(row=1, column=rec_lbl, value='Recoveries'); sty(ws.cell(row=1,column=rec_lbl), S['hdr'])
    for i in range(n_ms):
        ws.cell(row=1, column=rec_start+i, value=f'matrix spike_{i+1}'); sty(ws.cell(row=1,column=rec_start+i), S['hdr'])
    ws.cell(row=1, column=mid2, value=None)
    ws.cell(row=1, column=stat_col, value='average'); sty(ws.cell(row=1,column=stat_col), S['hdr'])
    ws.cell(row=1, column=sd_col, value='SD'); sty(ws.cell(row=1,column=sd_col), S['hdr'])
    ws.cell(row=1, column=se_col, value='SE'); sty(ws.cell(row=1,column=se_col), S['hdr'])

    ws.cell(row=2, column=1, value='分组'); sty(ws.cell(row=2,column=1), S['hdr'])
    for i in range(n_ms):
        ws.cell(row=2, column=ms_start+i, value='ppb'); sty(ws.cell(row=2,column=ms_start+i), S['hdr'])
    for i in range(n_ms):
        ws.cell(row=2, column=rec_start+i, value='%'); sty(ws.cell(row=2,column=rec_start+i), S['hdr'])

    rec_cl = get_column_letter(rec_start)
    rec_cr = get_column_letter(rec_start + n_ms - 1)

    row = 3
    for comp in all_compounds:
        ws.cell(row=row, column=1, value=comp); sty(ws.cell(row=row,column=1), S['cmpd'])
        for i, (_, cl, _) in enumerate(ms_cols):
            v = safe_float(raw_data.get(comp, {}).get(cl))
            c = ws.cell(row=row, column=ms_start+i)
            if v is not None: c.value = round6(v); c.number_format = '0.000000'
            sty(c, S['data'])
        ws.cell(row=row, column=rec_lbl, value=None)
        for i, (_, cl, _) in enumerate(ms_cols):
            v = safe_float(raw_data.get(comp, {}).get(cl))
            c = ws.cell(row=row, column=rec_start+i)
            if v is not None: c.value = round_int((v / spike) * 100); c.number_format = '0'
            sty(c, S['rec'])
        rr = f'{rec_cl}{row}:{rec_cr}{row}'
        ws.cell(row=row, column=stat_col, value=f'=ROUND(AVERAGE({rr}),0)'); ws.cell(row=row,column=stat_col).number_format = '0'
        ws.cell(row=row, column=sd_col, value=f'=ROUND(STDEV.S({rr}),0)'); ws.cell(row=row,column=sd_col).number_format = '0'
        ws.cell(row=row, column=se_col, value=f'=ROUND({get_column_letter(sd_col)}{row}/SQRT(COUNT({rr})),0)'); ws.cell(row=row,column=se_col).number_format = '0'
        sty(ws.cell(row=row, column=stat_col), S['data'])
        sty(ws.cell(row=row, column=sd_col), S['data'])
        sty(ws.cell(row=row, column=se_col), S['data'])
        row += 1

    # SS 回收率部分
    ss_comps = [c for c in all_compounds if re.match(r'd\d+-', c)]
    if ss_comps:
        row += 1
        ws.cell(row=row, column=1, value='SS recoveries, %  替代物回收率'); sty(ws.cell(row=row,column=1), S['yellL'])
        for c in range(2, last_col+1): sty(ws.cell(row=row,column=c), S['yell'])
        row += 1
        for ss in ss_comps:
            ws.cell(row=row, column=1, value=ss); sty(ws.cell(row=row,column=1), S['cmpd'])
            for i, (_, cl, _) in enumerate(ms_cols):
                v = safe_float(raw_data.get(ss, {}).get(cl))
                c = ws.cell(row=row, column=ms_start+i)
                if v is not None: c.value = round6(v / spike); c.number_format = '0.000000'
                sty(c, S['data'])
            for i, (_, cl, _) in enumerate(ms_cols):
                v = safe_float(raw_data.get(ss, {}).get(cl))
                c = ws.cell(row=row, column=rec_start+i)
                if v is not None: c.value = round_int((v / spike) * 100); c.number_format = '0'
                sty(c, S['rec'])
            row += 1

    row += 1
    note = rec_start + n_ms
    ws.cell(row=row, column=note, value='此表格计算方法：回收率，用测得浓度除以加标浓度'); sty(ws.cell(row=row,column=note), S['yell'])
    row += 1
    ws.cell(row=row, column=note, value=f'此处实验基质加标浓度都是{spike}，所以计算回收率时每个化合物的测得浓度除以{spike}'); sty(ws.cell(row=row,column=note), S['yell'])

    ws.row_dimensions[1].height = 19.5
    ws.row_dimensions[2].height = 17.25
    ws.column_dimensions['A'].width = 30.0
    return ws


# ============================================================
# Sheet 2: 空白基质检出限
# ============================================================
def build_sheet2(wb, raw_data, blank_cols, all_compounds, S, cf, unit):
    ws = wb.create_sheet('Blanks_MDL 空白基质检出限')
    n_b = len(blank_cols)

    ws.cell(row=1, column=1, value=cf); ws.cell(row=1,column=1).number_format = '0.000000'
    ws.cell(row=1, column=9, value='仅为检出率>50% ND取代'); sty(ws.cell(row=1,column=9), S['yell'])

    blank_end = 2 + n_b
    avg_c, mdl_c, half_c = blank_end + 1, blank_end + 2, blank_end + 3
    last_c = half_c

    ws.cell(row=2, column=1, value='化合物方法'); sty(ws.cell(row=2,column=1), S['hdr'])
    for i in range(n_b):
        ws.cell(row=2, column=2+i, value=f'blank_{i+1}'); sty(ws.cell(row=2,column=2+i), S['hdr'])
    ws.cell(row=2, column=avg_c, value='procedural blank average'); sty(ws.cell(row=2,column=avg_c), S['hdr'])
    ws.cell(row=2, column=mdl_c, value='MDL'); sty(ws.cell(row=2,column=mdl_c), S['hdr'])
    ws.cell(row=2, column=half_c, value='1/2 MDL'); sty(ws.cell(row=2,column=half_c), S['hdr'])

    ws.cell(row=3, column=1, value='分组'); sty(ws.cell(row=3,column=1), S['hdr'])
    for i in range(n_b):
        ws.cell(row=3, column=2+i, value='ppb'); sty(ws.cell(row=3,column=2+i), S['hdr'])
    ws.cell(row=3, column=avg_c, value='bottle ppb'); sty(ws.cell(row=3,column=avg_c), S['hdr'])
    ws.cell(row=3, column=mdl_c, value='bottle ppb'); sty(ws.cell(row=3,column=mdl_c), S['hdr'])
    ws.cell(row=3, column=half_c, value=unit); sty(ws.cell(row=3,column=half_c), S['hdr'])

    blank_l = get_column_letter(2)
    blank_r = get_column_letter(2 + n_b - 1)
    avg_l = get_column_letter(avg_c)
    mdl_l = get_column_letter(mdl_c)

    row = 4
    for comp in all_compounds:
        ws.cell(row=row, column=1, value=comp); sty(ws.cell(row=row,column=1), S['cmpd'])
        for i, (_, cl, _) in enumerate(blank_cols):
            v = safe_float(raw_data.get(comp, {}).get(cl))
            c = ws.cell(row=row, column=2+i)
            if v is not None: c.value = round6(v); c.number_format = '0.000000'
            sty(c, S['data'])
        br = f'{blank_l}{row}:{blank_r}{row}'
        ws.cell(row=row, column=avg_c, value=f'=AVERAGE({br})'); ws.cell(row=row,column=avg_c).number_format = '0.000000'
        ws.cell(row=row, column=mdl_c, value=f'=3*STDEVA({br})'); ws.cell(row=row,column=mdl_c).number_format = '0.000000'
        ws.cell(row=row, column=half_c, value=f'=ROUND({mdl_l}{row}/2*$A$1,6)'); ws.cell(row=row,column=half_c).number_format = '0.000000'
        sty(ws.cell(row=row, column=avg_c), S['data'])
        sty(ws.cell(row=row, column=mdl_c), S['data'])
        sty(ws.cell(row=row, column=half_c), S['data'])
        row += 1

    ws.row_dimensions[1].height = 19.5
    ws.row_dimensions[2].height = 17.25
    ws.column_dimensions['A'].width = 30.0

    return ws, {'avg_c': avg_c, 'mdl_c': mdl_c, 'half_c': half_c, 'avg_l': avg_l, 'mdl_l': mdl_l, 'half_l': get_column_letter(half_c)}


# ============================================================
# Sheet 3: 瓶内实测浓度
# ============================================================
def build_sheet3(wb, raw_data, sample_cols, target_compounds, S, mh_unit):
    ws = wb.create_sheet('Conc. in bottle 瓶内实测浓度')
    n_s = len(sample_cols)

    ws.cell(row=1, column=1, value='化合物方法'); sty(ws.cell(row=1,column=1), S['hdr'])
    for i, (_, _, hdr) in enumerate(sample_cols):
        ws.cell(row=1, column=2+i, value=hdr); sty(ws.cell(row=1,column=2+i), S['hdr'])
    ws.cell(row=2, column=1, value='分组'); sty(ws.cell(row=2,column=1), S['hdr'])
    for i in range(n_s):
        ws.cell(row=2, column=2+i, value=mh_unit); sty(ws.cell(row=2,column=2+i), S['hdr'])

    row = 3
    for comp in target_compounds:
        ws.cell(row=row, column=1, value=comp); sty(ws.cell(row=row,column=1), S['cmpd'])
        for i, (_, cl, _) in enumerate(sample_cols):
            v = safe_float(raw_data.get(comp, {}).get(cl))
            c = ws.cell(row=row, column=2+i)
            if v is not None: c.value = round6(v); c.number_format = '0.000000'
            sty(c, S['data'])
        row += 1

    ws.row_dimensions[1].height = 19.5
    ws.row_dimensions[2].height = 17.25
    ws.column_dimensions['A'].width = 30.0
    return ws


# ============================================================
# Sheet 4: 最终计算浓度
# ============================================================
def build_sheet4(wb, raw_data, sample_cols, target_compounds, all_compounds, blank_info, S, cf, unit):
    ws = wb.create_sheet('Final. conc 最终计算浓度')
    n_s = len(sample_cols)
    sample_start = 16  # P列
    last_sample = sample_start + n_s - 1

    blanks_name = 'Blanks_MDL 空白基质检出限'
    bottle_name = 'Conc. in bottle 瓶内实测浓度'
    al, ml, hl = blank_info['avg_l'], blank_info['mdl_l'], blank_info['half_l']

    ws.cell(row=1, column=1, value='化合物方法'); sty(ws.cell(row=1,column=1), S['hdr'])
    ws.cell(row=1, column=2, value='bottle'); sty(ws.cell(row=1,column=2), S['hdr'])
    ws.cell(row=1, column=3, value='bottle'); sty(ws.cell(row=1,column=3), S['hdr'])
    for i, (_, _, hdr) in enumerate(sample_cols):
        ws.cell(row=1, column=sample_start+i, value=hdr); sty(ws.cell(row=1,column=sample_start+i), S['hdr'])

    stat_names = ['DF 检出率','MEAN','Geometric Mean','MEDIAN','MIN','MAX','5TH','25TH','75TH','95TH']
    ws.cell(row=2, column=1, value='分组'); sty(ws.cell(row=2,column=1), S['hdr'])
    ws.cell(row=2, column=2, value='BLANK average'); sty(ws.cell(row=2,column=2), S['hdr'])
    ws.cell(row=2, column=3, value='MDL'); sty(ws.cell(row=2,column=3), S['hdr'])
    for i, nm in enumerate(stat_names):
        ws.cell(row=2, column=4+i, value=nm); sty(ws.cell(row=2,column=4+i), S['stat'])
    for i in range(n_s):
        ws.cell(row=2, column=sample_start+i, value=1); sty(ws.cell(row=2,column=sample_start+i), S['data'])

    # 合并统计表头 Row1+Row2
    for col in range(4, 4 + len(stat_names)):
        ws.merge_cells(start_row=1, start_column=col, end_row=2, end_column=col)
        ws.cell(row=1, column=col, value=stat_names[col-4]); sty(ws.cell(row=1,column=col), S['stat'])

    ws.cell(row=3, column=1, value=''); sty(ws.cell(row=3,column=1), S['cmpd'])
    ws.cell(row=3, column=2, value='ng/mL'); sty(ws.cell(row=3,column=2), S['hdr'])
    ws.cell(row=3, column=3, value='ng/mL'); sty(ws.cell(row=3,column=3), S['hdr'])
    ws.cell(row=3, column=4, value='%'); sty(ws.cell(row=3,column=4), S['stat'])
    for c in range(5, 14):
        ws.cell(row=3, column=c, value=unit); sty(ws.cell(row=3,column=c), S['stat'])

    ws.cell(row=38, column=1, value='换算因子'); sty(ws.cell(row=38,column=1), S['bold'])
    ws.cell(row=38, column=2, value=cf); ws.cell(row=38,column=2).number_format = '0.000000'
    sty(ws.cell(row=38,column=2), S['data'])

    # 行号映射：Sheet2 从 row 4 开始，Sheet3 从 row 3 开始
    comp2s2row = {c: 4 + i for i, c in enumerate(all_compounds)}
    comp2s3row = {c: 3 + i for i, c in enumerate(target_compounds)}

    row = 4
    sl = get_column_letter(sample_start)
    el = get_column_letter(last_sample)

    for comp in target_compounds:
        ws.cell(row=row, column=1, value=comp); sty(ws.cell(row=row,column=1), S['cmpd'])
        s2r = comp2s2row.get(comp)
        s3r = comp2s3row.get(comp)

        if s2r:
            ws.cell(row=row, column=2, value=f"='{blanks_name}'!{al}{s2r}"); ws.cell(row=row,column=2).number_format = '0.000000'
            ws.cell(row=row, column=3, value=f"='{blanks_name}'!{ml}{s2r}"); ws.cell(row=row,column=3).number_format = '0.000000'
        sty(ws.cell(row=row, column=2), S['data'])
        sty(ws.cell(row=row, column=3), S['data'])

        sr = f'{sl}{row}:{el}{row}'

        ws.cell(row=row, column=4, value=f'=COUNT({sr})/COLUMNS({sr})'); ws.cell(row=row,column=4).number_format = '0.00%'
        sty(ws.cell(row=row, column=4), S['stat'])

        funcs = {5:'AVERAGE', 6:'GEOMEAN', 7:'MEDIAN', 8:'MIN', 9:'MAX'}
        for col, func in funcs.items():
            ws.cell(row=row, column=col, value=f'=IF($D{row}>50%,{func}({sr}),"NC")'); ws.cell(row=row,column=col).number_format = '0.000000'
            sty(ws.cell(row=row, column=col), S['stat'])

        for col, pct in [(10,0.05),(11,0.25),(12,0.75),(13,0.95)]:
            ws.cell(row=row, column=col, value=f'=IF($D{row}>50%,PERCENTILE.INC({sr},{pct}),"NC")'); ws.cell(row=row,column=col).number_format = '0.000000'
            sty(ws.cell(row=row, column=col), S['stat'])

        for i in range(n_s):
            col = sample_start + i
            s3_cl = get_column_letter(2 + i)
            if s2r and s3r:
                ws.cell(row=row, column=col, value=(
                    f"=IF('{bottle_name}'!{s3_cl}{s3r}=\"\",\"\","
                    f"IF('{bottle_name}'!{s3_cl}{s3r}>'{blanks_name}'!{al}{s2r},"
                    f"('{bottle_name}'!{s3_cl}{s3r}-'{blanks_name}'!{al}{s2r})*$B$38,"
                    f"'{blanks_name}'!{hl}{s2r}*$B$38))")
                ); ws.cell(row=row, column=col).number_format = '0.000000'
            sty(ws.cell(row=row, column=col), S['data'])
        row += 1

    ws.row_dimensions[1].height = 19.5
    ws.row_dimensions[2].height = 17.25
    ws.column_dimensions['A'].width = 30.0
    return ws


# ============================================================
# Sheet 5: 统计辅助
# ============================================================
def build_sheet5(wb, sample_cols, target_compounds, S):
    ws = wb.create_sheet('统计计算结果')
    n_s = len(sample_cols)
    fn = 'Final. conc 最终计算浓度'
    ss = 16

    ws.cell(row=1, column=1, value='统计数据源（供统计列使用）')
    ws.cell(row=2, column=1, value='样品名称'); sty(ws.cell(row=2,column=1), S['hdr'])
    for i in range(n_s):
        ws.cell(row=2, column=2+i, value=f"='{fn}'!{get_column_letter(ss+i)}$1"); sty(ws.cell(row=2,column=2+i), S['hdr'])

    row = 3
    for idx, comp in enumerate(target_compounds):
        fr = 4 + idx
        ws.cell(row=row, column=1, value=comp); sty(ws.cell(row=row,column=1), S['cmpd'])
        for i in range(n_s):
            ws.cell(row=row, column=2+i, value=f"='{fn}'!{get_column_letter(ss+i)}{fr}"); sty(ws.cell(row=row,column=2+i), S['data'])
        row += 1

    ws.column_dimensions['A'].width = 30.0
    return ws


# ============================================================
# 说明书
# ============================================================
def build_info(wb, S, cfg, n_target, n_all):
    ws = wb.create_sheet('计算说明')
    rows = [
        ['LC-MS/MS 数据处理说明'],
        ['区域','公式/规则'],
        ['Sheet1','基质加标浓度 + 回收率(%) + avg/SD/SE (整数)'],
        ['Sheet2','空白值 + MDL(=3*STDEVA) + 1/2MDL(=MDL/2*换算因子)'],
        ['Sheet3','瓶内实测浓度 (原始数据直接迁移)'],
        ['Sheet4','最终计算浓度: (瓶内值-空白avg)*换算因子, 否则用1/2MDL*换算因子'],
        ['Sheet4统计','DF=检出率; IF DF>50%则计算MEAN/GEOMEAN/MEDIAN/PERCENTILE等'],
        ['换算因子','Sheet2 $A$1 和 Sheet4 $B$38 两处独立存储'],
        ['本次参数',f'样本:{cfg.get("sample_type","")} 换算因子:{cfg.get("conversion_factor",1)} 加标:{cfg.get("spike_conc_ppb",10)}ppb'],
        [f'化合物数',f'目标:{n_target} 全部(含IS/SS):{n_all}'],
    ]
    for r, rd in enumerate(rows, 1):
        for c, val in enumerate(rd, 1):
            sty(ws.cell(row=r, column=c, value=val), S['hdr'] if r<=2 else S['data'])
    for c, w in zip([1,2], [35,65]):
        ws.column_dimensions[get_column_letter(c)].width = w
    return ws


# ============================================================
# 主处理函数
# ============================================================
def process(config=None, return_bytes=False):
    cfg = config or {}
    spike = cfg.get('spike_conc_ppb', 10)
    cf = cfg.get('conversion_factor', 1)
    unit = cfg.get('output_unit', 'ng/mL')
    mh_unit = cfg.get('masshunter_unit', 'ppb')
    input_src = cfg.get('input_file', '')
    input_bytes = cfg.get('input_bytes', None)

    # 读取数据（支持文件路径或 bytes）
    src = input_bytes if input_bytes else input_src
    raw_data, blanks, mss, samps, target, is_c, ss_c, all_c = read_raw(src)

    S = make_styles()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    build_sheet1(wb, raw_data, mss, all_c, S, spike)
    _, binfo = build_sheet2(wb, raw_data, blanks, all_c, S, cf, unit)
    build_sheet3(wb, raw_data, samps, target, S, mh_unit)
    build_sheet4(wb, raw_data, samps, target, all_c, binfo, S, cf, unit)
    build_sheet5(wb, samps, target, S)
    build_info(wb, S, cfg, len(target), len(all_c))

    if return_bytes:
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        wb.close()
        return output.getvalue(), cfg.get('output_file', 'processed_data.xlsx')
    else:
        out = cfg.get('output_file', 'output.xlsx')
        wb.save(out)
        wb.close()
        return out


if __name__ == '__main__':
    import os
    process({
        'sample_type': '尿液',
        'sample_volume_ml': 2,
        'final_volume_ml': 0.5,
        'conversion_factor': 1,
        'spike_conc_ppb': 10,
        'masshunter_unit': 'ppb',
        'output_unit': 'ng/mL',
        'input_file': r'C:\Users\HP-PC\Desktop\未整理 - 副本.xlsx',
        'output_file': r'C:\Users\HP-PC\Desktop\已处理_尿液数据.xlsx',
    })
