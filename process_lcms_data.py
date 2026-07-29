#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LC-MS/MS 数据处理智能体 v2
从 MassHunter 导出的原始数据生成最终计算浓度表格

用法：
  python process_lcms_data.py              # 使用 CONFIG 中的参数直接运行
  from process_lcms_data import process     # 作为模块导入，供 Streamlit 调用
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import statistics
import math
import re
import io

# ============================================================
# 可配置参数
# ============================================================
CONFIG = {
    'sample_type': '尿液',
    'sample_volume_ml': 2,
    'final_volume_ml': 0.5,
    'extra_dilution': 1,
    'conversion_factor': 1,
    'spike_conc_ppb': 10,
    'masshunter_unit': 'ppb',
    'output_unit': 'ng/mL',
    'blank_handling': 'ND',
    'input_file': r'C:\Users\HP-PC\Desktop\未整理 - 副本.xlsx',
    'output_file': r'C:\Users\HP-PC\Desktop\已处理_尿液数据.xlsx',
}

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
# 化合物分组（按输出顺序）
# ============================================================
BAC_SERIES   = ['C8-BAC','C10-BAC','C12-BAC','C14-BAC','C16-BAC','C18-BAC']
DDAC_SERIES  = ['C8-DDAC','C8-10-DDAC','C10-DDAC','C12-DDAC','C14-DDAC','C16-DDAC','C18-DDAC']
ATMAC_SERIES = ['C8-ATMAC','C10-ATMAC','C12-ATMAC','C14-ATMAC','C16-ATMAC','C18-ATMAC']
BAC_COOH     = ['C10-BAC+2O-2H','C12-BAC+2O-2H','C14-BAC+2O-2H']
BAC_OH       = ['C10-BAC+O','C12-BAC+O','C14-BAC+O']
IS_COMPS     = ['C10-BAC-COOH[C13]','C12-BAC+COOH-d6','C12-BAC+OH-d3']
SS_COMPS     = ['d7-C12-BAC','d9-C10-ATMAC']

TARGET_COMPS = BAC_SERIES + DDAC_SERIES + ATMAC_SERIES + BAC_COOH + BAC_OH
ALL_COMPS    = TARGET_COMPS + IS_COMPS + SS_COMPS


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
# 样式
# ============================================================
def make_styles():
    base = Font(name='宋体', size=11, color=FONT_COLOR)
    bold = Font(name='宋体', size=11, color=FONT_COLOR, bold=True)
    red_bold = Font(name='宋体', size=11, color=RED_FONT, bold=True)
    blue = Font(name='宋体', size=11, color=BLUE_FONT)
    gold = Font(name='宋体', size=11, color=GOLD_FONT)

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
        'statB':{'font': blue, 'fill': blue_fill, 'alignment': ca},  # stat header bold
        'gold': {'font': gold, 'fill': gold_fill, 'alignment': ca},
        'yell': {'font': base, 'fill': yell_fill, 'alignment': ca},
        'yellL':{'font': bold, 'fill': yell_fill, 'alignment': la},
        'bold': {'font': bold, 'fill': no_fill,   'alignment': ca},
    }

def sty(cell, s):  # apply style dict
    for k in ('font','fill','alignment'):
        if s.get(k): setattr(cell, k, s[k])


# ============================================================
# 原始数据读取
# ============================================================
def read_raw(filepath_or_bytes):
    if isinstance(filepath_or_bytes, bytes):
        wb = openpyxl.load_workbook(io.BytesIO(filepath_or_bytes), data_only=True)
    else:
        wb = openpyxl.load_workbook(filepath_or_bytes, data_only=True)
    ws = wb['Sheet1']

    compound_col = 2; data_start_col = 4
    for col in range(1, ws.max_column + 1):
        v = str(ws.cell(row=2, column=col).value or '').strip()
        if v == '名称':
            compound_col = col; data_start_col = col + 1
            next_v = str(ws.cell(row=2, column=col+1).value or '').strip()
            if '离子' in next_v or next_v == '':
                data_start_col = col + 2 if col + 2 <= ws.max_column else col + 1
            break

    blanks, mss, samps = [], [], []
    for col in range(data_start_col, ws.max_column + 1):
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

    data = {}
    for row in range(3, ws.max_row + 1):
        nm = str(ws.cell(row=row, column=compound_col).value or '').strip()
        if not nm: continue
        data[nm] = {}
        for col in range(data_start_col, ws.max_column + 1):
            data[nm][get_column_letter(col)] = ws.cell(row=row, column=col).value
    # 读取化合物名称用于分类
    compounds_raw = []
    for row in range(3, ws.max_row + 1):
        nm = str(ws.cell(row=row, column=compound_col).value or '').strip()
        if nm: compounds_raw.append(nm)
    wb.close()

    target, is_c, ss_c = classify_compounds(compounds_raw)
    all_c = target + is_c + ss_c
    return data, blanks, mss, samps, target, is_c, ss_c, all_c


# ============================================================
# 化合物自动分类
# ============================================================
def classify_compounds(compounds):
    bac, ddac, atmac, metabolites, is_list, ss_list, others = [], [], [], [], [], [], []
    for c in compounds:
        cn = str(c).strip()
        if not cn: continue
        if any(kw in cn for kw in ['[C13]', 'COOH', '-d6', '-d3', 'OH-d3']): is_list.append(cn)
        elif re.match(r'd\d+-', cn): ss_list.append(cn)
        elif 'DDAC' in cn.upper(): ddac.append(cn)
        elif 'ATMAC' in cn.upper(): atmac.append(cn)
        elif 'BAC' in cn.upper(): bac.append(cn)
        else: others.append(cn)
    def sk(n):
        nums = re.findall(r'C(\d+)', n)
        return int(nums[0]) if nums else 0
    bac.sort(key=sk); ddac.sort(key=sk); atmac.sort(key=sk)
    metabolites.sort(key=sk); is_list.sort(key=sk); ss_list.sort(key=sk)
    pure_bac = []; meta = []
    for c in bac:
        if '+O' in c or '+2O' in c: meta.append(c)
        else: pure_bac.append(c)
    metabolites = meta + metabolites
    target = pure_bac + ddac + atmac + metabolites + others
    return target, is_list, ss_list


# ============================================================
# Sheet 1: 基质加标浓度
# ============================================================
def build_sheet1(wb, raw_data, ms_cols, S, cfg):
    ws = wb.create_sheet('Matrix spike  基质加标浓度')
    n_ms = len(ms_cols)
    spike = cfg.get('spike_conc_ppb', 10)
    ss_spike_d7 = cfg.get('ss_spike_d7_ppb', 4)
    ss_spike_d9 = cfg.get('ss_spike_d9_ppb', 4)

    # 列布局: A | B~(n_ms) MS data | 空 | Recoveries标签 | n_ms个回收率% | 空 | avg | SD | SE
    ms_start = 2
    mid1 = ms_start + n_ms           # 空列
    rec_lbl = mid1 + 1               # "Recoveries"
    rec_start = rec_lbl + 1          # 回收率% 起始
    mid2 = rec_start + n_ms          # 空列
    stat_col = mid2 + 1              # avg
    sd_col = stat_col + 1            # SD
    se_col = stat_col + 2            # SE
    last_col = se_col

    # Row 1
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

    # Row 2
    ws.cell(row=2, column=1, value='分组'); sty(ws.cell(row=2,column=1), S['hdr'])
    for i in range(n_ms):
        ws.cell(row=2, column=ms_start+i, value='ppb'); sty(ws.cell(row=2,column=ms_start+i), S['hdr'])
    ws.cell(row=2, column=mid1, value=None)
    ws.cell(row=2, column=rec_lbl, value=None)
    for i in range(n_ms):
        ws.cell(row=2, column=rec_start+i, value='%'); sty(ws.cell(row=2,column=rec_start+i), S['hdr'])
    ws.cell(row=2, column=mid2, value=None)
    ws.cell(row=2, column=stat_col, value='%'); sty(ws.cell(row=2,column=stat_col), S['hdr'])
    ws.cell(row=2, column=sd_col, value='%'); sty(ws.cell(row=2,column=sd_col), S['hdr'])
    ws.cell(row=2, column=se_col, value='%'); sty(ws.cell(row=2,column=se_col), S['hdr'])

    rec_cl = get_column_letter(rec_start)
    rec_cr = get_column_letter(rec_start + n_ms - 1)

    # 数据行（跳过 SS，SS 单独放底部）
    row = 3
    non_ss = [c for c in ALL_COMPS if c not in SS_COMPS]
    for comp in non_ss:
        ws.cell(row=row, column=1, value=comp); sty(ws.cell(row=row,column=1), S['cmpd'])

        # MS 原始数据
        for i, (_, cl, _) in enumerate(ms_cols):
            v = safe_float(raw_data.get(comp, {}).get(cl))
            c = ws.cell(row=row, column=ms_start+i)
            if v is not None:
                c.value = round6(v); c.number_format = '0.000000'
            sty(c, S['data'])

        ws.cell(row=row, column=mid1, value=None)
        ws.cell(row=row, column=rec_lbl, value=None)

        # 回收率 % = (MS值/spike)*100, 四舍五入取整
        for i, (_, cl, _) in enumerate(ms_cols):
            v = safe_float(raw_data.get(comp, {}).get(cl))
            c = ws.cell(row=row, column=rec_start+i)
            if v is not None:
                c.value = round_int((v / spike) * 100)
                c.number_format = '0'
            sty(c, S['rec'])

        ws.cell(row=row, column=mid2, value=None)

        # 统计公式 (基于回收率%列)
        rr = f'{rec_cl}{row}:{rec_cr}{row}'
        c_avg = ws.cell(row=row, column=stat_col)
        c_avg.value = f'=ROUND(AVERAGE({rr}),0)'; c_avg.number_format = '0'
        sty(c_avg, S['data'])

        c_sd = ws.cell(row=row, column=sd_col)
        c_sd.value = f'=ROUND(STDEV({rr}),0)'; c_sd.number_format = '0'
        sty(c_sd, S['data'])

        c_se = ws.cell(row=row, column=se_col)
        c_se.value = f'=ROUND({get_column_letter(sd_col)}{row}/SQRT(COUNT({rr})),0)'; c_se.number_format = '0'
        sty(c_se, S['data'])

        row += 1

    # --- SS 替代物回收率部分（独立于上方主列表） ---
    row += 1
    ws.cell(row=row, column=1, value='SS recoveries, %  替代物回收率')
    sty(ws.cell(row=row,column=1), S['yellL'])
    for c in range(2, last_col+1):
        sty(ws.cell(row=row,column=c), S['yell'])
    row += 1

    for ss in SS_COMPS:
        ws.cell(row=row, column=1, value=ss); sty(ws.cell(row=row,column=1), S['cmpd'])
        this_ss_spike = ss_spike_d7 if 'd7' in ss else ss_spike_d9
        # MS数据列：照抄原始浓度（不除以任何值）
        for i, (_, cl, _) in enumerate(ms_cols):
            v = safe_float(raw_data.get(ss, {}).get(cl))
            c = ws.cell(row=row, column=ms_start+i)
            if v is not None:
                c.value = round6(v); c.number_format = '0.000000'
            sty(c, S['data'])
        ws.cell(row=row, column=mid1, value=None)
        ws.cell(row=row, column=rec_lbl, value=None)
        # 回收率列：SS实测浓度 ÷ SS理论加标浓度 × 100%
        for i, (_, cl, _) in enumerate(ms_cols):
            v = safe_float(raw_data.get(ss, {}).get(cl))
            c = ws.cell(row=row, column=rec_start+i)
            if v is not None:
                c.value = round_int((v / this_ss_spike) * 100)
                c.number_format = '0'
            sty(c, S['rec'])
        for c in range(mid2, last_col+1):
            sty(ws.cell(row=row,column=c), S['data'])
        row += 1

    # 说明
    row += 1
    note = rec_lbl + n_ms
    ws.cell(row=row, column=note, value='此表格计算方法：回收率，用测得浓度除以加标浓度')
    sty(ws.cell(row=row,column=note), S['yell'])
    row += 1
    ws.cell(row=row, column=note, value=f'此处实验基质加标浓度都是{spike}，所以计算回收率时每个化合物的测得浓度除以{spike}')
    sty(ws.cell(row=row,column=note), S['yell'])

    ws.row_dimensions[1].height = 19.5
    ws.row_dimensions[2].height = 17.25
    ws.column_dimensions['A'].width = 30.0

    return ws


# ============================================================
# Sheet 2: 空白基质检出限
# ============================================================
def build_sheet2(wb, raw_data, blank_cols, S, cfg):
    ws = wb.create_sheet('Blanks_MDL 空白基质检出限')
    n_b = len(blank_cols)
    cf = cfg['conversion_factor']
    unit = cfg['output_unit']

    # 换算因子存放在 A1
    ws.cell(row=1, column=1, value=cf)
    ws.cell(row=1, column=1).number_format = '0.000000'

    # Row 1 说明
    ws.cell(row=1, column=9, value='仅为检出率>50% ND取代')
    sty(ws.cell(row=1,column=9), S['yell'])

    # Row 2 表头
    ws.cell(row=2, column=1, value='化合物方法'); sty(ws.cell(row=2,column=1), S['hdr'])
    for i in range(n_b):
        ws.cell(row=2, column=2+i, value=f'blank_{i+1}'); sty(ws.cell(row=2,column=2+i), S['hdr'])

    blank_end = 2 + n_b         # 空列
    avg_c = blank_end + 1       # I: average
    mdl_c = blank_end + 2       # J: MDL
    half_c = blank_end + 3      # K: 1/2 MDL
    last_c = half_c

    ws.cell(row=2, column=avg_c, value='procedural blank average'); sty(ws.cell(row=2,column=avg_c), S['hdr'])
    ws.cell(row=2, column=mdl_c, value='MDL'); sty(ws.cell(row=2,column=mdl_c), S['hdr'])
    ws.cell(row=2, column=half_c, value='1/2 MDL'); sty(ws.cell(row=2,column=half_c), S['hdr'])

    # Row 3 单位
    ws.cell(row=3, column=1, value='分组'); sty(ws.cell(row=3,column=1), S['hdr'])
    for i in range(n_b):
        ws.cell(row=3, column=2+i, value='ppb'); sty(ws.cell(row=3,column=2+i), S['hdr'])
    ws.cell(row=3, column=blank_end, value=None)
    ws.cell(row=3, column=avg_c, value='bottle ppb'); sty(ws.cell(row=3,column=avg_c), S['hdr'])
    ws.cell(row=3, column=mdl_c, value='bottle ppb'); sty(ws.cell(row=3,column=mdl_c), S['hdr'])
    ws.cell(row=3, column=half_c, value=unit); sty(ws.cell(row=3,column=half_c), S['hdr'])

    # Row 4 样本标签
    row = 4
    ws.cell(row=row, column=1, value=cfg['sample_type']); sty(ws.cell(row=row,column=1), S['cmpd'])
    for c in range(2, last_c+1): sty(ws.cell(row=row,column=c), S['data'])
    row += 1

    avg_l = get_column_letter(avg_c)
    blank_l = get_column_letter(2)
    blank_r = get_column_letter(2 + n_b - 1)

    first_data_row = row
    for comp in ALL_COMPS:
        ws.cell(row=row, column=1, value=comp); sty(ws.cell(row=row,column=1), S['cmpd'])

        # Blank 值
        for i, (_, cl, _) in enumerate(blank_cols):
            v = safe_float(raw_data.get(comp, {}).get(cl))
            c = ws.cell(row=row, column=2+i)
            if v is not None:
                c.value = round6(v); c.number_format = '0.000000'
            sty(c, S['data'])

        ws.cell(row=row, column=blank_end, value=None)

        # I: Average (公式)
        br = f'{blank_l}{row}:{blank_r}{row}'
        c_avg = ws.cell(row=row, column=avg_c)
        c_avg.value = f'=AVERAGE({br})'
        c_avg.number_format = '0.000000'
        sty(c_avg, S['data'])

        # J: MDL = 3*STDEVA(blanks) — 公式
        c_mdl = ws.cell(row=row, column=mdl_c)
        c_mdl.value = f'=3*STDEVA({br})'
        c_mdl.number_format = '0.000000'
        sty(c_mdl, S['data'])

        # K: 1/2 MDL = ROUND(MDL/2 * $A$1, 6) — 公式，引用A1换算因子
        mdl_l = get_column_letter(mdl_c)
        c_half = ws.cell(row=row, column=half_c)
        c_half.value = f'=ROUND({mdl_l}{row}/2*$A$1,6)'
        c_half.number_format = '0.000000'
        sty(c_half, S['data'])

        row += 1

    # SS 替代物部分
    row += 1
    ws.cell(row=row, column=1, value='SS recoveries, %替代物回收率')
    sty(ws.cell(row=row,column=1), S['yellL'])
    for c in range(2, last_c+1):
        sty(ws.cell(row=row,column=c), S['yell'])
    row += 1

    for ss in SS_COMPS:
        ws.cell(row=row, column=1, value=ss); sty(ws.cell(row=row,column=1), S['cmpd'])
        for i, (_, cl, _) in enumerate(blank_cols):
            v = safe_float(raw_data.get(ss, {}).get(cl))
            c = ws.cell(row=row, column=2+i)
            if v is not None:
                c.value = round6(v); c.number_format = '0.000000'
            sty(c, S['data'])
        for c in range(blank_end, last_c+1):
            sty(ws.cell(row=row,column=c), S['data'])
        row += 1

    ws.row_dimensions[1].height = 19.5
    ws.row_dimensions[2].height = 17.25
    ws.column_dimensions['A'].width = 30.0

    info = {
        'avg_c': avg_c, 'mdl_c': mdl_c, 'half_c': half_c,
        'avg_l': avg_l, 'mdl_l': get_column_letter(mdl_c), 'half_l': get_column_letter(half_c),
        'first_row': first_data_row,
    }
    return ws, info


# ============================================================
# Sheet 3: 瓶内实测浓度
# ============================================================
def build_sheet3(wb, raw_data, sample_cols, S, cfg):
    ws = wb.create_sheet('Conc. in bottle 瓶内实测浓度')
    n_s = len(sample_cols)

    ws.cell(row=1, column=1, value='化合物方法'); sty(ws.cell(row=1,column=1), S['hdr'])
    for i, (_, _, hdr) in enumerate(sample_cols):
        ws.cell(row=1, column=2+i, value=hdr.replace('0527 Urine-',''))
        sty(ws.cell(row=1,column=2+i), S['hdr'])

    ws.cell(row=2, column=1, value='分组'); sty(ws.cell(row=2,column=1), S['hdr'])
    for i in range(n_s):
        ws.cell(row=2, column=2+i, value=cfg['masshunter_unit']); sty(ws.cell(row=2,column=2+i), S['hdr'])

    row = 3
    ws.cell(row=row, column=1, value=cfg['sample_type']); sty(ws.cell(row=row,column=1), S['cmpd'])
    for i in range(n_s): sty(ws.cell(row=row,column=2+i), S['data'])
    row += 1
    first_row = row

    for comp in TARGET_COMPS:
        ws.cell(row=row, column=1, value=comp); sty(ws.cell(row=row,column=1), S['cmpd'])
        for i, (_, cl, _) in enumerate(sample_cols):
            v = safe_float(raw_data.get(comp, {}).get(cl))
            c = ws.cell(row=row, column=2+i)
            if v is not None:
                c.value = round6(v); c.number_format = '0.000000'
            sty(c, S['data'])
        row += 1

    ws.row_dimensions[1].height = 19.5
    ws.row_dimensions[2].height = 17.25
    ws.column_dimensions['A'].width = 30.0
    return ws, first_row


# ============================================================
# Sheet 4: 最终计算浓度 (v2: 删除D/E列, 合并统计表头, A38标注)
# ============================================================
def build_sheet4(wb, raw_data, sample_cols, blank_info, s3_first, S, cfg):
    ws = wb.create_sheet('Final. conc 最终计算浓度')
    n_s = len(sample_cols)
    cf = cfg['conversion_factor']
    unit = cfg['output_unit']

    # 列布局 (删除D/E后):
    # A:化合物 B:BLANK avg C:MDL D:DF E:MEAN F:Geomean G:MEDIAN H:MIN I:MAX
    # J:5TH K:25TH L:75TH M:95TH N:空 O:说明 P~CX:样品数据(88列)
    sample_start = 16  # P列
    last_sample = sample_start + n_s - 1

    blanks_name = 'Blanks_MDL 空白基质检出限'
    bottle_name = 'Conc. in bottle 瓶内实测浓度'
    al = blank_info['avg_l']
    ml = blank_info['mdl_l']
    hl = blank_info['half_l']

    # Row 1
    ws.cell(row=1, column=1, value='化合物方法'); sty(ws.cell(row=1,column=1), S['hdr'])
    ws.cell(row=1, column=2, value='bottle'); sty(ws.cell(row=1,column=2), S['hdr'])
    ws.cell(row=1, column=3, value='bottle'); sty(ws.cell(row=1,column=3), S['hdr'])
    for c in range(4, 14): ws.cell(row=1, column=c, value=None)
    ws.cell(row=1, column=14, value=None)
    ws.cell(row=1, column=15, value=None)

    # 样品列头
    for i, (_, _, hdr) in enumerate(sample_cols):
        ws.cell(row=1, column=sample_start+i, value=hdr.replace('0527 Urine-',''))
        sty(ws.cell(row=1,column=sample_start+i), S['hdr'])

    # Row 2
    ws.cell(row=2, column=1, value='分组'); sty(ws.cell(row=2,column=1), S['hdr'])
    ws.cell(row=2, column=2, value='BLANK average'); sty(ws.cell(row=2,column=2), S['hdr'])
    ws.cell(row=2, column=3, value='MDL'); sty(ws.cell(row=2,column=3), S['hdr'])

    stat_names = ['DF 检出率','MEAN','Geometric Mean','MEDIAN','MIN','MAX',
                  '5TH','25TH','75TH','95TH']
    for i, nm in enumerate(stat_names):
        ws.cell(row=2, column=4+i, value=nm)
        sty(ws.cell(row=2,column=4+i), S['stat'])

    ws.cell(row=2, column=15, value='称样量 g, 请根据你的实际称样量填写')
    sty(ws.cell(row=2,column=15), S['yell'])

    for i in range(n_s):
        ws.cell(row=2, column=sample_start+i, value=cfg.get('sample_volume_ml', 2))
        sty(ws.cell(row=2,column=sample_start+i), S['data'])

    # 统计列 D-M (cols 4-13): Row1+Row2 合并
    stat_end = 4 + len(stat_names) - 1  # col 13
    for col in range(4, stat_end + 1):
        ws.merge_cells(start_row=1, start_column=col, end_row=2, end_column=col)
        # 合并后，合并区域的左上角单元格保留值，其他被清除
        # 把 row2 的值移到 row1
        cl = get_column_letter(col)
        merged_cell = ws.cell(row=1, column=col)
        merged_cell.value = stat_names[col - 4]
        sty(merged_cell, S['stat'])

    # Row 3 单位
    ws.cell(row=3, column=1, value=cfg['sample_type']); sty(ws.cell(row=3,column=1), S['cmpd'])
    ws.cell(row=3, column=2, value='ng/mL'); sty(ws.cell(row=3,column=2), S['hdr'])
    ws.cell(row=3, column=3, value='ng/mL'); sty(ws.cell(row=3,column=3), S['hdr'])
    ws.cell(row=3, column=4, value='%'); sty(ws.cell(row=3,column=4), S['stat'])
    for c in range(5, 14):
        ws.cell(row=3, column=c, value=unit); sty(ws.cell(row=3,column=c), S['stat'])

    # A38 标注, B38 存储换算因子
    ws.cell(row=38, column=1, value='换算因子')
    sty(ws.cell(row=38,column=1), S['bold'])
    ws.cell(row=38, column=2, value=cf)
    ws.cell(row=38,column=2).number_format = '0.000000'
    sty(ws.cell(row=38,column=2), S['data'])

    # 行号映射
    s2_first = blank_info['first_row']
    comp2row_s2 = {c: s2_first + i for i, c in enumerate(ALL_COMPS)}
    comp2row_s3 = {c: s3_first + i for i, c in enumerate(TARGET_COMPS)}

    row = 4
    for comp in TARGET_COMPS:
        ws.cell(row=row, column=1, value=comp); sty(ws.cell(row=row,column=1), S['cmpd'])
        s2r = comp2row_s2.get(comp)
        s3r = comp2row_s3.get(comp)

        # B: BLANK average
        cb = ws.cell(row=row, column=2)
        if s2r: cb.value = f"='{blanks_name}'!{al}{s2r}"; cb.number_format = '0.000000'
        sty(cb, S['data'])

        # C: MDL
        cc = ws.cell(row=row, column=3)
        if s2r: cc.value = f"='{blanks_name}'!{ml}{s2r}"; cc.number_format = '0.000000'
        sty(cc, S['data'])

        # 样品数据范围 (P~CX)
        sl = get_column_letter(sample_start)
        el = get_column_letter(last_sample)
        sr = f'{sl}{row}:{el}{row}'

        # D: DF 检出率
        cd = ws.cell(row=row, column=4)
        cd.value = f'=COUNT({sr})/COLUMNS({sr})'
        cd.number_format = '0.00%'
        sty(cd, S['stat'])

        # E~I: MEAN, GEOMEAN, MEDIAN, MIN, MAX
        funcs = {5:'AVERAGE', 6:'GEOMEAN', 7:'MEDIAN', 8:'MIN', 9:'MAX'}
        for col, func in funcs.items():
            c = ws.cell(row=row, column=col)
            c.value = f'=IF($D{row}>50%,{func}({sr}),"NC")'
            c.number_format = '0.000000'
            sty(c, S['stat'])

        # J~M: Percentiles
        for col, pct in [(10,0.05),(11,0.25),(12,0.75),(13,0.95)]:
            c = ws.cell(row=row, column=col)
            c.value = f'=IF($D{row}>50%,PERCENTILE({sr},{pct}),"NC")'
            c.number_format = '0.000000'
            sty(c, S['stat'])

        # 样品数据列 P~CX
        for i in range(n_s):
            col = sample_start + i
            s3_cl = get_column_letter(2 + i)

            if s2r and s3r:
                formula = (
                    f"=IF('{bottle_name}'!{s3_cl}{s3r}=\"\",\"\","
                    f"IF('{bottle_name}'!{s3_cl}{s3r}>'{blanks_name}'!{al}{s2r},"
                    f"('{bottle_name}'!{s3_cl}{s3r}-'{blanks_name}'!{al}{s2r})*$B$38,"
                    f"'{blanks_name}'!{hl}{s2r}*$B$38))"
                )
                ws.cell(row=row, column=col).value = formula
                ws.cell(row=row, column=col).number_format = '0.000000'
            sty(ws.cell(row=row, column=col), S['data'])
        row += 1

    ws.row_dimensions[1].height = 19.5
    ws.row_dimensions[2].height = 17.25
    ws.column_dimensions['A'].width = 30.0
    return ws


# ============================================================
# Sheet 5: 统计计算数据
# ============================================================
def build_sheet5(wb, sample_cols, S):
    ws = wb.create_sheet('统计计算结果')
    n_s = len(sample_cols)
    fn = 'Final. conc 最终计算浓度'
    ss = 16  # sample start col P

    ws.cell(row=1, column=1, value='统计数据源（供F:O统计，当DF>50%时ND替换为该化合物1/2 MDL）')

    ws.cell(row=2, column=1, value='样品名称'); sty(ws.cell(row=2,column=1), S['hdr'])
    for i in range(n_s):
        ws.cell(row=2, column=2+i, value=f"='{fn}'!{get_column_letter(ss+i)}$1")
        sty(ws.cell(row=2,column=2+i), S['hdr'])

    ws.cell(row=3, column=1, value='称样量(g)'); sty(ws.cell(row=3,column=1), S['hdr'])
    for i in range(n_s):
        ws.cell(row=3, column=2+i, value=f"='{fn}'!{get_column_letter(ss+i)}$2")
        sty(ws.cell(row=3,column=2+i), S['data'])

    row = 4
    for idx, comp in enumerate(TARGET_COMPS):
        fr = 4 + idx
        ws.cell(row=row, column=1, value=comp); sty(ws.cell(row=row,column=1), S['cmpd'])
        for i in range(n_s):
            ws.cell(row=row, column=2+i, value=f"='{fn}'!{get_column_letter(ss+i)}{fr}")
            sty(ws.cell(row=row,column=2+i), S['data'])
        row += 1

    ws.column_dimensions['A'].width = 30.0
    return ws


# ============================================================
# 说明书
# ============================================================
def build_info_sheet(wb, S, cfg):
    ws = wb.create_sheet('计算说明')
    rows = [
        ['LC-MS/MS 数据处理说明 v2'],
        ['区域','公式/规则','来源','说明'],
        ['Sheet1: MS数据','从原始数据含"MS"列提取','原始数据',f'单位:{cfg["masshunter_unit"]}'],
        ['Sheet1: 回收率%',f'MS值÷{cfg["spike_conc_ppb"]}×100,四舍五入取整','Sheet1 MS列','单位:%'],
        ['Sheet1: avg/SD/SE','AVERAGE/STDEV/SD/SQRT(COUNT) 公式','Sheet1 回收率%列','基于回收率百分比值'],
        ['Sheet2: Blank avg','=AVERAGE(空白) 公式','Sheet2 B~G列',f'单位:bottle {cfg["masshunter_unit"]}'],
        ['Sheet2: MDL','=3*STDEVA(空白) 公式','Sheet2 J列',f'单位:bottle {cfg["masshunter_unit"]}'],
        ['Sheet2: 1/2 MDL','=ROUND(MDL/2*$A$1,6) 公式','Sheet2 K列',f'单位:{cfg["output_unit"]}，$A$1=换算因子={cfg["conversion_factor"]}'],
        ['Sheet3','原始数据实际样品列直接迁移','原始数据','空白单元格保持空白'],
        ['Sheet4 B/C','引用Sheet2 I/J列','Sheet2',''],
        ['Sheet4 R~DA','IF(瓶内值>空白平均,(瓶内值-空白平均)×$B$38, 1/2MDL×$B$38)','Sheet2/Sheet3','$B$38=换算因子'],
        ['Sheet4 DF','COUNT/COLUMNS 公式','Sheet4 样品列',''],
        ['Sheet4 统计','IF(DF>50%,统计函数,"NC")','Sheet4','空值自动忽略'],
        ['换算因子位置','Sheet2 $A$1 + Sheet4 $B$38','','两处均可独立修改'],
        ['本次参数',f'样本:{cfg["sample_type"]} 取样:{cfg["sample_volume_ml"]}mL 定容:{cfg["final_volume_ml"]}mL 换算因子:{cfg["conversion_factor"]}','',''],
    ]
    for r, rd in enumerate(rows, 1):
        for c, val in enumerate(rd, 1):
            cell = ws.cell(row=r, column=c, value=val)
            sty(cell, S['hdr'] if r <= 2 else S['data'])
    for c, w in zip([1,2,3,4], [35,65,20,60]):
        ws.column_dimensions[get_column_letter(c)].width = w
    return ws


# ============================================================
# 主处理函数 (可被 Streamlit 调用)
# ============================================================
def process(config=None, return_bytes=False):
    """
    执行完整的数据处理流程。

    Args:
        config: 配置字典，None 则使用默认 CONFIG
        return_bytes: True 返回 bytes (供 Streamlit), False 保存到文件

    Returns:
        如果 return_bytes=True，返回 (bytes, filename)
        否则返回 output_filepath
    """
    cfg = config or CONFIG
    print(f"Reading: {cfg['input_file']}")
    raw_data, blanks, mss, samps, target, is_c, ss_c, all_c = read_raw(cfg['input_file'])

    # 验证
    missing = [c for c in ALL_COMPS if c not in raw_data]
    if missing:
        print(f"WARNING: {len(missing)} compounds not found: {missing}")

    S = make_styles()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    print("[1/6] Matrix spike...")
    build_sheet1(wb, raw_data, mss, S, cfg)

    print("[2/6] Blanks_MDL...")
    _, binfo = build_sheet2(wb, raw_data, blanks, S, cfg)

    print("[3/6] Conc. in bottle...")
    _, s3_first = build_sheet3(wb, raw_data, samps, S, cfg)

    print("[4/6] Final conc...")
    build_sheet4(wb, raw_data, samps, binfo, s3_first, S, cfg)

    print("[5/6] Stats helper...")
    build_sheet5(wb, samps, S)

    print("[6/6] Info sheet...")
    build_info_sheet(wb, S, cfg)

    if return_bytes:
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        wb.close()
        return output.getvalue(), cfg.get('output_file', 'processed_data.xlsx')
    else:
        out = cfg['output_file']
        print(f"Saving to: {out}")
        wb.save(out)
        wb.close()
        print("Done!")
        return out


# ============================================================
# 入口
# ============================================================
if __name__ == '__main__':
    process()
