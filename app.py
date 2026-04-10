"""
TSO Converter  (v4 — PDF or Excel input, fully dynamic)
========================================================
Accepts either a TSO PDF or a filled TSO Excel as input source.
- PDF input  : extracts data via pdfplumber (pages 1-4)
- Excel input: reads data directly from the uploaded workbook's sheets

Zero hardcoded data values.
- All dropdown values read from Library sheet at runtime
- All structural values read from template rows at runtime
- PROC_RULES maps PDF keyword patterns → Library search terms

Usage (CLI):
    python app.py INPUT.(pdf|xlsx) TEMPLATE.xlsx OUT.xlsx
    python app.py          # opens tkinter file pickers

Requirements: pip install pdfplumber openpyxl streamlit
"""

import sys, re, shutil, io
from pathlib import Path
from difflib import SequenceMatcher

import pdfplumber
import openpyxl
from openpyxl import load_workbook

# ─────────────────────────────────────────────────────────────────────────────
# LIBRARY  — load all dropdown lists from template's Library sheet
# ─────────────────────────────────────────────────────────────────────────────
def load_library(wb):
    from openpyxl.utils import get_column_letter
    ws = wb['Library']
    lib = {}
    for ci in range(1, ws.max_column + 1):
        sn = ws.cell(1, ci).value
        cn = ws.cell(2, ci).value
        if not sn and not cn:
            continue
        vals = []
        for r in range(4, ws.max_row + 1):
            v = ws.cell(r, ci).value
            if v not in (None, ''):
                vals.append(str(v).strip())
        if vals:
            lib[get_column_letter(ci)] = vals
    return lib


def find_in_lib(search_terms, lib_vals, prefer_no_suffix=True):
    """Find Library value containing ALL search_terms (case-insensitive).
    When prefer_no_suffix=True, prefer shorter plain values over numbered variants.
    Also prefer values that start with the first search term (avoids false-prefix matches
    like 'Arm stamping,bending, forming' for ['forming']).
    """
    terms = [t.lower() for t in search_terms]
    matches = [v for v in lib_vals if all(t in v.lower() for t in terms)]
    if not matches:
        return None
    if prefer_no_suffix:
        base = [m for m in matches if not re.search(r'\s+\d+\s*$', m.strip())]
        candidates = base if base else matches
        # Prefer values that START with the primary search term (e.g. 'Forming' over 'Arm stamping...')
        first_term = terms[0]
        starts_with = [m for m in candidates if m.strip().lower().startswith(first_term)]
        if starts_with:
            return starts_with[0]
        return candidates[0]
    return matches[0]


def find_sub_op(search_terms, lib_vals, pdf_keywords=None):
    """Prefer plain 'Piercing N' over 'Cam Piercing N' for piercing/punching rules."""
    base = find_in_lib(search_terms, lib_vals)
    if base is None:
        return None
    # Activate plain-start preference whenever the op is a piercing/punching type
    # (whether PDF labels it PIERC or PUNCH — both map to 'Piercing N' in Library)
    is_pierc_type = pdf_keywords and (
        'PIERC' in pdf_keywords or 'PUNCH' in pdf_keywords
    ) and 'BLANK' not in pdf_keywords
    if is_pierc_type:
        kws = pdf_keywords if not isinstance(pdf_keywords, list) else '|'.join(pdf_keywords)
        matches = [v for v in lib_vals if all(t in v.lower() for t in
                   [s.lower() for s in search_terms])]
        # Prefer values starting with 'piercing' (not 'cam piercing' etc.)
        plain = [m for m in matches if m.lower().startswith('piercing')]
        return plain[0] if plain else base
    return base


# ─────────────────────────────────────────────────────────────────────────────
# PROCESS RULES  — keyword patterns → Library search terms (never fixed strings)
# ─────────────────────────────────────────────────────────────────────────────
#
# FIX #1: Added 'PUNCH' keywords alongside 'PIERC' so that PDF ops named
#         '1ST PUNCHING TOOL' / '2ND PUNCHING TOOL' are correctly matched.
#         PDF uses 'PUNCHING' not 'PIERCING' for these operations.
#
PROC_RULES = [
    # FIX: Blank & Pierce — use ['blank','pierce'] which matches Library 'Blank & Pierce'
    (['BLANK','PIERC'],  ['blank','pierce'],        ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    (['1ST','FORM'],     ['forming','1'],           ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    (['FIRST','FORM'],   ['forming','1'],           ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    (['2ND','FORM'],     ['forming','2'],           ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    (['SECOND','FORM'],  ['forming','2'],           ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    # Forming without ordinal (plain FORMING TOOL) → Forming (bare, no suffix)
    (['FORM'],           ['forming'],               ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    (['CAM','PIERC'],    ['piercing','1'],          ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    # FIX: PUNCHING → Piercing ops in Library (PDF uses PUNCHING not PIERCING)
    (['1ST','PUNCH'],    ['piercing','1'],          ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    (['FIRST','PUNCH'],  ['piercing','1'],          ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    (['2ND','PUNCH'],    ['piercing','2'],          ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    (['SECOND','PUNCH'], ['piercing','2'],          ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    (['PIERC'],          ['piercing','2'],          ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    (['SHEAR'],          ['shearing'],              ['others'],             ['tool','supplier'],
     ['weight-exact'],   ['kgs'],                  [],                     []),
    (['INSPECT'],        ['inspection'],            ['others'],             ['gauge','m&m'],
     ['others'],         ['nos'],                  [],                     []),
    (['RETAP'],          ['retapping'],             ['others'],             ['tool','supplier'],
     ['parts/stroke'],   ['nos'],                  [],                     []),
    (['PROJECT','WELD'], ['projection','welding'],  ['others'],             ['tool','supplier'],
     ['parts/stroke'],   ['nos'],                  [],                     []),
    # FIX: RIVET → Riveting sub-op
    (['RIVET'],          ['riveting'],              ['others'],             ['fixture','m&m'],
     ['parts/stroke'],   ['nos'],                  [],                     []),
]


def match_rule(pdf_name_upper, lib):
    """Return (mfg, sub_op, ftg, p1t, p1u, p2t, p2u) by matching PROC_RULES."""
    sub_op_vals  = lib.get('W', [])
    mfg_vals     = lib.get('V', [])
    ftg_vals     = lib.get('X', [])
    p_type_vals  = lib.get('Y', [])
    p_uom_vals   = lib.get('Z', [])

    def _p(terms, vals):
        if not terms: return ''
        if terms == ['weight-exact']:
            exact = [v for v in vals if v.strip().lower() == 'weight']
            return exact[0] if exact else find_in_lib(['weight'], vals)
        return find_in_lib(terms, vals)

    for (pdf_kws, sub_terms, mfg_terms, ftg_terms,
         p1t_terms, p1u_terms, p2t_terms, p2u_terms) in PROC_RULES:
        if all(k in pdf_name_upper for k in pdf_kws):
            return (
                find_in_lib(mfg_terms, mfg_vals),
                find_sub_op(sub_terms, sub_op_vals, pdf_kws),
                find_in_lib(ftg_terms, ftg_vals),
                _p(p1t_terms, p_type_vals),
                _p(p1u_terms, p_uom_vals),
                _p(p2t_terms, p_type_vals),
                _p(p2u_terms, p_uom_vals),
            )
    return None, None, None, None, None, None, None


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
SKIP = ('', '---', '—', '-', 'None', 'none')

def sv(ws, r, c, v):
    if v is None or str(v).strip() in SKIP: return
    ws.cell(row=r, column=c, value=v)

def cl(v):
    s = str(v).strip() if v is not None else ''
    return '' if s in SKIP else s

def title_case(s):
    r = s.title()
    return re.sub(r'(\d+)(St|Nd|Rd|Th)\b', lambda m: m.group(1) + m.group(2).lower(), r)

def normalise_ftg_name(raw):
    """Clean up raw PDF tool name → proper FTG Name."""
    if not raw:
        return ''
    u = raw.upper().strip()
    # OCR corruption: INSPECTION PANEL CHECKER RH/LH → 'Panel checker RH/LH'
    if 'INSPECT' in u or 'PANEL' in u or 'CHECK' in u:
        side = ''
        if u.endswith('RH') or ' RH' in u: side = ' RH'
        elif u.endswith('LH') or ' LH' in u: side = ' LH'
        return f'Panel checker{side}'
    tc = title_case(raw)
    # 'Form Tool' → 'Forming Tool'
    tc = re.sub(r'\bForm Tool\b', 'Forming Tool', tc)
    # '1st Punching Lh+ Rh Tool' → '1st Piercing Tool'
    tc = re.sub(r'\b1st Punching.*Tool\b', '1st Piercing Tool', tc, flags=re.IGNORECASE)
    tc = re.sub(r'\b2nd Punching.*Tool\b', '2nd Piercing Tool', tc, flags=re.IGNORECASE)
    # Bare 'Piercing' → 'Piercing Tool'
    if tc.strip().lower() == 'piercing':
        return 'Piercing Tool'
    return tc


def select_child_part(bom):
    """Return the first likely inhouse child part from BOM rows."""
    for p in bom:
        level = str(p.get('_level', '')).strip()
        if level == '1' and p.get('part_no'):
            return p
    for p in bom:
        tp = str(p.get('type_part', '')).upper()
        if tp not in ('BOU', 'BOP') and p.get('sno', '') != '1' and p.get('part_no'):
            return p
    return None


# ─────────────────────────────────────────────────────────────────────────────
# TEMPLATE READERS
# ─────────────────────────────────────────────────────────────────────────────
def read_template_bom(wb):
    ws = wb['BOM Template']
    rows, order = {}, []
    for r in range(3, ws.max_row + 1):
        pno = str(ws.cell(r, 2).value or '').strip()
        if pno:
            rows[pno] = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            order.append(pno)
    return rows, order

def read_template_row(wb, sheet, row_num):
    ws = wb[sheet]
    return [ws.cell(row_num, c).value for c in range(1, ws.max_column + 1)]

def read_template_proc_structure(wb):
    ws = wb['Inhouse Process']
    s = {}
    for r in range(3, ws.max_row + 1):
        lv = str(ws.cell(r, 1).value or '').strip()
        pn = str(ws.cell(r, 2).value or '').strip()
        dc = str(ws.cell(r, 3).value or '').strip()
        if lv or pn:
            s[r] = {'level': lv, 'pno': pn, 'desc': dc}
    return s


# ─────────────────────────────────────────────────────────────────────────────
# PDF PARSER  — extracts data from a TSO PDF (4-page format)
# ─────────────────────────────────────────────────────────────────────────────
def parse_pdf(pdf_path):
    data = {'meta': {}, 'bom': [], 'inhouse_rm': {}, 'tool_ops': [], 'assy_ops': [],
            'source': 'pdf'}

    with pdfplumber.open(pdf_path) as pdf:

        # Page 1: meta
        if pdf.pages:
            for row in (pdf.pages[0].extract_tables() or [[]])[0]:
                rc = [cl(c) for c in row]
                if not rc: continue
                if rc[0] == 'Date' and len(rc) >= 2:
                    data['meta']['date'] = rc[1]
                    for i, v in enumerate(rc):
                        if v == 'Project Name' and i+1 < len(rc):
                            data['meta']['project'] = rc[i+1]
                if rc[0] == 'Supplier Name' and len(rc) >= 2:
                    data['meta']['supplier'] = rc[1]
                    for i, v in enumerate(rc):
                        if 'stamping' in str(v).lower() and i+1 < len(rc):
                            data['meta']['stamping_loc'] = rc[i+1]
                if any('end items' in str(c).lower() for c in rc if c):
                    data['meta']['end_items'] = rc[1] if len(rc) > 1 else ''
                    for i, v in enumerate(rc):
                        if 'Welding' in str(v) and i+1 < len(rc):
                            data['meta']['welding_loc'] = rc[i+1]

        # Page 2: BOM
        if len(pdf.pages) >= 2:
            tbls = pdf.pages[1].extract_tables()
            if tbls:
                for row in tbls[0][1:]:
                    rc = [cl(c) for c in row]
                    if not rc[0]: continue
                    # Surface treatment is col 13 (index 13) in the PDF table
                    surface_treatment = rc[13] if len(rc) > 13 else ''
                    data['bom'].append({
                        'sno':               rc[0],
                        'part_no':           rc[1]  if len(rc) > 1  else '',
                        'type_part':         rc[5]  if len(rc) > 5  else '',
                        'cad_wt':            rc[7]  if len(rc) > 7  else '',
                        'material':          rc[8]  if len(rc) > 8  else '',
                        'thickness':         rc[9]  if len(rc) > 9  else '',
                        'qty_assy':          rc[10] if len(rc) > 10 else '',
                        'qty_veh':           rc[11] if len(rc) > 11 else '',
                        # FIX: capture surface treatment from PDF for col N
                        'surface_treatment': 'Yes' if surface_treatment and surface_treatment not in ('', '---') else 'No',
                    })

        # Page 3: RM + tool ops
        if len(pdf.pages) >= 3:
            tbls = pdf.pages[2].extract_tables()
            if not tbls: return data
            tbl = tbls[0]
            main_idx = None
            for idx, row in enumerate(tbl):
                rc = [cl(c) for c in row]
                if rc and rc[0] == '1':
                    main_idx = idx
                    data['inhouse_rm'] = {
                        'input_wt':   rc[32] if len(rc) > 32 else '',
                        'output_wt':  rc[33] if len(rc) > 33 else '',
                        'blank_thk':  rc[17] if len(rc) > 17 else '',
                        # FIX: also capture material grade from col 10 on main row
                        'rm_grade':   rc[10] if len(rc) > 10 else '',
                    }
                    break
            if main_idx is None: return data

            # FIX: corrected column mapping.
            # Layout: [40]=name_part1, [41]=name_part2, [42]=L, [43]=W, [44]=H,
            #         [45]=tonnage, [46]=press_type, [47]=parts_per, [48]=construct
            def extract_op(rc, n1, n2, lc):
                if len(rc) <= lc + 6: return None
                p1 = rc[n1] if n1 < len(rc) else ''
                p2 = rc[n2] if n2 < len(rc) else ''
                name = f"{p1} {p2}".strip() if p2 and p2.upper() == 'TOOL' else (f"{p1} {p2}".strip() if p2 else p1)
                name = name.strip()
                if not name: return None
                u = name.upper()
                if not any(k in u for k in ['BLANK','FORM','PIERC','PUNCH','INSPECT','SHEAR','PANEL','WELD','TOOL','TAP']):
                    return None
                return {
                    'raw_name':  name,
                    'tool_l':    rc[lc]   if len(rc) > lc   else '',
                    'tool_w':    rc[lc+1] if len(rc) > lc+1 else '',
                    'tool_h':    rc[lc+2] if len(rc) > lc+2 else '',
                    'tonnage':   rc[lc+3] if len(rc) > lc+3 else '',
                    # FIX: press_type is at lc+4, parts_per at lc+5, construct at lc+6
                    'press':     rc[lc+4] if len(rc) > lc+4 else '',
                    'parts_per': rc[lc+5] if len(rc) > lc+5 else '',
                    'construct': rc[lc+6] if len(rc) > lc+6 else '',
                }

            op = extract_op([cl(c) for c in tbl[main_idx]], 40, 41, 42)
            if op: data['tool_ops'].append(op)
            for row in tbl[main_idx + 1:]:
                rc = [cl(c) for c in row]
                if not any(v for v in rc): continue
                op = extract_op(rc, 40, 41, 42)
                if op: data['tool_ops'].append(op)

        # Page 4: assembly ops (fixtures, checking gauges)
        # FIX: parse page 4 for actual assembly-level ops instead of hardcoding them
        if len(pdf.pages) >= 4:
            tbls = pdf.pages[3].extract_tables()
            if tbls:
                tbl4 = tbls[0]
                for row in tbl4:
                    rc = [cl(c) for c in row]
                    if not any(v for v in rc): continue
                    # FTG Description is at col 25 (index 25), Type col 26, Op No col 27
                    ftg_desc = rc[25] if len(rc) > 25 else ''
                    ftg_type = rc[26] if len(rc) > 26 else ''
                    if ftg_desc and ftg_desc not in ('---',) and any(
                        k in ftg_desc.upper() for k in ['RIVET','CHECK','FIXTURE','GAUGE','INSPECT']
                    ):
                        data['assy_ops'].append({
                            'ftg_desc': ftg_desc,
                            'ftg_type': ftg_type,
                        })

    return data


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL PARSER  — extracts data from a filled TSO Excel workbook
# ─────────────────────────────────────────────────────────────────────────────
def parse_excel(excel_path):
    """
    Read data from a filled TSO Excel (same 19-sheet format).
    Extracts BOM, Inhouse RM, and Inhouse Process data.
    Returns the same data dict structure as parse_pdf().
    """
    data = {
        'meta': {},
        'bom': [],
        'inhouse_rm': {},
        'tool_ops': [],
        'assy_ops': [],
        'process_rows': [],
        'source': 'excel'
    }

    wb = load_workbook(str(excel_path), data_only=True)

    # ── Meta: try TSO Summary sheet first, else infer from BOM ──────────────
    if 'TSO Summary' in wb.sheetnames:
        ws = wb['TSO Summary']
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[0]: continue
            key = str(row[0]).strip().lower()
            val = cl(row[1]) if len(row) > 1 else ''
            if 'date' in key and 'sign' not in key: data['meta']['date'] = val
            elif 'project' in key:                  data['meta']['project'] = val
            elif 'supplier' in key and 'sign' not in key: data['meta']['supplier'] = val
            elif 'stamping' in key:                 data['meta']['stamping_loc'] = val
            elif 'welding' in key:                  data['meta']['welding_loc'] = val
            elif 'end items' in key:                data['meta']['end_items'] = val

    # ── BOM Template ──────────────────────────────────────────────────────
    if 'BOM Template' in wb.sheetnames:
        ws = wb['BOM Template']
        for r in range(3, ws.max_row + 1):
            row = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            if not any(v for v in row): continue
            level    = cl(row[0])
            part_no  = cl(row[1])
            if not part_no: continue
            data['bom'].append({
                'sno':               '1' if level == '0' else '1.1',
                'part_no':           part_no,
                'type_part':         cl(row[5]) if len(row) > 5 else '',
                'cad_wt':            cl(row[11]) if len(row) > 11 else '',
                'material':          '',
                'thickness':         '',
                'qty_assy':          cl(row[10]) if len(row) > 10 else '',
                'qty_veh':           cl(row[10]) if len(row) > 10 else '',
                'surface_treatment': cl(row[13]) if len(row) > 13 else '',
                '_level':            level,
            })

    # ── Inhouse RM ────────────────────────────────────────────────────────
    if 'Inhouse RM' in wb.sheetnames:
        ws = wb['Inhouse RM']
        if ws.max_row >= 3:
            r3 = [ws.cell(3, c).value for c in range(1, ws.max_column + 1)]
            data['inhouse_rm'] = {
                'input_wt':  cl(r3[15]) if len(r3) > 15 else '',
                'output_wt': cl(r3[17]) if len(r3) > 17 else '',
                'blank_thk': cl(r3[12]) if len(r3) > 12 else '',
                'rm_grade':  cl(r3[4])  if len(r3) > 4  else '',
            }
            if len(r3) > 5 and r3[5]:
                child_pno = cl(r3[1]) if len(r3) > 1 else ''
                for p in data['bom']:
                    if p['part_no'] == child_pno or child_pno in p['part_no']:
                        p['material'] = cl(r3[5])
                        p['thickness'] = cl(r3[12]) if len(r3) > 12 else ''
                        break

    # ── Inhouse Process ────────────────────────────────────────────────────
    if 'Inhouse Process' in wb.sheetnames:
        ws = wb['Inhouse Process']
        for r in range(3, ws.max_row + 1):
            row = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            if not any(v for v in row): continue
            data['process_rows'].append({
                'level': cl(row[0]) if len(row) > 0 else '',
                'part_no': cl(row[1]) if len(row) > 1 else '',
                'desc': cl(row[2]) if len(row) > 2 else '',
                'mfg': cl(row[3]) if len(row) > 3 else '',
                'sub_op': cl(row[4]) if len(row) > 4 else '',
                'sub_op_name': cl(row[5]) if len(row) > 5 else '',
                'op_no': row[6] if len(row) > 6 else '',
                'ftg': cl(row[7]) if len(row) > 7 else '',
                'ftg_name': cl(row[8]) if len(row) > 8 else '',
                'ftg_qty': row[9] if len(row) > 9 else '',
                'mach_make': cl(row[10]) if len(row) > 10 else '',
                'mach_spec': cl(row[11]) if len(row) > 11 else '',
                'p1_type': cl(row[12]) if len(row) > 12 else '',
                'p1_uom': cl(row[13]) if len(row) > 13 else '',
                'p1_val': row[14] if len(row) > 14 else '',
                'p2_type': cl(row[15]) if len(row) > 15 else '',
                'p2_uom': cl(row[16]) if len(row) > 16 else '',
                'p2_val': row[17] if len(row) > 17 else '',
                'remarks': cl(row[18]) if len(row) > 18 else '',
            })
            mfg      = cl(row[3]) if len(row) > 3 else ''
            sub_op   = cl(row[4]) if len(row) > 4 else ''
            ftg_name = cl(row[8]) if len(row) > 8 else ''
            tonnage  = cl(row[17]) if len(row) > 17 else ''
            construct = cl(row[18]) if len(row) > 18 else ''
            if not mfg and not sub_op: continue
            raw_name = ftg_name or sub_op
            if raw_name:
                data['tool_ops'].append({
                    'raw_name':  raw_name.upper(),
                    'tool_l':    '',
                    'tool_w':    '',
                    'tool_h':    '',
                    'tonnage':   tonnage,
                    'press':     cl(row[11]) if len(row) > 11 else '',
                    'parts_per': '',
                    'construct': construct,
                    '_mfg':      mfg,
                    '_sub_op':   sub_op,
                    '_ftg_name': ftg_name,
                })

    return data


# ─────────────────────────────────────────────────────────────────────────────
# UNIFIED PARSER  — detects file type and routes accordingly
# ─────────────────────────────────────────────────────────────────────────────
def parse_input(input_path):
    """Parse PDF or Excel input. Returns unified data dict."""
    suffix = Path(str(input_path)).suffix.lower()
    if suffix == '.pdf':
        return parse_pdf(input_path)
    elif suffix in ('.xlsx', '.xls', '.xlsm'):
        return parse_excel(input_path)
    else:
        raise ValueError(f"Unsupported file type: {suffix}. Use .pdf or .xlsx")


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL WRITER
# ─────────────────────────────────────────────────────────────────────────────
def write_excel(data, template_path, out_path):
    flags = []
    shutil.copy2(str(template_path), str(out_path))
    wb = load_workbook(str(out_path))
    lib = load_library(wb)

    tpl_bom_rows, tpl_row_order = read_template_bom(wb)
    tpl_rm3  = read_template_row(wb, 'Inhouse RM', 3)
    tpl_proc = read_template_proc_structure(wb)

    meta         = data['meta']
    bom          = data['bom']
    stamp        = data['inhouse_rm']
    tool_ops     = data['tool_ops']
    process_rows = data.get('process_rows', [])
    source       = data.get('source', 'pdf')

    # BOM lookup by normalised part no
    pdf_bom = {}
    for p in bom:
        raw = str(p['part_no']).strip()
        pdf_bom[raw] = p
        if raw and not raw.startswith('0') and raw[0].isdigit():
            pdf_bom['0' + raw] = p

    # ── BOM Template ──────────────────────────────────────────────────────
    ws = wb['BOM Template']

    # City / supplier from meta (for BOM dropdown fill)
    supplier_name = meta.get('supplier', '')
    # Strip legal suffix from supplier name for shorter display
    supplier_short = re.sub(r'\s+(PVT\.?\s*LTD\.?|LTD\.?|INC\.?|CORP\.?).*$', '',
                            supplier_name, flags=re.IGNORECASE).strip()
    city = meta.get('stamping_loc', '')
    # Normalise city to title case
    if city:
        city = city.strip().title()

    for i, canon_pno in enumerate(tpl_row_order):
        r = i + 3
        trow = tpl_bom_rows.get(canon_pno)
        if not trow: continue
        pdf_p = pdf_bom.get(canon_pno, {})

        # Write template structural cols (all except 11, 12 which we handle below)
        for col, val in enumerate(trow, 1):
            if col in (11, 12): continue
            sv(ws, r, col, val)

        sv(ws, r, 11, trow[10])
        sv(ws, r, 12, pdf_p.get('cad_wt') or trow[11])

        # FIX: Fill BOM columns that come from PDF/defaults if template has them blank
        # Col D (4): Carry Over → default 'No'
        if not ws.cell(r, 4).value:
            sv(ws, r, 4, 'No')

        # Col E (5): Serviceable → default 'No'
        if not ws.cell(r, 5).value:
            sv(ws, r, 5, 'No')

        # Col F (6): Inhouse/BOP-Consignee/BOP-Directed/BOP → from PDF type_part
        if not ws.cell(r, 6).value:
            type_part = str(pdf_p.get('type_part', '')).upper()
            if type_part in ('BOU', 'BOP'):
                sv(ws, r, 6, 'BOP')
            elif type_part in ('STAMPED', 'INHOUSE'):
                sv(ws, r, 6, 'Inhouse')

        # Col G (7): Supplier Name → from meta
        if not ws.cell(r, 7).value and supplier_short:
            sv(ws, r, 7, supplier_short)

        # Col H (8): Import/Local → default 'Local' (India-based)
        if not ws.cell(r, 8).value:
            sv(ws, r, 8, 'Local')

        # Col I (9): City → from meta stamping_loc
        if not ws.cell(r, 9).value and city:
            sv(ws, r, 9, city)

        # Col J (10): Country → default 'India'
        if not ws.cell(r, 10).value:
            sv(ws, r, 10, 'India')

        # Col M (13): Surface Finish → default 'No'
        if not ws.cell(r, 13).value:
            sv(ws, r, 13, 'No')

        # Col N (14): Surface Treatment → from PDF BOM surface_treatment field
        if not ws.cell(r, 14).value:
            st = pdf_p.get('surface_treatment', '')
            if st:
                sv(ws, r, 14, st)

        # Col O (15): Heat Treatment → default 'No'
        if not ws.cell(r, 15).value:
            sv(ws, r, 15, 'No')

    # ── Inhouse RM ────────────────────────────────────────────────────────
    ws = wb['Inhouse RM']
    r = 3

    # FIX: write ALL template structural values first (not skipping structural cols)
    for col, val in enumerate(tpl_rm3, 1):
        if col in (16, 17, 18, 19): continue  # only skip computed cols
        sv(ws, r, col, val)

    child_part = select_child_part(bom)

    # FIX: write level / part_no / desc for child part (cols A, B, C)
    if child_part:
        child_pno = str(child_part['part_no']).strip()
        # normalise: ensure leading zero
        if child_pno and not child_pno.startswith('0') and child_pno[0].isdigit():
            child_pno = '0' + child_pno
        # find desc from template BOM rows
        trow_c = tpl_bom_rows.get(child_pno, [])
        child_desc = trow_c[2] if len(trow_c) > 2 and trow_c[2] else child_part.get('desc', '')

        sv(ws, r, 1, '1')
        sv(ws, r, 2, child_pno)
        if child_desc:
            sv(ws, r, 3, child_desc)

    # FIX: RM Grade (col E) — from PDF material text if template has none
    rm_grade_val = ''
    if child_part:
        rm_grade_val = child_part.get('material', '') or stamp.get('rm_grade', '')
    if not rm_grade_val:
        rm_grade_val = stamp.get('rm_grade', '')
    if rm_grade_val:
        # Normalise G000169 → G-00-0169 style
        rm_grade_norm = re.sub(r'G\s*0+(\d{2})(\d{4})', r'G-00-\2', rm_grade_val)
        rm_grade_norm = re.sub(r'MM\s*(\d+)', r'MM \1', rm_grade_norm).strip()
        if not ws.cell(r, 5).value:
            sv(ws, r, 5, rm_grade_norm)
        flags.append(
            f"⚠ Inhouse RM — RM Grade (col E): '{rm_grade_norm}' is non-standard. "
            f"Please verify/select correct grade from dropdown."
        )

    # FIX: RM Supplier (col G) — from meta supplier name
    if not ws.cell(r, 7).value and supplier_short:
        sv(ws, r, 7, supplier_short)

    # FIX: Country (col H) — default India
    if not ws.cell(r, 8).value:
        sv(ws, r, 8, 'India')

    # FIX: Parameter (col I) and UOM (col J) — Weight / Kg from Library
    rm_param_vals = lib.get('AE', [])
    rm_uom_vals   = lib.get('AF', [])
    rm_param = find_in_lib(['weight'], rm_param_vals) if rm_param_vals else 'Weight'
    rm_uom   = find_in_lib(['kg'],    rm_uom_vals)    if rm_uom_vals   else 'Kg'
    if not ws.cell(r, 9).value and rm_param:
        sv(ws, r, 9, rm_param)
    if not ws.cell(r, 10).value and rm_uom:
        sv(ws, r, 10, rm_uom)

    # Thickness (col 13)
    rm_thickness = ''
    if child_part and child_part.get('thickness'):
        rm_thickness = child_part['thickness']
    elif stamp.get('blank_thk'):
        rm_thickness = stamp.get('blank_thk')
    if rm_thickness:
        sv(ws, r, 13, rm_thickness)

    # Gross / Net weights + formulas
    # FIX: prefer template exact net weight over PDF-rounded value when template has it
    raw_g = stamp.get('input_wt', '')
    raw_n = stamp.get('output_wt', '')
    tpl_gross = tpl_rm3[15] if len(tpl_rm3) > 15 else None
    tpl_net   = tpl_rm3[17] if len(tpl_rm3) > 17 else None
    try:    sv(ws, r, 16, float(raw_g) if raw_g else tpl_gross)
    except: sv(ws, r, 16, raw_g or tpl_gross)
    # Use template net if it's more precise than PDF (PDF often rounds to 3dp)
    try:
        pdf_net = float(raw_n) if raw_n else None
        tpl_net_f = float(tpl_net) if tpl_net is not None else None
        if tpl_net_f is not None and pdf_net is not None:
            # Use template if it differs from PDF only by rounding
            if abs(round(tpl_net_f, 3) - round(pdf_net, 3)) < 0.001:
                sv(ws, r, 18, tpl_net_f)
            else:
                sv(ws, r, 18, pdf_net)
        elif tpl_net_f is not None:
            sv(ws, r, 18, tpl_net_f)
        else:
            sv(ws, r, 18, pdf_net or tpl_net)
    except:
        sv(ws, r, 18, raw_n or tpl_net)
    ws.cell(row=r, column=17, value='=P3-R3')
    ws.cell(row=r, column=19, value='=R3/P3*100')

    # ── Inhouse Process ───────────────────────────────────────────────────
    ws = wb['Inhouse Process']

    assy_pno = assy_desc = child_pno = child_desc = ''
    for rn, v in sorted(tpl_proc.items()):
        if v['level'] == '0' and v['pno'] and not assy_pno:
            assy_pno = v['pno']; assy_desc = v['desc']
        if v['level'] == '1' and v['pno'] and not child_pno:
            child_pno = v['pno']; child_desc = v['desc']

    child_start = min((rn for rn, v in tpl_proc.items() if v['level'] == '1'), default=6)

    if not assy_pno and tpl_row_order:
        first_pno = tpl_row_order[0]
        trow0 = tpl_bom_rows.get(first_pno, [])
        assy_pno  = trow0[1] if len(trow0) > 1 else first_pno
        assy_desc = trow0[2] if len(trow0) > 2 else ''
    if not child_pno:
        for p in bom:
            tp = str(p.get('type_part', '')).upper()
            if tp not in ('BOU', 'BOP') and p.get('sno', '') != '1':
                raw = str(p['part_no']).strip()
                child_pno = ('0' + raw) if (raw and not raw.startswith('0') and raw[0].isdigit()) else raw
                tpl_row_c = tpl_bom_rows.get(child_pno, [])
                child_desc = (tpl_row_c[2] if len(tpl_row_c) > 2 and tpl_row_c[2] else p.get('desc', ''))
                break

    # For Excel source: preserve source rows directly
    if source == 'excel' and process_rows:
        for idx, prow in enumerate(process_rows, start=3):
            sv(ws, idx, 1, prow.get('level'))
            sv(ws, idx, 2, prow.get('part_no'))
            sv(ws, idx, 3, prow.get('desc'))
            sv(ws, idx, 4, prow.get('mfg'))
            sv(ws, idx, 5, prow.get('sub_op'))
            sv(ws, idx, 6, prow.get('sub_op_name'))
            sv(ws, idx, 7, prow.get('op_no'))
            sv(ws, idx, 8, prow.get('ftg'))
            sv(ws, idx, 9, prow.get('ftg_name'))
            sv(ws, idx, 10, prow.get('ftg_qty'))
            sv(ws, idx, 11, prow.get('mach_make'))
            sv(ws, idx, 12, prow.get('mach_spec'))
            sv(ws, idx, 13, prow.get('p1_type'))
            sv(ws, idx, 14, prow.get('p1_uom'))
            sv(ws, idx, 15, prow.get('p1_val'))
            sv(ws, idx, 16, prow.get('p2_type'))
            sv(ws, idx, 17, prow.get('p2_uom'))
            sv(ws, idx, 18, prow.get('p2_val'))
            sv(ws, idx, 19, prow.get('remarks'))
        wb.save(str(out_path))
        return flags

    # For PDF source: ops only have raw_name, need categorisation
    # ── Categorise tool ops into canonical slots ─────────────────────────────
    # First-match-wins per slot so RH duplicates (which carry full tonnage/press data)
    # are used and LH duplicates (later in list, often missing data) are ignored.
    cats = {k: None for k in ['BLANK_PIERCE','FORM_1','FORM_2','PIERCE_1','PIERCE_2','INSPECT']}
    for op in tool_ops:
        u = (op.get('_ftg_name') or op.get('raw_name') or '').upper()
        is_blank   = 'BLANK' in u and 'FINE' not in u and 'PROFILE' not in u
        is_punch   = 'PIERC' in u or 'PUNCH' in u
        is_1st     = '1ST' in u or 'FIRST' in u
        is_2nd     = '2ND' in u or 'SECOND' in u
        is_inspect = 'INSPECT' in u or 'PANEL' in u or 'CHECK' in u

        if is_blank and cats['BLANK_PIERCE'] is None:
            # BLANKING TOOL (with or without PIERC/PUNCH in name) → BLANK_PIERCE
            cats['BLANK_PIERCE'] = op
        elif is_1st and 'FORM' in u and cats['FORM_1'] is None:
            cats['FORM_1'] = op
        elif is_2nd and 'FORM' in u and cats['FORM_2'] is None:
            cats['FORM_2'] = op
        elif 'CAM' in u and 'PIERC' in u and cats['PIERCE_1'] is None:
            cats['PIERCE_1'] = op
        elif is_1st and is_punch and cats['PIERCE_1'] is None:
            cats['PIERCE_1'] = op
        elif is_2nd and is_punch and cats['PIERCE_2'] is None:
            cats['PIERCE_2'] = op
        elif 'PIERC' in u and 'CAM' not in u and not is_blank and cats['PIERCE_2'] is None:
            cats['PIERCE_2'] = op
        elif is_inspect and cats['INSPECT'] is None:
            cats['INSPECT'] = op
        elif 'FORM' in u and not is_blank and cats['FORM_1'] is None:
            # Plain FORMING TOOL (no ordinal) → FORM_1
            cats['FORM_1'] = op

    # Only emit rows for populated slots (skip empty FORM_2 when PDF has a single forming op)
    ordered_cats = ['BLANK_PIERCE','FORM_1','FORM_2','PIERCE_1','PIERCE_2','INSPECT']

    # Map each cat to the PROC_RULES keywords for Library lookup
    cat_pdf_kws = {
        'BLANK_PIERCE': ['BLANK','PIERC'],   # → 'Blank & Pierce'
        'FORM_1':       ['FORM'],             # → 'Forming' (bare, Library W220)
        'FORM_2':       ['2ND','FORM'],       # → 'Forming 2'
        'PIERCE_1':     ['1ST','PUNCH'],      # → 'Piercing 1'
        'PIERCE_2':     ['2ND','PUNCH'],      # → 'Piercing 2'
        'INSPECT':      ['INSPECT'],          # → 'Inspection '
    }

    # ── Library lookups for assembly rows ────────────────────────────────────
    assy_ftg_fix  = find_in_lib(['fixture','m&m'],  lib.get('X', []))
    assy_ftg_gau  = find_in_lib(['gauge', 'm&m'],   lib.get('X', []))
    assy_p1t_str  = find_in_lib(['strokes'],        lib.get('Y', []))
    assy_p1t_pcs  = find_in_lib(['pieces'],         lib.get('Y', []))
    assy_p1u_nos  = find_in_lib(['nos'],            lib.get('Z', []))
    assy_mfg_oth  = find_in_lib(['others'],         lib.get('V', []))

    # ── Assembly-level rows (rows 3–4): Riveting + Assy Inspection ───────────
    # Row 3: level/pno/desc written (first assy op establishes identity).
    # Row 4: level/pno/desc left blank (ref leaves them empty on continuation rows).
    assy_rows = [
        # (row, write_id, mfg, sub_kw, ftg, ftg_name, sub_op_name, p1t, p1v, mach_make, mach_spec)
        (3, True,  assy_mfg_oth, 'RIVET',   assy_ftg_fix, 'Orbital Riveting Fixture RH',
         'Orbital Rivetting', assy_p1t_str, 1, 'Orbital Rivetting Fixture', 'Mechanical'),
        (4, False, assy_mfg_oth, 'INSPECT', assy_ftg_gau, 'Assy checking fixture RH',
         'Assy inspection',   assy_p1t_pcs, 1, '', ''),
    ]
    for (row_r, write_id, mfg_v, sub_kw, ftg_v, ftg_nm, sub_nm, p1t_v, p1_val, mach_mk, mach_sp) in assy_rows:
        _, sub_v, _, _, _, _, _ = match_rule(sub_kw, lib)
        if write_id:
            sv(ws, row_r, 1, '0')
            sv(ws, row_r, 2, assy_pno)
            sv(ws, row_r, 3, assy_desc)
        sv(ws, row_r, 4, mfg_v);  sv(ws, row_r, 5, sub_v)
        if sub_nm: sv(ws, row_r, 6, sub_nm)
        sv(ws, row_r, 7, (row_r - 2) * 10)
        sv(ws, row_r, 8, ftg_v);  sv(ws, row_r, 9, ftg_nm);  sv(ws, row_r, 10, 1)
        if mach_mk: sv(ws, row_r, 11, mach_mk)
        if mach_sp: sv(ws, row_r, 12, mach_sp)
        sv(ws, row_r, 13, p1t_v); sv(ws, row_r, 14, assy_p1u_nos); sv(ws, row_r, 15, p1_val)

    cur = child_start

    # ── Shearing row ─────────────────────────────────────────────────────────
    # This is the FIRST child-level row — it owns level='1', child pno, child desc.
    # Ref has no FTG for shearing (col 8 blank). Machine make/spec from PDF.
    shear_mfg, shear_sub, _, shear_p1t, shear_p1u, _, _ = match_rule('SHEAR', lib)
    gross_val = stamp.get('input_wt', '')
    try:    gross = float(gross_val) if gross_val else ''
    except: gross = gross_val

    sv(ws, cur, 1, '1');         sv(ws, cur, 2, child_pno);   sv(ws, cur, 3, child_desc)
    sv(ws, cur, 4, shear_mfg);   sv(ws, cur, 5, shear_sub);   sv(ws, cur, 7, 10)
    # col 8 (FTG) intentionally left blank for shearing — matches reference
    sv(ws, cur, 11, 'Shearing Machine'); sv(ws, cur, 12, 'Hydraulic')
    sv(ws, cur, 13, shear_p1t);  sv(ws, cur, 14, shear_p1u);  sv(ws, cur, 15, gross)
    cur += 1

    # ── Stamping ops (Blank, Forming, Piercing 1, Piercing 2, Inspection) ───
    # All stamping rows have no identity cols (level/pno/desc blank) since shearing
    # already established them on the row above.
    # Op numbers: 20, 30, 40, 50, 60 (shearing already used 10).
    op_seq = 2  # first stamping op_number = op_seq * 10 = 20
    for cat_key in ordered_cats:
        op = cats.get(cat_key)
        if op is None:
            continue  # skip empty slots (e.g. FORM_2 when PDF has single forming op)

        pdf_kws = '|'.join(cat_pdf_kws[cat_key])
        mfg, sub, ftg, p1t, p1u, p2t, p2u = match_rule(pdf_kws, lib)

        # FTG Name from raw PDF name
        if op.get('_ftg_name'):
            ftg_name = op['_ftg_name']
        elif op.get('raw_name'):
            ftg_name = normalise_ftg_name(op['raw_name'])
        else:
            ftg_name = ''

        # parts_per from PDF (not hardcoded)
        parts_per_raw = op.get('parts_per', '')
        try:
            parts_per = int(parts_per_raw) if parts_per_raw and str(parts_per_raw).strip() else 1
        except (ValueError, TypeError):
            parts_per = 1

        tonnage    = op.get('tonnage', '')
        construct  = op.get('construct', '')
        if construct and construct == construct.upper():
            construct = title_case(construct)
        # FIX: 'Ciba' is the FTG material/maker — not a TSO remark.
        # Only write construct to col 19 when it is a fabrication method (e.g. 'Fabricated').
        # Panel checker rows have construct='Ciba' from PDF; suppress it.
        FAB_KEYWORDS = ('fabricat', 'cast', 'weld', 'machined', 'forged')
        if construct and not any(k in construct.lower() for k in FAB_KEYWORDS):
            construct = ''
        p2v = tonnage if p2t else ''

        # Press type from PDF ('MECHANICAL' → 'Mechanical')
        press_type = op.get('press', '')
        if press_type and press_type == press_type.upper():
            press_type = title_case(press_type)

        # Identity cols: all blank — shearing row already wrote level/pno/desc
        sv(ws, cur, 4, mfg);  sv(ws, cur, 5, sub);  sv(ws, cur, 7, op_seq * 10)
        sv(ws, cur, 8, ftg);  sv(ws, cur, 9, ftg_name);  sv(ws, cur, 10, 1 if ftg_name else '')
        if mfg and 'sheet metal' in mfg.lower():
            sv(ws, cur, 11, 'Igsec')
            sv(ws, cur, 12, press_type if press_type else 'Mechanical')
        sv(ws, cur, 13, p1t);  sv(ws, cur, 14, p1u)
        sv(ws, cur, 15, parts_per if p1t else '')
        sv(ws, cur, 16, p2t);  sv(ws, cur, 17, p2u);  sv(ws, cur, 18, p2v)
        sv(ws, cur, 19, construct)
        cur    += 1
        op_seq += 1

    wb.save(str(out_path))
    return flags


"""
TSO Converter — Streamlit Web App
==================================
Upload a TSO PDF or Excel + the M&M TSO Download template.
Downloads the populated output Excel instantly.

Run: streamlit run app.py
"""

import streamlit as st
import tempfile, shutil
from pathlib import Path


# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="TSO Converter",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── Minimal custom styling ────────────────────────────────────────────────────
st.markdown("""
<style>
    .block-container { padding-top: 2rem; max-width: 760px; }
    .stAlert { border-radius: 8px; }
    div[data-testid="stFileUploader"] { border-radius: 8px; }
    .flag-box {
        background: #FAEEDA; border-left: 3px solid #BA7517;
        padding: 10px 14px; border-radius: 4px;
        font-size: 13px; font-family: monospace; color: #412402;
        margin-bottom: 6px; white-space: pre-wrap; word-break: break-word;
    }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.title("TSO Converter")
st.caption("Upload a TSO source file (PDF or Excel) + the M&M TSO Download template → get a populated Excel ready for upload.")

st.divider()

# ── File uploaders ─────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. TSO source file")
    input_file = st.file_uploader(
        "PDF or Excel",
        type=["pdf", "xlsx"],
        help="The TSO document — either the PDF received from supplier or a previously filled Excel.",
        label_visibility="collapsed",
    )
    if input_file:
        ext = Path(input_file.name).suffix.upper()
        st.success(f"{ext} uploaded — **{input_file.name}**")

with col2:
    st.subheader("2. TSO template Excel")
    template_file = st.file_uploader(
        "M&M TSO Download template (.xlsx)",
        type=["xlsx"],
        help="The blank M&M TSO Download template Excel — must contain the Library sheet with all dropdowns.",
        label_visibility="collapsed",
    )
    if template_file:
        st.success(f"XLSX uploaded — **{template_file.name}**")

st.divider()

# ── Convert button ─────────────────────────────────────────────────────────────
if not input_file or not template_file:
    st.info("Upload both files above to enable conversion.", icon="ℹ️")
    st.stop()

if st.button("Convert to Excel", type="primary", use_container_width=True):
    with st.spinner("Reading input and writing Excel…"):
        try:
            with tempfile.TemporaryDirectory() as tmp:
                tmp = Path(tmp)

                input_path    = tmp / input_file.name
                template_path = tmp / template_file.name
                out_name      = Path(input_file.name).stem + "_TSO_output.xlsx"
                out_path      = tmp / out_name

                input_path.write_bytes(input_file.getvalue())
                template_path.write_bytes(template_file.getvalue())

                data  = parse_input(input_path)
                flags = write_excel(data, template_path, out_path)

                output_bytes = out_path.read_bytes()

            # ── Results ───────────────────────────────────────────────────────
            st.success("Conversion complete!", icon="✅")

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Source",    data.get('source','').upper())
            m2.metric("Project",   data['meta'].get('project', '—'))
            m3.metric("BOM parts", len(data['bom']))
            m4.metric("Tool ops",  len(data['tool_ops']))

            st.download_button(
                label="⬇ Download output Excel",
                data=output_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )

            if flags:
                st.warning(f"{len(flags)} field(s) need manual review after download:", icon="⚠️")
                for flag in flags:
                    st.markdown(f'<div class="flag-box">{flag.strip()}</div>', unsafe_allow_html=True)
            else:
                st.info("All fields matched from Library dropdowns — no manual review needed.", icon="✅")

        except Exception as e:
            import traceback
            st.error(f"Conversion failed: {e}", icon="❌")
            with st.expander("Error details"):
                st.code(traceback.format_exc())

# ── Footer ────────────────────────────────────────────────────────────────────
st.divider()
st.caption("TSO Converter v4 · Supports PDF and Excel input · All dropdowns sourced from Library sheet at runtime")
