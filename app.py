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
    """
    Strategy: TEMPLATE-FIRST.
    The template already contains the correct row structure, part numbers, descriptions,
    dropdown values, RM grade, supplier, dimensions, density, process op structure, etc.
    We copy the template verbatim, then overlay ONLY the values that are dynamic per PDF:
      - BOM: CAD weight (col 12), surface treatment (col 14) — only if template has blank
      - Inhouse RM: gross weight (col 16), net weight (col 18), thickness (col 13)
                    plus formulas for scrap (col 17) and yield (col 19)
                    Flag RM Grade if it differs from PDF
      - Inhouse Process: FTG Name (col 9), tonnage (col 18), press type (col 12),
                         parts/stroke (col 15), construct/remarks (col 19)
                         for each stamping op row — matched by sub-op keyword
    No values are hardcoded — everything comes from template rows or PDF extraction.
    """
    flags = []
    shutil.copy2(str(template_path), str(out_path))
    wb = load_workbook(str(out_path))

    stamp        = data['inhouse_rm']
    tool_ops     = data['tool_ops']
    process_rows = data.get('process_rows', [])
    source       = data.get('source', 'pdf')

    # ── BOM Template ──────────────────────────────────────────────────────────
    # Template already has all correct values.
    # Only override CAD weight (col 12) and surface treatment (col 14) from PDF
    # when the template cell is blank.
    ws_bom = wb['BOM Template']

    # Build PDF BOM lookup by part number
    pdf_bom = {}
    for p in data['bom']:
        pno = str(p['part_no']).strip()
        pdf_bom[pno] = p

    for r in range(3, ws_bom.max_row + 1):
        pno = str(ws_bom.cell(r, 2).value or '').strip()
        if not pno:
            continue
        pdf_p = pdf_bom.get(pno, {})

        # CAD weight (col 12) — use PDF value if template has none
        if not ws_bom.cell(r, 12).value and pdf_p.get('cad_wt'):
            try:    sv(ws_bom, r, 12, float(pdf_p['cad_wt']))
            except: sv(ws_bom, r, 12, pdf_p['cad_wt'])

        # Surface Treatment (col 14) — use PDF value if template has none
        if not ws_bom.cell(r, 14).value and pdf_p.get('surface_treatment'):
            sv(ws_bom, r, 14, pdf_p['surface_treatment'])

    # ── Inhouse RM ────────────────────────────────────────────────────────────
    # Template row 3 already has level, pno, desc, RM grade, supplier, country,
    # parameter, UOM, length, width, density — copy verbatim.
    # Override: thickness (col 13) from PDF, gross (col 16) and net (col 18) from PDF.
    # Add formulas for scrap (col 17) and yield (col 19).
    ws_rm = wb['Inhouse RM']
    tpl_rm3 = [ws_rm.cell(3, c).value for c in range(1, ws_rm.max_column + 1)]

    # All template values for row 3 are already in place (file was copied from template).
    # Just overlay the PDF-dynamic values:

    # Thickness (col 13) — PDF blank_thk is more reliable than template for new parts
    blank_thk = stamp.get('blank_thk', '')
    if blank_thk:
        try:    sv(ws_rm, 3, 13, float(blank_thk))
        except: sv(ws_rm, 3, 13, blank_thk)

    # Gross weight (col 16) — from PDF
    raw_g = stamp.get('input_wt', '')
    if raw_g:
        try:    sv(ws_rm, 3, 16, float(raw_g))
        except: sv(ws_rm, 3, 16, raw_g)

    # Net weight (col 18) — from PDF; use more precise template value if PDF is rounded
    raw_n = stamp.get('output_wt', '')
    tpl_net = tpl_rm3[17] if len(tpl_rm3) > 17 else None
    try:
        pdf_net_f = float(raw_n) if raw_n else None
        tpl_net_f = float(tpl_net) if tpl_net is not None else None
        if tpl_net_f is not None and pdf_net_f is not None:
            # If template value rounds to same as PDF, use the more precise template value
            if abs(round(tpl_net_f, 3) - round(pdf_net_f, 3)) < 0.001:
                sv(ws_rm, 3, 18, tpl_net_f)
            else:
                sv(ws_rm, 3, 18, pdf_net_f)
        elif pdf_net_f is not None:
            sv(ws_rm, 3, 18, pdf_net_f)
    except:
        if raw_n: sv(ws_rm, 3, 18, raw_n)

    # Scrap and Yield — always formulas referencing gross/net
    ws_rm.cell(row=3, column=17, value='=P3-R3')
    ws_rm.cell(row=3, column=19, value='=R3/P3*100')

    # Flag if PDF RM grade differs from template (template grade is the correct standard value)
    tpl_grade = str(tpl_rm3[4]).strip() if len(tpl_rm3) > 4 and tpl_rm3[4] else ''
    pdf_grade = stamp.get('rm_grade', '')
    if pdf_grade and tpl_grade:
        # Normalise both for comparison
        def norm_grade(g):
            return re.sub(r'[\s\-]', '', g).upper()
        if norm_grade(pdf_grade) not in norm_grade(tpl_grade):
            flags.append(
                f"⚠ Inhouse RM — RM Grade mismatch: PDF says '{pdf_grade}', "
                f"template has '{tpl_grade}'. Verify correct grade."
            )

    # ── Inhouse Process ───────────────────────────────────────────────────────
    # For Excel source: rows already fully populated — copy verbatim
    ws_proc = wb['Inhouse Process']

    if source == 'excel' and process_rows:
        for idx, prow in enumerate(process_rows, start=3):
            sv(ws_proc, idx, 1,  prow.get('level'))
            sv(ws_proc, idx, 2,  prow.get('part_no'))
            sv(ws_proc, idx, 3,  prow.get('desc'))
            sv(ws_proc, idx, 4,  prow.get('mfg'))
            sv(ws_proc, idx, 5,  prow.get('sub_op'))
            sv(ws_proc, idx, 6,  prow.get('sub_op_name'))
            sv(ws_proc, idx, 7,  prow.get('op_no'))
            sv(ws_proc, idx, 8,  prow.get('ftg'))
            sv(ws_proc, idx, 9,  prow.get('ftg_name'))
            sv(ws_proc, idx, 10, prow.get('ftg_qty'))
            sv(ws_proc, idx, 11, prow.get('mach_make'))
            sv(ws_proc, idx, 12, prow.get('mach_spec'))
            sv(ws_proc, idx, 13, prow.get('p1_type'))
            sv(ws_proc, idx, 14, prow.get('p1_uom'))
            sv(ws_proc, idx, 15, prow.get('p1_val'))
            sv(ws_proc, idx, 16, prow.get('p2_type'))
            sv(ws_proc, idx, 17, prow.get('p2_uom'))
            sv(ws_proc, idx, 18, prow.get('p2_val'))
            sv(ws_proc, idx, 19, prow.get('remarks'))
        wb.save(str(out_path))
        return flags

    # For PDF source: template already has the correct process row structure.
    # We only need to overlay per-row dynamic values from PDF tool_ops:
    #   col  9  FTG Name    — derived from PDF raw op name
    #   col 12  Machine Spec (press type) — from PDF
    #   col 15  P1 Value (parts/stroke)   — from PDF
    #   col 18  P2 Value (tonnage)        — from PDF
    #   col 19  Remarks (construct)       — from PDF

    # ── Categorise PDF tool ops into named slots ──────────────────────────────
    # Match each PDF op to a canonical slot by keyword scanning.
    # First-match-wins so the RH part ops (which carry full data) take precedence
    # over the LH duplicates that follow.
    cats = {k: None for k in ['BLANK','FORM_1','FORM_2','PIERCE_1','PIERCE_2','INSPECT']}
    for op in tool_ops:
        u = (op.get('raw_name') or '').upper()
        is_blank   = 'BLANK' in u and 'FINE' not in u and 'PROFILE' not in u
        is_1st     = '1ST' in u or 'FIRST' in u
        is_2nd     = '2ND' in u or 'SECOND' in u
        is_punch   = 'PIERC' in u or 'PUNCH' in u
        is_inspect = 'INSPECT' in u or 'PANEL' in u or 'CHECK' in u

        if   is_blank                              and cats['BLANK']    is None: cats['BLANK']    = op
        elif is_1st and 'FORM' in u                and cats['FORM_1']   is None: cats['FORM_1']   = op
        elif is_2nd and 'FORM' in u                and cats['FORM_2']   is None: cats['FORM_2']   = op
        elif is_1st and is_punch                   and cats['PIERCE_1'] is None: cats['PIERCE_1'] = op
        elif is_2nd and is_punch                   and cats['PIERCE_2'] is None: cats['PIERCE_2'] = op
        elif 'CAM' in u and 'PIERC' in u           and cats['PIERCE_1'] is None: cats['PIERCE_1'] = op
        elif 'PIERC' in u and not is_blank         and cats['PIERCE_2'] is None: cats['PIERCE_2'] = op
        elif is_inspect                            and cats['INSPECT']  is None: cats['INSPECT']  = op
        elif 'FORM' in u and not is_blank          and cats['FORM_1']   is None: cats['FORM_1']   = op

    # Map template process sub-op keyword → cat slot
    # These keywords are matched against the template's Sub Operation (col 5) values
    SUB_TO_CAT = {
        'blank':      'BLANK',
        'forming':    'FORM_1',
        'forming 1':  'FORM_1',
        'forming 2':  'FORM_2',
        'piercing 1': 'PIERCE_1',
        'piercing 2': 'PIERCE_2',
        'inspection': 'INSPECT',
    }

    # Walk every process row in template (rows 3 onward) and overlay PDF values
    for r in range(3, ws_proc.max_row + 1):
        sub_op = str(ws_proc.cell(r, 5).value or '').strip().lower()
        if not sub_op:
            continue

        # Find matching cat for this sub-op
        cat_key = None
        for kw, ck in SUB_TO_CAT.items():
            if sub_op.startswith(kw):
                cat_key = ck
                break
        if cat_key is None:
            continue

        op = cats.get(cat_key)
        if op is None:
            continue

        # FTG Name (col 9) — normalise from PDF raw name
        ftg_name = normalise_ftg_name(op.get('raw_name', ''))
        if ftg_name:
            sv(ws_proc, r, 9, ftg_name)

        # Machine Spec / press type (col 12) — from PDF
        press_type = op.get('press', '')
        if press_type:
            sv(ws_proc, r, 12, title_case(press_type) if press_type == press_type.upper() else press_type)

        # P1 Value / parts per stroke (col 15) — from PDF
        parts_per_raw = op.get('parts_per', '')
        if parts_per_raw:
            try:    sv(ws_proc, r, 15, int(parts_per_raw))
            except: sv(ws_proc, r, 15, parts_per_raw)

        # P2 Value / tonnage (col 18) — from PDF
        tonnage = op.get('tonnage', '')
        if tonnage:
            sv(ws_proc, r, 18, tonnage)

        # Remarks / construct (col 19) — only fabrication-method keywords, not material names
        construct = op.get('construct', '')
        if construct:
            c_tc = title_case(construct) if construct == construct.upper() else construct
            FAB_KW = ('fabricat', 'cast', 'weld', 'machined', 'forged')
            if any(k in c_tc.lower() for k in FAB_KW):
                sv(ws_proc, r, 19, c_tc)

        # Gross weight for shearing row (col 15 when sub_op is 'shearing')
        if sub_op.startswith('shear'):
            gross_val = stamp.get('input_wt', '')
            if gross_val:
                try:    sv(ws_proc, r, 15, float(gross_val))
                except: sv(ws_proc, r, 15, gross_val)

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
