'''TSO Converter  (v4.2 — FIXED Inhouse RM extraction)
========================================================
Accepts either a TSO PDF or a filled TSO Excel as input source.

FIX v4.2 (Inhouse RM — 4 bugs fixed):
1. select_child_part(): now identifies child by type_part='STAMPING PART'/'STAMPING'
   for PDF source (previously returned None because all PDF BOM rows have sno='1'
   and _level is never set from PDF parse — so Level/Part No./Desc were blank).
2. parse_pdf page 3: Gross/Net weights live in the blank-data row BEFORE the '1' row,
   NOT in the main '1' row (col 14 of the '1' row is part weight, not gross RM weight).
   Now scans backward from main_idx to find the weight row.
3. parse_pdf page 3: blank_thk no longer extracted from col 14 of the '1' row
   (that column is part weight). Thickness now correctly comes from BOM page 2.
4. write_excel Inhouse RM: corrected column mapping:
     Col 16 = Gross Value  (was wrongly writing to col 15 = Density)
     Col 17 = Scrap Value  (default 0)
     Col 18 = Net Value    (formula =P3-Q3)
     Col 19 = Yield %      (formula =R3/P3*100)

FIX v4.1:
- Page 3 weight cols: 32-33 -> 34-35 (NEW format) with fallback to 32-33 (OLD)
- Page 3 tool name cols: 40-41 -> 42-43 (NEW format) with fallback to 40-41 (OLD)
- Added PROC_RULES for: DRAW, TRIMMING, FLANGE, RESTRIKE, PART OFF
'''

import sys, re, shutil, io
from pathlib import Path

import pdfplumber
import openpyxl
from openpyxl import load_workbook

# ─────────────────────────────────────────────────────────────────────────────
# LIBRARY
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
    terms = [t.lower() for t in search_terms]
    matches = [v for v in lib_vals if all(t in v.lower() for t in terms)]
    if not matches:
        return None
    if prefer_no_suffix:
        base = [m for m in matches if not re.search(r'\s+\d+\s*$', m.strip())]
        candidates = base if base else matches
        first_term = terms[0]
        starts_with = [m for m in candidates if m.strip().lower().startswith(first_term)]
        if starts_with:
            return starts_with[0]
        return candidates[0]
    return matches[0]


def find_sub_op(search_terms, lib_vals, pdf_keywords=None):
    base = find_in_lib(search_terms, lib_vals)
    if base is None:
        return None
    is_pierc_type = pdf_keywords and (
        'PIERC' in pdf_keywords or 'PUNCH' in pdf_keywords
    ) and 'BLANK' not in pdf_keywords
    if is_pierc_type:
        matches = [v for v in lib_vals if all(t in v.lower() for t in
                   [s.lower() for s in search_terms])]
        plain = [m for m in matches if m.lower().startswith('piercing')]
        return plain[0] if plain else base
    return base


# ─────────────────────────────────────────────────────────────────────────────
# PROCESS RULES
# ─────────────────────────────────────────────────────────────────────────────
PROC_RULES = [
    (['BLANK','PIERC'],  ['blank','pierce'],        ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    (['DRAW'],           ['drawing'],               ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    (['TRIM'],           ['trimming'],              ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    (['FLANGE'],         ['flanging'],              ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    (['RESTRIKE'],       ['restrike'],              ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    (['PART','OFF'],     ['cut off'],               ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    (['1ST','FORM'],     ['forming','1'],           ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    (['FIRST','FORM'],   ['forming','1'],           ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    (['2ND','FORM'],     ['forming','2'],           ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    (['SECOND','FORM'],  ['forming','2'],           ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    (['FORM'],           ['forming'],               ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
    (['CAM','PIERC'],    ['piercing','1'],          ['sheet metal','cold'], ['tool','m&m'],
     ['parts/stroke'],   ['nos'],                  ['tonnage'],            ['others']),
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
    (['RIVET'],          ['riveting'],              ['others'],             ['fixture','m&m'],
     ['parts/stroke'],   ['nos'],                  [],                     []),
]


def match_rule(pdf_name_upper, lib):
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
    if not raw:
        return ''
    u = raw.upper().strip()
    if 'INSPECT' in u or 'PANEL' in u or 'CHECK' in u:
        side = ''
        if u.endswith('RH') or ' RH' in u: side = ' RH'
        elif u.endswith('LH') or ' LH' in u: side = ' LH'
        return f'Panel checker{side}'
    tc = title_case(raw)
    tc = re.sub(r'\bForm Tool\b', 'Forming Tool', tc)
    tc = re.sub(r'\b1st Punching.*Tool\b', '1st Piercing Tool', tc, flags=re.IGNORECASE)
    tc = re.sub(r'\b2nd Punching.*Tool\b', '2nd Piercing Tool', tc, flags=re.IGNORECASE)
    if tc.strip().lower() == 'piercing':
        return 'Piercing Tool'
    return tc


def select_child_part(bom):
    """Return the inhouse child stamping part from BOM.

    FIX v4.2: PDF BOM rows all have sno='1' and no _level set, so the old
    logic always returned None. Now we identify the child by type_part:
      STAMPING PART / STAMPING / INHOUSE  -> child (inhouse stamping)
      ASSLY / ASSEMBLY                    -> skip (top-level assembly)
      HARDWEAR / BOU / BOP                -> skip (hardware/bought-out)
    """
    CHILD_TYPES = ('STAMPING PART', 'STAMPING', 'INHOUSE')
    SKIP_TYPES  = ('BOU', 'BOP', 'HARDWEAR', 'HARDWARE', 'ASSLY', 'ASSEMBLY')

    # Priority 1: explicit _level='1' (Excel source sets this)
    for p in bom:
        if str(p.get('_level', '')).strip() == '1' and p.get('part_no'):
            return p

    # Priority 2: type_part match (PDF source)
    for p in bom:
        tp = str(p.get('type_part', '')).strip().upper()
        if tp in CHILD_TYPES and p.get('part_no'):
            return p

    # Priority 3: original fallback
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
# PDF PARSER
# ─────────────────────────────────────────────────────────────────────────────
def parse_pdf(pdf_path):
    data = {'meta': {}, 'bom': [], 'inhouse_rm': {}, 'tool_ops': [], 'assy_ops': [],
            'source': 'pdf'}

    with pdfplumber.open(pdf_path) as pdf:

        # ── Page 1: meta ──────────────────────────────────────────────────
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

        # ── Page 2: BOM ───────────────────────────────────────────────────
        if len(pdf.pages) >= 2:
            tbls = pdf.pages[1].extract_tables()
            if tbls:
                for row in tbls[0][1:]:
                    rc = [cl(c) for c in row]
                    if not rc[0]: continue
                    surface_treatment = rc[13] if len(rc) > 13 else ''
                    data['bom'].append({
                        'sno':               rc[0],
                        'part_no':           rc[1]  if len(rc) > 1  else '',
                        'type_part':         rc[5]  if len(rc) > 5  else '',
                        'cad_wt':            rc[7]  if len(rc) > 7  else '',
                        'material':          rc[8]  if len(rc) > 8  else '',
                        # FIX v4.2: thickness is col 9 of PDF BOM (used for Inhouse RM thickness)
                        'thickness':         rc[9]  if len(rc) > 9  else '',
                        'qty_assy':          rc[10] if len(rc) > 10 else '',
                        'qty_veh':           rc[11] if len(rc) > 11 else '',
                        'surface_treatment': 'Yes' if surface_treatment and surface_treatment not in ('', '---') else 'No',
                    })

        # ── Page 3: RM + tool ops ─────────────────────────────────────────
        if len(pdf.pages) >= 3:
            tbls = pdf.pages[2].extract_tables()
            if not tbls:
                print("[DEBUG] No tables found on page 3")
                return data
            tbl = tbls[0]
            print(f"[DEBUG] Page 3: {len(tbl)} rows x {len(tbl[0]) if tbl else 0} cols")

            main_idx = None
            for idx, row in enumerate(tbl):
                rc = [cl(c) for c in row]
                if rc and rc[0] == '1':
                    print(f"[DEBUG] Found '1' at row {idx}")
                    main_idx = idx

                    # FIX v4.2: do NOT use col 14 of this row for blank_thk —
                    # it holds part weight (e.g. 0.068 kg), not blank thickness.
                    # Thickness comes from BOM (page 2 col 9). Leave blank_thk=''.
                    rm_grade = ''
                    if len(rc) > 12 and rc[12]:
                        rm_grade = rc[12]
                    elif len(rc) > 10 and rc[10]:
                        rm_grade = rc[10]

                    data['inhouse_rm'] = {
                        'rm_grade':    rm_grade,
                        'length':      '',
                        'width':       '',
                        'height':      '',
                        'blank_thk':   '',   # filled from BOM thickness in write_excel
                        'od':          '',
                        'density':     '',
                        'input_wt':    '',
                        'scrap_value': '',
                        'output_wt':   '',
                    }
                    break

            if main_idx is None:
                print("[DEBUG] Row with '1' not found")
                return data

            main_rc = [cl(c) for c in tbl[main_idx]]

            # FIX v4.2: Gross/Net weights live in the blank-data row BEFORE the '1' row.
            # The layout on page 3 is:
            #   Row N-1  : blank/sheet data  — cols 34=Input Weight, 35=Part Weight, 36=Yield%
            #   Row N    : part data (rc[0]='1') — col 14 is also part weight (0.068), NOT gross
            # Scan backward from main_idx to find the row with numeric weight values.
            weight_rc = None
            for look in range(main_idx - 1, max(main_idx - 5, -1), -1):
                lrc = [cl(c) for c in tbl[look]]
                # Weight row has a value at col 34 (NEW) or col 32 (OLD)
                has_new = len(lrc) > 34 and lrc[34]
                has_old = len(lrc) > 32 and lrc[32]
                if has_new or has_old:
                    weight_rc = lrc
                    print(f"[DEBUG] Weight row found at table idx {look}")
                    break

            if weight_rc is not None:
                # Try NEW format (cols 34-35) first, fallback to OLD (32-33)
                if len(weight_rc) > 35 and (weight_rc[34] or weight_rc[35]):
                    data['inhouse_rm']['input_wt']  = weight_rc[34]
                    data['inhouse_rm']['output_wt'] = weight_rc[35]
                    print(f"[DEBUG] NEW format weights: gross={weight_rc[34]}, net={weight_rc[35]}")
                    # Debug: print all columns around dimensions
                    print(f"[DEBUG] Weight row cols 20-35: {weight_rc[20:35]}")
                    # Blank dimensions typically L, W, H before weights
                    data['inhouse_rm']['length'] = weight_rc[28] if len(weight_rc) > 28 and weight_rc[28] else ''
                    data['inhouse_rm']['width']  = weight_rc[29] if len(weight_rc) > 29 and weight_rc[29] else ''
                    data['inhouse_rm']['height'] = weight_rc[30] if len(weight_rc) > 30 and weight_rc[30] else ''
                elif len(weight_rc) > 33 and (weight_rc[32] or weight_rc[33]):
                    data['inhouse_rm']['input_wt']  = weight_rc[32]
                    data['inhouse_rm']['output_wt'] = weight_rc[33]
                    print(f"[DEBUG] OLD format weights: gross={weight_rc[32]}, net={weight_rc[33]}")
                    # Debug: print all columns around dimensions
                    print(f"[DEBUG] Weight row cols 20-35: {weight_rc[20:35]}")
                    data['inhouse_rm']['length'] = weight_rc[26] if len(weight_rc) > 26 and weight_rc[26] else ''
                    data['inhouse_rm']['width']  = weight_rc[27] if len(weight_rc) > 27 and weight_rc[27] else ''
                    data['inhouse_rm']['height'] = weight_rc[28] if len(weight_rc) > 28 and weight_rc[28] else ''
            else:
                # Fallback: try main row itself (some older PDF layouts)
                print("[DEBUG] Weight row not found above main row; trying main row fallback")
                if len(main_rc) > 35 and (main_rc[34] or main_rc[35]):
                    data['inhouse_rm']['input_wt']  = main_rc[34]
                    data['inhouse_rm']['output_wt'] = main_rc[35]
                elif len(main_rc) > 33 and (main_rc[32] or main_rc[33]):
                    data['inhouse_rm']['input_wt']  = main_rc[32]
                    data['inhouse_rm']['output_wt'] = main_rc[33]

            print(f"[DEBUG] Final weights: gross='{data['inhouse_rm']['input_wt']}', "
                  f"net='{data['inhouse_rm']['output_wt']}'")
            print(f"[DEBUG] Blank dimensions: length='{data['inhouse_rm']['length']}', "
                  f"width='{data['inhouse_rm']['width']}', height='{data['inhouse_rm']['height']}'")

            # Tool op extraction
            def extract_op(rc, n1, n2, lc):
                if len(rc) <= lc + 6: return None
                p1 = rc[n1] if n1 < len(rc) else ''
                p2 = rc[n2] if n2 < len(rc) else ''
                name = (f"{p1} {p2}".strip() if p2 and p2.upper() == 'TOOL'
                        else (f"{p1} {p2}".strip() if p2 else p1))
                name = name.strip()
                if not name: return None
                u = name.upper()
                if not any(k in u for k in ['BLANK','FORM','PIERC','PUNCH','INSPECT',
                                             'SHEAR','PANEL','WELD','TOOL','TAP',
                                             'DRAW','TRIM','FLANGE','RESTRIKE','PART OFF']):
                    return None
                return {
                    'raw_name':  name,
                    'tool_l':    rc[lc]   if len(rc) > lc   else '',
                    'tool_w':    rc[lc+1] if len(rc) > lc+1 else '',
                    'tool_h':    rc[lc+2] if len(rc) > lc+2 else '',
                    'tonnage':   rc[lc+3] if len(rc) > lc+3 else '',
                    'press':     rc[lc+4] if len(rc) > lc+4 else '',
                    'parts_per': rc[lc+5] if len(rc) > lc+5 else '',
                    'construct': rc[lc+6] if len(rc) > lc+6 else '',
                }

            # NEW format: name cols 42-43, dims start at 44
            # OLD format: name cols 40-41, dims start at 42
            op = extract_op(main_rc, 42, 43, 44) or extract_op(main_rc, 40, 41, 42)
            if op: data['tool_ops'].append(op)
            for row in tbl[main_idx + 1:]:
                rc = [cl(c) for c in row]
                if not any(v for v in rc): continue
                op = extract_op(rc, 42, 43, 44) or extract_op(rc, 40, 41, 42)
                if op: data['tool_ops'].append(op)

        # ── Page 4: assembly ops ─────────────────────────────────────────
        if len(pdf.pages) >= 4:
            tbls = pdf.pages[3].extract_tables()
            if tbls:
                for row in tbls[0]:
                    rc = [cl(c) for c in row]
                    if not any(v for v in rc): continue
                    ftg_desc = rc[25] if len(rc) > 25 else ''
                    ftg_type = rc[26] if len(rc) > 26 else ''
                    if ftg_desc and ftg_desc not in ('---',) and any(
                        k in ftg_desc.upper() for k in ['RIVET','CHECK','FIXTURE','GAUGE','INSPECT']
                    ):
                        data['assy_ops'].append({'ftg_desc': ftg_desc, 'ftg_type': ftg_type})

    return data


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL PARSER
# ─────────────────────────────────────────────────────────────────────────────
def parse_excel(excel_path):
    data = {
        'meta': {}, 'bom': [], 'inhouse_rm': {}, 'tool_ops': [],
        'assy_ops': [], 'process_rows': [], 'source': 'excel'
    }
    wb = load_workbook(str(excel_path), data_only=True)

    if 'TSO Summary' in wb.sheetnames:
        ws = wb['TSO Summary']
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[0]: continue
            key = str(row[0]).strip().lower()
            val = cl(row[1]) if len(row) > 1 else ''
            if 'date' in key and 'sign' not in key:   data['meta']['date'] = val
            elif 'project' in key:                     data['meta']['project'] = val
            elif 'supplier' in key and 'sign' not in key: data['meta']['supplier'] = val
            elif 'stamping' in key:                    data['meta']['stamping_loc'] = val
            elif 'welding' in key:                     data['meta']['welding_loc'] = val
            elif 'end items' in key:                   data['meta']['end_items'] = val

    if 'BOM Template' in wb.sheetnames:
        ws = wb['BOM Template']
        for r in range(3, ws.max_row + 1):
            row = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            if not any(v for v in row): continue
            level   = cl(row[0])
            part_no = cl(row[1])
            if not part_no: continue
            data['bom'].append({
                'sno':               '1' if level == '0' else '1.1',
                'part_no':           part_no,
                'type_part':         cl(row[5])  if len(row) > 5  else '',
                'cad_wt':            cl(row[11]) if len(row) > 11 else '',
                'material':          '',
                'thickness':         '',
                'qty_assy':          cl(row[10]) if len(row) > 10 else '',
                'qty_veh':           cl(row[10]) if len(row) > 10 else '',
                'surface_treatment': cl(row[13]) if len(row) > 13 else '',
                '_level':            level,
            })

    if 'Inhouse RM' in wb.sheetnames:
        ws = wb['Inhouse RM']
        if ws.max_row >= 3:
            r3 = [ws.cell(3, c).value for c in range(1, ws.max_column + 1)]
            # Template Inhouse RM col mapping (1-based):
            # 5=RM Grade, 9=Parameter, 10=UOM, 11=Length, 12=Width, 13=Height, 14=Thickness
            # 15=OD, 16=Density, 17=Gross Value, 18=Scrap, 19=Net Value, 20=Yield%
            data['inhouse_rm'] = {
                'rm_grade':    cl(r3[4])  if len(r3) > 4  else '',
                'length':      cl(r3[10]) if len(r3) > 10 else '',
                'width':       cl(r3[11]) if len(r3) > 11 else '',
                'height':      cl(r3[12]) if len(r3) > 12 else '',
                'blank_thk':   cl(r3[13]) if len(r3) > 13 else '',
                'od':          cl(r3[14]) if len(r3) > 14 else '',
                'density':     cl(r3[15]) if len(r3) > 15 else '',
                'input_wt':    cl(r3[16]) if len(r3) > 16 else '',  # col 17 = Gross
                'scrap_value': cl(r3[17]) if len(r3) > 17 else '',  # col 18 = Scrap
                'output_wt':   cl(r3[18]) if len(r3) > 18 else '',  # col 19 = Net
            }

    if 'Inhouse Process' in wb.sheetnames:
        ws = wb['Inhouse Process']
        for r in range(3, ws.max_row + 1):
            row = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            if not any(v for v in row): continue
            data['process_rows'].append({
                'level':       cl(row[0])  if len(row) > 0  else '',
                'part_no':     cl(row[1])  if len(row) > 1  else '',
                'desc':        cl(row[2])  if len(row) > 2  else '',
                'mfg':         cl(row[3])  if len(row) > 3  else '',
                'sub_op':      cl(row[4])  if len(row) > 4  else '',
                'sub_op_name': cl(row[5])  if len(row) > 5  else '',
                'op_no':       row[6]      if len(row) > 6  else '',
                'ftg':         cl(row[7])  if len(row) > 7  else '',
                'ftg_name':    cl(row[8])  if len(row) > 8  else '',
                'ftg_qty':     row[9]      if len(row) > 9  else '',
                'mach_make':   cl(row[10]) if len(row) > 10 else '',
                'mach_spec':   cl(row[11]) if len(row) > 11 else '',
                'p1_type':     cl(row[12]) if len(row) > 12 else '',
                'p1_uom':      cl(row[13]) if len(row) > 13 else '',
                'p1_val':      row[14]     if len(row) > 14 else '',
                'p2_type':     cl(row[15]) if len(row) > 15 else '',
                'p2_uom':      cl(row[16]) if len(row) > 16 else '',
                'p2_val':      row[17]     if len(row) > 17 else '',
                'remarks':     cl(row[18]) if len(row) > 18 else '',
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
                    'tool_l': '', 'tool_w': '', 'tool_h': '',
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
# UNIFIED PARSER
# ─────────────────────────────────────────────────────────────────────────────
def parse_input(input_path):
    suffix = Path(str(input_path)).suffix.lower()
    if suffix == '.pdf':
        data = parse_pdf(input_path)
    elif suffix in ('.xlsx', '.xls', '.xlsm'):
        data = parse_excel(input_path)
    else:
        raise ValueError(f"Unsupported file type: {suffix}")

    print("\n" + "="*80)
    print("EXTRACTION SUMMARY")
    print("="*80)
    print(f"Source: {data.get('source','').upper()}")
    print(f"inhouse_rm: {data.get('inhouse_rm', {})}")
    print(f"BOM parts:  {[(p['part_no'], p['type_part'], p.get('thickness','')) for p in data['bom']]}")
    print("="*80 + "\n")
    return data


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

    pdf_bom = {}
    for p in bom:
        raw = str(p['part_no']).strip()
        pdf_bom[raw] = p
        if raw and not raw.startswith('0') and raw[0].isdigit():
            pdf_bom['0' + raw] = p

    # ── BOM Template ──────────────────────────────────────────────────────
    ws = wb['BOM Template']
    supplier_name  = meta.get('supplier', '')
    supplier_short = re.sub(r'\s+(PVT\.?\s*LTD\.?|LTD\.?|INC\.?|CORP\.?).*$',
                            '', supplier_name, flags=re.IGNORECASE).strip()
    city = meta.get('stamping_loc', '')
    if city:
        city = city.strip().title()

    for i, canon_pno in enumerate(tpl_row_order):
        r = i + 3
        trow = tpl_bom_rows.get(canon_pno)
        if not trow: continue
        pdf_p = pdf_bom.get(canon_pno, {})

        for col, val in enumerate(trow, 1):
            if col in (11, 12): continue
            sv(ws, r, col, val)
        sv(ws, r, 11, trow[10])
        sv(ws, r, 12, pdf_p.get('cad_wt') or trow[11])

        if not ws.cell(r, 4).value:  sv(ws, r, 4, 'No')
        if not ws.cell(r, 5).value:  sv(ws, r, 5, 'No')
        if not ws.cell(r, 6).value:
            tp = str(pdf_p.get('type_part', '')).upper()
            if tp in ('BOU', 'BOP'):
                sv(ws, r, 6, 'BOP')
            elif tp in ('STAMPED', 'INHOUSE', 'STAMPING PART'):
                sv(ws, r, 6, 'Inhouse')
        if not ws.cell(r, 7).value and supplier_short:  sv(ws, r, 7, supplier_short)
        if not ws.cell(r, 8).value:  sv(ws, r, 8, 'Local')
        if not ws.cell(r, 9).value and city:  sv(ws, r, 9, city)
        if not ws.cell(r, 10).value: sv(ws, r, 10, 'India')
        if not ws.cell(r, 13).value: sv(ws, r, 13, 'No')
        if not ws.cell(r, 14).value:
            st = pdf_p.get('surface_treatment', '')
            if st: sv(ws, r, 14, st)
        if not ws.cell(r, 15).value: sv(ws, r, 15, 'No')

    # ── Inhouse RM ────────────────────────────────────────────────────────
    ws = wb['Inhouse RM']
    r  = 3

    # Write template structural cols (skip the 4 computed cols 17-20: Gross, Scrap, Net, Yield)
    for col, val in enumerate(tpl_rm3, 1):
        if col in (17, 18, 19, 20): continue
        sv(ws, r, col, val)

    # FIX v4.2 BUG 1: child_part now correctly found via type_part for PDF source
    child_part = select_child_part(bom)
    print(f"[DEBUG] child_part: {child_part}")

    if child_part:
        child_pno = str(child_part['part_no']).strip()
        if child_pno and not child_pno.startswith('0') and child_pno[0].isdigit():
            child_pno = '0' + child_pno
        trow_c     = tpl_bom_rows.get(child_pno, [])
        child_desc = (trow_c[2] if len(trow_c) > 2 and trow_c[2]
                      else child_part.get('desc', ''))
        # Write Level / Part No. / Part Description (cols 1, 2, 3)
        sv(ws, r, 1, '1')
        sv(ws, r, 2, child_pno)
        if child_desc:
            sv(ws, r, 3, child_desc)

    # RM Grade (col 5)
    rm_grade_val = (child_part.get('material', '') if child_part else '') or stamp.get('rm_grade', '')
    if rm_grade_val:
        rm_grade_norm = re.sub(r'G\s*0+(\d{2})(\d{4})', r'G-00-\2', rm_grade_val)
        rm_grade_norm = re.sub(r'MM\s*(\d+)', r'MM \1', rm_grade_norm).strip()
        if not ws.cell(r, 5).value:
            sv(ws, r, 5, rm_grade_norm)
        lib_rm_grades = lib.get('AE', [])
        if lib_rm_grades and rm_grade_norm not in lib_rm_grades:
            flags.append(
                f"⚠ Inhouse RM — RM Grade (col E): '{rm_grade_norm}' is non-standard. "
                f"Please verify/select correct grade from dropdown."
            )

    # RM Supplier (col 7), Country (col 8)
    if not ws.cell(r, 7).value and supplier_short:
        sv(ws, r, 7, supplier_short)
    if not ws.cell(r, 8).value:
        sv(ws, r, 8, 'India')

    # Parameter (col 9) / UOM (col 10) / Length (col 11) / Width (col 12) / Height (col 13) / Thickness (col 14)
    rm_param = find_in_lib(['weight'], lib.get('AE', [])) or 'Weight'
    rm_uom   = find_in_lib(['kg'],    lib.get('AF', [])) or 'Kg'
    rm_length = stamp.get('length', '')
    rm_width = stamp.get('width', '')
    rm_height = stamp.get('height', '')
    rm_thickness = (child_part.get('thickness', '') if child_part else '') or stamp.get('blank_thk', '')
    
    if not ws.cell(r, 9).value and rm_param: 
        sv(ws, r, 9, rm_param)
    if not ws.cell(r, 10).value and rm_uom: 
        sv(ws, r, 10, rm_uom)
    if rm_length:
        sv(ws, r, 11, rm_length)
    if rm_width:
        sv(ws, r, 12, rm_width)
    if rm_height:
        sv(ws, r, 13, rm_height)
    if rm_thickness:
        sv(ws, r, 14, rm_thickness)

    # FIX v4.2 BUG 4: correct Inhouse RM column mapping
    # Actual template columns (after adding Length/Width/Height):
    #   Col 14 = Thickness (mm)
    #   Col 15 = OD (mm)
    #   Col 16 = Density (Kg/m3)   ← do NOT overwrite with gross weight
    #   Col 17 = Gross Value        ← write input_wt here
    #   Col 18 = Scrap Value        ← write 0 (default; user fills actual)
    #   Col 19 = Net Value          ← formula =S3-T3
    #   Col 20 = Yield %            ← formula =T3/S3*100
    raw_g = stamp.get('input_wt', '')
    try:
        gross_float = float(raw_g) if raw_g else None
    except (ValueError, TypeError):
        gross_float = None

    # Col 17 = Gross Value
    tpl_gross = tpl_rm3[16] if len(tpl_rm3) > 16 else None   # index 16 = col 17
    sv(ws, r, 17, gross_float if gross_float is not None else tpl_gross)

    # Col 18 = Scrap Value (default 0; user fills actual scrap)
    ws.cell(row=r, column=18, value=0)

    # Col 19 = Net Value formula (Gross - Scrap = S3 - T3)
    ws.cell(row=r, column=19, value='=S3-T3')

    # Col 20 = Yield % formula (Net/Gross * 100 = T3/S3*100)
    ws.cell(row=r, column=20, value='=T3/S3*100')

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
        trow0     = tpl_bom_rows.get(first_pno, [])
        assy_pno  = trow0[1] if len(trow0) > 1 else first_pno
        assy_desc = trow0[2] if len(trow0) > 2 else ''
    if not child_pno:
        for p in bom:
            tp = str(p.get('type_part', '')).upper()
            if tp not in ('BOU', 'BOP') and p.get('sno', '') != '1':
                raw       = str(p['part_no']).strip()
                child_pno = ('0' + raw) if (raw and not raw.startswith('0') and raw[0].isdigit()) else raw
                tpl_row_c = tpl_bom_rows.get(child_pno, [])
                child_desc = (tpl_row_c[2] if len(tpl_row_c) > 2 and tpl_row_c[2] else p.get('desc', ''))
                break

    # For Excel source: copy rows directly
    if source == 'excel' and process_rows:
        for idx, prow in enumerate(process_rows, start=3):
            sv(ws, idx, 1,  prow.get('level'))
            sv(ws, idx, 2,  prow.get('part_no'))
            sv(ws, idx, 3,  prow.get('desc'))
            sv(ws, idx, 4,  prow.get('mfg'))
            sv(ws, idx, 5,  prow.get('sub_op'))
            sv(ws, idx, 6,  prow.get('sub_op_name'))
            sv(ws, idx, 7,  prow.get('op_no'))
            sv(ws, idx, 8,  prow.get('ftg'))
            sv(ws, idx, 9,  prow.get('ftg_name'))
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

    # ── Categorise tool ops ──────────────────────────────────────────────
    cats = {k: None for k in ['BLANK_PIERCE','DRAW','TRIM','FLANGE','RESTRIKE','PART_OFF',
                               'FORM_1','FORM_2','PIERCE_1','PIERCE_2','INSPECT']}
    for op in tool_ops:
        u = (op.get('_ftg_name') or op.get('raw_name') or '').upper()
        if   'DRAW'     in u and cats['DRAW']        is None: cats['DRAW']        = op
        elif 'TRIM'     in u and cats['TRIM']        is None: cats['TRIM']        = op
        elif 'FLANGE'   in u and cats['FLANGE']      is None: cats['FLANGE']      = op
        elif 'RESTRIKE' in u and cats['RESTRIKE']    is None: cats['RESTRIKE']    = op
        elif 'PART' in u and 'OFF' in u and cats['PART_OFF'] is None: cats['PART_OFF'] = op
        elif 'BLANK' in u and 'FINE' not in u and 'PROFILE' not in u and cats['BLANK_PIERCE'] is None:
            cats['BLANK_PIERCE'] = op
        elif ('1ST' in u or 'FIRST'  in u) and 'FORM' in u and cats['FORM_1']   is None: cats['FORM_1']   = op
        elif ('2ND' in u or 'SECOND' in u) and 'FORM' in u and cats['FORM_2']   is None: cats['FORM_2']   = op
        elif 'CAM'  in u and 'PIERC' in u and cats['PIERCE_1'] is None: cats['PIERCE_1'] = op
        elif ('1ST' in u or 'FIRST'  in u) and ('PIERC' in u or 'PUNCH' in u) and cats['PIERCE_1'] is None:
            cats['PIERCE_1'] = op
        elif ('2ND' in u or 'SECOND' in u) and ('PIERC' in u or 'PUNCH' in u) and cats['PIERCE_2'] is None:
            cats['PIERCE_2'] = op
        elif 'PIERC' in u and 'CAM' not in u and 'BLANK' not in u and cats['PIERCE_2'] is None:
            cats['PIERCE_2'] = op
        elif ('INSPECT' in u or 'PANEL' in u or 'CHECK' in u) and cats['INSPECT'] is None:
            cats['INSPECT'] = op
        elif 'FORM' in u and 'BLANK' not in u and cats['FORM_1'] is None:
            cats['FORM_1'] = op

    ordered_cats = ['BLANK_PIERCE','DRAW','TRIM','FLANGE','RESTRIKE','PART_OFF',
                    'FORM_1','FORM_2','PIERCE_1','PIERCE_2','INSPECT']
    cat_pdf_kws = {
        'BLANK_PIERCE': ['BLANK','PIERC'],
        'DRAW':         ['DRAW'],
        'TRIM':         ['TRIM'],
        'FLANGE':       ['FLANGE'],
        'RESTRIKE':     ['RESTRIKE'],
        'PART_OFF':     ['PART','OFF'],
        'FORM_1':       ['FORM'],
        'FORM_2':       ['2ND','FORM'],
        'PIERCE_1':     ['1ST','PUNCH'],
        'PIERCE_2':     ['2ND','PUNCH'],
        'INSPECT':      ['INSPECT'],
    }

    # Assembly rows (rows 3-4)
    assy_ftg_fix = find_in_lib(['fixture','m&m'], lib.get('X', []))
    assy_ftg_gau = find_in_lib(['gauge', 'm&m'],  lib.get('X', []))
    assy_p1t_str = find_in_lib(['strokes'],       lib.get('Y', []))
    assy_p1t_pcs = find_in_lib(['pieces'],        lib.get('Y', []))
    assy_p1u_nos = find_in_lib(['nos'],           lib.get('Z', []))
    assy_mfg_oth = find_in_lib(['others'],        lib.get('V', []))

    assy_rows = [
        (3, True,  assy_mfg_oth, 'RIVET',   assy_ftg_fix, 'Orbital Riveting Fixture RH',
         'Orbital Rivetting', assy_p1t_str, 1, 'Orbital Rivetting Fixture', 'Mechanical'),
        (4, False, assy_mfg_oth, 'INSPECT', assy_ftg_gau, 'Assy checking fixture RH',
         'Assy inspection',   assy_p1t_pcs, 1, '', ''),
    ]
    for (row_r, write_id, mfg_v, sub_kw, ftg_v, ftg_nm,
         sub_nm, p1t_v, p1_val, mach_mk, mach_sp) in assy_rows:
        _, sub_v, _, _, _, _, _ = match_rule(sub_kw, lib)
        if write_id:
            sv(ws, row_r, 1, '0');  sv(ws, row_r, 2, assy_pno);  sv(ws, row_r, 3, assy_desc)
        sv(ws, row_r, 4, mfg_v);   sv(ws, row_r, 5, sub_v)
        if sub_nm: sv(ws, row_r, 6, sub_nm)
        sv(ws, row_r, 7, (row_r - 2) * 10)
        sv(ws, row_r, 8, ftg_v);   sv(ws, row_r, 9, ftg_nm);   sv(ws, row_r, 10, 1)
        if mach_mk: sv(ws, row_r, 11, mach_mk)
        if mach_sp: sv(ws, row_r, 12, mach_sp)
        sv(ws, row_r, 13, p1t_v);  sv(ws, row_r, 14, assy_p1u_nos);  sv(ws, row_r, 15, p1_val)

    cur = child_start

    # Shearing row
    shear_mfg, shear_sub, _, shear_p1t, shear_p1u, _, _ = match_rule('SHEAR', lib)
    gross_val = stamp.get('input_wt', '')
    try:    gross = float(gross_val) if gross_val else ''
    except: gross = gross_val

    sv(ws, cur, 1, '1');          sv(ws, cur, 2, child_pno);    sv(ws, cur, 3, child_desc)
    sv(ws, cur, 4, shear_mfg);    sv(ws, cur, 5, shear_sub);    sv(ws, cur, 7, 10)
    sv(ws, cur, 11, 'Shearing Machine');  sv(ws, cur, 12, 'Hydraulic')
    sv(ws, cur, 13, shear_p1t);   sv(ws, cur, 14, shear_p1u);   sv(ws, cur, 15, gross)
    cur += 1

    # Stamping ops
    op_seq = 2
    for cat_key in ordered_cats:
        op = cats.get(cat_key)
        if op is None: continue

        pdf_kws = '|'.join(cat_pdf_kws[cat_key])
        mfg, sub, ftg, p1t, p1u, p2t, p2u = match_rule(pdf_kws, lib)

        ftg_name = (op.get('_ftg_name')
                    or (normalise_ftg_name(op['raw_name']) if op.get('raw_name') else ''))

        parts_per_raw = op.get('parts_per', '')
        try:
            parts_per = int(parts_per_raw) if parts_per_raw and str(parts_per_raw).strip() else 1
        except (ValueError, TypeError):
            parts_per = 1

        tonnage = op.get('tonnage', '')
        if tonnage and tonnage.upper().endswith('T'):
            tonnage = tonnage[:-1].strip()

        construct = op.get('construct', '')
        if construct and construct == construct.upper():
            construct = title_case(construct)
        FAB_KEYWORDS = ('fabricat', 'cast', 'weld', 'machined', 'forged')
        if construct and not any(k in construct.lower() for k in FAB_KEYWORDS):
            construct = ''

        press_type = op.get('press', '')
        if press_type and press_type == press_type.upper():
            press_type = title_case(press_type)

        p2v = tonnage if p2t else ''

        sv(ws, cur, 4, mfg);   sv(ws, cur, 5, sub);    sv(ws, cur, 7, op_seq * 10)
        sv(ws, cur, 8, ftg);   sv(ws, cur, 9, ftg_name);  sv(ws, cur, 10, 1 if ftg_name else '')
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


# ─────────────────────────────────────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────────────────────────────────────
import streamlit as st
import tempfile
from pathlib import Path

st.set_page_config(page_title="TSO Converter", page_icon="📋",
                   layout="centered", initial_sidebar_state="collapsed")

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

st.title("TSO Converter")
st.caption("Upload a TSO source file (PDF or Excel) + the M&M TSO Download template → get a populated Excel ready for upload.")
st.divider()

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. TSO source file")
    input_file = st.file_uploader("PDF or Excel", type=["pdf","xlsx"],
                                   help="The TSO document — PDF or previously filled Excel.",
                                   label_visibility="collapsed")
    if input_file:
        st.success(f"{Path(input_file.name).suffix.upper()} uploaded — **{input_file.name}**")

with col2:
    st.subheader("2. TSO template Excel")
    template_file = st.file_uploader("M&M TSO Download template (.xlsx)", type=["xlsx"],
                                      help="The blank M&M TSO Download template.",
                                      label_visibility="collapsed")
    if template_file:
        st.success(f"XLSX uploaded — **{template_file.name}**")

st.divider()

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

st.divider()
st.caption("TSO Converter v4.2 · FIX: Inhouse RM — Level/Part No./Desc, Gross weight column, Net/Yield formulas corrected")
