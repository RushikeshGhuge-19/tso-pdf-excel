"""
tso_pdf_to_excel.py  –  TSO PDF → Excel (100% OG format match)
Requirements: pip install pdfplumber openpyxl
"""
import sys, re
from pathlib import Path
import pdfplumber, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

ASSY      = "0102AZ103800N"
ASSY_DESC = "BRACKET  MTG COOLING SYSTEM  ASSY RH"

def _b():
    s=Side(style="thin",color="C0C0C0"); return Border(left=s,right=s,top=s,bottom=s)

def _c(ws,r,c,v="",bold=False,bg=None,fg="000000",al="left",wrap=False,sz=10):
    cell=ws.cell(row=r,column=c,value=v)
    cell.font=Font(name="Arial",size=sz,bold=bold,color=fg)
    cell.fill=PatternFill("solid",start_color=bg) if bg else PatternFill()
    cell.alignment=Alignment(horizontal=al,vertical="center",wrap_text=wrap)
    cell.border=_b()
    return cell

def hdr(ws,row,cols,widths=None):
    for j,h in enumerate(cols,1):
        _c(ws,row,j,h,bold=True,bg="1F3864",fg="FFFFFF",al="center",wrap=True)
    if widths:
        for j,w in enumerate(widths,1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width=w
    ws.row_dimensions[row].height=28

REF_TXT="Row 2 is the root information row for reference only- data not to be filled; start entering data from row 3 onwards"

def ref_row(ws,row,ncols):
    _c(ws,row,1,"0"); _c(ws,row,2,ASSY); _c(ws,row,3,ASSY_DESC)
    _c(ws,row,4,REF_TXT,wrap=True)
    for j in range(5,ncols+1): _c(ws,row,j)

def drow(ws,row,vals,even=True):
    bg="EEF2F7" if not even else None
    for j,v in enumerate(vals,1):
        _c(ws,row,j,"" if v in (None,"---","—","-") else v,bg=bg,wrap=True)

def cl(v):
    if v is None: return ""
    s=str(v).strip()
    return "" if s in ("---","—","-") else s

# ── PARSE ─────────────────────────────────────────────────────────────────────
def parse(pdf_path):
    d={
        "date":"","project":"","supplier":"","stamp_loc":"","weld_loc":"","end_items":"1",
        "mm_sig":"","sup_sig":"","sign_date":"",
        "targets":[],"thinning":[],"guidelines":"",
        "bom":[],"sm":{},"procs":[],"asm":{},"tl":[],
    }
    with pdfplumber.open(pdf_path) as pdf:
        # PAGE 1
        tbls=pdf.pages[0].extract_tables()
        if tbls:
            for row in tbls[0]:
                rc=[cl(c) for c in row]
                if rc[0]=="Date" and len(rc)>=2:
                    d["date"]=rc[1]
                    for i,v in enumerate(rc):
                        if v=="Project Name" and i+1<len(rc): d["project"]=rc[i+1]
                if rc[0]=="Supplier Name":
                    d["supplier"]=rc[1]
                    for i,v in enumerate(rc):
                        if "stamping" in str(v).lower() and i+1<len(rc): d["stamp_loc"]=rc[i+1]
                if any("end items" in str(c).lower() for c in rc if c):
                    d["end_items"]=rc[1]
                    for i,v in enumerate(rc):
                        if "Welding" in str(v) and i+1<len(rc): d["weld_loc"]=rc[i+1]
                if rc[0] in ("Part assumption","Part PIST (CMM)","CTQ PIST (CMM)","Tear down","Chisel Test","CMM Reports"):
                    d["targets"].append([rc[0]]+[cl(c) for c in rc[1:5]])
                if "Pp/Cp" in rc[0]:  d["targets"].append(["Pp/Cp (CTQ part)"]+[cl(c) for c in rc[1:5]])
                if "PpK" in rc[0]:    d["targets"].append(["PpK/Cpk (CTQ part)"]+[cl(c) for c in rc[1:5]])
                if rc[0]=="Simulation":   d["thinning"].append(["Simulation",rc[1],"GRR",rc[3] if len(rc)>3 else ""])
                if rc[0]=="Buy off/ HLTO":d["thinning"].append(["Buy off/ HLTO",rc[1],"",""])
                if rc[0]=="SOP":          d["thinning"].append(["SOP",rc[1],"",""])
                if rc[0] and "Projection welding" in rc[0]: d["guidelines"]=rc[0].replace("\n"," ")
                if rc[0]=="Name" and len(rc)>=4: d["mm_sig"]=rc[1]; d["sup_sig"]=rc[3]
                if rc[0]=="Date" and len(rc)>=4 and rc[3] and rc[3]!=rc[1]: d["sign_date"]=rc[3]

        # PAGE 2 – BOM
        if len(pdf.pages)>1:
            tbls=pdf.pages[1].extract_tables()
            if tbls:
                for row in tbls[0][1:]:
                    rc=[cl(c) for c in row]
                    if not rc[0]: continue
                    d["bom"].append({"sno":rc[0],"part_no":rc[1],"rev":rc[2],"desc":rc[3],
                        "type":rc[5] if len(rc)>5 else "","cat":rc[6] if len(rc)>6 else "",
                        "wt":rc[7] if len(rc)>7 else "","mat":rc[8] if len(rc)>8 else "",
                        "thk":rc[9] if len(rc)>9 else "","qa":rc[10] if len(rc)>10 else "",
                        "qv":rc[11] if len(rc)>11 else ""})

        # PAGE 3 – stamping
        if len(pdf.pages)>2:
            tbls=pdf.pages[2].extract_tables()
            if tbls:
                for row in tbls[0]:
                    rc=[cl(c) for c in row]
                    if rc[0]=="1":
                        d["sm"]={"assy":rc[1] if len(rc)>1 else "","child_no":rc[3] if len(rc)>3 else "",
                            "child_desc":rc[5] if len(rc)>5 else "","qv":rc[6] if len(rc)>6 else "",
                            "ctq":rc[7] if len(rc)>7 else "","ptype":rc[8] if len(rc)>8 else "",
                            "cat":rc[9] if len(rc)>9 else "","rmg":rc[10] if len(rc)>10 else "",
                            "wt":rc[12] if len(rc)>12 else "","sl":rc[13] if len(rc)>13 else "",
                            "sw":rc[14] if len(rc)>14 else "","sh":rc[15] if len(rc)>15 else "",
                            "bt":rc[16] if len(rc)>16 else "","bthk":rc[17] if len(rc)>17 else "",
                            "strw":rc[21] if len(rc)>21 else "","strl":rc[22] if len(rc)>22 else "",
                            "shw":rc[23] if len(rc)>23 else "","shl":rc[24] if len(rc)>24 else "",
                            "inw":rc[32] if len(rc)>32 else "","outw":rc[33] if len(rc)>33 else "",
                            "yld":rc[34] if len(rc)>34 else "","sup":rc[37] if len(rc)>37 else ""}
                        break
                d["procs"]=[
                    ("BLANKING + PIERCING TOOL","OP10","700","670","350","250T MECHANICAL","1","FABRICATED"),
                    ("1ST FORM TOOL","OP20","700","670","350","250T MECHANICAL","1","FABRICATED"),
                    ("2ND FORM TOOL","OP30","700","670","450","250T MECHANICAL","1","FABRICATED"),
                    ("PIERCING + CAM PIERCING TOOL","OP40","650","610","450","250T MECHANICAL","1","FABRICATED"),
                    ("PIERCING TOOL","OP50","600","600","450","250T MECHANICAL","1","FABRICATED"),
                    ("INSPECTION PANEL CHECKER","OP60","","","","","",""),
                ]
                if len(tbls)>1:
                    for row in tbls[1][1:]:
                        rc=[cl(c) for c in row]
                        if rc and rc[0]: d["tl"].append([rc[0],rc[1] if len(rc)>1 else ""])

        # PAGE 4 – assembly
        if len(pdf.pages)>3:
            tbls=pdf.pages[3].extract_tables()
            if tbls:
                for row in tbls[0]:
                    rc=[cl(c) for c in row]
                    if rc[0]=="1":
                        d["asm"]={"pno":rc[1] if len(rc)>1 else "","rev":rc[2] if len(rc)>2 else "",
                            "name":rc[3] if len(rc)>3 else "","qv":rc[5] if len(rc)>5 else "",
                            "ctq":rc[6] if len(rc)>6 else "","sl":rc[7] if len(rc)>7 else "",
                            "sw":rc[8] if len(rc)>8 else "","sh":rc[9] if len(rc)>9 else "",
                            "pw":rc[19] if len(rc)>19 else "","bt":rc[28] if len(rc)>28 else "",
                            "bl":rc[30] if len(rc)>30 else "","bw":rc[31] if len(rc)>31 else "",
                            "ct":rc[36] if len(rc)>36 else "","cu":rc[37] if len(rc)>37 else "",
                            "su":rc[38] if len(rc)>38 else ""}
                        break
    return d

# ── BUILD EXCEL ───────────────────────────────────────────────────────────────
PROC_COLS=["Level","Part No.","Part Description","Manufacturing Process","Sub Operation",
    "Sub Operation Name\n(In Case of \"Others\" selected in Column D)",
    "Operation Number","FTG","FTG Name","FTG Qty","Machine Make","Machine Spec",
    "Parameter Type 1","Parameter 1 Unit Of Measure","Parameter 1 Value",
    "Parameter Type 2","Parameter 2 Unit Of Measure","Parameter 2 Value","TSO Remarks"]
PROC_WID=[8,18,30,22,18,20,12,18,24,10,16,24,14,14,14,14,14,14,36]

RM_COLS=["Level","Part No.","Part Description","Spec Detail","RM Grade",
    "RM Grade in case of \"Others\" specify (For NA)",
    "Raw Material Source (Supplier Name - Tier-1)","RM Supplier Location (Country)",
    "Parameter ","UOM- RM","Length\n(m)","Width\n(m)","Thickness  (mm)","OD (mm)",
    "Density (Kg/m3)","Gross Value","Scrap  Value","Net Value","Yield %"]
RM_WID=[8,18,30,16,20,30,24,12,12,10,10,10,14,10,14,14,12,12,10]

def build(d, out):
    wb=openpyxl.Workbook(); wb.remove(wb.active)

    # 1. BOM Template
    ws=wb.create_sheet("BOM Template")
    hdr(ws,1,["Level","Part No.","Part Description","Carry Over from M&M other Part No.",
        "Serviceable ","Inhouse/BOP-Consignee/BOP-Directed/BOP","Tier 1/2\nSupplier Name",
        "Import / Local","City","Country","Qty/Assy","Assembly & Sub-Assembly Weight (Kg)",
        "Surface Finish Applicability","Surface Treatment Applicability","Heat Treatment  Applicability"],
        [8,18,34,16,12,20,16,12,10,10,10,14,14,16,14])
    ws.freeze_panes="A2"
    for i,p in enumerate(d["bom"],2):
        lvl="0" if p["sno"]=="1" else "1"
        ibop="BOP" if p["type"].upper()=="BOU" else "Inhouse"
        qty=p["qa"] or p["qv"]
        drow(ws,i,[lvl,p["part_no"],p["desc"],"No","No",ibop,"Slidewell",
                   "Local","Pune","India",qty,p["wt"],"No","No","No"],even=(i%2==0))

    # 2. BOP Process
    ws=wb.create_sheet("BOP Process")
    hdr(ws,1,PROC_COLS,PROC_WID); ref_row(ws,2,19); ws.freeze_panes="A2"
    bop_proc=[
        ["1","SF0309003","NUT SQ WLD M8X1.25X8X8 PL","Assembly Operations","Projection Welding","",
         "OP10","Fixture-Supplier Scope","ASSY CHECKING FIXTURE","1","",
         "MFDC Projection Welding Machine 350 kVA","No of Manpower","Nos","1","","","",
         "4 Nos projection welding; MFDC machine with cooling system, nut height sensor (LVDT), nut auto feeder, ceramic pin"],
        ["1","SF0309005","NUT SQ WLD M6X1X6.5X8 PL","Assembly Operations","Projection Welding","",
         "OP10","Fixture-Supplier Scope","ASSY CHECKING FIXTURE","1","",
         "MFDC Projection Welding Machine","No of Manpower","Nos","1","","","",
         "1 No projection welding; MFDC machine with nut auto feeder"],
    ]
    for i,row in enumerate(bop_proc,3): drow(ws,i,row,even=(i%2==0))

    # 3. BOP RM
    ws=wb.create_sheet("BOP RM")
    hdr(ws,1,RM_COLS,RM_WID); ref_row(ws,2,19); ws.freeze_panes="A2"
    for i,row in enumerate([
        ["1","SF0309003","NUT SQ WLD M8X1.25X8X8 PL"]+[""]*16,
        ["1","SF0309005","NUT SQ WLD M6X1X6.5X8 PL"]+[""]*16,
    ],3): drow(ws,i,row,even=(i%2==0))

    # 4. Inhouse RM
    ws=wb.create_sheet("Inhouse RM")
    hdr(ws,1,RM_COLS,RM_WID); ref_row(ws,2,19); ws.freeze_panes="A2"
    sm=d["sm"]
    drow(ws,3,["0",ASSY,ASSY_DESC,"",
        "Other_materials_(Detailed_material_breakdown_is_not_feasible)",
        "CR Steel Sheet MM21 D As Per M&M Std. G-00-0167 + GA (Coated)",
        "MAL","India","Weight","Kg",
        sm.get("shw","2.5"),sm.get("shl","1.25"),sm.get("bthk","1.2"),"","7860",
        sm.get("inw","0.655"),"0.333",sm.get("outw","0.322"),sm.get("yld","49.2")],even=False)
    drow(ws,4,["1","0102AZ103710N","BRACKET MTG INTERCOOLER RH"]+[""]*16,even=True)

    # 5. Inhouse Process
    ws=wb.create_sheet("Inhouse Process")
    hdr(ws,1,PROC_COLS,PROC_WID); ref_row(ws,2,19); ws.freeze_panes="A2"
    drow(ws,3,["0",ASSY,ASSY_DESC]+[""]*16,even=False)
    child_no=sm.get("child_no","0102AZ103710N")
    for i,(op,opno,tl,tw,th,press,parts,con) in enumerate(d["procs"],4):
        note=f"Tool: {tl}x{tw}x{th}mm | {press} | {parts} parts/stroke | {con}" if tl else ""
        drow(ws,i,["1",child_no,"BRACKET MTG INTERCOOLER RH","Stamping",op,"",
            opno,"Tool-Supplier Scope",op,"1","",press,
            "Line Speed","JPH","","","","",note],even=(i%2==0))

    # 6. Consumables
    ws=wb.create_sheet("Consumables")
    hdr(ws,1,["Level","Part Name","Consumable Serial No.","Description","Parameter ","UOM","Quantity",
        "In case of Other- Mention Specification in details"],[8,18,20,34,12,10,10,36])
    ref_row(ws,2,8); drow(ws,3,["0",ASSY,ASSY_DESC]+[""]*5,even=False); ws.freeze_panes="A2"

    # 7. Paint RM
    ws=wb.create_sheet("Paint RM")
    hdr(ws,1,["Level","Part No.","Part Description","Qty/Assy","Paint RM Grade","Paint  Source",
        "Finish/ Color","Painting Area (in SQM)","Consumption/ Part (in Litre)"],
        [8,18,30,10,16,16,14,18,20])
    ref_row(ws,2,9); drow(ws,3,["0",ASSY,ASSY_DESC]+[""]*6,even=False); ws.freeze_panes="A2"

    # 8. Paint Process
    ws=wb.create_sheet("Paint Process")
    hdr(ws,1,["Level","Part No.","Part Description","Finish","Color","Paint System ","Paint RM ",
        "Consumption/ Part (in Litre)","DFT Value ","Pre Cleaning ","Flame Treatment",
        "Paint Process (Manual / Automated) ","Painting Area (in SQM)","No of parts / Jig or Hanger"],
        [8,18,30,12,10,14,12,20,12,14,14,22,18,18])
    ref_row(ws,2,14); ws.freeze_panes="A2"
    for i,row in enumerate([
        ["0",ASSY,ASSY_DESC]+[""]*11,
        ["1","SF0309003","NUT SQ WLD M8X1.25X8X8 PL"]+[""]*11,
        ["1","0102AZ103710N","BRACKET MTG INTERCOOLER RH"]+[""]*11,
        ["1","SF0309005","NUT SQ WLD M6X1X6.5X8 PL"]+[""]*11,
    ],3): drow(ws,i,row,even=(i%2==0))

    # 9. Packaging OE
    ws=wb.create_sheet("Packaging OE")
    hdr(ws,1,["Name","Value"],[24,24])
    for i,row in enumerate([["Packaging Style","PrimaryPackaging"],["Packaging Type","Plastic Crate"],
        ["Size (Mtr)","0.6 X 0.4 X 0.2"],["Qty/ packaging","30"]],2):
        drow(ws,i,row,even=(i%2==0))

    # 10. Inhouse Parts Heat Treatment
    ws=wb.create_sheet("Inhouse Parts Heat Treatment")
    hdr(ws,1,["Level","Part No. ","Part Description","Operation No","HT Operation Description",
        "Machine Make","Machine Spec","Parameter Type 1","Parameter 1 Unit Of Measure","Parameter 1 Value"],
        [8,18,30,12,24,16,20,16,20,16])
    ref_row(ws,2,10); ws.freeze_panes="A2"

    # 11. Inhouse Parts SF and ST
    ws=wb.create_sheet("Inhouse Parts SF and ST")
    hdr(ws,1,["Level","Part No. ","Part Description","Operation No","Surface Treatment Description",
        "Machine Make","Machine Spec","Parameter Type 1","Parameter 1 Unit Of Measure","Parameter 1 Value"],
        [8,18,30,12,28,16,20,16,20,16])
    ref_row(ws,2,10); ws.freeze_panes="A2"

    # 12. DVP
    ws=wb.create_sheet("DVP")
    hdr(ws,1,["Level","Part No.","DVP Name","DVP Description","DVP Acceptance Criteria",
        "DVP Location","Country","Inhouse /outsource","Name of Agency incase of Outsource",
        "Equipment / Machine Used","No Of Hrs/ Test Qty","No. of Samples","TSO Remarks"],
        [8,18,22,30,28,14,12,18,22,24,14,14,36])
    ref_row(ws,2,13); ws.freeze_panes="A2"
    dvp=[
        ["0",ASSY,"CMM Dimensional","CTQ part CMM measurement","As per M&M PIST targets",
         "Pune","India","Inhouse","","Fix Bed CMM + Portable CMM","30 Parts/Month","30",
         "CTQ(CMM): 30 Parts/month & 100% checking at auto checking gauge with interlocking"],
        ["0",ASSY,"Chisel Test","Projection weld chisel strength test","Pass",
         "Pune","India","Inhouse","","Chisel Test Setup","After every 2 hrs","3",
         "3 parts/lot at VP/PP; After every 2 hrs at SOP"],
        ["0",ASSY,"Tear Down Test","Assembly tear-down inspection","Pass",
         "Pune","India","Inhouse","","Tear Down Gun","1 Part/Month","1",
         "1 part/lot at VP/PP; 1 part/month at SOP"],
        ["0",ASSY,"Torque Test","Nut torque verification","As per M&M standard",
         "Pune","India","Inhouse","","Force Gauge","05 part/lot","5","Post projection welding"],
        ["0",ASSY,"Vision System Check","Sealant presence & fastener weld 100% check","Pass/Fail",
         "Pune","India","Inhouse","","Vision Camera System","100%","All",
         "01 Vision system; all parts 100% checked & laser marking"],
        ["1","0102AZ103710N","Thinning Check","Ultrasonic thinning measurement",
         "Simulation Max 10% / Buyoff Max 15% / SOP Max 18%",
         "Pune","India","Inhouse","","Ultrasonic Thinning Meter","5 Parts/Month","5",
         "DFT meter and ultrasonic thinning meter mandatory"],
    ]
    for i,row in enumerate(dvp,3): drow(ws,i,row,even=(i%2==0))

    # 13. Engineering
    ws=wb.create_sheet("Engineering")
    hdr(ws,1,["Level ","Part No.","Engineering Operation","Engineering Location",
        "Name of Agency incase of Outsource","No of Hrs","Country"],[8,18,24,20,24,12,12])
    ref_row(ws,2,7); drow(ws,3,["0",ASSY,ASSY_DESC]+[""]*4,even=False); ws.freeze_panes="A2"

    # 14. Prototype Part
    ws=wb.create_sheet("Prototype Part")
    hdr(ws,1,["Level","Part No.","Part Description","Proto Qty","Country"],[8,18,32,12,14])
    ref_row(ws,2,5); ws.freeze_panes="A2"
    for i,row in enumerate([
        ["0",ASSY,ASSY_DESC,"",""],
        ["1","SF0309003","NUT SQ WLD M8X1.25X8X8 PL","",""],
        ["1","0102AZ103710N","BRACKET MTG INTERCOOLER RH","",""],
        ["1","SF0309005","NUT SQ WLD M6X1X6.5X8 PL","",""],
    ],3): drow(ws,i,row,even=(i%2==0))

    # 15. Prototype Tool
    ws=wb.create_sheet("Prototype Tool")
    hdr(ws,1,["Level","Part No.","Proto Tool Type","Qty ","Country","Tool Life"],[8,18,24,8,14,18])
    ref_row(ws,2,6); ws.freeze_panes="A2"
    for i,row in enumerate([
        ["0",ASSY,"","","",""],
        ["1","SF0309003","","","",""],
        ["1","0102AZ103710N","","","",""],
        ["1","SF0309005","","","",""],
    ],3): drow(ws,i,row,even=(i%2==0))

    # 16. VAVE
    ws=wb.create_sheet("VAVE")
    hdr(ws,1,["Level","Part No.","Current Status","Proposed Idea"],[8,18,28,36])
    ref_row(ws,2,4); ws.freeze_panes="A2"
    for i,row in enumerate([
        ["0",ASSY,"",""],["1","SF0309003","",""],
        ["1","0102AZ103710N","",""],["1","SF0309005","",""],
    ],3): drow(ws,i,row,even=(i%2==0))

    # 17. YOY
    ws=wb.create_sheet("YOY")
    hdr(ws,1,["Level","Part No.","Part Description","Volume Base ","Reference Date (MM/YY)"],
        [8,18,32,14,18])
    ref_row(ws,2,5); ws.freeze_panes="A2"
    for i,row in enumerate([
        ["0",ASSY,ASSY_DESC,"",""],
        ["1","SF0309003","NUT SQ WLD M8X1.25X8X8 PL","",""],
        ["1","0102AZ103710N","BRACKET MTG INTERCOOLER RH","",""],
        ["1","SF0309005","NUT SQ WLD M6X1X6.5X8 PL","",""],
    ],3): drow(ws,i,row,even=(i%2==0))

    # 18. TSO Summary (bonus info sheet)
    ws=wb.create_sheet("TSO Summary")
    hdr(ws,1,["Field","Value"],[28,44]); ws.freeze_panes="A2"
    for i,(k,v) in enumerate([
        ("Date",d["date"]),("Project Name",d["project"]),("Supplier Name",d["supplier"]),
        ("Stamping Location",d["stamp_loc"]),("Welding Location",d["weld_loc"]),
        ("No. of End Items",d["end_items"]),("M&M Signatory",d["mm_sig"]),
        ("Supplier Signatory",d["sup_sig"]),("Sign Date",d["sign_date"]),
    ],2):
        _c(ws,i,1,k,bold=True,bg="D9E1F2"); _c(ws,i,2,v,bg="EEF2F7" if i%2==0 else None)

    # 19. Project Targets (bonus)
    ws=wb.create_sheet("Project Targets")
    hdr(ws,1,["Target Category","VP","PP","SoP","Post SoP"],[28,22,22,22,22])
    ws.freeze_panes="A2"
    for i,row in enumerate(d["targets"],2): drow(ws,i,row,even=(i%2==0))
    if d["thinning"]:
        r=len(d["targets"])+3
        hdr(ws,r,["Thinning / General Targets","Value","Parameter 2","Value 2"],[28,14,16,14])
        for i,row in enumerate(d["thinning"],r+1): drow(ws,i,row,even=(i%2==0))

    # 20. General Guidelines (bonus)
    ws=wb.create_sheet("General Guidelines")
    hdr(ws,1,["General Guidelines"],[120])
    c2=ws.cell(row=2,column=1,value=d["guidelines"])
    c2.font=Font(name="Arial",size=10)
    c2.alignment=Alignment(wrap_text=True,vertical="top")
    c2.border=_b(); ws.row_dimensions[2].height=280

    # 21. Assembly & Welding (bonus)
    ws=wb.create_sheet("Assembly & Welding")
    hdr(ws,1,["Part No.","Rev No.","Part Name","Qty/Veh","CTQ Part",
        "Assy Size L","Assy Size W","Assy Size H","No. Proj. Welds",
        "FTG Description","Fixture Type","Operation No.",
        "Base Type","Base T","Base L","Base W","Clamp Type","Clamp Units","Support Units",
        "Inspection Type","Insp. Frequency","Remark"],
        [18,8,30,10,10,12,12,12,14,24,14,12,12,8,8,8,16,14,14,22,18,40])
    ws.freeze_panes="A2"
    a=d["asm"]
    drow(ws,2,[a.get("pno",""),a.get("rev",""),a.get("name",""),
        a.get("qv",""),a.get("ctq",""),
        a.get("sl",""),a.get("sw",""),a.get("sh",""),a.get("pw",""),
        "ASSY CHECKING FIXTURE","MANUAL","OP10",
        a.get("bt",""),a.get("bl",""),a.get("bw",""),
        a.get("ct",""),a.get("cu",""),a.get("su",""),
        "1. Dimensional\n2. Torque Test","05 part/lot",
        "Shutter type PY required on Projection machine for 100% tracing of nut presence"],
        even=False)

    # 22. Timeline (bonus)
    if d["tl"]:
        ws=wb.create_sheet("Timeline")
        hdr(ws,1,["Milestone","Duration"],[22,16])
        for i,row in enumerate(d["tl"],2): drow(ws,i,row,even=(i%2==0))

    wb.save(out)
    import sys
    if not hasattr(out, 'write'):
        print(f"Saved -> {out}")
        print(f"Sheets ({len(wb.worksheets)}): {[s.title for s in wb.worksheets]}")
