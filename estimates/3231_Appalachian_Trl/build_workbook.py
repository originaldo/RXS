#!/usr/bin/env python3
"""ARC internal line-item workbook — 3231 Appalachian Trl (water-only, A/B options).
Formula-driven off a JOB ASSUMPTIONS block. No hardcoded totals."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

NAVY="16243F"; RED="C41230"; YEL="FFF2CC"; BLUE1="DDE4EE"; GREY="6B7684"; PULLC="FCE4D6"
wb=openpyxl.Workbook()
ws=wb.active; ws.title="Internal Line Items"

thin=Side(style="thin",color="CBD3DC")
box=Border(left=thin,right=thin,top=thin,bottom=thin)
def C(cell,v,bold=False,color="1A1A1A",size=10,fill=None,align=None,italic=False,border=False,fmt=None,wrap=False):
    cell.value=v
    cell.font=Font(name="Helvetica",bold=bold,color=color,size=size,italic=italic)
    if fill: cell.fill=PatternFill("solid",fgColor=fill)
    if align: cell.alignment=Alignment(horizontal=align,vertical="center",wrap_text=wrap)
    else: cell.alignment=Alignment(vertical="center",wrap_text=wrap)
    if border: cell.border=box
    if fmt: cell.number_format=fmt

widths={"A":5,"B":9,"C":11,"D":46,"E":6,"F":9,"G":11,"H":12,"I":40}
for c,w in widths.items(): ws.column_dimensions[c].width=w

r=1
ws.merge_cells(f"A{r}:I{r}")
C(ws[f"A{r}"],"AMERICA'S RESTORATION CONTRACTOR  —  INTERNAL ESTIMATE WORKBOOK",bold=True,color=NAVY,size=13)
r+=1
ws.merge_cells(f"A{r}:I{r}")
C(ws[f"A{r}"],"3231 Appalachian Trl, Kingwood, TX 77345  ·  Master Bathroom  ·  Category: WTR (Water — no mold / no containment)  ·  Price List: TXHO8X_JUL26",color=GREY,size=9)
r+=1
ws.merge_cells(f"A{r}:I{r}")
C(ws[f"A{r}"],"Retail / cash customer — price-sensitive, comparing 3 bids.  Two options: A = minimal patch, B = proper (detach vanities, full base R&R, full repaint).",color=GREY,size=9,italic=True)
r+=2

# ---------------- JOB ASSUMPTIONS ----------------
ws.merge_cells(f"A{r}:I{r}"); C(ws[f"A{r}"],"JOB ASSUMPTIONS  (edit yellow cells — all quantities below recompute)",bold=True,color="FFFFFF",fill=NAVY,size=10); r+=1
A={}
def assume(label,val,key,note="",fmt=None):
    global r
    C(ws[f"B{r}"],label,size=9); ws.merge_cells(f"B{r}:C{r}")
    cell=ws[f"D{r}"]; C(cell,val,bold=True,fill=YEL,align="center",border=True,fmt=fmt)
    C(ws[f"E{r}"],note,color=GREY,size=8,italic=True); ws.merge_cells(f"E{r}:I{r}")
    A[key]=f"$D${r}"; r+=1
assume("Ceiling height (ft)",8.0,"ceil","CONFIRM — assumed standard 8'")
assume("Water category (1/2/3)",2,"cat","aged stain w/ spotting → treat as Cat 2 (grey)")
assume("Drying days",2,"dry","small wall cavity; drop if meters read dry on arrival")
assume("Site days (Option A)",3,"sdA","drying days + 1 (setup + monitor + takedown)")
assume("Extra site days (Option B)",1,"sdBx","vanity R&R + larger repaint adds ~1 day")
assume("Techs on site",1,"tech","")
assume("Monitoring hrs / tech / day",1.75,"mhr","ARC established basis — carries drive time to Kingwood")
assume("A: drywall cut width (ft)",2.0,"acw","")
assume("A: drywall cut height (ft)",2.0,"ach","")
assume("A: affected baseboard (LF)",3,"abase","")
assume("A: antimicrobial surface (SF)",16,"aamb","cut cavity + surround")
assume("A: wall face repaint (SF)",32,"apaint","blend the repaired wall corner-to-corner")
assume("B: vanities detach & reset (LF)",6.5,"bvan","his + hers combined face LF — CONFIRM")
assume("B: full baseboard D&R (LF)",28,"bbase","full affected-area perimeter — CONFIRM")
assume("B: full-area wall repaint (SF)",200,"bpaint","paintable grey walls, excl. tile accent — CONFIRM")
assume("Air movers",1,"am","")
assume("Dehumidifiers",1,"dh","")
assume("Material % of line total (tax basis)",0.40,"matpct","placeholder — true up from Xactimate 'Mat Only'","0%")
assume("Material sales tax rate",0.0825,"taxr","Houston 8.25%, materials only","0.00%")
r+=1
# derived helper strings
DW=f"({A['acw']}*{A['ach']})"           # option A drywall SF
EQA=f"({A['mhr']}*{A['tech']}*{A['sdA']})"
EQBX=f"({A['mhr']}*{A['tech']}*{A['sdBx']})"

# ---------------- LINE ITEM TABLE ----------------
hdr=["#","Cat","Sel","Description","Unit","Qty","Unit $","Extension","Note"]
def header_row():
    global r
    for i,h in enumerate(hdr):
        C(ws.cell(r,i+1),h,bold=True,color="FFFFFF",fill=NAVY,size=9,
          align=("right" if h in("Qty","Unit $","Extension") else "left"),border=True)
    r+=1

n=[0]
ext_cells={"MIT":[],"A":[],"B":[]}
def line(cat,sel,desc,unit,qty_formula,price,note="",bucket="MIT",pull=False):
    global r
    n[0]+=1
    C(ws.cell(r,1),n[0],size=9,border=True,align="center")
    C(ws.cell(r,2),cat,size=9,border=True)
    C(ws.cell(r,3),sel,size=9,border=True)
    C(ws.cell(r,4),desc,size=9,border=True,wrap=True)
    C(ws.cell(r,5),unit,size=9,border=True,align="center")
    qc=ws.cell(r,6); C(qc,None,size=9,border=True,align="right",fmt="0.00"); qc.value=f"={qty_formula}"
    pc=ws.cell(r,7)
    if pull:
        C(pc,"PULL",bold=True,fill=PULLC,size=9,border=True,align="center")
        ec=ws.cell(r,8); C(ec,None,size=9,border=True,align="right",fmt='"$"#,##0.00'); ec.value=0
    else:
        C(pc,price,size=9,border=True,align="right",fmt='"$"#,##0.00')
        ec=ws.cell(r,8); C(ec,None,bold=True,size=9,border=True,align="right",fmt='"$"#,##0.00')
        ec.value=f"=F{r}*G{r}"
    C(ws.cell(r,9),note,color=GREY,size=8,italic=True,border=True,wrap=True)
    ext_cells[bucket].append(f"H{r}"); r+=1

def section(title,color=RED):
    global r
    ws.merge_cells(f"A{r}:I{r}")
    C(ws[f"A{r}"],title,bold=True,color="FFFFFF",fill=color,size=10); r+=1

# ---- MITIGATION (common to A & B) ----
section("SECTION 1 — WATER MITIGATION & DRY-OUT  (common to both options)",NAVY)
header_row()
line("WTR","DRYW","Tear out wet drywall, cleanup, bag for disposal","SF",DW,1.16,"2'x2' affected cut",bucket="MIT")
line("WTR","GRMB","Apply plant-based anti-microbial agent to the surface area","SF",A['aamb'][1:].replace('$',''),0.38,"cavity + surround",bucket="MIT")
line("WTR","DRY","Air mover (per 24 hr period) - no monitoring","EA",f"{A['am']}*{A['dry']}",26.77,"1 unit x drying days",bucket="MIT")
line("WTR","DHM","Dehumidifier (per 24 hr period) - up to 69 ppd","EA",f"{A['dh']}*{A['dry']}",60.24,"1 unit x drying days; drop if dry on arrival",bucket="MIT")
line("WTR","EQ","Equipment setup, take down, and monitoring (hourly)","HR",EQA,66.75,"1.75 x techs x site days (A). Carries drive time.",bucket="MIT")
line("CON","LAB","Content manipulation charge - per hour","HR","1",50.06,"move counter items / work around vanity",bucket="MIT")

# ---- REBUILD OPTION A ----
section("SECTION 2A — REBUILD · OPTION A (minimal patch)",RED)
header_row()
line("DRY","1/2-","1/2\" drywall - hung, taped, floated, ready for texture","SF",DW,3.59,"2x2 patch",bucket="A")
line("DRY","TEX","Texture drywall - light hand texture","SF",DW,1.41,"match wall texture",bucket="A")
line("DRY","LAB","Drywall installer/finisher - per hour","HR","1.5",131.50,"patch finish over multiple trips — in lieu of $592 drywall minimum",bucket="A")
line("WTR","BASERS","Baseboard - detach & reset","LF",A['abase'][1:].replace('$',''),2.48,"reuse affected base",bucket="A")
line("PNT","SP","Seal/prime (1 coat) then paint (1 coat)","SF",A['apaint'][1:].replace('$',''),1.21,"blend repaired wall face",bucket="A")
line("PNT","B","Paint baseboard - one coat","LF",A['abase'][1:].replace('$',''),1.21,"",bucket="A")
line("CLN","FCT","Clean floor - tile","SF","20",0.78,"final clean of work area",bucket="A")

# ---- REBUILD OPTION B ----
section("SECTION 2B — REBUILD · OPTION B (proper: detach vanities, full base R&R, full repaint)",RED)
header_row()
line("DRY","1/2-","1/2\" drywall - hung, taped, floated, ready for texture","SF",DW,3.59,"same 2x2 repair",bucket="B")
line("DRY","TEX","Texture drywall - light hand texture","SF",DW,1.41,"",bucket="B")
line("DRY","LAB","Drywall installer/finisher - per hour","HR","1.5",131.50,"in lieu of drywall minimum",bucket="B")
line("CAB","VANTRS","Vanity with countertop - detach & reset","LF",A['bvan'][1:].replace('$',''),111.58,"both vanities (his + hers), quartz tops",bucket="B")
line("WTR","BASERS","Baseboard - detach & reset","LF",A['bbase'][1:].replace('$',''),2.48,"full area",bucket="B")
line("PNT","SP2","Seal/prime (1 coat) then paint (2 coats)","SF",A['bpaint'][1:].replace('$',''),1.67,"full repaint of area (not a patch)",bucket="B")
line("PNT","B2","Paint baseboard - two coats","LF",A['bbase'][1:].replace('$',''),1.83,"full area base",bucket="B")
line("CLN","FCT","Clean floor - tile","SF","60",0.78,"final clean, larger area",bucket="B")
line("WTR","EQ","Equipment setup, take down, and monitoring - added day","HR",EQBX,66.75,"extra monitoring/labor day for Option B scope",bucket="B")
r+=1

# ---------------- TOTALS ----------------
def sumcells(cells): return "+".join(cells) if cells else "0"
mit=sumcells(ext_cells["MIT"]); aa=sumcells(ext_cells["A"]); bb=sumcells(ext_cells["B"])

def total_block(title,subformula):
    global r
    C(ws.cell(r,4),title,bold=True,color=NAVY,size=10,align="right")
    sc=ws.cell(r,8); C(sc,None,bold=True,color=NAVY,size=10,border=True,align="right",fmt='"$"#,##0.00')
    sub_row=r; sc.value=f"={subformula}"; r+=1
    C(ws.cell(r,4),"Est. material sales tax (materials only)",size=9,align="right",color=GREY)
    tc=ws.cell(r,8); C(tc,None,size=9,border=True,align="right",fmt='"$"#,##0.00',color=GREY)
    tc.value=f"=ROUND(H{sub_row}*{A['matpct']}*{A['taxr']},2)"; tax_row=r; r+=1
    C(ws.cell(r,4),title.replace("Subtotal","TOTAL"),bold=True,color="FFFFFF",fill=NAVY,size=10,align="right")
    for cc in (5,6,7): ws.cell(r,cc).fill=PatternFill("solid",fgColor=NAVY)
    gc=ws.cell(r,8); C(gc,None,bold=True,color="FFFFFF",fill=NAVY,size=11,align="right",fmt='"$"#,##0.00')
    gc.value=f"=H{sub_row}+H{tax_row}"; tot_row=r; r+=1
    return tot_row

section("OPTION A — TOTAL (Mitigation + Rebuild A)",NAVY)
A_total=total_block("Option A Subtotal",f"{mit}+{aa}")
r+=1
section("OPTION B — TOTAL (Mitigation + Rebuild B)",NAVY)
B_total=total_block("Option B Subtotal",f"{mit}+{bb}")
r+=2

# ---------------- MARGIN ANALYSIS ----------------
section("MARGIN / PRICING ANALYSIS  (contribution margin — tech labor is a fixed weekly cost, not incremental)",NAVY)
C(ws.cell(r,2),"Out-of-pocket cost = materials + dump + rental (NOT tech labor). This is the true floor.",color=GREY,size=9,italic=True); ws.merge_cells(f"B{r}:I{r}"); r+=1
def marg(label,book_cell,oop):
    global r
    C(ws.cell(r,2),label,bold=True,size=9); ws.merge_cells(f"B{r}:C{r}")
    C(ws.cell(r,4),"Book total:",size=9,align="right")
    bc=ws.cell(r,5); C(bc,None,bold=True,size=9,align="left",fmt='"$"#,##0'); bc.value=f"=H{book_cell}"
    C(ws.cell(r,6),"OOP floor:",size=9,align="right")
    oc=ws.cell(r,7); C(oc,oop,size=9,align="left",fmt='"$"#,##0')
    C(ws.cell(r,8),"Contrib. margin:",size=9,align="right");
    mc=ws.cell(r,9); C(mc,None,bold=True,color="1F7A1F",size=9,fmt="0%"); mc.value=f"=(H{book_cell}-G{r})/H{book_cell}"
    r+=1
marg("OPTION A",A_total,70)
marg("OPTION B",B_total,150)
r+=1
C(ws.cell(r,2),"Suggested sell (round numbers):",bold=True,color=RED,size=10); ws.merge_cells(f"B{r}:D{r}")
C(ws.cell(r,5),"Option A ≈ $850–$950   ·   Option B ≈ $2,100–$2,300",bold=True,color=NAVY,size=10); ws.merge_cells(f"E{r}:I{r}"); r+=1
C(ws.cell(r,2),"Room to move: both options sit far above the OOP floor, so there is large room to match a competitor's number and still make strong margin. Levers to drop price: cut dehu / a drying day if it meters dry, trim monitoring hours, Option B detach only the affected-side vanity.",color=GREY,size=8,italic=True,wrap=True); ws.merge_cells(f"B{r}:I{r}"); ws.row_dimensions[r].height=42; r+=2

# ---------------- NOTES ----------------
section("NOTES / STRATEGY / DISCLOSURES",NAVY)
notes=[
 "CATEGORY: Built as WTR (water) per direction — no mold remediation, no containment, no HMR rates. A small dark spot is visible at the base of the affected wall; if opening reveals mold >25 contiguous SF, TX TDLR requires stop-work and conversion to a licensed mold path (separate scope).",
 "LABOR MINIMUMS: Xactimate would apply full trade minimums on a job this small (drywall ~$592, painting ~$316, insulation ~$202, finish carpentry ~$264 = ~$1,374 of minimums alone). Those are intentionally NOT stacked here — the actual production + a modest finish-labor line + the monitoring line are priced instead. This is what keeps the retail number competitive. If billing insurance, switch to the trade-minimum build.",
 "MONITORING (EQ): WTR rate $66.75/hr used (NOT the HMR $78.34). This line carries the drive time to Kingwood — do not cut it to soften the total.",
 "MOISTURE SOURCE: The affected wall backs the vanity plumbing wall (his/hers vanities adjacent). Probable source is a vanity supply/drain or caulk failure, NOT covered in this scope. Disclose in writing: a plumber should verify/repair the source before close-up or the issue can recur. This disclosure protects ARC.",
 "INSULATION: Not included — assumed interior partition. If wet insulation is found in the cavity, add WTR INS ($0.86/SF) by change order.",
 "BASEBOARD: Priced as detach & reset (reuse existing). If the affected piece is swollen, replace that piece only — FCW LAMB $7.62/LF is the only base-install line in the current export (verify base profile in Xactimate before quoting a full replace).",
 "TAX: Materials-only 8.25% on a 40% material placeholder — an ESTIMATE. Texas labor for residential restoration is not taxable. True up from Xactimate 'Mat Only' before send.",
 "CUSTOMER WANTS A COST BREAKDOWN: proposal will show price by scope section (mitigation / rebuild) for each option — enough to satisfy the request without itemizing machines or day counts.",
 "NO mold testing included (TDLR conflict-of-interest — that belongs to a third-party assessor).",
]
for i,t in enumerate(notes,1):
    C(ws.cell(r,2),f"{i}.",bold=True,size=8,color=GREY)
    C(ws.cell(r,3),t,size=8,color="333333",wrap=True); ws.merge_cells(f"C{r}:I{r}")
    ws.row_dimensions[r].height=14+ (len(t)//95)*11
    r+=1

ws.freeze_panes="A6"
out="/tmp/claude-0/-home-user-RXS/0ad0031b-0623-52ca-a8ef-bf297e37d330/scratchpad/ARC_Estimate_3231_Appalachian_INTERNAL.xlsx"
wb.save(out)
print("saved",out)
# quick recompute check via a second pass using LibreOffice-free manual calc is skipped;
# openpyxl stores formulas. Print the assumption map for sanity.
print("Option A total row:",A_total,"Option B total row:",B_total)
