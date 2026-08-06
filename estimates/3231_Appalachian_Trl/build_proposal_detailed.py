#!/usr/bin/env python3
"""ARC customer proposal — 3231 Appalachian Trl. DETAILED work-step breakdown, A/B options.

Granularity rule: every row is a COMPLETE WORK STEP with its labor baked in.
Never a standalone "labor" row, and never hours, machine counts, day counts, or
unit rates — those are the details a customer can audit against what he sees in
the driveway, and they turn a correct estimate into an argument.

Row amounts are allocated from the internal workbook and sum exactly to the same
option totals. Verify: 1 page, small non-negative bottom space, and RENDER it.
"""

# ================================ CONFIG ================================
OUT   = "/tmp/claude-0/-home-user-RXS/0ad0031b-0623-52ca-a8ef-bf297e37d330/scratchpad/ARC_Proposal_3231_Appalachian_DETAILED.pdf"
LOGO  = "/tmp/claude-0/-home-user-RXS/0ad0031b-0623-52ca-a8ef-bf297e37d330/scratchpad/arc-logo.png"

JOB_TITLE = "WATER DAMAGE REPAIR ESTIMATE — MASTER BATHROOM"
EST_NO    = "ARC-2026-0806-001"
DATE      = "August 6, 2026"
CUSTOMER  = ["David Miniter"]
LOCATION  = ["3231 Appalachian Trl", "Kingwood, TX 77345"]

SCOPE_HEADING = "PROJECT SCOPE: MASTER BATHROOM — WATER DAMAGE REPAIR"
SCOPE_TEXT = (
    "Repair of the water-damaged wall and baseboard at the vanity wall of the master bathroom. The "
    "affected material is removed, the wall cavity is dried and treated, and the wall is rebuilt and "
    "refinished. Each stage of the work is broken out below. Two options are presented. Plumbing "
    "repair, tile work, flooring, and third-party testing are not included.")

# (work step, description, Option A amount, Option B amount)  None = not included
STEPS_TABLE = [
 ("Site Setup, Protection &amp; Removal of Damaged Material",
  "Protect the surrounding area and fixtures, open the wall, remove and dispose of the "
  "water-damaged material, and treat the exposed area.",
  "$177.59", "$177.59"),
 ("Structural Drying &amp; Daily Monitoring",
  "Professional drying equipment placed on the wall cavity, with a technician on site to monitor "
  "and document progress until the wall reaches a dry standard.",
  "$407.65", "$407.65"),
 ("Vanities Removed &amp; Reset",
  "Both vanities carefully detached and reset afterward, so the wall and baseboard behind them can "
  "be finished properly rather than cut around.",
  None, "$842.08"),
 ("Wall Rebuilt &amp; Textured",
  "New drywall hung, taped, floated, and textured to match the existing wall finish.",
  "$611.16", "$611.16"),
 ("Baseboard Removed &amp; Reset",
  "Option A covers the affected section of baseboard. Option B covers all baseboard in the bathroom.",
  "$7.44", "$69.44"),
 ("Priming &amp; Painting",
  "Option A primes and paints the repaired area. Option B repaints the walls and baseboard wall to "
  "wall for a uniform finish. A painter's minimum charge applies either way, so the full repaint "
  "adds very little.",
  "$316.08", "$385.24"),
 ("Final Detailed Cleaning",
  "Detailed cleaning of the bathroom once the work is complete.",
  None, "$46.80"),
]

SUB_A, TAX_A, TOT_A = "$1,519.92", "$50.16", "$1,570.08"
SUB_B, TAX_B, TOT_B = "$2,539.96", "$83.82", "$2,623.78"

OPT_A_LABEL, OPT_A_SUB = "OPTION A", "Repair affected area"
OPT_B_LABEL, OPT_B_SUB = "OPTION B", "Full finish, no patch"

NOTE = (
    "Note: Labor to repair, remodel, or restore residential property is not subject to Texas sales tax; only "
    "materials are taxed. The material tax shown is an estimate and the exact amount will be reflected on your "
    "final invoice. This estimate is based on visible conditions at the time of inspection and is subject to "
    "change once the wall is opened; any additional conditions found will be priced by written change order and "
    "approved by you before that work proceeds. Please note the damaged wall backs the vanity plumbing. The "
    "source of the moisture is not included in this estimate and should be verified and repaired by a licensed "
    "plumber; without that repair the damage can return. A 3.5% surcharge applies to credit card payments; no "
    "surcharge on check, ACH, or debit.")

STEPS = [
 ("STEP 1: APPROVE",
  "Let us know which option you would like to move forward with and we will confirm your project scope."),
 ("STEP 2: SCHEDULE",
  "We will coordinate a start date that works for you and confirm it in writing before we arrive."),
 ("STEP 3: CLEAR AREA",
  "Please clear the vanity countertops and remove personal items from the bathroom before our crew arrives."),
]
# ============================== END CONFIG ==============================

from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_RIGHT, TA_JUSTIFY, TA_CENTER
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph,
                                Spacer, Table, TableStyle, Image, KeepTogether)

NAVY=colors.HexColor("#16243F"); RED=colors.HexColor("#C41230")
BLUE1=colors.HexColor("#DDE4EE"); GREY=colors.HexColor("#6B7684")
LINE=colors.HexColor("#CBD3DC"); INK=colors.HexColor("#1A1A1A")
W,H=LETTER; LM=RM=0.55*inch
S=lambda **k: ParagraphStyle(**k)

st_lab  = S(name="lab",fontName="Helvetica-Bold",fontSize=6.4,leading=8,textColor=GREY)
st_val  = S(name="val",fontName="Helvetica-Bold",fontSize=8.6,leading=11,textColor=NAVY)
st_val2 = S(name="val2",fontName="Helvetica",fontSize=8,leading=10.5,textColor=INK)
st_redhd= S(name="rh",fontName="Helvetica-Bold",fontSize=8.6,leading=11,textColor=RED,spaceBefore=5,spaceAfter=2)
st_scope= S(name="sc",fontName="Helvetica",fontSize=7.6,leading=10.0,textColor=INK,alignment=TA_JUSTIFY)
st_stepn= S(name="sn",fontName="Helvetica-Bold",fontSize=8.0,leading=9.8,textColor=NAVY)
st_stepd= S(name="sd",fontName="Helvetica",fontSize=6.9,leading=8.6,textColor=colors.HexColor("#44505F"))
st_amt  = S(name="am",fontName="Helvetica-Bold",fontSize=8.4,leading=10.5,textColor=NAVY,alignment=TA_CENTER)
st_na   = S(name="na",fontName="Helvetica-Oblique",fontSize=6.6,leading=9,textColor=GREY,alignment=TA_CENTER)
st_totk = S(name="tk",fontName="Helvetica",fontSize=7.8,leading=10.5,textColor=INK,alignment=TA_RIGHT)
st_totv = S(name="tv",fontName="Helvetica-Bold",fontSize=7.8,leading=10.5,textColor=NAVY,alignment=TA_CENTER)
st_note = S(name="nt",fontName="Helvetica",fontSize=6.0,leading=7.7,textColor=GREY,alignment=TA_JUSTIFY)
st_stpH = S(name="sh",fontName="Helvetica-Bold",fontSize=7.4,leading=9.6,textColor=RED)
st_stpB = S(name="sb",fontName="Helvetica",fontSize=6.9,leading=9.0,textColor=INK)

def deco(canv,doc):
    canv.saveState(); canv.setFillColor(colors.HexColor("#8A96A3")); canv.setFont("Helvetica",6.6)
    canv.drawCentredString(W/2.0,0.34*inch,
        "RXS Group LLC dba America's Restoration Contractor   |   Restoring Lives, Rebuilding Communities")
    canv.restoreState()

doc=BaseDocTemplate(OUT,pagesize=LETTER,leftMargin=LM,rightMargin=RM,topMargin=0.34*inch,
                    bottomMargin=0.42*inch,title="ARC Estimate",author="America's Restoration Contractor")
doc.addPageTemplates([PageTemplate(id="p",frames=[Frame(LM,doc.bottomMargin,W-LM-RM,
                      H-doc.topMargin-doc.bottomMargin,id="f")],onPage=deco)])
CW=W-LM-RM; E=[]

contact=Paragraph("<b>Phone:</b> (832) 650-6500<br/><b>Email:</b> help@arcsvcs.com<br/>"
    "<b>Web:</b> arcsvcs.com<br/><font color='#C41230'><b>LICENSED &amp; INSURED</b></font>",
    S(name="rt",fontName="Helvetica",fontSize=7.3,leading=10.0,textColor=INK,alignment=TA_RIGHT))
head=Table([[Image(LOGO,width=1.22*inch,height=0.645*inch),contact]],colWidths=[1.36*inch,CW-1.36*inch])
head.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(0,0),0),
    ("RIGHTPADDING",(-1,0),(-1,0),0),("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0)]))
E.append(head); E.append(Spacer(1,6))

bar=Table([[Paragraph("<font color='#FFFFFF'><b>%s</b></font>"%JOB_TITLE,
              S(name="bt",fontName="Helvetica-Bold",fontSize=10.0,leading=12.5,textColor=colors.white)),
            Paragraph("<font color='#FFFFFF'>Estimate #: %s</font>"%EST_NO,
              S(name="bn",fontName="Helvetica",fontSize=7.3,leading=12.5,textColor=colors.white,alignment=TA_RIGHT))]],
          colWidths=[CW*0.72,CW*0.28])
bar.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),NAVY),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ("LEFTPADDING",(0,0),(0,0),9),("RIGHTPADDING",(-1,0),(-1,0),9),
    ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5)]))
E.append(bar)

def block(label,lines):
    cells=[[Paragraph(label,st_lab)]]
    for i,t in enumerate(lines): cells.append([Paragraph(t,st_val if i==0 else st_val2)])
    t=Table(cells,colWidths=[CW/3-0.28*inch])
    t.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
        ("TOPPADDING",(0,0),(-1,-1),1),("BOTTOMPADDING",(0,0),(-1,-1),1)]))
    return t
info=Table([[block("DATE",[DATE]),block("PREPARED FOR",CUSTOMER),block("PROPERTY LOCATION",LOCATION)]],
           colWidths=[CW/3]*3)
info.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),9),
    ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4)]))
E.append(info)

E.append(Paragraph(SCOPE_HEADING,st_redhd))
sb=Table([[Paragraph(SCOPE_TEXT,st_scope)]],colWidths=[CW])
sb.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#F4F6F9")),
    ("LINEBEFORE",(0,0),(0,-1),2.5,RED),("LEFTPADDING",(0,0),(-1,-1),9),
    ("RIGHTPADDING",(0,0),(-1,-1),9),("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5)]))
E.append(sb)

# ---- detailed work-step breakdown ----
E.append(Paragraph("SCOPE OF WORK &amp; PRICING",st_redhd))
AW=1.06*inch; DW_=CW-2*AW
hw=lambda t: Paragraph("<font color='#FFFFFF'><b>%s</b></font>"%t,
      S(name="h",fontName="Helvetica-Bold",fontSize=7.5,leading=9.2,textColor=colors.white,alignment=TA_CENTER))
sub=lambda t: Paragraph("<font color='#FFFFFF'>%s</font>"%t,
      S(name="hs",fontName="Helvetica",fontSize=5.7,leading=6.8,textColor=colors.white,alignment=TA_CENTER))
rows=[[Paragraph("<font color='#FFFFFF'><b>Scope of Work</b></font>",
        S(name="hl",fontName="Helvetica-Bold",fontSize=7.5,leading=9.2,textColor=colors.white)),
       [hw(OPT_A_LABEL),sub(OPT_A_SUB)],[hw(OPT_B_LABEL),sub(OPT_B_SUB)]]]
amt=lambda v: Paragraph(v,st_amt) if v else Paragraph("Not included",st_na)
body=[]
for i,(nm,desc,a,b) in enumerate(STEPS_TABLE,start=1):
    rows.append([[Paragraph(nm,st_stepn),Paragraph(desc,st_stepd)],amt(a),amt(b)])
    body.append(i)
sr=len(rows); rows.append([Paragraph("Line Item Subtotal",st_totk),
                           Paragraph(SUB_A,st_totv),Paragraph(SUB_B,st_totv)])
tr=len(rows); rows.append([Paragraph("Est. Material Sales Tax (8.25%, materials only)",st_totk),
                           Paragraph(TAX_A,st_totv),Paragraph(TAX_B,st_totv)])
ev=lambda t: Paragraph("<font color='#FFFFFF'><b>%s</b></font>"%t,
      S(name="ev",fontName="Helvetica-Bold",fontSize=10.2,leading=12.6,textColor=colors.white,alignment=TA_CENTER))
totr=len(rows); rows.append([Paragraph("<font color='#FFFFFF'><b>ESTIMATED TOTAL</b></font>",
    S(name="et",fontName="Helvetica-Bold",fontSize=10.2,leading=12.6,textColor=colors.white,alignment=TA_RIGHT)),
    ev(TOT_A),ev(TOT_B)])

tbl=Table(rows,colWidths=[DW_,AW,AW],repeatRows=1)
sty=[("BACKGROUND",(0,0),(-1,0),NAVY),("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8),
     ("TOPPADDING",(0,0),(-1,0),4),("BOTTOMPADDING",(0,0),(-1,0),4),("VALIGN",(0,0),(-1,-1),"MIDDLE")]
for i in body:
    sty+=[("TOPPADDING",(0,i),(-1,i),4),("BOTTOMPADDING",(0,i),(-1,i),4),
          ("LINEBELOW",(0,i),(-1,i),0.4,LINE),("VALIGN",(0,i),(0,i),"TOP")]
    if i%2==1: sty+=[("BACKGROUND",(0,i),(-1,i),colors.HexColor("#F7F9FC"))]
sty+=[("TOPPADDING",(0,sr),(-1,tr),2),("BOTTOMPADDING",(0,sr),(-1,tr),2),
      ("BACKGROUND",(0,totr),(-1,totr),NAVY),("TOPPADDING",(0,totr),(-1,totr),5),
      ("BOTTOMPADDING",(0,totr),(-1,totr),5),
      ("LINEAFTER",(0,0),(0,totr-1),0.4,LINE),("LINEAFTER",(1,0),(1,totr-1),0.4,LINE)]
tbl.setStyle(TableStyle(sty))
E.append(tbl); E.append(Spacer(1,5))
E.append(Paragraph(NOTE,st_note)); E.append(Spacer(1,5))

hd=Paragraph("NEXT STEPS",st_redhd)
rule=Table([[""]],colWidths=[CW],rowHeights=[1.6])
rule.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),RED),("LEFTPADDING",(0,0),(-1,-1),0),
    ("RIGHTPADDING",(0,0),(-1,-1),0),("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0)]))
cells=[[Paragraph(t,st_stpH),Paragraph(b,st_stpB)] for t,b in STEPS]
cols=Table([[Table([[c[0]],[c[1]]],colWidths=[CW/3-0.18*inch]) for c in cells]],colWidths=[CW/3]*3)
cols.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),0),
    ("RIGHTPADDING",(0,0),(-1,-1),9),("TOPPADDING",(0,0),(-1,-1),3)]))
E.append(KeepTogether([hd,rule,Spacer(1,3),cols]))

doc.build(E)
print("built",OUT)

# ---- arithmetic check: rows must sum to the stated subtotals ----
def money(s): return float(s.replace("$","").replace(",",""))
a=sum(money(x[2]) for x in STEPS_TABLE if x[2])
b=sum(money(x[3]) for x in STEPS_TABLE if x[3])
print(f"Option A rows: {a:.2f}  vs subtotal {money(SUB_A):.2f}  ->", "OK" if abs(a-money(SUB_A))<0.01 else "MISMATCH")
print(f"Option B rows: {b:.2f}  vs subtotal {money(SUB_B):.2f}  ->", "OK" if abs(b-money(SUB_B))<0.01 else "MISMATCH")
