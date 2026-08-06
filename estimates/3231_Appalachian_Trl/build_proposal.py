#!/usr/bin/env python3
"""ARC customer proposal — 3231 Appalachian Trl. Two-option (A/B) comparison variant.
Edit CONFIG only. Verify: 1 page, small non-negative bottom space, and RENDER it."""

# ================================ CONFIG ================================
OUT   = "/tmp/claude-0/-home-user-RXS/0ad0031b-0623-52ca-a8ef-bf297e37d330/scratchpad/ARC_Proposal_3231_Appalachian.pdf"
LOGO  = "/tmp/claude-0/-home-user-RXS/0ad0031b-0623-52ca-a8ef-bf297e37d330/scratchpad/arc-logo.png"

JOB_TITLE = "WATER DAMAGE REPAIR ESTIMATE — MASTER BATHROOM"
EST_NO    = "ARC-2026-0806-001"
DATE      = "August 6, 2026"
CUSTOMER  = ["Homeowner"]          # TODO: replace with the homeowner's name
LOCATION  = ["3231 Appalachian Trl", "Kingwood, TX 77345"]

INCLUDED_LEFT  = ["Removal of water-damaged wall material",
                  "Anti-microbial treatment of the affected area",
                  "Professional drying of the wall cavity",
                  "Daily monitoring to a documented dry standard"]
INCLUDED_RIGHT = ["Wall rebuilt, textured, and refinished",
                  "Baseboard removed and reset",
                  "Debris removal and disposal",
                  "Final cleanup of the work area"]

SCOPE_HEADING = "PROJECT SCOPE: MASTER BATHROOM — WATER DAMAGE REPAIR"
SCOPE_TEXT = (
    "Repair of water-damaged wall and baseboard at the vanity wall of the master bathroom. The "
    "affected material is removed, the wall cavity is dried and treated, and the wall is rebuilt "
    "and refinished. Two options are presented below. Plumbing repair, tile work, flooring, and "
    "third-party testing are not included.")

# (section name, description, option A amount, option B amount)
GROUPS = [
 ("Water Mitigation &amp; Dry-Out",
  "Remove the water-damaged wall material, treat the area, and dry the wall cavity with "
  "professional drying equipment until it reaches a documented dry standard.",
  "$585.24", "$585.24"),
 ("Repairs &amp; Finish",
  "<b>Option A</b> — rebuild and refinish the repaired section only, and reset the affected "
  "baseboard.<br/><b>Option B</b> — both vanities carefully removed and reset for full access, "
  "all baseboard removed and reset, and the bathroom walls and baseboard repainted wall-to-wall "
  "for a uniform finish rather than a patch.",
  "$934.68", "$1,954.72"),
]

SUB_A, TAX_A, TOT_A = "$1,519.92", "$50.16", "$1,570.08"
SUB_B, TAX_B, TOT_B = "$2,539.96", "$83.82", "$2,623.78"

OPT_A_LABEL = "OPTION A"
OPT_B_LABEL = "OPTION B"
OPT_A_SUB   = "Repair affected area"
OPT_B_SUB   = "Full finish, no patch"

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
                                Spacer, Table, TableStyle, Image, KeepTogether, Flowable)

NAVY=colors.HexColor("#16243F"); RED=colors.HexColor("#C41230")
BLUE1=colors.HexColor("#DDE4EE"); GREY=colors.HexColor("#6B7684")
LINE=colors.HexColor("#CBD3DC"); INK=colors.HexColor("#1A1A1A")
W,H=LETTER; LM=RM=0.55*inch
S=lambda **k: ParagraphStyle(**k)

st_lab  = S(name="lab",fontName="Helvetica-Bold",fontSize=6.4,leading=8,textColor=GREY)
st_val  = S(name="val",fontName="Helvetica-Bold",fontSize=8.6,leading=11,textColor=NAVY)
st_val2 = S(name="val2",fontName="Helvetica",fontSize=8,leading=10.5,textColor=INK)
st_redhd= S(name="rh",fontName="Helvetica-Bold",fontSize=8.6,leading=11,textColor=RED,spaceBefore=6,spaceAfter=2)
st_scope= S(name="sc",fontName="Helvetica",fontSize=7.8,leading=10.4,textColor=INK,alignment=TA_JUSTIFY)
st_grpn = S(name="gn",fontName="Helvetica-Bold",fontSize=8.6,leading=11,textColor=NAVY)
st_amt  = S(name="am",fontName="Helvetica-Bold",fontSize=8.6,leading=11,textColor=NAVY,alignment=TA_CENTER)
st_desc = S(name="ds",fontName="Helvetica",fontSize=7.6,leading=9.8,textColor=INK)
st_totk = S(name="tk",fontName="Helvetica",fontSize=7.8,leading=10.5,textColor=INK,alignment=TA_RIGHT)
st_totv = S(name="tv",fontName="Helvetica-Bold",fontSize=7.8,leading=10.5,textColor=NAVY,alignment=TA_CENTER)
st_note = S(name="nt",fontName="Helvetica",fontSize=6.1,leading=7.9,textColor=GREY,alignment=TA_JUSTIFY)
st_stpH = S(name="sh",fontName="Helvetica-Bold",fontSize=7.6,leading=10,textColor=RED)
st_stpB = S(name="sb",fontName="Helvetica",fontSize=7,leading=9.4,textColor=INK)

def deco(canv,doc):
    canv.saveState(); canv.setFillColor(colors.HexColor("#8A96A3")); canv.setFont("Helvetica",6.6)
    canv.drawCentredString(W/2.0,0.34*inch,
        "RXS Group LLC dba America's Restoration Contractor   |   Restoring Lives, Rebuilding Communities")
    canv.restoreState()

doc=BaseDocTemplate(OUT,pagesize=LETTER,leftMargin=LM,rightMargin=RM,topMargin=0.36*inch,
                    bottomMargin=0.44*inch,title="ARC Estimate",author="America's Restoration Contractor")
doc.addPageTemplates([PageTemplate(id="p",frames=[Frame(LM,doc.bottomMargin,W-LM-RM,
                      H-doc.topMargin-doc.bottomMargin,id="f")],onPage=deco)])
CW=W-LM-RM; E=[]

# ---- letterhead (logo image only; no recreated wordmark) ----
contact=Paragraph("<b>Phone:</b> (832) 650-6500<br/><b>Email:</b> help@arcsvcs.com<br/>"
    "<b>Web:</b> arcsvcs.com<br/><font color='#C41230'><b>LICENSED &amp; INSURED</b></font>",
    S(name="rt",fontName="Helvetica",fontSize=7.4,leading=10.2,textColor=INK,alignment=TA_RIGHT))
logo=Image(LOGO,width=1.30*inch,height=0.687*inch)
head=Table([[logo,contact]],colWidths=[1.45*inch,CW-1.45*inch])
head.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(0,0),0),
                          ("RIGHTPADDING",(-1,0),(-1,0),0),("TOPPADDING",(0,0),(-1,-1),0),
                          ("BOTTOMPADDING",(0,0),(-1,-1),0)]))
E.append(head); E.append(Spacer(1,7))

bar=Table([[Paragraph("<font color='#FFFFFF'><b>%s</b></font>"%JOB_TITLE,
              S(name="bt",fontName="Helvetica-Bold",fontSize=10.2,leading=13,textColor=colors.white)),
            Paragraph("<font color='#FFFFFF'>Estimate #: %s</font>"%EST_NO,
              S(name="bn",fontName="Helvetica",fontSize=7.4,leading=13,textColor=colors.white,alignment=TA_RIGHT))]],
          colWidths=[CW*0.72,CW*0.28])
bar.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),NAVY),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ("LEFTPADDING",(0,0),(0,0),9),("RIGHTPADDING",(-1,0),(-1,0),9),
    ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6)]))
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
                          ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5)]))
E.append(info)

E.append(Paragraph(SCOPE_HEADING,st_redhd))
sb=Table([[Paragraph(SCOPE_TEXT,st_scope)]],colWidths=[CW])
sb.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#F4F6F9")),
    ("LINEBEFORE",(0,0),(0,-1),2.5,RED),("LEFTPADDING",(0,0),(-1,-1),9),
    ("RIGHTPADDING",(0,0),(-1,-1),9),("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5)]))
E.append(sb)

# ---- what's included ----
# Checkmarks are drawn as vector paths, NOT ZapfDingbats "3": that glyph falls back to a
# notdef box in viewers without the base-14 font, which is what a customer would then see.
class Check(Flowable):
    def __init__(self,size=5.6,color=RED,weight=1.15):
        Flowable.__init__(self); self.size=size; self.color=color; self.weight=weight
        self.width=size; self.height=size
    def draw(self):
        c=self.canv; s=self.size
        c.setStrokeColor(self.color); c.setLineWidth(self.weight)
        c.setLineCap(1); c.setLineJoin(1)
        p=c.beginPath()
        p.moveTo(0,s*0.50); p.lineTo(s*0.36,s*0.13); p.lineTo(s*1.02,s*0.88)
        c.drawPath(p)

E.append(Paragraph("WHAT'S INCLUDED",st_redhd))
st_chk=S(name="ck",fontName="Helvetica",fontSize=7.8,leading=10.4,textColor=INK)
CKW=0.17*inch; TXW=CW/2-CKW
inc=Table([[Check(),Paragraph(a,st_chk),Check(),Paragraph(b,st_chk)]
           for a,b in zip(INCLUDED_LEFT,INCLUDED_RIGHT)],
          colWidths=[CKW,TXW,CKW,TXW])
inc.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
                         ("LEFTPADDING",(0,0),(0,-1),2),("LEFTPADDING",(2,0),(2,-1),2),
                         ("TOPPADDING",(0,0),(-1,-1),1),("BOTTOMPADDING",(0,0),(-1,-1),1),
                         ("VALIGN",(0,0),(-1,-1),"TOP"),
                         ("TOPPADDING",(0,0),(0,-1),3),("TOPPADDING",(2,0),(2,-1),3)]))
E.append(inc)

# ---- pricing comparison ----
E.append(Paragraph("SCOPE OF WORK &amp; PRICING",st_redhd))
AW=1.10*inch; DW_=CW-2*AW
hw=lambda t,sz=7.6: Paragraph("<font color='#FFFFFF'><b>%s</b></font>"%t,
      S(name="h",fontName="Helvetica-Bold",fontSize=sz,leading=9.5,textColor=colors.white,alignment=TA_CENTER))
hl=Paragraph("<font color='#FFFFFF'><b>Description</b></font>",
      S(name="hl",fontName="Helvetica-Bold",fontSize=7.6,leading=9.5,textColor=colors.white))
sub=lambda t: Paragraph("<font color='#FFFFFF' size='5.8'>%s</font>"%t,
      S(name="hs",fontName="Helvetica",fontSize=5.8,leading=7,textColor=colors.white,alignment=TA_CENTER))

rows=[[hl,[hw(OPT_A_LABEL),sub(OPT_A_SUB)],[hw(OPT_B_LABEL),sub(OPT_B_SUB)]]]
grp_rows=[];desc_rows=[];i=1
for name,desc,a,b in GROUPS:
    rows.append([Paragraph(name,st_grpn),Paragraph(a,st_amt),Paragraph(b,st_amt)]); grp_rows.append(i); i+=1
    rows.append([Paragraph(desc,st_desc),"",""]); desc_rows.append(i); i+=1
rows.append([Paragraph("Line Item Subtotal",st_totk),Paragraph(SUB_A,st_totv),Paragraph(SUB_B,st_totv)]); sr=i; i+=1
rows.append([Paragraph("Est. Material Sales Tax (8.25%, materials only)",st_totk),
             Paragraph(TAX_A,st_totv),Paragraph(TAX_B,st_totv)]); tr=i; i+=1
ev=lambda t: Paragraph("<font color='#FFFFFF'><b>%s</b></font>"%t,
      S(name="ev",fontName="Helvetica-Bold",fontSize=10.5,leading=13,textColor=colors.white,alignment=TA_CENTER))
rows.append([Paragraph("<font color='#FFFFFF'><b>ESTIMATED TOTAL</b></font>",
    S(name="et",fontName="Helvetica-Bold",fontSize=10.5,leading=13,textColor=colors.white,alignment=TA_RIGHT)),
    ev(TOT_A),ev(TOT_B)]); totr=i

tbl=Table(rows,colWidths=[DW_,AW,AW])
sty=[("BACKGROUND",(0,0),(-1,0),NAVY),("LEFTPADDING",(0,0),(-1,-1),9),("RIGHTPADDING",(0,0),(-1,-1),9),
     ("TOPPADDING",(0,0),(-1,0),5),("BOTTOMPADDING",(0,0),(-1,0),5),("VALIGN",(0,0),(-1,-1),"MIDDLE")]
for gr in grp_rows:
    sty+=[("BACKGROUND",(0,gr),(-1,gr),BLUE1),("TOPPADDING",(0,gr),(-1,gr),4),("BOTTOMPADDING",(0,gr),(-1,gr),4)]
for dr in desc_rows:
    sty+=[("TOPPADDING",(0,dr),(-1,dr),4),("BOTTOMPADDING",(0,dr),(-1,dr),5),
          ("LINEBELOW",(0,dr),(-1,dr),0.4,LINE)]
sty+=[("TOPPADDING",(0,sr),(-1,tr),2),("BOTTOMPADDING",(0,sr),(-1,tr),2),
      ("BACKGROUND",(0,totr),(-1,totr),NAVY),("TOPPADDING",(0,totr),(-1,totr),6),
      ("BOTTOMPADDING",(0,totr),(-1,totr),6),
      ("LINEAFTER",(0,0),(0,totr-1),0.4,LINE),("LINEAFTER",(1,0),(1,totr-1),0.4,LINE)]
tbl.setStyle(TableStyle(sty))
E.append(tbl); E.append(Spacer(1,6))

E.append(Paragraph(NOTE,st_note)); E.append(Spacer(1,6))

hd=Paragraph("NEXT STEPS",st_redhd)
rule=Table([[""]],colWidths=[CW],rowHeights=[1.6])
rule.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),RED),("LEFTPADDING",(0,0),(-1,-1),0),
                          ("RIGHTPADDING",(0,0),(-1,-1),0),("TOPPADDING",(0,0),(-1,-1),0),
                          ("BOTTOMPADDING",(0,0),(-1,-1),0)]))
cells=[[Paragraph(t,st_stpH),Paragraph(b,st_stpB)] for t,b in STEPS]
cols=Table([[Table([[c[0]],[c[1]]],colWidths=[CW/3-0.18*inch]) for c in cells]],colWidths=[CW/3]*3)
cols.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),0),
                          ("RIGHTPADDING",(0,0),(-1,-1),9),("TOPPADDING",(0,0),(-1,-1),4)]))
E.append(KeepTogether([hd,rule,Spacer(1,4),cols]))

doc.build(E)
print("built",OUT)
