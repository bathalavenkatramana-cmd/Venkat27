"""Monica's exact incremental-revenue model for TIME PG/PD.
Hero table (section x slot, her columns) + 3 scenarios (5/10/15%) + 30/60/90 ramp.
Target (high-value) slots only: prog share >= 60% AND direct share <= 35%.
Period Jan-May 2026 (5 mo). Real per-slot programmatic CPM. Fixed PG/PD $8-$10.

Monica's formula:
  current revenue      = (current impressions / 1000) x current programmatic eCPM
  projected revenue    = ((remaining prog imp /1000) x current eCPM) + ((shifted imp /1000) x PG/PD eCPM)
  incremental revenue  = projected - current  =  (shifted imp / 1000) x (PG/PD eCPM - current prog eCPM)
"""
import os, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

HERE=os.path.dirname(os.path.abspath(__file__))
MONTHS=5; PA,PB=8.0,10.0

# DATA[section][slot] = (std_imp, prog_imp, house_imp, std_cpm, prog_cpm)
DATA = {
 "Entertainment": {"stickyfooter1":(13767927,47410845,2166739,21.22,1.92),"leaderboard1":(9918394,11875493,753998,22.99,1.88),
   "inline1":(5880365,11630907,396956,21.88,2.93),"inline2":(3876484,10920299,471422,26.00,2.64),
   "inline3":(2914244,10019326,511614,23.61,2.55),"rightrail1":(1998657,4846780,210342,19.14,2.88)},
 "Health": {"stickyfooter1":(11067220,20185172,960572,26.29,2.53),"leaderboard1":(8665005,7007562,756373,32.25,2.81),
   "inline1":(4831301,6850100,414509,27.80,5.15),"inline2":(3772312,5874817,368795,28.50,4.24),
   "inline3":(2530253,4930438,336651,29.55,4.11),"rightrail1":(3808220,4880198,408881,22.90,4.11)},
 "Science": {"stickyfooter1":(723615,2006324,98956,27.67,2.65),"leaderboard1":(1012980,1612613,69612,44.32,2.39),
   "inline1":(453383,1280246,40846,26.84,3.19),"inline2":(312130,1046394,43054,27.89,2.89),
   "inline3":(229424,781024,44550,27.19,2.77),"rightrail1":(292628,571444,48350,23.72,3.06)},
 "Tech": {"stickyfooter1":(1502361,2725359,145460,28.67,2.62),"leaderboard1":(2502857,1551790,105365,20.64,2.97),
   "inline1":(1245304,1137367,53511,23.65,5.53),"inline2":(777475,885597,54644,25.27,4.24),
   "inline3":(501022,673770,43319,25.62,3.95),"rightrail1":(866947,821619,46627,27.96,3.85)},
 "U.S.": {"stickyfooter1":(5470947,19228301,502892,22.92,2.13),"leaderboard1":(5864010,6609878,435768,25.97,2.13),
   "inline1":(2054682,4331009,223459,24.71,3.22),"inline2":(1201113,3319301,247599,27.76,2.85),
   "inline3":(798265,2359398,168507,28.08,2.66),"rightrail1":(1317962,2677384,181060,19.87,2.65)},
}
SLOTS=["stickyfooter1","leaderboard1","inline1","inline2","inline3","rightrail1"]
SECS=list(DATA.keys())

def is_target(std,prog,house):
    tot=std+prog+house; return (prog/tot>=0.60) and (std/tot<=0.35)

# target rows, grouped by section, monthly programmatic impressions
T=[]  # dicts
for s in SECS:
    for sl in SLOTS:
        std,prog,house,scpm,pcpm=DATA[s][sl]
        if is_target(std,prog,house):
            T.append(dict(sec=s,slot=sl,prog_m=prog/MONTHS,pcpm=pcpm,scpm=scpm))
# order sections as-is; within section by prog_m desc
T.sort(key=lambda r:(SECS.index(r["sec"]),-r["prog_m"]))

def inc(prog_m,pcpm,price,pct): return (prog_m*pct)/1000*(price-pcpm)
def pgrev(prog_m,price,pct):    return (prog_m*pct)/1000*price
def curr(prog_m,pcpm,pct):      return (prog_m*pct)/1000*pcpm

tot_prog_m=sum(r["prog_m"] for r in T)
blend_ecpm=sum(r["prog_m"]*r["pcpm"] for r in T)/tot_prog_m
def scen_tot(pct):
    return (sum(inc(r["prog_m"],r["pcpm"],PA,pct) for r in T),
            sum(inc(r["prog_m"],r["pcpm"],PB,pct) for r in T),
            sum(pgrev(r["prog_m"],PA,pct) for r in T),
            sum(pgrev(r["prog_m"],PB,pct) for r in T))
print(f"targets {len(T)} | prog/mo {tot_prog_m:,.0f} | blended eCPM ${blend_ecpm:.2f}")
for p in (0.05,0.10,0.15):
    i8,i10,r8,r10=scen_tot(p); print(f"  {int(p*100)}%: inc @8 ${i8:,.0f} @10 ${i10:,.0f} | pgrev @8 ${r8:,.0f} @10 ${r10:,.0f}")

def rng(a,b): return f"${a:,.0f}\u2013${b:,.0f}"

# ================= DOCX =================
TEALc=RGBColor(0x0F,0x7D,0x88); TEALD=RGBColor(0x0B,0x5C,0x65); INK=RGBColor(0x1F,0x2A,0x30)
MUT=RGBColor(0x5B,0x6B,0x73); WHT=RGBColor(0xFF,0xFF,0xFF)
doc=Document()
sec=doc.sections[0]; sec.left_margin=sec.right_margin=Inches(0.7); sec.top_margin=sec.bottom_margin=Inches(0.7)
nm=doc.styles["Normal"]; nm.font.name="Calibri"; nm.font.size=Pt(11); nm.font.color.rgb=INK
nm.paragraph_format.space_after=Pt(6); nm.paragraph_format.line_spacing=1.12
def shade(cell,hx):
    tcPr=cell._tc.get_or_add_tcPr(); sh=OxmlElement("w:shd"); sh.set(qn("w:val"),"clear"); sh.set(qn("w:fill"),hx); tcPr.append(sh)
def borders(t,color="C9D2D4",sz=6):
    b=OxmlElement("w:tblBorders")
    for e in ("top","left","bottom","right","insideH","insideV"):
        x=OxmlElement(f"w:{e}"); x.set(qn("w:val"),"single"); x.set(qn("w:sz"),str(sz)); x.set(qn("w:space"),"0"); x.set(qn("w:color"),color); b.append(x)
    t._tbl.tblPr.append(b)
def setc(c,t,bold=False,color=None,size=9.5,align="left",fill=None):
    c.text=""; p=c.paragraphs[0]
    p.alignment={"left":WD_ALIGN_PARAGRAPH.LEFT,"center":WD_ALIGN_PARAGRAPH.CENTER,"right":WD_ALIGN_PARAGRAPH.RIGHT}[align]
    rn=p.add_run(t); rn.bold=bold; rn.font.size=Pt(size)
    if color is not None: rn.font.color.rgb=color
    if fill: shade(c,fill)
    p.paragraph_format.space_after=Pt(1); p.paragraph_format.space_before=Pt(1)
def H(t,size=14,color=TEALD,sb=12):
    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(sb); p.paragraph_format.space_after=Pt(3)
    rn=p.add_run(t); rn.bold=True; rn.font.color.rgb=color; rn.font.size=Pt(size); return p
def body(t,size=10.5,italic=False,color=INK,after=5,bold=False):
    p=doc.add_paragraph(); rn=p.add_run(t); rn.italic=italic; rn.bold=bold; rn.font.size=Pt(size); rn.font.color.rgb=color
    p.paragraph_format.space_after=Pt(after); return p
def callout(bold,rest):
    t=doc.add_table(rows=1,cols=1); c=t.cell(0,0); shade(c,"E8F4F5"); c.text=""
    p=c.paragraphs[0]; p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(6)
    rb=p.add_run(bold+"  "); rb.bold=True; rb.font.color.rgb=TEALD; rb.font.size=Pt(11.5)
    rr=p.add_run(rest); rr.font.size=Pt(11); rr.font.color.rgb=INK; borders(t,"0F7D88",12)
    doc.add_paragraph().paragraph_format.space_after=Pt(2)

def mk_table(headers,rows,widths,rf,fs=9.3):
    t=doc.add_table(rows=1,cols=len(headers)); t.alignment=WD_TABLE_ALIGNMENT.CENTER; borders(t)
    for j,h in enumerate(headers): setc(t.rows[0].cells[j],h,bold=True,color=WHT,size=8.8,align="left" if j<rf else "right",fill="0B5C65")
    for i,row in enumerate(rows):
        tag=row[0] if isinstance(row,tuple) else None
        vals=row[1] if isinstance(row,tuple) else row
        cells=t.add_row().cells
        fill=("DCEAEC" if tag=="SUB" else ("0B5C65" if tag=="GRAND" else ("F7FAFA" if i%2 else None)))
        for j,v in enumerate(vals):
            setc(cells[j],v,bold=(tag in ("SUB","GRAND")),color=(WHT if tag=="GRAND" else None),
                 size=fs,align="left" if j<rf else "right",fill=fill)
    for j,w in enumerate(widths):
        for rr in t.rows: rr.cells[j].width=Inches(w)
    doc.add_paragraph().paragraph_format.space_after=Pt(2)

# ---- title ----
tp=doc.add_paragraph(); rn=tp.add_run("DATABEAT"); rn.bold=True; rn.font.size=Pt(16); rn.font.color.rgb=TEALD
r2=tp.add_run("   |   Prepared for TIME"); r2.font.size=Pt(11); r2.font.color.rgb=MUT
H("PG/PD Incremental Revenue Model",size=20,sb=6)
body("The actual monthly uplift from shifting programmatic inventory into Preferred Deals / Programmatic Guaranteed, "
     "by section and slot. US inventory \u00b7 high-value (target) slots \u00b7 Jan\u2013May 2026.",italic=True,size=11)

i8,i10,r8,r10=scen_tot(0.10)
callout("Bottom line:",
    f"converting 10% of programmatic impressions in these high-value section/slot pairs into PG/PD at $8\u2013$10 generates "
    f"{rng(i8,i10)} in incremental revenue per month (~{rng(i8*12,i10*12)}/yr) \u2014 on {tot_prog_m*0.10:,.0f} shifted "
    f"impressions/month that currently clear at just ${blend_ecpm:.2f}.")

# ---- HERO TABLE (Monica's exact columns) ----
H("Incremental Revenue by Section \u00d7 Slot  (Base case: 10% shift)",size=13)
rows=[]
for s in SECS:
    grp=[r for r in T if r["sec"]==s]
    if not grp: continue
    for r in grp:
        rows.append([r["sec"],r["slot"],f"{r['prog_m']:,.0f}",f"${r['pcpm']:.2f}","$8\u2013$10","10%",
                     rng(inc(r['prog_m'],r['pcpm'],PA,0.10),inc(r['prog_m'],r['pcpm'],PB,0.10))])
    gp=sum(r["prog_m"] for r in grp); ge=sum(r["prog_m"]*r["pcpm"] for r in grp)/gp
    rows.append(("SUB",[f"{s} subtotal","",f"{gp:,.0f}",f"${ge:.2f}","$8\u2013$10","10%",
                 rng(sum(inc(r['prog_m'],r['pcpm'],PA,0.10) for r in grp),sum(inc(r['prog_m'],r['pcpm'],PB,0.10) for r in grp))]))
rows.append(("GRAND",["ALL TARGETS","",f"{tot_prog_m:,.0f}",f"${blend_ecpm:.2f}","$8\u2013$10","10%",rng(i8,i10)]))
mk_table(["Section","Slot","Current prog.\nimpressions (mo)","Current\nprog. eCPM","Proposed\nPG/PD eCPM","%\nshifted","Incremental\nmonthly revenue"],
         rows,widths=[1.25,1.15,1.25,0.95,0.95,0.6,1.55],rf=2,fs=9.0)
body("Formula:  incremental revenue = (shifted impressions \u00f7 1000) \u00d7 (PG/PD eCPM \u2212 current programmatic eCPM). "
     "\u201cCurrent programmatic impressions\u201d are monthly (period \u00f7 5). Range reflects PG/PD at $8 (low) to $10 (high).",
     size=9,italic=True,color=MUT)

# ---- SCENARIOS ----
H("Three Scenarios",size=13)
srows=[]
for lbl,p in [("Low case",0.05),("Base case",0.10),("High case",0.15)]:
    a8,a10,_,_=scen_tot(p)
    srows.append([f"{lbl} \u2014 {int(p*100)}% shift",rng(a8,a10),rng(a8*12,a10*12)])
mk_table(["Scenario","Incremental monthly revenue","Annualized"],srows,widths=[2.6,2.2,2.2],rf=1,fs=10)

# ---- RAMP ----
H("30 / 60 / 90-Day Ramp",size=13)
body("Monica asked for a phased build, not instant impact \u2014 activation scales over roughly a quarter:")
ramp=[("30 days","Small pilot",0.05),("60 days","Broader rollout",0.10),("90 days","Scaled rollout",0.15)]
rrows=[]
for period,stage,p in ramp:
    i8,i10,rr8,rr10=scen_tot(p)
    rrows.append([period,stage,f"{int(p*100)}%",rng(rr8,rr10),rng(i8,i10)])
mk_table(["Period","Stage","% inventory activated","Estimated PG/PD revenue / mo","Incremental uplift / mo"],
         rrows,widths=[1.0,1.6,1.5,1.9,1.9],rf=2,fs=9.5)

callout("One-line answer for Monica:",
    f"\u201cConvert 10% of today\u2019s programmatic impressions in our high-value section/slot pairs into PG/PD at $8\u2013$10 and "
    f"we add {rng(i8_10:=scen_tot(0.10)[0],scen_tot(0.10)[1])} per month (~{rng(scen_tot(0.10)[0]*12,scen_tot(0.10)[1]*12)}/yr), "
    f"phased over 30/60/90 days \u2014 these impressions clear at ~${blend_ecpm:.2f} today.\u201d")

fp=doc.add_paragraph(); fp.paragraph_format.space_before=Pt(10)
rn=fp.add_run("Method: 18 high-value slots (programmatic share \u2265 60% and direct share \u2264 35%) across 5 sections, US, "
    "Jan\u2013May 2026 (5 months). Impressions and eCPMs from GAM; every figure reconciled to grand totals. Direct-heavy "
    "slots (leaderboard1 and most of Tech) are excluded and reserved for the sales team. World & Politics not in this "
    "dataset (additive). Prepared by Databeat for TIME.")
rn.italic=True; rn.font.size=Pt(8); rn.font.color.rgb=MUT
out_doc=os.path.join(HERE,"TIME_PG-PD_Incremental_Model_Monica.docx")
doc.save(out_doc); print("saved:",out_doc)

# ================= XLSX (live formulas, Monica's columns) =================
wb=openpyxl.Workbook(); ws=wb.active; ws.title="Incremental Model"
TEAL="0F7D88"; TEALDX="0B5C65"; GOLD="FCEFCB"; SUB="DCEAEC"
def sty(c,bold=False,color="1F2A30",size=10,fill=None,align="left",fmt=None,white=False):
    c.font=Font(bold=bold,color=("FFFFFF" if white else color),size=size,name="Calibri")
    c.alignment=Alignment(horizontal=align,vertical="center"); 
    if fill: c.fill=PatternFill("solid",fgColor=fill)
    if fmt: c.number_format=fmt
thin=Side(style="thin",color="C9D2D4")
def bd(rng):
    for row in ws[rng]:
        for c in row: c.border=Border(thin,thin,thin,thin)
r=1
ws.merge_cells(f"A{r}:H{r}"); sty(ws[f"A{r}"],bold=True,size=14,white=True,fill=TEALDX)
ws[f"A{r}"]="TIME | PG/PD Incremental Revenue Model \u2014 Section \u00d7 Slot (Monica format)"; ws.row_dimensions[r].height=22; r+=2
# assumptions
sty(ws[f"A{r}"],bold=True,color=TEALDX); ws[f"A{r}"]="Assumptions (editable)"; r+=1
asum=[("PG/PD eCPM \u2013 low ($)",PA,"$#,##0.00"),("PG/PD eCPM \u2013 high ($)",PB,"$#,##0.00"),("% shifted (base)",0.10,"0%")]
cells={}
for lbl,val,fmt in asum:
    sty(ws[f"A{r}"]); ws[f"A{r}"]=lbl; sty(ws[f"B{r}"],bold=True,fill=GOLD,align="center",fmt=fmt); ws[f"B{r}"]=val
    bd(f"A{r}:B{r}"); cells[lbl]=f"$B${r}"; r+=1
PAx=cells["PG/PD eCPM \u2013 low ($)"]; PBx=cells["PG/PD eCPM \u2013 high ($)"]; PCT=cells["% shifted (base)"]
r+=1
hdr=["Section","Slot","Current prog. impressions (mo)","Current prog. eCPM","Proposed PG/PD eCPM",
     "% shifted","Incremental monthly rev @ low","Incremental monthly rev @ high"]
hrow=r
for j,h in enumerate(hdr): sty(ws.cell(row=r,column=j+1,value=h),bold=True,white=True,fill=TEALDX,size=9,align="center")
ws.row_dimensions[r].height=30; r+=1
first=r
for rr in T:
    ws.cell(row=r,column=1,value=rr["sec"]); ws.cell(row=r,column=2,value=rr["slot"])
    ws.cell(row=r,column=3,value=round(rr["prog_m"])); ws.cell(row=r,column=4,value=rr["pcpm"])
    ws.cell(row=r,column=5,value=f'=$B${hrow-4}&" \u2013 "&$B${hrow-3}')  # display range text (informational)
    ws.cell(row=r,column=6,value=f"={PCT}")
    ws.cell(row=r,column=7,value=f"=C{r}*F{r}/1000*({PAx}-D{r})")
    ws.cell(row=r,column=8,value=f"=C{r}*F{r}/1000*({PBx}-D{r})")
    for col in range(1,9):
        fmt=None
        if col==3: fmt='#,##0'
        elif col==4: fmt='$#,##0.00'
        elif col==6: fmt='0%'
        elif col in(7,8): fmt='$#,##0'
        sty(ws.cell(row=r,column=col),size=9.5,align="left" if col<=2 else ("center" if col==5 else "right"),fmt=fmt)
    r+=1
last=r-1
sty(ws.cell(row=r,column=1,value="ALL TARGETS"),bold=True,fill=SUB)
for col in (2,5): sty(ws.cell(row=r,column=col,value=""),fill=SUB)
sty(ws.cell(row=r,column=3,value=f"=SUM(C{first}:C{last})"),bold=True,fill=SUB,align="right",fmt='#,##0')
sty(ws.cell(row=r,column=4,value=f"=SUMPRODUCT(C{first}:C{last},D{first}:D{last})/C{r}"),bold=True,fill=SUB,align="right",fmt='$#,##0.00')
sty(ws.cell(row=r,column=6,value=f"={PCT}"),bold=True,fill=SUB,align="right",fmt='0%')
sty(ws.cell(row=r,column=7,value=f"=SUM(G{first}:G{last})"),bold=True,fill=SUB,align="right",fmt='$#,##0')
sty(ws.cell(row=r,column=8,value=f"=SUM(H{first}:H{last})"),bold=True,fill=SUB,align="right",fmt='$#,##0')
bd(f"A{hrow}:H{r}"); grand=r; r+=2
# scenarios block
sty(ws[f"A{r}"],bold=True,color=TEALDX); ws[f"A{r}"]="Scenarios (recompute from the % cell above)"; r+=1
for j,h in enumerate(["Scenario","Shift %","Incremental / mo @ low","Incremental / mo @ high","Annualized @ low","Annualized @ high"]):
    sty(ws.cell(row=r,column=j+1,value=h),bold=True,white=True,fill=TEAL,size=9,align="center")
r+=1
mprog=f"C{grand}"; me=f"D{grand}"
for lbl,p in [("Low",0.05),("Base",0.10),("High",0.15)]:
    ws.cell(row=r,column=1,value=lbl); ws.cell(row=r,column=2,value=p)
    ws.cell(row=r,column=3,value=f"={mprog}*B{r}/1000*({PAx}-{me})")
    ws.cell(row=r,column=4,value=f"={mprog}*B{r}/1000*({PBx}-{me})")
    ws.cell(row=r,column=5,value=f"=C{r}*12"); ws.cell(row=r,column=6,value=f"=D{r}*12")
    for col in range(1,7): sty(ws.cell(row=r,column=col),align="center" if col<=2 else "right",fmt=('0%' if col==2 else ('$#,##0' if col>=3 else None)))
    r+=1
bd(f"A{r-3}:F{r-1}")
for i,w in enumerate([13,13,16,13,14,10,16,16]): ws.column_dimensions[get_column_letter(i+1)].width=w
ws.freeze_panes=f"A{hrow+1}"
out_x=os.path.join(HERE,"TIME_PG-PD_Incremental_Model_Monica.xlsx"); wb.save(out_x); print("saved:",out_x)
