"""Client-ready SECTION x SLOT PG/PD conversion-TARGET document for TIME.
Period Jan-May 2026 (5 mo). Real per-slot CPMs. TARGET = prog>=60% & direct<=35%;
RESERVED = direct>=40% (leave for direct sales); else SECONDARY.
Model: shift 5/10/15% of a slot's programmatic impressions to PG/PD at fixed $8/$10,
baselined on that slot's ACTUAL programmatic CPM.
"""
import os, matplotlib
matplotlib.use("Agg"); import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

HERE=os.path.dirname(os.path.abspath(__file__))
MONTHS=5; PA,PB=8.0,10.0; BASE=0.10

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

def cls(d,p):
    if p>=0.60 and d<=0.35: return "Target"
    if d>=0.40: return "Reserved"
    return "Secondary"

R=[]
for s in SECS:
    for sl in SLOTS:
        std,prog,house,scpm,pcpm=DATA[s][sl]; tot=std+prog+house
        R.append(dict(sec=s,slot=sl,std=std,prog=prog,house=house,tot=tot,scpm=scpm,pcpm=pcpm,
                      dsh=std/tot,psh=prog/tot,cls=cls(std/tot,prog/tot)))
def up_m(r,price,pct=BASE): return (r["prog"]/MONTHS)*pct/1000*(price-r["pcpm"])
def cur_m(r,pct=BASE):      return (r["prog"]/MONTHS)*pct*r["pcpm"]/1000
def pg_m(r,price,pct=BASE): return (r["prog"]/MONTHS)*pct*price/1000

TGT=[r for r in R if r["cls"]=="Target"]
RES=[r for r in R if r["cls"]=="Reserved"]
gprog=sum(r["prog"] for r in R)
tprog=sum(r["prog"] for r in TGT); tprog_m=tprog/MONTHS
tcur_m=sum(r["prog"]/MONTHS*r["pcpm"]/1000 for r in TGT)
def scen(pct):
    u8=sum(up_m(r,PA,pct) for r in TGT); u10=sum(up_m(r,PB,pct) for r in TGT)
    cur=sum(cur_m(r,pct) for r in TGT); r8=sum(pg_m(r,PA,pct) for r in TGT); r10=sum(pg_m(r,PB,pct) for r in TGT)
    return cur,r8,r10,u8,u10
cur10,r8_10,r10_10,u8_10,u10_10=scen(BASE)
print(f"targets {len(TGT)} prog {tprog:,} ({tprog/gprog*100:.1f}%) cur/mo ${tcur_m:,.0f} | 10%: +${u8_10:,.0f}/+${u10_10:,.0f}/mo")

# =============== CHARTS ===============
TEAL="#0f7d88"; GOLD="#e0a83c"; GREY="#b9c2c5"; RED="#c65b4e"; INK="#1f2a30"; TEALD="#0b5c65"
plt.rcParams.update({"font.family":"DejaVu Sans","font.size":11,"axes.grid":True,
                     "grid.color":"#e6eef0","axes.axisbelow":True,"axes.edgecolor":"#c9d2d4"})
def save(fig,n): fig.savefig(os.path.join(HERE,n),dpi=165,bbox_inches="tight",facecolor="white"); plt.close(fig)

# Chart 1: top target slots by monthly prog impressions
top=sorted(TGT,key=lambda r:-r["prog"])[:12]
labels=[f"{r['sec'][:4]}·{r['slot']}" for r in top]; vals=[r["prog"]/MONTHS/1e6 for r in top]
fig,ax=plt.subplots(figsize=(8.4,4.3))
ax.barh(range(len(top)),vals,color=TEAL)
ax.set_yticks(range(len(top))); ax.set_yticklabels(labels,fontsize=9); ax.invert_yaxis()
for i,r in enumerate(top):
    ax.text(vals[i]+0.05,i,f"{vals[i]:.1f}M  ({r['psh']*100:.0f}% prog / {r['dsh']*100:.0f}% dir)  ${r['pcpm']:.2f}",
            va="center",fontsize=8,color=INK)
ax.set_xlabel("Monthly programmatic impressions (millions)"); ax.set_xlim(0,max(vals)*1.5)
ax.set_title("Top PG/PD conversion targets — high programmatic, low direct",
             fontsize=12,fontweight="bold",color=TEALD,pad=8); ax.grid(axis="y",visible=False)
save(fig,"t1_targets.png")

# Chart 2: incremental monthly revenue by top target slot @ $8/$10 (10%)
top8=sorted(TGT,key=lambda r:-up_m(r,PA))[:10]
lab=[f"{r['sec'][:4]}·{r['slot']}" for r in top8]
i8=[up_m(r,PA) for r in top8]; i10=[up_m(r,PB) for r in top8]
x=range(len(top8)); w=0.4
fig,ax=plt.subplots(figsize=(8.6,4.0))
ax.bar([i-w/2 for i in x],i8,w,label="Incremental @ $8",color=TEAL)
ax.bar([i+w/2 for i in x],i10,w,label="Incremental @ $10",color=GOLD)
ax.set_xticks(list(x)); ax.set_xticklabels(lab,rotation=30,ha="right",fontsize=8)
ax.set_ylabel("Incremental revenue / month"); ax.yaxis.set_major_formatter(FuncFormatter(lambda y,_:f"${y:,.0f}"))
ax.set_title("Incremental monthly revenue by target slot — 10% shifted to PG/PD",
             fontsize=12,fontweight="bold",color=TEALD,pad=8)
ax.legend(frameon=False,fontsize=9.5); ax.grid(axis="x",visible=False)
save(fig,"t2_uplift.png")

# Chart 3: section split target vs reserved vs secondary (prog imp share)
fig,ax=plt.subplots(figsize=(8.4,3.5))
cats=["Target","Secondary","Reserved"]; colors={"Target":TEAL,"Secondary":GOLD,"Reserved":RED}
bottom=[0]*len(SECS)
for cat in cats:
    vals=[sum(r["prog"] for r in R if r["sec"]==s and r["cls"]==cat)/1e6 for s in SECS]
    ax.bar(SECS,vals,bottom=bottom,label=cat,color=colors[cat])
    bottom=[b+v for b,v in zip(bottom,vals)]
ax.set_ylabel("Programmatic impressions (millions, period)")
ax.set_title("Programmatic inventory by section — convertible (target) vs reserved-for-direct",
             fontsize=11.5,fontweight="bold",color=TEALD,pad=8)
ax.legend(frameon=False,fontsize=9.5,ncol=3,loc="upper right"); ax.grid(axis="x",visible=False)
save(fig,"t3_sections.png")
print("charts saved")

# =============== DOCX ===============
TEALc=RGBColor(0x0F,0x7D,0x88); TEALD_c=RGBColor(0x0B,0x5C,0x65); INKc=RGBColor(0x1F,0x2A,0x30)
MUT=RGBColor(0x5B,0x6B,0x73); WHT=RGBColor(0xFF,0xFF,0xFF); REDc=RGBColor(0xB0,0x43,0x38); GOLDd=RGBColor(0x8A,0x6D,0x1F)
doc=Document()
sec=doc.sections[0]; sec.orientation=WD_ORIENT.LANDSCAPE
sec.page_width,sec.page_height=Inches(11),Inches(8.5)
sec.left_margin=sec.right_margin=Inches(0.7); sec.top_margin=sec.bottom_margin=Inches(0.6)
nm=doc.styles["Normal"]; nm.font.name="Calibri"; nm.font.size=Pt(11); nm.font.color.rgb=INKc
nm.paragraph_format.space_after=Pt(6); nm.paragraph_format.line_spacing=1.1
def shade(cell,hx):
    tcPr=cell._tc.get_or_add_tcPr(); sh=OxmlElement("w:shd"); sh.set(qn("w:val"),"clear")
    sh.set(qn("w:fill"),hx); tcPr.append(sh)
def borders(table,color="C9D2D4",sz=6):
    b=OxmlElement("w:tblBorders")
    for e in ("top","left","bottom","right","insideH","insideV"):
        x=OxmlElement(f"w:{e}"); x.set(qn("w:val"),"single"); x.set(qn("w:sz"),str(sz))
        x.set(qn("w:space"),"0"); x.set(qn("w:color"),color); b.append(x)
    table._tbl.tblPr.append(b)
def setc(cell,t,bold=False,color=None,size=9,align="left",fill=None):
    cell.text=""; p=cell.paragraphs[0]
    p.alignment={"left":WD_ALIGN_PARAGRAPH.LEFT,"center":WD_ALIGN_PARAGRAPH.CENTER,"right":WD_ALIGN_PARAGRAPH.RIGHT}[align]
    rn=p.add_run(t); rn.bold=bold; rn.font.size=Pt(size)
    if color is not None: rn.font.color.rgb=color
    if fill: shade(cell,fill)
    p.paragraph_format.space_after=Pt(1); p.paragraph_format.space_before=Pt(1)
def H(t,size=15,color=TEALD_c,sb=12):
    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(sb); p.paragraph_format.space_after=Pt(3)
    rn=p.add_run(t); rn.bold=True; rn.font.color.rgb=color; rn.font.size=Pt(size); return p
def body(t,size=10.5,italic=False,color=INKc,after=5,bold=False):
    p=doc.add_paragraph(); rn=p.add_run(t); rn.italic=italic; rn.bold=bold; rn.font.size=Pt(size); rn.font.color.rgb=color
    p.paragraph_format.space_after=Pt(after); return p
def bullet(t,lead=None):
    p=doc.add_paragraph(style="List Bullet")
    if lead: rn=p.add_run(lead+" "); rn.bold=True; rn.font.color.rgb=TEALD_c; rn.font.size=Pt(10.5)
    r2=p.add_run(t); r2.font.size=Pt(10.5); r2.font.color.rgb=INKc; p.paragraph_format.space_after=Pt(2); return p
def callout(bold,rest,fill="E8F4F5",bar="0F7D88"):
    t=doc.add_table(rows=1,cols=1); c=t.cell(0,0); shade(c,fill); c.text=""
    p=c.paragraphs[0]; p.paragraph_format.space_before=Pt(5); p.paragraph_format.space_after=Pt(5)
    if bold: rb=p.add_run(bold+"  "); rb.bold=True; rb.font.color.rgb=TEALD_c; rb.font.size=Pt(11)
    rr=p.add_run(rest); rr.font.size=Pt(10.5); rr.font.color.rgb=INKc; borders(t,bar,12)
    doc.add_paragraph().paragraph_format.space_after=Pt(1); return t
def img(n,w=9.3,cap=None):
    doc.add_picture(os.path.join(HERE,n),width=Inches(w)); doc.paragraphs[-1].alignment=WD_ALIGN_PARAGRAPH.CENTER
    if cap:
        p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        rn=p.add_run(cap); rn.italic=True; rn.font.size=Pt(8.5); rn.font.color.rgb=MUT; p.paragraph_format.space_after=Pt(6)
CLSFILL={"Target":None,"Reserved":"FBE7E4","Secondary":"FCF4DE"}
CLSCOLOR={"Target":TEALD_c,"Reserved":REDc,"Secondary":GOLDd}
def table(headers,data,widths=None,rf=2,fs=8.6,classcol=None):
    t=doc.add_table(rows=1,cols=len(headers)); t.alignment=WD_TABLE_ALIGNMENT.CENTER; borders(t)
    for j,h in enumerate(headers): setc(t.rows[0].cells[j],h,bold=True,color=WHT,size=8.2,align="left" if j<rf else "right",fill="0B5C65")
    for i,row in enumerate(data):
        tag=row[0] if isinstance(row,tuple) else None
        vals=row[1] if isinstance(row,tuple) else row
        cells=t.add_row().cells
        if tag in ("SUB","GRAND"): fillbase="DCEAEC" if tag=="SUB" else "0B5C65"
        else: fillbase=None
        for j,v in enumerate(vals):
            fill=fillbase; color=None; bold=(tag in ("SUB","GRAND"))
            if tag=="GRAND": color=WHT
            if fill is None:
                fill="F7FAFA" if i%2 else None
            if classcol is not None and j==classcol and tag not in ("SUB","GRAND"):
                fill=CLSFILL.get(v) or fill; color=CLSCOLOR.get(v); bold=True
            setc(cells[j],v,bold=bold,size=fs,align="left" if j<rf else "right",fill=fill,color=color)
    if widths:
        for j,w in enumerate(widths):
            for rr in t.rows: rr.cells[j].width=Inches(w)
    doc.add_paragraph().paragraph_format.space_after=Pt(1); return t
def m(v): return f"${v:,.0f}"

# ---- title ----
tp=doc.add_paragraph(); tp.paragraph_format.space_before=Pt(2)
rn=tp.add_run("DATABEAT"); rn.bold=True; rn.font.size=Pt(17); rn.font.color.rgb=TEALD_c
r2=tp.add_run("   |   Prepared for TIME"); r2.font.size=Pt(12); r2.font.color.rgb=MUT
H("PG/PD Conversion Targets — Section × Slot Analysis",size=21,sb=6)
p=doc.add_paragraph(); rn=p.add_run("A granular, ad-unit-level map of where TIME's premium inventory is being sold at "
    "open-auction programmatic prices instead of direct — and how much incremental revenue converting that slice to "
    "Preferred Deals / Programmatic Guaranteed would generate. US inventory · top-6 slots · Jan–May 2026.")
rn.italic=True; rn.font.size=Pt(11); rn.font.color.rgb=INKc

callout("Bottom line:",
    f"{tprog/gprog*100:.0f}% of programmatic impressions in these top slots ({tprog/1e6:.0f}M) sit in "
    f"“target” pools — heavily programmatic (≥60%), barely sold direct (≤35%) — clearing at just ${tcur_m*MONTHS/tprog*1000:.2f} "
    f"eCPM while direct on the same units fetches $19–$44. Converting 10% of that to PG/PD at $8–$10 adds "
    f"{m(u8_10)}–{m(u10_10)}/month (~{m(u8_10*12)}–{m(u10_10*12)}/yr) — without touching the "
    f"{sum(r['prog'] for r in RES)/gprog*100:.0f}% of inventory that is direct-heavy and reserved for the sales team.")

# ---- scope ----
H("Scope & Method",size=12)
bullet("US inventory, six top slots × five sections (Entertainment, U.S., Health, Tech, Science), Jan–May 2026 (5 months).",lead="What —")
bullet("GAM pivot with impressions AND CPM per section × slot for Standard (direct), Programmatic and House; reconciled to grand totals.",lead="Data —")
bullet("Programmatic share ≥ 60% AND direct share ≤ 35%. These are re-sellable via PG/PD with minimal direct cannibalization.",lead="Target rule —")
bullet("Direct share ≥ 40% — the sales team already owns these; excluded from conversion.",lead="Reserved rule —")
bullet("shift 5% / 10% / 15% of a slot's programmatic impressions to PG/PD at a fixed $8 and $10 eCPM; uplift is measured against that slot's actual programmatic CPM. Monthly = period ÷ 5.",lead="Model —")

# ---- section 1: TARGETS ----
H("1.  Conversion Targets — Ranked",size=12)
body(f"The {len(TGT)} slots that meet the target rule, ranked by programmatic volume. Uplift is monthly at the base 10% shift.")
rows=[]
for r in sorted(TGT,key=lambda x:-x["prog"]):
    rows.append([f"{r['sec']} · {r['slot']}",f"{r['prog']:,.0f}",f"{r['psh']*100:.0f}%",f"{r['dsh']*100:.0f}%",
                 f"${r['pcpm']:.2f}",f"${r['scpm']:.2f}",m(up_m(r,PA)),m(up_m(r,PB))])
rows.append(("GRAND",["18 target slots",f"{tprog:,.0f}",f"{tprog/sum(r['tot'] for r in TGT)*100:.0f}%","—",
             f"${tcur_m*MONTHS/tprog*1000:.2f}","—",m(u8_10),m(u10_10)]))
table(["Section · Slot","Prog. imp","Prog %","Dir %","Prog CPM","Direct CPM","Uplift/mo @ $8","Uplift/mo @ $10"],rows,
      widths=[2.6,1.2,0.7,0.7,0.9,0.95,1.15,1.2],rf=1,fs=8.8)
img("t1_targets.png",w=9.2,cap="Figure 1 — Top targets by monthly programmatic impressions (prog% / dir% and current prog CPM labeled).")

# ---- section 2: RESERVED ----
H("2.  Reserved for Direct Sales",size=12)
body("These slots are direct-dominant (direct share ≥ 40%) — the sales team already monetizes them well. We explicitly "
     "leave them out of the PG/PD push to avoid competing with direct.")
rows=[]
for r in sorted(RES,key=lambda x:-x["dsh"]):
    rows.append([f"{r['sec']} · {r['slot']}",f"{r['std']:,.0f}",f"{r['dsh']*100:.0f}%",f"{r['psh']*100:.0f}%",
                 f"${r['scpm']:.2f}",f"${r['pcpm']:.2f}"])
table(["Section · Slot","Direct imp","Dir %","Prog %","Direct CPM","Prog CPM"],rows,
      widths=[2.6,1.3,0.8,0.8,1.1,1.1],rf=1,fs=8.8)
body("Pattern: leaderboard1 across sections and most of Tech are direct territory — consistent with leaderboard being the "
     "premium billboard the sales team leads with.",italic=True,size=9.5,color=MUT)

# ---- section 3: per-section granular ----
H("3.  Section-by-Section Detail — Every Slot",size=12)
body("The full granular view: for each section, every slot's impressions, the CPMs it is selling at (direct vs programmatic), "
     "its programmatic/direct share, its classification, and the monthly PG/PD opportunity at 10%.")
for s in SECS:
    H(s,size=11,color=TEALc,sb=8)
    rows=[]; sub_prog=sub_std=0; sub_u8=sub_u10=0
    for sl in SLOTS:
        r=next(x for x in R if x["sec"]==s and x["slot"]==sl)
        u8=up_m(r,PA) if r["cls"]=="Target" else 0; u10=up_m(r,PB) if r["cls"]=="Target" else 0
        rows.append([sl,f"{r['std']:,.0f}",f"{r['prog']:,.0f}",f"{r['psh']*100:.0f}%",f"{r['dsh']*100:.0f}%",
                     f"${r['scpm']:.2f}",f"${r['pcpm']:.2f}",r["cls"],
                     (m(u8) if u8 else "—"),(m(u10) if u10 else "—")])
        sub_prog+=r["prog"]; sub_std+=r["std"]; sub_u8+=u8; sub_u10+=u10
    rows.append(("SUB",[f"{s} total",f"{sub_std:,.0f}",f"{sub_prog:,.0f}","—","—","—","—","",
                        (m(sub_u8) if sub_u8 else "—"),(m(sub_u10) if sub_u10 else "—")]))
    table(["Slot","Direct imp","Prog. imp","Prog %","Dir %","Dir CPM","Prog CPM","Class","Opp/mo @$8","Opp/mo @$10"],rows,
          widths=[1.25,1.05,1.05,0.6,0.6,0.75,0.8,0.95,0.95,0.98],rf=1,fs=8.3,classcol=7)

# ---- section 4: model ----
H("4.  Conversion Model — Targets Only",size=12)
body(f"Applying the shift to the {len(TGT)} target slots (30.1M programmatic impressions/month, ${tcur_m:,.0f}/mo today):")
sc=[]
for pct in (0.05,0.10,0.15):
    cur,r8,r10,u8,u10=scen(pct)
    sc.append([f"{int(pct*100)}%",m(u8),m(u10),m(u8*12),m(u10*12)])
table(["Shift level","Incremental / mo @ $8","Incremental / mo @ $10","Annualized @ $8","Annualized @ $10"],
      [tuple(["GRAND",row]) if False else row for row in sc],
      widths=[1.6,2.3,2.3,2.0,2.0],rf=1,fs=10)
img("t3_sections.png",w=8.8,cap="Figure 2 — Programmatic inventory by section: convertible target pool (teal) vs reserved-for-direct (red).")
img("t2_uplift.png",w=9.2,cap="Figure 3 — Incremental monthly revenue by target slot at $8 / $10 (10% shift).")

# ---- section 5 ----
H("5.  What This Means",size=12)
bullet(f"three-quarters of programmatic in these slots is convertible — high prog, low direct — so the opportunity is broad, not niche.",lead="Scale —")
bullet(f"stickyfooter1 (Entertainment, U.S., Health) alone drives ~{ (up_m(next(r for r in TGT if r['sec']=='Entertainment' and r['slot']=='stickyfooter1'),PA)+up_m(next(r for r in TGT if r['sec']=='U.S.' and r['slot']=='stickyfooter1'),PA)+up_m(next(r for r in TGT if r['sec']=='Health' and r['slot']=='stickyfooter1'),PA))/u8_10*100:.0f}% of the 10% uplift — start there.",lead="Focus —")
bullet("target pools clear at ~$2.44 vs direct at $19–$44; PG/PD at $8–$10 sits comfortably between — capturing value without undercutting direct.",lead="Headroom —")
bullet("leaderboard1 and most of Tech stay with direct sales; no channel conflict.",lead="No conflict —")
callout("One-line answer:",
    f"“Three-quarters of our programmatic in the top slots — {tprog/1e6:.0f}M impressions clearing at ~${tcur_m*MONTHS/tprog*1000:.2f} — "
    f"is barely sold direct. Converting 10% to PG/PD at $8–$10 adds {m(u8_10)}–{m(u10_10)} a month "
    f"(~{m(u8_10*12)}–{m(u10_10*12)}/yr), led by stickyfooter1, while leaderboard1 stays with the direct team.”")

fp=doc.add_paragraph(); fp.paragraph_format.space_before=Pt(10)
rn=fp.add_run("Scope: 5 sections × 6 top slots, US, Jan–May 2026 (5 months). Impressions and CPMs from GAM pivot; every "
    "figure reconciled to grand totals (200.05M programmatic impressions; $2.63 blended prog CPM). Target = prog ≥ 60% & "
    "direct ≤ 35%; reserved = direct ≥ 40%. World & Politics not in this dataset (additive). Prepared by Databeat for TIME.")
rn.italic=True; rn.font.size=Pt(8); rn.font.color.rgb=MUT

out=os.path.join(HERE,"TIME_PG-PD_Conversion_Targets_by_Slot.docx")
doc.save(out); print("saved:",out)
