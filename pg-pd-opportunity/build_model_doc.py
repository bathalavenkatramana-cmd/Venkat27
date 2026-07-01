"""FINAL client-ready PG/PD Incremental Revenue Model doc for TIME.
Fixed $8/$10 PG/PD CPM. Sections: Entertainment, US, Health, Science & Tech.
Data: 2026 Jan-May (5 months), US, TOP-6 slots (GAM). Monthly = period / 5.
Answers Monica: avg monthly programmatic impressions -> current revenue at current prog eCPM
vs revenue if 10% is converted to PG/PD at $8 and $10, plus 5/10/15% scenarios and 30/60/90 ramp.
"""
import os, matplotlib
matplotlib.use("Agg"); import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

HERE=os.path.dirname(os.path.abspath(__file__)); MONTHS=5

# section -> (prog period imp, prog period rev, direct eCPM benchmark)
SEC=[("Entertainment",22835006,53795,16.8),
     ("U.S.",15568972,40869,19.2),
     ("Health",10085819,40761,23.8),
     ("Science & Tech",1875841+1462654,6356+5119,30.0)]  # 30.0 = impression-weighted direct eCPM of science($26.5)+tech($32.7)

def ecpm(rev,imp): return rev/imp*1000
rows=[]
for name,imp,rev,std in SEC:
    m_imp=imp/MONTHS; c=ecpm(rev,imp); m_rev=rev/MONTHS
    rows.append(dict(name=name,imp=imp,rev=rev,std=std,m_imp=m_imp,c=c,m_rev=m_rev))

tot_imp=sum(r["imp"] for r in rows); tot_rev=sum(r["rev"] for r in rows)
blend_c=ecpm(tot_rev,tot_imp); tot_m_imp=tot_imp/MONTHS; tot_m_rev=tot_rev/MONTHS

# ---- 10% conversion (base) ----
def at_shift(pct):
    cur=sum(r["m_imp"]*pct*r["c"]/1000 for r in rows)
    r8 =sum(r["m_imp"]*pct*8/1000 for r in rows)
    r10=sum(r["m_imp"]*pct*10/1000 for r in rows)
    return cur,r8,r10
cur10,r8_10,r10_10=at_shift(0.10)

# cross-check asserts
assert abs(tot_m_imp-10365658)<2, tot_m_imp
assert abs(blend_c-2.8344)<0.01, blend_c
assert abs((r8_10-cur10)-5355)<5, (r8_10-cur10)      # incremental @$8 10%
assert abs((r10_10-cur10)-7428)<5, (r10_10-cur10)    # incremental @$10 10%
print("CROSS-CHECK PASSED")
print(f"Total avg monthly programmatic impressions (4 sections, top-6 slots): {tot_m_imp:,.0f}")
print(f"Blended current programmatic eCPM: ${blend_c:.2f}")
print(f"10% shift: current rev ${cur10:,.0f}/mo | @$8 ${r8_10:,.0f} (+${r8_10-cur10:,.0f}) | @$10 ${r10_10:,.0f} (+${r10_10-cur10:,.0f})")
for pct in (0.05,0.10,0.15):
    cur,r8,r10=at_shift(pct)
    print(f"  {int(pct*100)}% -> inc @$8 ${r8-cur:,.0f}/mo (${(r8-cur)*12:,.0f}/yr) | inc @$10 ${r10-cur:,.0f}/mo (${(r10-cur)*12:,.0f}/yr)")

# =============== CHARTS ===============
TEAL="#0f7d88"; GOLD="#e0a83c"; GREY="#b9c2c5"; INK="#1f2a30"; TEALD="#0b5c65"
plt.rcParams.update({"font.family":"DejaVu Sans","font.size":11,"axes.grid":True,
                     "grid.color":"#e6eef0","axes.axisbelow":True,"axes.edgecolor":"#c9d2d4"})
def save(fig,n): fig.savefig(os.path.join(HERE,n),dpi=180,bbox_inches="tight",facecolor="white"); plt.close(fig)

# Chart A: per-section monthly revenue on the 10% shifted imp: current vs $8 vs $10
names=[r["name"] for r in rows]
cur=[r["m_imp"]*0.10*r["c"]/1000 for r in rows]
v8=[r["m_imp"]*0.10*8/1000 for r in rows]
v10=[r["m_imp"]*0.10*10/1000 for r in rows]
x=range(len(names)); w=0.26
fig,ax=plt.subplots(figsize=(8.4,3.9))
ax.bar([i-w for i in x],cur,w,label="Current (at prog. eCPM)",color=GREY)
ax.bar(list(x),v8,w,label="At $8 PG/PD",color=TEAL)
ax.bar([i+w for i in x],v10,w,label="At $10 PG/PD",color=GOLD)
for i,(a,b,c) in enumerate(zip(cur,v8,v10)):
    for off,val in ((-w,a),(0,b),(w,c)):
        ax.text(i+off,val+40,f"${val:,.0f}",ha="center",fontsize=8.5,fontweight="bold",color=INK)
ax.set_xticks(list(x)); ax.set_xticklabels(names,fontsize=10)
ax.set_ylabel("Monthly revenue on the 10% shifted"); ax.set_ylim(0,5200)
ax.yaxis.set_major_formatter(FuncFormatter(lambda y,_:f"${y:,.0f}"))
ax.set_title(r"Revenue on 10% of programmatic impressions: today vs PG/PD at \$8 / \$10",
             fontsize=11.5,fontweight="bold",color=TEALD,pad=8)
ax.legend(frameon=False,fontsize=9.5,ncol=3,loc="upper right"); ax.grid(axis="x",visible=False)
save(fig,"m1_conversion.png")

# Chart B: total incremental/mo by scenario
scen=["5% shift","10% shift","15% shift"]
inc8=[at_shift(p)[1]-at_shift(p)[0] for p in (0.05,0.10,0.15)]
inc10=[at_shift(p)[2]-at_shift(p)[0] for p in (0.05,0.10,0.15)]
x=range(len(scen)); w=0.36
fig,ax=plt.subplots(figsize=(8.0,3.7))
b1=ax.bar([i-w/2 for i in x],inc8,w,label="Incremental @ $8",color=TEAL)
b2=ax.bar([i+w/2 for i in x],inc10,w,label="Incremental @ $10",color=GOLD)
for bars in (b1,b2):
    for b in bars: ax.text(b.get_x()+b.get_width()/2,b.get_height()+120,f"+${b.get_height():,.0f}",ha="center",fontsize=9.5,fontweight="bold",color=INK)
ax.set_xticks(list(x)); ax.set_xticklabels(scen,fontsize=10.5)
ax.set_ylabel("Incremental revenue / month"); ax.set_ylim(0,12500)
ax.yaxis.set_major_formatter(FuncFormatter(lambda y,_:f"${y:,.0f}"))
ax.set_title("Total incremental monthly revenue by scenario (4 sections, top-6 slots)",
             fontsize=11.5,fontweight="bold",color=TEALD,pad=8)
ax.legend(frameon=False,fontsize=10,loc="upper left"); ax.grid(axis="x",visible=False)
save(fig,"m2_scenarios.png")
print("charts saved")

# =============== DOCX ===============
TEALc=RGBColor(0x0F,0x7D,0x88); TEALD_c=RGBColor(0x0B,0x5C,0x65); INKc=RGBColor(0x1F,0x2A,0x30)
MUT=RGBColor(0x5B,0x6B,0x73); WHT=RGBColor(0xFF,0xFF,0xFF)
doc=Document(); nm=doc.styles["Normal"]; nm.font.name="Calibri"; nm.font.size=Pt(11); nm.font.color.rgb=INKc
nm.paragraph_format.space_after=Pt(6); nm.paragraph_format.line_spacing=1.12

def shade(cell,hx):
    tcPr=cell._tc.get_or_add_tcPr(); sh=OxmlElement("w:shd")
    sh.set(qn("w:val"),"clear"); sh.set(qn("w:fill"),hx); tcPr.append(sh)
def borders(table,color="C9D2D4",sz=6):
    tblPr=table._tbl.tblPr; b=OxmlElement("w:tblBorders")
    for e in ("top","left","bottom","right","insideH","insideV"):
        x=OxmlElement(f"w:{e}"); x.set(qn("w:val"),"single"); x.set(qn("w:sz"),str(sz))
        x.set(qn("w:space"),"0"); x.set(qn("w:color"),color); b.append(x)
    tblPr.append(b)
def setc(cell,t,bold=False,color=None,size=10,align="left",fill=None):
    cell.text=""; p=cell.paragraphs[0]
    p.alignment={"left":WD_ALIGN_PARAGRAPH.LEFT,"center":WD_ALIGN_PARAGRAPH.CENTER,"right":WD_ALIGN_PARAGRAPH.RIGHT}[align]
    r=p.add_run(t); r.bold=bold; r.font.size=Pt(size)
    if color is not None: r.font.color.rgb=color
    if fill: shade(cell,fill)
    p.paragraph_format.space_after=Pt(2); p.paragraph_format.space_before=Pt(2)
def H(t,size=16,color=TEALD_c,sb=14):
    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(sb); p.paragraph_format.space_after=Pt(4)
    r=p.add_run(t); r.bold=True; r.font.color.rgb=color; r.font.size=Pt(size); return p
def body(t,size=11,italic=False,color=INKc,after=6,bold=False):
    p=doc.add_paragraph(); r=p.add_run(t); r.italic=italic; r.bold=bold; r.font.size=Pt(size); r.font.color.rgb=color
    p.paragraph_format.space_after=Pt(after); return p
def bullet(t,lead=None):
    p=doc.add_paragraph(style="List Bullet")
    if lead: r=p.add_run(lead+" "); r.bold=True; r.font.color.rgb=TEALD_c; r.font.size=Pt(11)
    r2=p.add_run(t); r2.font.size=Pt(11); r2.font.color.rgb=INKc; p.paragraph_format.space_after=Pt(3); return p
def callout(bold,rest):
    t=doc.add_table(rows=1,cols=1); c=t.cell(0,0); shade(c,"E8F4F5"); c.text=""
    p=c.paragraphs[0]; p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(6)
    if bold: rb=p.add_run(bold+"  "); rb.bold=True; rb.font.color.rgb=TEALD_c; rb.font.size=Pt(11.5)
    rr=p.add_run(rest); rr.font.size=Pt(11); rr.font.color.rgb=INKc; borders(t,"0F7D88",12)
    doc.add_paragraph().paragraph_format.space_after=Pt(2); return t
def img(n,w=6.5,cap=None):
    doc.add_picture(os.path.join(HERE,n),width=Inches(w)); doc.paragraphs[-1].alignment=WD_ALIGN_PARAGRAPH.CENTER
    if cap:
        p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        r=p.add_run(cap); r.italic=True; r.font.size=Pt(9); r.font.color.rgb=MUT; p.paragraph_format.space_after=Pt(8)
def table(headers,data,widths=None,hi=None,rf=1):
    hi=hi or []; t=doc.add_table(rows=1,cols=len(headers)); t.alignment=WD_TABLE_ALIGNMENT.CENTER; borders(t)
    for j,h in enumerate(headers): setc(t.rows[0].cells[j],h,bold=True,color=WHT,size=9,align="left" if j==0 else "right",fill="0B5C65")
    for i,row in enumerate(data):
        cells=t.add_row().cells; fill="E8F4F5" if i in hi else ("F7FAFA" if i%2 else None)
        for j,v in enumerate(row): setc(cells[j],v,bold=(i in hi),size=9.5,align="left" if j<rf else "right",fill=fill)
    if widths:
        for j,w in enumerate(widths):
            for rr in t.rows: rr.cells[j].width=Inches(w)
    doc.add_paragraph().paragraph_format.space_after=Pt(2); return t

def m(v): return f"${v:,.0f}"
# ---- title ----
tp=doc.add_paragraph(); tp.paragraph_format.space_before=Pt(6)
r=tp.add_run("DATABEAT"); r.bold=True; r.font.size=Pt(18); r.font.color.rgb=TEALD_c
r2=tp.add_run("   |   Prepared for TIME"); r2.font.size=Pt(12); r2.font.color.rgb=MUT
H("PG/PD Incremental Revenue Model",size=24,sb=16)
p=doc.add_paragraph(); r=p.add_run("How much additional revenue TIME can generate by converting a share of "
     "open-auction programmatic impressions into Preferred Deals / Programmatic Guaranteed (US inventory)")
r.italic=True; r.font.size=Pt(12); r.font.color.rgb=INKc

# ---- headline callout ----
callout("Bottom line:",
        f"across Entertainment, U.S., Health and Science & Tech, the top-6 ad slots serve about "
        f"{tot_m_imp/1e6:.1f}M programmatic impressions per month that earn only {m(tot_m_rev)}/month today "
        f"(~${blend_c:.2f} eCPM). Converting 10% of that inventory into PG/PD at $8–$10 generates "
        f"{m(r8_10)}–{m(r10_10)}/month versus {m(cur10)} today — an incremental "
        f"{m(r8_10-cur10)}–{m(r10_10-cur10)}/month (~{m((r8_10-cur10)*12)}–{m((r10_10-cur10)*12)}/year), "
        f"built gradually over 30/60/90 days.")

# ---- scope ----
H("Scope & Method",size=13)
bullet("U.S. inventory, top-6 high-impact ad slots (stickyfooter1, leaderboard1, inline1–3, rightrail1).",lead="What —")
bullet("Google Ad Manager, January–May 2026 (5 months); monthly figures = period ÷ 5. June 2026 excluded (a CPD deal was mis-booked as CPM).",lead="Period —")
bullet("shift = share of programmatic impressions re-sold as PG/PD. PG/PD priced at a fixed $8 and $10 eCPM. Scenarios: 5% / 10% / 15%.",lead="Assumptions —")
bullet("incremental revenue = (shifted impressions ÷ 1000) × (PG/PD eCPM − current programmatic eCPM).",lead="Formula —")

# ---- current state ----
H("1.  Current State — Average Monthly Programmatic Inventory",size=13)
body("These are the average monthly programmatic impressions on the top-6 slots for each section, and what "
     "they earn today at current open-auction eCPMs. Direct sales on the same sections clear far higher — "
     "shown for reference — which is exactly why PG/PD at $8–$10 is realistic and non-cannibalizing.")
table(["Section","Avg monthly prog. impressions","Current prog. eCPM","Current revenue / month","Direct eCPM (ref.)"],
      [[r["name"],f"{r['m_imp']:,.0f}",f"${r['c']:.2f}",m(r["m_rev"]),f"${r['std']:.2f}"] for r in rows]
      +[["Total (4 sections)",f"{tot_m_imp:,.0f}",f"${blend_c:.2f}",m(tot_m_rev),"—"]],
      widths=[1.9,2.1,1.3,1.6,1.3],hi=[len(rows)])
body("Context: Preferred Deals in the account already clear at roughly 2× the open-auction rate, and current "
     "PG/PD volume is under ~0.1% of impressions — so this is untapped, not unproven.",size=9.5,italic=True,color=MUT)

# ---- 10% conversion ----
H("2.  The 10% Conversion — Current vs PG/PD at $8 / $10",size=13)
body("Taking 10% of each section's monthly programmatic impressions and re-selling it as PG/PD:")
d=[]
for r in rows:
    sh=r["m_imp"]*0.10; cr=sh*r["c"]/1000; a8=sh*8/1000; a10=sh*10/1000
    d.append([r["name"],f"{sh:,.0f}",m(cr),m(a8),m(a10),m(a8-cr),m(a10-cr)])
d.append(["Total",f"{tot_m_imp*0.10:,.0f}",m(cur10),m(r8_10),m(r10_10),m(r8_10-cur10),m(r10_10-cur10)])
table(["Section","10% impressions","Current rev","Rev @ $8","Rev @ $10","Uplift @ $8","Uplift @ $10"],d,
      widths=[1.6,1.5,1.1,1.1,1.1,1.1,1.1],hi=[len(rows)])
img("m1_conversion.png",w=6.6,cap="Figure 1 — Monthly revenue on the 10% shifted impressions: today vs PG/PD at $8 / $10.")

# ---- scenarios ----
H("3.  Scenarios — 5% / 10% / 15% Shift",size=13)
body("Incremental monthly revenue across the four sections at each shift level:")
sc=[]
for pct in (0.05,0.10,0.15):
    cur,r8,r10=at_shift(pct)
    sc.append([f"{int(pct*100)}%",m(r8-cur),m(r10-cur),m((r8-cur)*12),m((r10-cur)*12)])
table(["Shift level","Incremental / mo @ $8","Incremental / mo @ $10","Annualized @ $8","Annualized @ $10"],sc,
      widths=[1.4,1.9,1.9,1.6,1.6],hi=[1])
img("m2_scenarios.png",w=6.4,cap="Figure 2 — Total incremental monthly revenue by scenario.")

# ---- ramp ----
H("4.  30 / 60 / 90-Day Ramp",size=13)
body("Monica asked for a phased build, not instant impact. PG/PD scales gradually as deals are signed and "
     "activated. The ramp maps to the shift scenarios above (blended across the four sections):")
ramp=[]
for period,stage,pct in [("30 days","Small pilot",0.05),("60 days","Broader rollout",0.10),("90 days","Scaled rollout",0.15)]:
    act=tot_m_imp*pct
    ramp.append([period,stage,f"{int(pct*100)}%",f"{act:,.0f}",m(act*8/1000),m(act*10/1000),
                 m(act*(8-blend_c)/1000),m(act*(10-blend_c)/1000)])
table(["Period","Stage","% activated","Prog. imp / mo","PG/PD rev @ $8","PG/PD rev @ $10","Uplift @ $8","Uplift @ $10"],ramp,
      widths=[0.9,1.3,0.9,1.2,1.1,1.1,1.0,1.0])

# ---- narrative ----
H("5.  What This Means (for Monica)",size=13)
bullet("PG/PD is under ~0.1% of impressions today — effectively an untapped channel.",lead="Current state —")
bullet("premium slots clear at ~$2–$4 programmatically while direct earns $17–$29; PG/PD at $8–$10 captures the middle without touching direct.",lead="Opportunity —")
bullet(f"converting 10% of programmatic in these sections adds {m(r8_10-cur10)}–{m(r10_10-cur10)}/month "
       f"(~{m((r8_10-cur10)*12)}–{m((r10_10-cur10)*12)}/year); 15% roughly scales that by 1.5×.",lead="Impact —")
bullet("gradual over 30/60/90 days — pilot, broaden, scale — not an overnight shift.",lead="Ramp —")
callout("One-line answer:",
        f"“If we convert 10% of today's programmatic impressions in these high-value sections into PG/PD at "
        f"$8–$10, we lift them from {m(cur10)} to {m(r8_10)}–{m(r10_10)} per month — roughly "
        f"{m(r8_10-cur10)}–{m(r10_10-cur10)} of incremental revenue every month, ramping over one quarter.”")

# ---- footer / scope note ----
fp=doc.add_paragraph(); fp.paragraph_format.space_before=Pt(14)
r=fp.add_run("Scope note: figures cover four sections (Entertainment, U.S., Health, Science & Tech) on the "
     "top-6 ad slots only; other sections (e.g., World, Politics) and remaining slots are additive and not "
     "included here. U.S. inventory, GAM Jan–May 2026, June excluded. Prepared by Databeat for TIME.")
r.italic=True; r.font.size=Pt(8.5); r.font.color.rgb=MUT

out=os.path.join(HERE,"TIME_PG-PD_Incremental_Revenue_Model.docx")
doc.save(out); print("saved:",out)
