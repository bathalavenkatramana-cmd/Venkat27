"""FINAL client-ready SLOT-WISE PG/PD Incremental Revenue Model doc for TIME.
Section x slot, US, top-6 slots. Fixed $8/$10 PG/PD eCPM. 5/10/15% shift scenarios.
Impressions from client GAM pivot (reconciled); current prog eCPM from confirmed section rate card.
Period = 18 months (editable assumption; implied by ~3.86x volume vs the 5-month slice).
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

HERE=os.path.dirname(os.path.abspath(__file__))
MONTHS=18; PA=8.0; PB=10.0; BASE=0.10

PROG = {
 "Entertainment": {"stickyfooter1":47410845,"leaderboard1":11875493,"inline1":11630907,"inline2":10920299,"inline3":10019326,"rightrail1":4846780},
 "Health":        {"stickyfooter1":20185172,"leaderboard1":7007562,"inline1":6850100,"inline2":5874817,"inline3":4930438,"rightrail1":4880198},
 "U.S.":          {"stickyfooter1":19228301,"leaderboard1":6609878,"inline1":4331009,"inline2":3319301,"inline3":2359398,"rightrail1":2677384},
 "Tech":          {"stickyfooter1":2725359,"leaderboard1":1551790,"inline1":1137367,"inline2":885597,"inline3":673770,"rightrail1":821619},
 "Science":       {"stickyfooter1":2006324,"leaderboard1":1612613,"inline1":1280246,"inline2":1046394,"inline3":781024,"rightrail1":571444},
}
ECPM={"Entertainment":2.36,"U.S.":2.63,"Health":4.04,"Science":3.39,"Tech":3.50}
DIRECT={"Entertainment":16.8,"U.S.":19.2,"Health":23.8,"Science":26.5,"Tech":32.7}
SLOTS=["stickyfooter1","leaderboard1","inline1","inline2","inline3","rightrail1"]
SECS=list(PROG.keys())

# ---- aggregates ----
tot_prog=sum(PROG[s][sl] for s in SECS for sl in SLOTS)
assert tot_prog==200050755, tot_prog

# per-slot aggregate across sections
slot_imp={sl:sum(PROG[s][sl] for s in SECS) for sl in SLOTS}
slot_rev={sl:sum(PROG[s][sl]*ECPM[s]/1000 for s in SECS) for sl in SLOTS}   # period $
slot_ecpm={sl:slot_rev[sl]/slot_imp[sl]*1000 for sl in SLOTS}
assert abs(slot_imp["stickyfooter1"]-91556001)<2, slot_imp["stickyfooter1"]

# per-section aggregate
sec_imp={s:sum(PROG[s][sl] for sl in SLOTS) for s in SECS}
sec_rev={s:sum(PROG[s][sl]*ECPM[s]/1000 for sl in SLOTS) for s in SECS}

tot_rev=sum(slot_rev.values()); blend=tot_rev/tot_prog*1000
tot_m_imp=tot_prog/MONTHS; tot_m_rev=tot_rev/MONTHS
def uplift(imp_period, ecpm, price, pct):  # monthly uplift
    m=imp_period/MONTHS; return m*pct/1000*(price-ecpm)
def rev_at(imp_period, price, pct):
    m=imp_period/MONTHS; return m*pct*price/1000
def cur_rev(imp_period, ecpm, pct):
    m=imp_period/MONTHS; return m*pct*ecpm/1000

# totals at 10%
cur10=sum(cur_rev(slot_imp[sl],slot_ecpm[sl],BASE) for sl in SLOTS)
r8_10=sum(rev_at(slot_imp[sl],PA,BASE) for sl in SLOTS)
r10_10=sum(rev_at(slot_imp[sl],PB,BASE) for sl in SLOTS)
up8=r8_10-cur10; up10=r10_10-cur10
print("CROSS-CHECK")
print(f"total prog (period) {tot_prog:,} | monthly {tot_m_imp:,.0f} | blended eCPM ${blend:.2f}")
print(f"monthly prog rev ${tot_m_rev:,.0f}")
print(f"10% shift: current ${cur10:,.0f}/mo | @${PA:.0f} ${r8_10:,.0f} (+${up8:,.0f}) | @${PB:.0f} ${r10_10:,.0f} (+${up10:,.0f})")
print(f"annualized uplift: @${PA:.0f} ${up8*12:,.0f} | @${PB:.0f} ${up10*12:,.0f}")
print("stickyfooter1 share of prog:", f"{slot_imp['stickyfooter1']/tot_prog*100:.1f}%")
for sl in SLOTS:
    print(f"  {sl:14s} imp {slot_imp[sl]:>12,} ({slot_imp[sl]/tot_prog*100:4.1f}%) eCPM ${slot_ecpm[sl]:.2f}  uplift@8 +${uplift(slot_imp[sl],slot_ecpm[sl],PA,BASE):,.0f}/mo")

# =============== CHARTS ===============
TEAL="#0f7d88"; GOLD="#e0a83c"; GREY="#b9c2c5"; INK="#1f2a30"; TEALD="#0b5c65"
plt.rcParams.update({"font.family":"DejaVu Sans","font.size":11,"axes.grid":True,
                     "grid.color":"#e6eef0","axes.axisbelow":True,"axes.edgecolor":"#c9d2d4"})
def save(fig,n): fig.savefig(os.path.join(HERE,n),dpi=170,bbox_inches="tight",facecolor="white"); plt.close(fig)

# Chart A: programmatic impressions by slot (share) - horizontal bar
order=sorted(SLOTS,key=lambda s:slot_imp[s],reverse=True)
vals=[slot_imp[s]/1e6 for s in order]; shares=[slot_imp[s]/tot_prog*100 for s in order]
fig,ax=plt.subplots(figsize=(8.2,3.6))
bars=ax.barh(range(len(order)),vals,color=[TEAL if s=="stickyfooter1" else GREY for s in order])
ax.set_yticks(range(len(order))); ax.set_yticklabels(order); ax.invert_yaxis()
for i,(v,sh) in enumerate(zip(vals,shares)):
    ax.text(v+1,i,f"{v:.1f}M  ({sh:.0f}%)",va="center",fontsize=9.5,fontweight="bold",color=INK)
ax.set_xlabel("Programmatic impressions (period, millions)"); ax.set_xlim(0,105)
ax.set_title("Programmatic impressions by slot — stickyfooter1 is ~46% of the convertible pool",
             fontsize=11.5,fontweight="bold",color=TEALD,pad=8); ax.grid(axis="y",visible=False)
save(fig,"s1_slot_impressions.png")

# Chart B: incremental monthly revenue by slot at $8/$10 (10% shift)
inc8=[uplift(slot_imp[s],slot_ecpm[s],PA,BASE) for s in order]
inc10=[uplift(slot_imp[s],slot_ecpm[s],PB,BASE) for s in order]
x=range(len(order)); w=0.38
fig,ax=plt.subplots(figsize=(8.4,3.8))
b1=ax.bar([i-w/2 for i in x],inc8,w,label="Incremental @ $8",color=TEAL)
b2=ax.bar([i+w/2 for i in x],inc10,w,label="Incremental @ $10",color=GOLD)
for bars in (b1,b2):
    for b in bars: ax.text(b.get_x()+b.get_width()/2,b.get_height()+40,f"+${b.get_height():,.0f}",ha="center",fontsize=8.5,fontweight="bold",color=INK)
ax.set_xticks(list(x)); ax.set_xticklabels(order,fontsize=9.5)
ax.set_ylabel("Incremental revenue / month"); ax.set_ylim(0,max(inc10)*1.2)
ax.yaxis.set_major_formatter(FuncFormatter(lambda y,_:f"${y:,.0f}"))
ax.set_title("Incremental monthly revenue by slot — 10% of programmatic shifted to PG/PD",
             fontsize=11.5,fontweight="bold",color=TEALD,pad=8)
ax.legend(frameon=False,fontsize=10,loc="upper right"); ax.grid(axis="x",visible=False)
save(fig,"s2_slot_uplift.png")
print("charts saved")

# =============== DOCX ===============
TEALc=RGBColor(0x0F,0x7D,0x88); TEALD_c=RGBColor(0x0B,0x5C,0x65); INKc=RGBColor(0x1F,0x2A,0x30)
MUT=RGBColor(0x5B,0x6B,0x73); WHT=RGBColor(0xFF,0xFF,0xFF)
doc=Document(); nm=doc.styles["Normal"]; nm.font.name="Calibri"; nm.font.size=Pt(11); nm.font.color.rgb=INKc
nm.paragraph_format.space_after=Pt(6); nm.paragraph_format.line_spacing=1.12
def shade(cell,hx):
    tcPr=cell._tc.get_or_add_tcPr(); sh=OxmlElement("w:shd"); sh.set(qn("w:val"),"clear")
    sh.set(qn("w:fill"),hx); tcPr.append(sh)
def borders(table,color="C9D2D4",sz=6):
    tblPr=table._tbl.tblPr; b=OxmlElement("w:tblBorders")
    for e in ("top","left","bottom","right","insideH","insideV"):
        x=OxmlElement(f"w:{e}"); x.set(qn("w:val"),"single"); x.set(qn("w:sz"),str(sz))
        x.set(qn("w:space"),"0"); x.set(qn("w:color"),color); b.append(x)
    tblPr.append(b)
def setc(cell,t,bold=False,color=None,size=10,align="left",fill=None):
    cell.text=""; p=cell.paragraphs[0]
    p.alignment={"left":WD_ALIGN_PARAGRAPH.LEFT,"center":WD_ALIGN_PARAGRAPH.CENTER,"right":WD_ALIGN_PARAGRAPH.RIGHT}[align]
    rn=p.add_run(t); rn.bold=bold; rn.font.size=Pt(size)
    if color is not None: rn.font.color.rgb=color
    if fill: shade(cell,fill)
    p.paragraph_format.space_after=Pt(2); p.paragraph_format.space_before=Pt(2)
def H(t,size=16,color=TEALD_c,sb=14):
    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(sb); p.paragraph_format.space_after=Pt(4)
    rn=p.add_run(t); rn.bold=True; rn.font.color.rgb=color; rn.font.size=Pt(size); return p
def body(t,size=11,italic=False,color=INKc,after=6,bold=False):
    p=doc.add_paragraph(); rn=p.add_run(t); rn.italic=italic; rn.bold=bold; rn.font.size=Pt(size); rn.font.color.rgb=color
    p.paragraph_format.space_after=Pt(after); return p
def bullet(t,lead=None):
    p=doc.add_paragraph(style="List Bullet")
    if lead: rn=p.add_run(lead+" "); rn.bold=True; rn.font.color.rgb=TEALD_c; rn.font.size=Pt(11)
    r2=p.add_run(t); r2.font.size=Pt(11); r2.font.color.rgb=INKc; p.paragraph_format.space_after=Pt(3); return p
def callout(bold,rest):
    t=doc.add_table(rows=1,cols=1); c=t.cell(0,0); shade(c,"E8F4F5"); c.text=""
    p=c.paragraphs[0]; p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(6)
    if bold: rb=p.add_run(bold+"  "); rb.bold=True; rb.font.color.rgb=TEALD_c; rb.font.size=Pt(11.5)
    rr=p.add_run(rest); rr.font.size=Pt(11); rr.font.color.rgb=INKc; borders(t,"0F7D88",12)
    doc.add_paragraph().paragraph_format.space_after=Pt(2); return t
def img(n,w=6.6,cap=None):
    doc.add_picture(os.path.join(HERE,n),width=Inches(w)); doc.paragraphs[-1].alignment=WD_ALIGN_PARAGRAPH.CENTER
    if cap:
        p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        rn=p.add_run(cap); rn.italic=True; rn.font.size=Pt(9); rn.font.color.rgb=MUT; p.paragraph_format.space_after=Pt(8)
def table(headers,data,widths=None,hi=None,rf=2,fs=9.0):
    hi=hi or []; t=doc.add_table(rows=1,cols=len(headers)); t.alignment=WD_TABLE_ALIGNMENT.CENTER; borders(t)
    for j,h in enumerate(headers): setc(t.rows[0].cells[j],h,bold=True,color=WHT,size=8.5,align="left" if j<rf else "right",fill="0B5C65")
    for i,row in enumerate(data):
        cells=t.add_row().cells
        sub=isinstance(row,tuple) and row[0]=="SUB"; grand=isinstance(row,tuple) and row[0]=="GRAND"
        vals=row[1] if (sub or grand) else row
        fill="DCEAEC" if sub else ("E8F4F5" if (i in hi or grand) else ("F7FAFA" if i%2 else None))
        for j,v in enumerate(vals): setc(cells[j],v,bold=(sub or grand or i in hi),size=fs,align="left" if j<rf else "right",fill=fill)
    if widths:
        for j,w in enumerate(widths):
            for rr in t.rows: rr.cells[j].width=Inches(w)
    doc.add_paragraph().paragraph_format.space_after=Pt(2); return t
def m(v): return f"${v:,.0f}"

# ---- title ----
tp=doc.add_paragraph(); tp.paragraph_format.space_before=Pt(6)
rn=tp.add_run("DATABEAT"); rn.bold=True; rn.font.size=Pt(18); rn.font.color.rgb=TEALD_c
r2=tp.add_run("   |   Prepared for TIME"); r2.font.size=Pt(12); r2.font.color.rgb=MUT
H("PG/PD Incremental Revenue Model — Slot-wise",size=23,sb=14)
p=doc.add_paragraph(); rn=p.add_run("Section × slot view: how much additional revenue TIME can generate by converting a "
    "share of open-auction programmatic impressions into Preferred Deals / Programmatic Guaranteed, broken down "
    "to the individual ad slot (US inventory, top-6 slots)")
rn.italic=True; rn.font.size=Pt(12); rn.font.color.rgb=INKc

callout("Bottom line:",
    f"across the five sections and six top slots, TIME serves ~{tot_m_imp/1e6:.1f}M programmatic impressions/month "
    f"earning ~{m(tot_m_rev)}/month today (~${blend:.2f} eCPM). Converting 10% into PG/PD at $8–$10 yields "
    f"{m(r8_10)}–{m(r10_10)}/month vs {m(cur10)} today — incremental {m(up8)}–{m(up10)}/month "
    f"(~{m(up8*12)}–{m(up10*12)}/year). One slot — stickyfooter1 — is ~46% of the entire convertible pool, so it "
    f"is the single highest-leverage place to start.")

# ---- scope ----
H("Scope & Method",size=13)
bullet("US inventory, six top slots (stickyfooter1, leaderboard1, inline1–3, rightrail1) across five sections (Entertainment, U.S., Health, Tech, Science).",lead="What —")
bullet("GAM impressions pivot (section × slot), reconciled to the grand total (200.05M programmatic impressions).",lead="Data —")
bullet("shift = share of programmatic impressions re-sold as PG/PD, priced at a fixed $8 and $10 eCPM; scenarios 5% / 10% / 15%.",lead="Assumptions —")
bullet("incremental = (shifted impressions ÷ 1000) × (PG/PD eCPM − current programmatic eCPM).",lead="Formula —")
body("Two transparency notes (both editable in the companion spreadsheet): (1) this pivot carries impressions "
     "only, so current programmatic eCPM is applied from the confirmed section-level rate card — the same rate "
     "across a section's slots; slot-level eCPMs would refine the split but not the totals. (2) The programmatic "
     "volume implies an ~18-month window (Jan-2025–Jun-2026); “monthly” figures divide by 18. Confirm the export "
     "period and the per-month figures rescale exactly.",size=9.5,italic=True,color=MUT)

# ---- 1. slot current state ----
H("1.  Current State — Programmatic Inventory by Slot",size=13)
body("Aggregated across the five sections, here is how the programmatic impression pool and its current "
     "open-auction revenue split across the six slots:")
rows=[]
for sl in order:
    rows.append([sl,f"{slot_imp[sl]:,.0f}",f"{slot_imp[sl]/tot_prog*100:.1f}%",f"{slot_imp[sl]/MONTHS:,.0f}",
                 f"${slot_ecpm[sl]:.2f}",m(slot_rev[sl]/MONTHS)])
rows.append(("GRAND",["All 6 slots",f"{tot_prog:,.0f}","100%",f"{tot_m_imp:,.0f}",f"${blend:.2f}",m(tot_m_rev)]))
table(["Slot","Prog. imp (period)","Share","Monthly prog. imp","Current eCPM","Current rev / mo"],rows,
      widths=[1.5,1.6,0.8,1.5,1.1,1.2],rf=1)
img("s1_slot_impressions.png",w=6.6,cap="Figure 1 — Programmatic impressions by slot. stickyfooter1 alone is ~46% of the convertible pool.")

# ---- 2. 10% conversion by slot ----
H("2.  The 10% Conversion — by Slot",size=13)
body("Taking 10% of each slot's monthly programmatic impressions and re-selling as PG/PD at $8 / $10:")
rows=[]
for sl in order:
    sh=slot_imp[sl]/MONTHS*BASE
    rows.append([sl,f"{sh:,.0f}",m(cur_rev(slot_imp[sl],slot_ecpm[sl],BASE)),
                 m(rev_at(slot_imp[sl],PA,BASE)),m(rev_at(slot_imp[sl],PB,BASE)),
                 m(uplift(slot_imp[sl],slot_ecpm[sl],PA,BASE)),m(uplift(slot_imp[sl],slot_ecpm[sl],PB,BASE))])
rows.append(("GRAND",["All 6 slots",f"{tot_m_imp*BASE:,.0f}",m(cur10),m(r8_10),m(r10_10),m(up8),m(up10)]))
table(["Slot","10% imp / mo","Current rev","Rev @ $8","Rev @ $10","Uplift @ $8","Uplift @ $10"],rows,
      widths=[1.5,1.2,1.1,1.05,1.05,1.1,1.1],rf=1)
img("s2_slot_uplift.png",w=6.6,cap="Figure 2 — Incremental monthly revenue by slot at $8 / $10 (10% shift).")
callout("Slot takeaway:",
    f"stickyfooter1 alone contributes ~{m(uplift(slot_imp['stickyfooter1'],slot_ecpm['stickyfooter1'],PA,BASE))}–"
    f"{m(uplift(slot_imp['stickyfooter1'],slot_ecpm['stickyfooter1'],PB,BASE))}/month of the total uplift — roughly "
    f"46%. Packaging stickyfooter1 as the first PG/PD product captures nearly half the opportunity from a single, "
    f"high-viewability unit.")

# ---- 3. full section x slot detail ----
H("3.  Full Detail — Section × Slot",size=13)
body("The complete granular view Monica requested. Uplift is monthly, at the base 10% shift.")
rows=[]
for s in SECS:
    for sl in SLOTS:
        imp=PROG[s][sl]
        rows.append([s,sl,f"{imp:,.0f}",f"{imp/MONTHS:,.0f}",f"${ECPM[s]:.2f}",
                     m(uplift(imp,ECPM[s],PA,BASE)),m(uplift(imp,ECPM[s],PB,BASE))])
    rows.append(("SUB",[f"{s} subtotal","",f"{sec_imp[s]:,.0f}",f"{sec_imp[s]/MONTHS:,.0f}",f"${sec_rev[s]/sec_imp[s]*1000:.2f}",
                        m(sum(uplift(PROG[s][sl],ECPM[s],PA,BASE) for sl in SLOTS)),
                        m(sum(uplift(PROG[s][sl],ECPM[s],PB,BASE) for sl in SLOTS))]))
rows.append(("GRAND",["GRAND TOTAL","",f"{tot_prog:,.0f}",f"{tot_m_imp:,.0f}",f"${blend:.2f}",m(up8),m(up10)]))
table(["Section","Slot","Prog. imp (period)","Monthly imp","Cur. eCPM","Uplift/mo @ $8","Uplift/mo @ $10"],rows,
      widths=[1.5,1.3,1.5,1.2,0.95,1.2,1.2],rf=2,fs=8.3)

# ---- 4. scenarios ----
H("4.  Scenarios — 5% / 10% / 15% Shift",size=13)
sc=[]
for pct in (0.05,0.10,0.15):
    c=sum(cur_rev(slot_imp[sl],slot_ecpm[sl],pct) for sl in SLOTS)
    a8=sum(rev_at(slot_imp[sl],PA,pct) for sl in SLOTS); a10=sum(rev_at(slot_imp[sl],PB,pct) for sl in SLOTS)
    sc.append([f"{int(pct*100)}%",m(a8-c),m(a10-c),m((a8-c)*12),m((a10-c)*12)])
table(["Shift level","Incremental / mo @ $8","Incremental / mo @ $10","Annualized @ $8","Annualized @ $10"],sc,
      widths=[1.4,1.9,1.9,1.6,1.6],hi=[1],rf=1)

# ---- 5. what it means ----
H("5.  What This Means (for Monica)",size=13)
bullet("go slot-first, not section-first: stickyfooter1 is ~46% of the programmatic pool — one product captures nearly half the upside.",lead="Focus —")
bullet(f"converting 10% of programmatic across these slots adds {m(up8)}–{m(up10)}/month (~{m(up8*12)}–{m(up10*12)}/year); 15% scales it ~1.5×.",lead="Impact —")
bullet("premium slots clear at ~$2–$4 programmatically vs direct at $17–$33; PG/PD at $8–$10 captures the middle without cannibalizing direct.",lead="Headroom —")
bullet("stickyfooter1 → leaderboard1 → inline units, phased over 30/60/90 days as deals are signed.",lead="Sequence —")
callout("One-line answer:",
    f"“Slot-by-slot, converting 10% of programmatic in these top-6 units to PG/PD at $8–$10 lifts them from "
    f"{m(cur10)} to {m(r8_10)}–{m(r10_10)} per month — ~{m(up8)}–{m(up10)} incremental every month — and nearly "
    f"half of that comes from a single unit, stickyfooter1.”")

fp=doc.add_paragraph(); fp.paragraph_format.space_before=Pt(12)
rn=fp.add_run("Scope note: five sections (Entertainment, U.S., Health, Tech, Science) × six top slots; US inventory. "
    "Impressions from GAM pivot; current programmatic eCPM applied from the confirmed section rate card (impressions-only "
    "source). Monthly = period ÷ 18 (assumed window Jan-2025–Jun-2026; confirm). World & Politics not included (additive). "
    "Prepared by Databeat for TIME.")
rn.italic=True; rn.font.size=Pt(8.5); rn.font.color.rgb=MUT

out=os.path.join(HERE,"TIME_PG-PD_SlotWise_Incremental_Model.docx")
doc.save(out); print("saved:",out)
