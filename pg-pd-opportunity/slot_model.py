"""TIME PG/PD SLOT-WISE incremental-revenue model (section x slot).
Builds an interactive .xlsx with live formulas. US, top-6 slots.

Data source: client GAM impressions pivots (section x slot), reconciled in slot_reconcile.py.
This pivot carries IMPRESSIONS only (no revenue), so current programmatic eCPM is applied
from the confirmed SECTION-level rate card (editable). Period defaults to 18 months
(full Jan-2025..Jun-2026 window implied by the ~3.86x volume vs the 5-month slice) -> editable.

Formula (Monica): incremental = (shifted prog impressions / 1000) x (PG/PD eCPM - current prog eCPM)
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

HERE = os.path.dirname(os.path.abspath(__file__))

# (section, slot) -> programmatic impressions (period). Reconciled.
PROG = {
 "Entertainment": {"stickyfooter1":47410845,"leaderboard1":11875493,"inline1":11630907,"inline2":10920299,"inline3":10019326,"rightrail1":4846780},
 "Health":        {"stickyfooter1":20185172,"leaderboard1":7007562,"inline1":6850100,"inline2":5874817,"inline3":4930438,"rightrail1":4880198},
 "U.S.":          {"stickyfooter1":19228301,"leaderboard1":6609878,"inline1":4331009,"inline2":3319301,"inline3":2359398,"rightrail1":2677384},
 "Tech":          {"stickyfooter1":2725359,"leaderboard1":1551790,"inline1":1137367,"inline2":885597,"inline3":673770,"rightrail1":821619},
 "Science":       {"stickyfooter1":2006324,"leaderboard1":1612613,"inline1":1280246,"inline2":1046394,"inline3":781024,"rightrail1":571444},
}
# confirmed section programmatic eCPM (proxy applied across that section's slots)
ECPM = {"Entertainment":2.36,"U.S.":2.63,"Health":4.04,"Science":3.39,"Tech":3.50}
SLOT_ORDER = ["stickyfooter1","leaderboard1","inline1","inline2","inline3","rightrail1"]

wb=openpyxl.Workbook(); ws=wb.active; ws.title="Slot-wise PG-PD Model"
TEAL="0F7D88"; TEALD="0B5C65"; LIGHT="E8F4F5"; GOLD="FCEFCB"; SUB="DCEAEC"
def style(cell,bold=False,color="1F2A30",size=10,fill=None,align="left",fmt=None,white=False):
    cell.font=Font(bold=bold,color=("FFFFFF" if white else color),size=size,name="Calibri")
    cell.alignment=Alignment(horizontal=align,vertical="center")
    if fill: cell.fill=PatternFill("solid",fgColor=fill)
    if fmt: cell.number_format=fmt
thin=Side(style="thin",color="C9D2D4")
def border(rng):
    for row in ws[rng]:
        for c in row: c.border=Border(thin,thin,thin,thin)

r=1
ws.merge_cells(f"A{r}:K{r}")
style(ws[f"A{r}"],bold=True,size=15,white=True,fill=TEALD)
ws[f"A{r}"]="TIME  |  PG/PD Incremental Revenue Model — Slot-wise (Section × Slot), US, Top-6 Slots"
ws.row_dimensions[r].height=24; r+=2

# ---------------- assumptions ----------------
style(ws[f"A{r}"],bold=True,size=11,color=TEALD); ws[f"A{r}"]="Assumptions (editable — everything below recalculates)"; r+=1
assume=[("Months in reference period",18,"0"),
        ("PG/PD price A (eCPM)",8,"$#,##0.00"),
        ("PG/PD price B (eCPM)",10,"$#,##0.00"),
        ("Low case — % programmatic shifted",0.05,"0%"),
        ("Base case — % shifted",0.10,"0%"),
        ("High case — % shifted",0.15,"0%")]
A={}
for lbl,val,fmt in assume:
    style(ws[f"A{r}"]); ws[f"A{r}"]=lbl
    style(ws[f"B{r}"],bold=True,fill=GOLD,align="center",fmt=fmt); ws[f"B{r}"]=val
    border(f"A{r}:B{r}"); A[lbl]=f"$B${r}"; r+=1
MO=A["Months in reference period"]; PA=A["PG/PD price A (eCPM)"]; PB=A["PG/PD price B (eCPM)"]
LOW=A["Low case — % programmatic shifted"]; BASE=A["Base case — % shifted"]; HIGH=A["High case — % shifted"]
r+=1

# rate card
style(ws[f"A{r}"],bold=True,size=11,color=TEALD); ws[f"A{r}"]="Current programmatic eCPM rate card (by section — editable)"; r+=1
ratecell={}
style(ws.cell(row=r,column=1,value="Section"),bold=True,white=True,fill=TEAL,size=9)
style(ws.cell(row=r,column=2,value="Prog. eCPM"),bold=True,white=True,fill=TEAL,size=9,align="center")
r+=1
for sec in PROG:
    style(ws.cell(row=r,column=1,value=sec),size=10)
    style(ws.cell(row=r,column=2,value=ECPM[sec]),bold=True,fill=GOLD,align="center",fmt="$#,##0.00")
    ratecell[sec]=f"$B${r}"; border(f"A{r}:B{r}"); r+=1
r+=1

# ---------------- main slot table ----------------
hdr=["Section","Slot","Prog. imp (period)","Monthly prog. imp","Current eCPM",
     "Current rev / mo","Base-shift imp / mo","Rev / mo @ A","Rev / mo @ B","Uplift / mo @ A","Uplift / mo @ B"]
hr=r
for j,h in enumerate(hdr):
    style(ws.cell(row=r,column=j+1,value=h),bold=True,white=True,fill=TEALD,size=8.5,align="center")
ws.row_dimensions[r].height=34; r+=1

data_first=r
section_subtotal_rows=[]
for sec in PROG:
    sec_first=r
    for slot in SLOT_ORDER:
        imp=PROG[sec][slot]
        ws.cell(row=r,column=1,value=sec)
        ws.cell(row=r,column=2,value=slot)
        ws.cell(row=r,column=3,value=imp)
        ws.cell(row=r,column=4,value=f"=C{r}/{MO}")
        ws.cell(row=r,column=5,value=f"={ratecell[sec]}")
        ws.cell(row=r,column=6,value=f"=D{r}*E{r}/1000")
        ws.cell(row=r,column=7,value=f"=D{r}*{BASE}")
        ws.cell(row=r,column=8,value=f"=G{r}*{PA}/1000")
        ws.cell(row=r,column=9,value=f"=G{r}*{PB}/1000")
        ws.cell(row=r,column=10,value=f"=G{r}/1000*({PA}-E{r})")
        ws.cell(row=r,column=11,value=f"=G{r}/1000*({PB}-E{r})")
        for col in range(1,12):
            fmt=None
            if col in(3,4,7): fmt='#,##0'
            elif col==5: fmt='$#,##0.00'
            elif col in(6,8,9,10,11): fmt='$#,##0'
            style(ws.cell(row=r,column=col),size=9.5,align="left" if col<=2 else "right",fmt=fmt)
        r+=1
    # section subtotal
    style(ws.cell(row=r,column=1,value=f"{sec} — subtotal"),bold=True,fill=SUB)
    style(ws.cell(row=r,column=2,value=""),fill=SUB)
    for col,letter in [(3,"C"),(4,"D"),(6,"F"),(7,"G"),(8,"H"),(9,"I"),(10,"J"),(11,"K")]:
        c=ws.cell(row=r,column=col,value=f"=SUM({letter}{sec_first}:{letter}{r-1})")
        style(c,bold=True,fill=SUB,align="right",fmt=('#,##0' if col in(3,4,7) else '$#,##0'))
    style(ws.cell(row=r,column=5,value=f"=F{r}/D{r}*1000"),bold=True,fill=SUB,align="right",fmt='$#,##0.00')
    section_subtotal_rows.append(r); r+=1
data_last=r-1

# grand total (sum of subtotal rows)
style(ws.cell(row=r,column=1,value="GRAND TOTAL (5 sections × 6 slots)"),bold=True,white=True,fill=TEALD)
style(ws.cell(row=r,column=2,value=""),fill=TEALD)
for col,letter in [(3,"C"),(4,"D"),(6,"F"),(7,"G"),(8,"H"),(9,"I"),(10,"J"),(11,"K")]:
    ref="+".join(f"{letter}{sr}" for sr in section_subtotal_rows)
    c=ws.cell(row=r,column=col,value=f"={ref}")
    style(c,bold=True,white=True,fill=TEALD,align="right",fmt=('#,##0' if col in(3,4,7) else '$#,##0'))
style(ws.cell(row=r,column=5,value=f"=F{r}/D{r}*1000"),bold=True,white=True,fill=TEALD,align="right",fmt='$#,##0.00')
grand_row=r
border(f"A{hr}:K{grand_row}"); r+=2

# annualized callouts
style(ws.cell(row=r,column=1,value="Annualized uplift @ price A (Base 10% × 12)"),bold=True,color=TEALD)
style(ws.cell(row=r,column=4,value=f"=J{grand_row}*12"),bold=True,fill=GOLD,fmt='$#,##0',align="center")
style(ws.cell(row=r+1,column=1,value="Annualized uplift @ price B (Base 10% × 12)"),bold=True,color=TEALD)
style(ws.cell(row=r+1,column=4,value=f"=K{grand_row}*12"),bold=True,fill=GOLD,fmt='$#,##0',align="center")
r+=3

# scenario summary (5/10/15%) at grand-total level
style(ws[f"A{r}"],bold=True,size=11,color=TEALD); ws[f"A{r}"]="Scenario summary — incremental revenue / month (all slots)"; r+=1
for j,h in enumerate(["Shift","Incremental / mo @ A","Incremental / mo @ B","Annualized @ A","Annualized @ B"]):
    style(ws.cell(row=r,column=j+1,value=h),bold=True,white=True,fill=TEAL,size=9,align="center")
r+=1
monthly_imp=f"D{grand_row}"; blend=f"E{grand_row}"
for pct in (LOW,BASE,HIGH):
    ws.cell(row=r,column=1,value=f"={pct}")
    ws.cell(row=r,column=2,value=f"={monthly_imp}*{pct}/1000*({PA}-{blend})")
    ws.cell(row=r,column=3,value=f"={monthly_imp}*{pct}/1000*({PB}-{blend})")
    ws.cell(row=r,column=4,value=f"=B{r}*12")
    ws.cell(row=r,column=5,value=f"=C{r}*12")
    for col in range(1,6):
        style(ws.cell(row=r,column=col),align="center" if col==1 else "right",fmt=('0%' if col==1 else '$#,##0'))
    r+=1
border(f"A{r-4}:E{r-1}"); r+=1

note=("Notes: (1) This pivot supplies IMPRESSIONS only; current programmatic eCPM is applied from the "
      "confirmed section rate card above (same rate across a section's slots) — replace with slot-level eCPM "
      "when a revenue export is available for slot-precise figures. (2) 'Monthly' = period ÷ months-in-period; "
      "the ~200M programmatic total implies an ~18-month window (Jan-2025..Jun-2026), not the 5-month slice — "
      "confirm the export period. (3) % shift = share of PROGRAMMATIC impressions re-sold as PG/PD. "
      "World & Politics are not in this dataset (additive).")
ws.merge_cells(f"A{r}:K{r+2}")
style(ws[f"A{r}"],size=9,color="8A6D1F",fill="FFF7E6")
ws[f"A{r}"]=note; ws[f"A{r}"].alignment=Alignment(horizontal="left",vertical="top",wrap_text=True)
ws.row_dimensions[r].height=54

widths=[24,14,17,15,12,14,15,12,12,13,13]
for i,w in enumerate(widths): ws.column_dimensions[get_column_letter(i+1)].width=w
ws.freeze_panes=f"A{hr+1}"

out=os.path.join(HERE,"TIME_PG-PD_SlotWise_Model.xlsx")
wb.save(out); print("saved:",out)
