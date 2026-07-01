"""Build Monica's PG/PD incremental-revenue model.
Data: 2026 Jan-May (5 months), US, TOP-6 slots, by section (GAM).
Formula (Monica): incremental = (shifted programmatic impressions / 1000) x (PG/PD eCPM - current programmatic eCPM)
Scenarios: 5% / 10% / 15% of PROGRAMMATIC impressions shifted.
PG/PD eCPM base = 2.2x current programmatic; also compare fixed $8 / $10.
Reference period = 5 months -> monthly = period / 5.
NOTE: World (a Monica priority) is NOT in this dataset -> flagged, to be added.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

HERE = os.path.dirname(os.path.abspath(__file__))
MONTHS = 5

# section -> (programmatic period impressions, programmatic period revenue, standard eCPM benchmark)
sec = {
 "Entertainment": (22835006, 53795, 16.8),
 "U.S.":          (15568972, 40869, 19.2),
 "Health":        (10085819, 40761, 23.8),
 "Science & Tech":(1875841+1462654, 6356+5119, 28.5),  # combined per Monica
}

def ecpm(rev, imp): return rev/imp*1000

print("=== Current programmatic (monthly, Jan-May 2026, US, top-6 slots) ===")
for s,(imp,rev,std) in sec.items():
    print(f"{s:15s} monthly prog imp {imp/MONTHS:>12,.0f}  eCPM ${ecpm(rev,imp):5.2f}  (direct benchmark ${std})")

print("\n=== Incremental MONTHLY revenue by scenario ===")
def block(title, pgpd_fn):
    print(f"\n-- {title} --")
    print(f"{'Section':15s} {'PG/PD eCPM':>10s} {'@5%':>9s} {'@10%':>9s} {'@15%':>9s}")
    tot={5:0,10:0,15:0}
    for s,(imp,rev,std) in sec.items():
        m=imp/MONTHS; c=ecpm(rev,imp); p=pgpd_fn(c)
        row=[]
        for pct in (5,10,15):
            inc=(m*pct/100/1000)*(p-c); tot[pct]+=inc; row.append(inc)
        print(f"{s:15s} ${p:9.2f} ${row[0]:8,.0f} ${row[1]:8,.0f} ${row[2]:8,.0f}")
    print(f"{'TOTAL (5 sec)':15s} {'':>10s} ${tot[5]:8,.0f} ${tot[10]:8,.0f} ${tot[15]:8,.0f}")
    print(f"{'  annualized':15s} {'':>10s} ${tot[5]*12:8,.0f} ${tot[10]*12:8,.0f} ${tot[15]*12:8,.0f}")
    return tot

block("PG/PD = 2.2x current programmatic (Monica base)", lambda c: c*2.2)
block("PG/PD = fixed $8", lambda c: 8.0)
block("PG/PD = fixed $10", lambda c: 10.0)

# blended for ramp
tot_imp=sum(v[0] for v in sec.values()); tot_rev=sum(v[1] for v in sec.values())
blend_c=ecpm(tot_rev,tot_imp); m_tot=tot_imp/MONTHS
print(f"\n=== Blended current programmatic: ${blend_c:.2f} eCPM, {m_tot:,.0f} imp/mo ===")
print("=== 30/60/90 ramp (PG/PD @ $8) ===")
for period,stage,pct in [("30 days","Small pilot",5),("60 days","Broader rollout",10),("90 days","Scaled rollout",15)]:
    act=m_tot*pct/100; rev=act*8/1000; inc=act*(8-blend_c)/1000
    print(f"{period:8s} {stage:16s} {pct:2d}%  activated {act:>11,.0f}/mo  PG/PD rev ${rev:8,.0f}  incremental ${inc:8,.0f}/mo")

# ======================= BUILD XLSX =======================
wb=openpyxl.Workbook(); ws=wb.active; ws.title="PG-PD Incremental Model"
TEAL="0F7D88"; TEALD="0B5C65"; LIGHT="E8F4F5"; GOLD="FCEFCB"; GREY="F2F5F6"
def style(cell,bold=False,color="1F2A30",size=10,fill=None,align="left",fmt=None,white=False):
    cell.font=Font(bold=bold,color=("FFFFFF" if white else color),size=size,name="Calibri")
    cell.alignment=Alignment(horizontal=align,vertical="center",wrap_text=False)
    if fill: cell.fill=PatternFill("solid",fgColor=fill)
    if fmt: cell.number_format=fmt
thin=Side(style="thin",color="C9D2D4")
def border(rng):
    for row in ws[rng]:
        for c in row: c.border=Border(thin,thin,thin,thin)

r=1
ws.merge_cells(f"A{r}:H{r}")
style(ws[f"A{r}"],bold=True,size=15,white=True,fill=TEALD,align="left")
ws[f"A{r}"]="TIME  |  PG/PD Incremental Revenue Model  —  US, Top-6 Slots (Jan–May 2026)"
ws.row_dimensions[r].height=24
r+=2

# assumptions
style(ws[f"A{r}"],bold=True,size=11,color=TEALD); ws[f"A{r}"]="Assumptions (editable)"; r+=1
assume=[("PG/PD eCPM multiple (× current programmatic)",2.2,"0.0"),
        ("Low case — % programmatic impressions shifted",0.05,"0%"),
        ("Base case — % shifted",0.10,"0%"),
        ("High case — % shifted",0.15,"0%"),
        ("Months in reference period",5,"0")]
assume_cells={}
for i,(lbl,val,fmt) in enumerate(assume):
    style(ws[f"A{r}"],size=10); ws[f"A{r}"]=lbl
    style(ws[f"B{r}"],bold=True,fill=GOLD,align="center",fmt=fmt); ws[f"B{r}"]=val
    border(f"A{r}:B{r}")
    assume_cells[lbl]=f"$B${r}"; r+=1
MULT=assume_cells["PG/PD eCPM multiple (× current programmatic)"]
LOW=assume_cells["Low case — % programmatic impressions shifted"]
BASE=assume_cells["Base case — % shifted"]
HIGH=assume_cells["High case — % shifted"]
MO=assume_cells["Months in reference period"]
r+=1

# main table
hdr=["Section","Prog. impressions (period)","Prog. revenue (period)","Monthly prog. impressions",
     "Current prog. eCPM","Proposed PG/PD eCPM","Inc./mo @ Low","Inc./mo @ Base","Inc./mo @ High"]
hr=r
for j,h in enumerate(hdr):
    c=ws.cell(row=r,column=j+1,value=h); style(c,bold=True,white=True,fill=TEALD,size=9,align="center")
ws.row_dimensions[r].height=30
r+=1
first=r
for s,(imp,rev,std) in sec.items():
    ws.cell(row=r,column=1,value=s)
    ws.cell(row=r,column=2,value=imp)
    ws.cell(row=r,column=3,value=rev)
    ws.cell(row=r,column=4,value=f"=B{r}/{MO}")
    ws.cell(row=r,column=5,value=f"=C{r}/B{r}*1000")
    ws.cell(row=r,column=6,value=f"=E{r}*{MULT}")
    ws.cell(row=r,column=7,value=f"=D{r}*{LOW}/1000*(F{r}-E{r})")
    ws.cell(row=r,column=8,value=f"=D{r}*{BASE}/1000*(F{r}-E{r})")
    ws.cell(row=r,column=9,value=f"=D{r}*{HIGH}/1000*(F{r}-E{r})")
    for col in range(1,10):
        cc=ws.cell(row=r,column=col)
        style(cc,size=10,align="left" if col==1 else "right",
              fmt=('#,##0' if col in(2,4) else ('$#,##0' if col==3 else ('$#,##0.00' if col in(5,6) else '$#,##0'))))
    r+=1
last=r-1
# total row
ws.cell(row=r,column=1,value="TOTAL (5 sections shown)")
for col,letter in [(7,"G"),(8,"H"),(9,"I")]:
    ws.cell(row=r,column=col,value=f"=SUM({letter}{first}:{letter}{last})")
ws.cell(row=r,column=4,value=f"=SUM(D{first}:D{last})")
for col in range(1,10):
    cc=ws.cell(row=r,column=col)
    style(cc,bold=True,fill=LIGHT,align="left" if col==1 else "right",
          fmt=('#,##0' if col==4 else ('$#,##0' if col in(7,8,9) else None)))
border(f"A{hr}:I{r}")
tot_row=r; r+=2

# annualized note
style(ws[f"A{r}"],bold=True,color=TEALD); ws[f"A{r}"]="Annualized (Base case × 12)"
ws[f"B{r}"]=f"=H{tot_row}*12"; style(ws[f"B{r}"],bold=True,fill=GOLD,fmt='$#,##0',align="center"); r+=2

# fixed price comparison
style(ws[f"A{r}"],bold=True,size=11,color=TEALD); ws[f"A{r}"]="Alternative: fixed PG/PD eCPM (Base-case 10% shift)"; r+=1
for j,h in enumerate(["Section","Inc./mo @ $8","Inc./mo @ $10"]):
    c=ws.cell(row=r,column=j+1,value=h); style(c,bold=True,white=True,fill=TEAL,size=9,align="center")
r+=1; f2=r
for s,(imp,rev,std) in sec.items():
    ref=first+list(sec).index(s)
    ws.cell(row=r,column=1,value=s)
    ws.cell(row=r,column=2,value=f"=D{ref}*{BASE}/1000*(8-E{ref})")
    ws.cell(row=r,column=3,value=f"=D{ref}*{BASE}/1000*(10-E{ref})")
    for col in range(1,4):
        style(ws.cell(row=r,column=col),align="left" if col==1 else "right",fmt=None if col==1 else '$#,##0')
    r+=1
ws.cell(row=r,column=1,value="TOTAL")
ws.cell(row=r,column=2,value=f"=SUM(B{f2}:B{r-1})")
ws.cell(row=r,column=3,value=f"=SUM(C{f2}:C{r-1})")
for col in range(1,4): style(ws.cell(row=r,column=col),bold=True,fill=LIGHT,align="left" if col==1 else "right",fmt=None if col==1 else '$#,##0')
border(f"A{f2-1}:C{r}"); r+=2

# ramp table
style(ws[f"A{r}"],bold=True,size=11,color=TEALD); ws[f"A{r}"]="30 / 60 / 90-day ramp (PG/PD @ $8, blended across sections)"; r+=1
for j,h in enumerate(["Period","Stage","% activated","Prog. imp activated /mo","Est. PG/PD revenue /mo","Incremental uplift /mo"]):
    c=ws.cell(row=r,column=j+1,value=h); style(c,bold=True,white=True,fill=TEAL,size=9,align="center")
r+=1
tot_month_imp=f"D{tot_row}"
blend_ecpm=f"(C{first}+C{first+1}+C{first+2}+C{first+3})/(B{first}+B{first+1}+B{first+2}+B{first+3})*1000"
ramp=[("30 days","Small pilot",LOW),("60 days","Broader rollout",BASE),("90 days","Scaled rollout",HIGH)]
for period,stage,pctcell in ramp:
    ws.cell(row=r,column=1,value=period)
    ws.cell(row=r,column=2,value=stage)
    ws.cell(row=r,column=3,value=f"={pctcell}")
    ws.cell(row=r,column=4,value=f"={tot_month_imp}*{pctcell}")
    ws.cell(row=r,column=5,value=f"={tot_month_imp}*{pctcell}*8/1000")
    ws.cell(row=r,column=6,value=f"={tot_month_imp}*{pctcell}*(8-{blend_ecpm})/1000")
    for col in range(1,7):
        style(ws.cell(row=r,column=col),align="left" if col<=2 else "right",
              fmt=('0%' if col==3 else ('#,##0' if col==4 else ('$#,##0' if col in(5,6) else None))))
    r+=1
border(f"A{r-4}:F{r-1}"); r+=1

style(ws[f"A{r}"],italic:=False)
ws[f"A{r}"]="Note: World (a Monica priority) and Politics are NOT in this dataset; adding them will materially increase the totals. % shift = share of PROGRAMMATIC impressions moved to PG/PD."
style(ws[f"A{r}"],size=9,color="8A6D1F",fill="FFF7E6"); ws.merge_cells(f"A{r}:I{r}")

# widths
widths=[26,22,20,22,16,18,15,15,15]
for i,w in enumerate(widths): ws.column_dimensions[get_column_letter(i+1)].width=w
ws.freeze_panes="A"+str(hr+1)

out=os.path.join(HERE,"TIME_PG-PD_Incremental_Model.xlsx")
wb.save(out); print("\nsaved:",out)
