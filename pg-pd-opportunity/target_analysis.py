"""TIME PG/PD conversion-TARGET analysis (section x slot, US, top-6 slots).
Period: Jan-May 2026 (5 months). Data: GAM pivot with impressions + CPM for
Standard (direct) / Programmatic / House per section x slot.

Goal: find high-volume inventory NOT sold direct (low direct share) but heavily
programmatic (high prog share) -> prime PG/PD conversion targets. Reserve
direct-heavy slots for the sales team.

TARGET rule:   Programmatic share >= 60%  AND  direct(Standard) share <= 35%
RESERVED rule: direct(Standard) share >= 40%   (leave for direct sales)
else SECONDARY.
Model: shift 5/10/15% of a slot's programmatic impressions to PG/PD at fixed $8/$10.
Baseline = that slot's ACTUAL programmatic CPM. Monthly = period / 5.
"""
MONTHS = 5
PA, PB = 8.0, 10.0

# DATA[section][slot] = (std_imp, prog_imp, house_imp, std_cpm, prog_cpm)
DATA = {
 "Entertainment": {
   "stickyfooter1": (13767927, 47410845, 2166739, 21.22, 1.92),
   "leaderboard1":  (9918394, 11875493, 753998, 22.99, 1.88),
   "inline1":       (5880365, 11630907, 396956, 21.88, 2.93),
   "inline2":       (3876484, 10920299, 471422, 26.00, 2.64),
   "inline3":       (2914244, 10019326, 511614, 23.61, 2.55),
   "rightrail1":    (1998657, 4846780, 210342, 19.14, 2.88),
 },
 "Health": {
   "stickyfooter1": (11067220, 20185172, 960572, 26.29, 2.53),
   "leaderboard1":  (8665005, 7007562, 756373, 32.25, 2.81),
   "inline1":       (4831301, 6850100, 414509, 27.80, 5.15),
   "inline2":       (3772312, 5874817, 368795, 28.50, 4.24),
   "inline3":       (2530253, 4930438, 336651, 29.55, 4.11),
   "rightrail1":    (3808220, 4880198, 408881, 22.90, 4.11),
 },
 "Science": {
   "stickyfooter1": (723615, 2006324, 98956, 27.67, 2.65),
   "leaderboard1":  (1012980, 1612613, 69612, 44.32, 2.39),
   "inline1":       (453383, 1280246, 40846, 26.84, 3.19),
   "inline2":       (312130, 1046394, 43054, 27.89, 2.89),
   "inline3":       (229424, 781024, 44550, 27.19, 2.77),
   "rightrail1":    (292628, 571444, 48350, 23.72, 3.06),
 },
 "Tech": {
   "stickyfooter1": (1502361, 2725359, 145460, 28.67, 2.62),
   "leaderboard1":  (2502857, 1551790, 105365, 20.64, 2.97),
   "inline1":       (1245304, 1137367, 53511, 23.65, 5.53),
   "inline2":       (777475, 885597, 54644, 25.27, 4.24),
   "inline3":       (501022, 673770, 43319, 25.62, 3.95),
   "rightrail1":    (866947, 821619, 46627, 27.96, 3.85),
 },
 "U.S.": {
   "stickyfooter1": (5470947, 19228301, 502892, 22.92, 2.13),
   "leaderboard1":  (5864010, 6609878, 435768, 25.97, 2.13),
   "inline1":       (2054682, 4331009, 223459, 24.71, 3.22),
   "inline2":       (1201113, 3319301, 247599, 27.76, 2.85),
   "inline3":       (798265, 2359398, 168507, 28.08, 2.66),
   "rightrail1":    (1317962, 2677384, 181060, 19.87, 2.65),
 },
}
SLOTS = ["stickyfooter1","leaderboard1","inline1","inline2","inline3","rightrail1"]
SECS = list(DATA.keys())

# ---- reconcile impressions + weighted CPM to pivot section/grand totals ----
SEC_PROG_CPM = {"Entertainment":2.23,"Health":3.44,"Science":2.77,"Tech":3.55,"U.S.":2.38}
gstd=gprog=ghouse=0; gprogrev=0
print("="*74); print("RECONCILE — weighted section programmatic CPM vs pivot"); print("="*74)
for s in SECS:
    pi=sum(DATA[s][sl][1] for sl in SLOTS)
    prev=sum(DATA[s][sl][1]*DATA[s][sl][4] for sl in SLOTS)/1000
    wcpm=prev/pi*1000
    print(f"{s:14s} prog imp {pi:>12,}  weighted prog CPM ${wcpm:5.2f}  (pivot ${SEC_PROG_CPM[s]:.2f})  "
          f"{'OK' if abs(wcpm-SEC_PROG_CPM[s])<0.03 else 'CHECK'}")
    gstd+=sum(DATA[s][sl][0] for sl in SLOTS); gprog+=pi; ghouse+=sum(DATA[s][sl][2] for sl in SLOTS); gprogrev+=prev
print(f"\nGRAND  Std {gstd:,} (exp 100,157,487 {'OK' if gstd==100157487 else 'X'}) | "
      f"Prog {gprog:,} (exp 200,050,755 {'OK' if gprog==200050755 else 'X'}) | "
      f"House {ghouse:,} (exp 10,310,431 {'OK' if ghouse==10310431 else 'X'})")
print(f"GRAND blended prog CPM ${gprogrev/gprog*1000:.2f} (pivot $2.63)  prog rev/period ${gprogrev:,.0f}  /mo ${gprogrev/MONTHS:,.0f}")

# ---- classify ----
def classify(direct_sh, prog_sh):
    if prog_sh>=0.60 and direct_sh<=0.35: return "TARGET"
    if direct_sh>=0.40: return "RESERVED"
    return "SECONDARY"

rows=[]
for s in SECS:
    for sl in SLOTS:
        std,prog,house,scpm,pcpm=DATA[s][sl]
        tot=std+prog+house
        d_sh=std/tot; p_sh=prog/tot
        rows.append(dict(sec=s,slot=sl,std=std,prog=prog,house=house,tot=tot,
                         scpm=scpm,pcpm=pcpm,d_sh=d_sh,p_sh=p_sh,cls=classify(d_sh,p_sh),
                         prog_m=prog/MONTHS, prog_rev_m=prog/MONTHS*pcpm/1000))

def uplift_m(prog, pcpm, price, pct):   # monthly incremental
    return (prog/MONTHS)*pct/1000*(price-pcpm)

print("\n"+"="*74); print("CLASSIFICATION (ranked by programmatic impressions)"); print("="*74)
for r in sorted(rows,key=lambda x:-x["prog"]):
    print(f"{r['cls']:9s} {r['sec']:13s} {r['slot']:13s} prog {r['prog']:>11,} "
          f"prog% {r['p_sh']*100:4.0f} dir% {r['d_sh']*100:4.0f} pCPM ${r['pcpm']:4.2f} "
          f"up@8 +${uplift_m(r['prog'],r['pcpm'],PA,0.10):>7,.0f}/mo")

# ---- aggregates by class ----
print("\n"+"="*74); print("SUMMARY BY CLASS"); print("="*74)
for cls in ("TARGET","SECONDARY","RESERVED"):
    grp=[r for r in rows if r["cls"]==cls]
    prog=sum(r["prog"] for r in grp); revm=sum(r["prog_rev_m"] for r in grp)
    print(f"{cls:9s} n={len(grp):2d}  prog imp {prog:>12,} ({prog/gprog*100:4.1f}% of prog)  "
          f"cur rev/mo ${revm:>9,.0f}  blended pCPM ${revm/prog*MONTHS*1000/prog*prog:.2f}" if prog else f"{cls} none")

# ---- TARGET conversion model ----
tgt=[r for r in rows if r["cls"]=="TARGET"]
tprog=sum(r["prog"] for r in tgt); tprog_m=tprog/MONTHS
tcur_m=sum(r["prog_rev_m"] for r in tgt)
print("\n"+"="*74); print(f"TARGET CONVERSION MODEL  ({len(tgt)} slots, period {MONTHS} mo)"); print("="*74)
print(f"target programmatic imp (period) {tprog:,} | /mo {tprog_m:,.0f} | current rev/mo ${tcur_m:,.0f} "
      f"| blended target pCPM ${tcur_m*MONTHS/tprog*1000:.2f}")
for pct in (0.05,0.10,0.15):
    u8=sum(uplift_m(r["prog"],r["pcpm"],PA,pct) for r in tgt)
    u10=sum(uplift_m(r["prog"],r["pcpm"],PB,pct) for r in tgt)
    cur=sum((r["prog"]/MONTHS)*pct*r["pcpm"]/1000 for r in tgt)
    r8=sum((r["prog"]/MONTHS)*pct*PA/1000 for r in tgt)
    r10=sum((r["prog"]/MONTHS)*pct*PB/1000 for r in tgt)
    print(f"  {int(pct*100):>2}%  shift/mo {tprog_m*pct:>10,.0f}  cur ${cur:>8,.0f}  "
          f"@$8 ${r8:>9,.0f} (+${u8:>8,.0f}, ann +${u8*12:>10,.0f})  "
          f"@$10 ${r10:>9,.0f} (+${u10:>8,.0f}, ann +${u10*12:>10,.0f})")
