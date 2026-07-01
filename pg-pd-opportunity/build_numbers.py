"""Reconcile and print all figures used in the TIME PG/PD resource-justification doc.
US inventory, Jan 2025 - Jun 2026. June 2026 EXCLUDED from revenue/eCPM (mis-booked CPD->CPM).
Definitions: Direct = Standard + Sponsorship ; Programmatic = AdX + Price Priority + Preferred Deal + EBDA(blank) ; House = House.
"""

def cpm(rev, imp):
    return rev / imp * 1000 if imp else 0

# ---------- FY2025 line-item revenue (clean full year) ----------
rev25 = {"Standard":3753960,"Sponsorship":1598339,"AdExchange":224363,
         "PricePriority":646664,"PreferredDeal":8,"EBDA":88959,"House":19902}
imp25 = {"Standard":212041407,"Sponsorship":53065402,"AdExchange":99494854,
         "PricePriority":271400995,"PreferredDeal":1682,"EBDA":41636716,"House":65454948}

direct_rev = rev25["Standard"]+rev25["Sponsorship"]
prog_rev   = rev25["AdExchange"]+rev25["PricePriority"]+rev25["PreferredDeal"]+rev25["EBDA"]
house_rev  = rev25["House"]
tot_rev    = direct_rev+prog_rev+house_rev

direct_imp = imp25["Standard"]+imp25["Sponsorship"]
prog_imp   = imp25["AdExchange"]+imp25["PricePriority"]+imp25["PreferredDeal"]+imp25["EBDA"]
house_imp  = imp25["House"]
tot_imp    = direct_imp+prog_imp+house_imp

print("=== FY2025 (clean) channel summary ===")
for name, r, i in [("Direct",direct_rev,direct_imp),("Programmatic",prog_rev,prog_imp),
                   ("House",house_rev,house_imp),("TOTAL",tot_rev,tot_imp)]:
    print(f"{name:13s} rev ${r:>12,.0f} ({r/tot_rev*100:5.1f}%)  imp {i:>13,.0f} ({i/tot_imp*100:5.1f}%)  eCPM ${cpm(r,i):5.2f}")

# ---------- Full-period impressions (valid; error is revenue-only) ----------
print("\n=== Full-period impressions (Jan25-Jun26, valid) ===")
fp_direct_imp = 314481004+73563380
fp_prog_imp   = 174894634+388317837+98736+56102950
fp_house_imp  = 94427632
fp_tot_imp    = fp_direct_imp+fp_prog_imp+fp_house_imp
for name,i in [("Direct",fp_direct_imp),("Programmatic",fp_prog_imp),("House",fp_house_imp),("TOTAL",fp_tot_imp)]:
    print(f"{name:13s} imp {i:>13,.0f} ({i/fp_tot_imp*100:5.1f}%)")

# ---------- Programmatic 2026 eCPM (clean; error is in direct) ----------
print("\n=== Programmatic eCPM trend ===")
print(f"2025 programmatic eCPM ${cpm(prog_rev,prog_imp):.2f}")
print(f"2026 programmatic eCPM ${cpm(575319,206879910):.2f}")
print(f"Preferred Deal eCPM 2025 $4.50 / 2026 $6.01  (vs prog ~$2.5) -> premium multiple {6.01/cpm(prog_rev,prog_imp):.1f}x")

# ---------- Top programmatic channels (SSPs), full period ----------
print("\n=== Top programmatic partners by revenue (full period) ===")
ssp = [("Media.net",593477),("Ad Exchange",408188),("Amazon/A9",319257),
       ("EBDA (Open Bidding)",130530),("Media.net Bytes",36196),("Magnite",20418),
       ("InfoLinks",13380),("SimpleFeed",13270)]
for n,v in ssp: print(f"{n:22s} ${v:>10,.0f}")
print(f"{'SUM top partners':22s} ${sum(v for _,v in ssp):>10,.0f}  (of $1,535,313 programmatic)")

# ---------- Top-6 ad units, 2025 clean: Standard vs Programmatic eCPM ----------
print("\n=== Top-6 ad units (FY2025 clean): Direct vs Programmatic eCPM ===")
units = [
 ("stickyfooter1",972615,42439067,212295,118862623),
 ("leaderboard1",760160,40740800,81456,45675851),
 ("inline1",338853,17917853,120527,37105159),
 ("inline2",237828,11458682,89609,31827607),
 ("rightrail1",177416,9071443,54306,18854932),
 ("inline3",148420,7401774,69780,25440371),
]
for n,sr,si,pr,pi in units:
    print(f"{n:14s} Direct ${cpm(sr,si):5.2f} ({si/1e6:5.1f}M)   Prog ${cpm(pr,pi):5.2f} ({pi/1e6:5.1f}M)   ratio {cpm(sr,si)/cpm(pr,pi):4.1f}x")
tot_std_imp = sum(u[2] for u in units); tot_std_rev=sum(u[1] for u in units)
tot_prg_imp = sum(u[4] for u in units); tot_prg_rev=sum(u[3] for u in units)
print(f"TOP-6 TOTAL   Direct ${cpm(tot_std_rev,tot_std_imp):.2f} ({tot_std_imp/1e6:.1f}M)   Prog ${cpm(tot_prg_rev,tot_prg_imp):.2f} ({tot_prg_imp/1e6:.1f}M)")

# ---------- Opportunity model: top-6 programmatic pool -> $8/$10 ----------
print("\n=== Opportunity: shifting top-6 programmatic pool to PG/PD ($8/$10) ===")
pool = tot_prg_imp            # 277.8M annual (2025)
base = cpm(tot_prg_rev,tot_prg_imp)  # ~$2.30
print(f"Annual programmatic pool (top-6): {pool:,.0f} imp @ current ${base:.2f} = ${pool*base/1000:,.0f}/yr")
for pct in (0.10,0.20,0.30):
    shifted = pool*pct
    for tgt in (8,10):
        inc = shifted*(tgt-base)/1000
        print(f"  shift {int(pct*100):2d}% ({shifted/1e6:4.1f}M) @ ${tgt} -> +${inc:,.0f}/yr")
# house upside
house6_imp = 14426179
print(f"House (top-6) {house6_imp:,.0f} imp @ $1 -> if PD $5: +${house6_imp*(5-1)/1000:,.0f}/yr ; if $6: +${house6_imp*(6-1)/1000:,.0f}/yr")
