"""Slot-wise (section x slot) reconciliation + scope check for TIME PG/PD.
Source: two GAM pivots supplied by client (US, top-6 slots) --
  (1) % share of impressions Programmatic/Standard/House by section x slot
  (2) absolute impressions House/Programmatic/Standard by section x slot
This script:
  - encodes pivot (2),
  - verifies each section's slot rows sum to the section total and to the grand total,
  - checks the %-shares in pivot (1) against the absolute numbers (transcription QA),
  - quantifies the scope/period gap vs the prior Jan-May 2026 top-6 section-level dataset.
"""

# (section, slot) -> (House, Programmatic, Standard)  -- from absolute-impressions pivot
DATA = {
 "Entertainment": {
   "stickyfooter1": (2166739, 47410845, 13767927),
   "leaderboard1":  (753998, 11875493, 9918394),
   "inline1":       (396956, 11630907, 5880365),
   "inline2":       (471422, 10920299, 3876484),
   "inline3":       (511614, 10019326, 2914244),
   "rightrail1":    (210342, 4846780, 1998657),
 },
 "Health": {
   "stickyfooter1": (960572, 20185172, 11067220),
   "leaderboard1":  (756373, 7007562, 8665005),
   "inline1":       (414509, 6850100, 4831301),
   "inline2":       (368795, 5874817, 3772312),
   "inline3":       (336651, 4930438, 2530253),
   "rightrail1":    (408881, 4880198, 3808220),
 },
 "U.S.": {
   "stickyfooter1": (502892, 19228301, 5470947),
   "leaderboard1":  (435768, 6609878, 5864010),
   "inline1":       (223459, 4331009, 2054682),
   "inline2":       (247599, 3319301, 1201113),
   "inline3":       (168507, 2359398, 798265),
   "rightrail1":    (181060, 2677384, 1317962),
 },
 "Tech": {
   "stickyfooter1": (145460, 2725359, 1502361),
   "leaderboard1":  (105365, 1551790, 2502857),
   "inline1":       (53511, 1137367, 1245304),
   "inline2":       (54644, 885597, 777475),
   "inline3":       (43319, 673770, 501022),
   "rightrail1":    (46627, 821619, 866947),
 },
 "Science": {
   "stickyfooter1": (98956, 2006324, 723615),
   "leaderboard1":  (69612, 1612613, 1012980),
   "inline1":       (40846, 1280246, 453383),
   "inline2":       (43054, 1046394, 312130),
   "inline3":       (44550, 781024, 229424),
   "rightrail1":    (48350, 571444, 292628),
 },
}

# section totals from pivot (2) for verification
SECTION_TOTALS = {
 "Entertainment": (4511071, 96703650, 38356071),
 "Health":        (3245781, 49728287, 34674311),
 "U.S.":          (1759285, 38525271, 16706979),
 "Tech":          (448926, 7795502, 7395966),
 "Science":       (345368, 7298045, 3024160),
}
GRAND = (10310431, 200050755, 100157487)

print("="*72)
print("STEP 1  -- Verify slot rows sum to section totals")
print("="*72)
gh=gp=gs=0
for sec, slots in DATA.items():
    h=sum(v[0] for v in slots.values()); p=sum(v[1] for v in slots.values()); s=sum(v[2] for v in slots.values())
    th,tp,ts=SECTION_TOTALS[sec]
    ok = (h==th and p==tp and s==ts)
    print(f"{sec:14s} House {h:>11,} (exp {th:>11,}) {'OK' if h==th else 'MISMATCH'} | "
          f"Prog {p:>12,} {'OK' if p==tp else 'MISMATCH'} | Std {s:>11,} {'OK' if s==ts else 'MISMATCH'}")
    gh+=h; gp+=p; gs+=s
print(f"\nGRAND  House {gh:,} (exp {GRAND[0]:,}) {'OK' if gh==GRAND[0] else 'MISMATCH'} | "
      f"Prog {gp:,} (exp {GRAND[1]:,}) {'OK' if gp==GRAND[1] else 'MISMATCH'} | "
      f"Std {gs:,} (exp {GRAND[2]:,}) {'OK' if gs==GRAND[2] else 'MISMATCH'}")

print("\n"+"="*72)
print("STEP 2  -- %-share (pivot 1) vs absolute (pivot 2) transcription QA (spot rows)")
print("="*72)
# expected programmatic % from pivot (1)
EXP_PROG_PCT = {
 ("Entertainment","stickyfooter1"):75,("Entertainment","leaderboard1"):53,("Entertainment","inline1"):65,
 ("Entertainment","inline2"):72,("Entertainment","inline3"):75,("Entertainment","rightrail1"):69,
 ("Health","stickyfooter1"):63,("U.S.","stickyfooter1"):76,("Tech","leaderboard1"):37,("Science","inline2"):75,
}
for (sec,slot),exp in EXP_PROG_PCT.items():
    h,p,s=DATA[sec][slot]; tot=h+p+s; act=round(p/tot*100)
    print(f"{sec:14s} {slot:14s} prog% computed {act:3d}%  pivot {exp:3d}%  {'OK' if abs(act-exp)<=1 else 'CHECK'}")

print("\n"+"="*72)
print("STEP 3  -- Scope / period gap vs prior Jan-May 2026 top-6 section-level dataset")
print("="*72)
PRIOR_PROG_PERIOD = {  # from build_model.py (5 months, Jan-May 2026, US, top-6 slots)
 "Entertainment":22835006,"U.S.":15568972,"Health":10085819,"Science":1875841,"Tech":1462654}
prior_total=sum(PRIOR_PROG_PERIOD.values())
new_total=GRAND[1]
print(f"Prior programmatic (5 sections, Jan-May 2026, 5 mo): {prior_total:,}")
print(f"New  programmatic (same 5 sections, top-6 slots):    {new_total:,}")
print(f"Ratio new/prior: {new_total/prior_total:.2f}x")
print(f"If prior monthly rate held, new total implies ~{new_total/(prior_total/5):.1f} months of data.")
print("Interpretation: new pivot ~ matches a full Jan-2025..Jun-2026 (~18 mo) window,")
print("NOT the 5-month Jan-May 2026 slice used in the section-level doc. Period MUST be confirmed.")
for sec in PRIOR_PROG_PERIOD:
    print(f"  {sec:14s} new {SECTION_TOTALS[sec][1]:>12,}  vs prior-5mo {PRIOR_PROG_PERIOD[sec]:>11,}  "
          f"({SECTION_TOTALS[sec][1]/PRIOR_PROG_PERIOD[sec]:.1f}x)")
