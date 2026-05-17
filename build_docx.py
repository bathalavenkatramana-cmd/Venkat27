"""
Builds Nativo_SSP_Optimization_Roadmap.docx using only the Python stdlib.
A .docx file is a ZIP archive of XML parts; we author the minimum set of
parts required for Word, Pages, and Google Docs to open the file cleanly.

Output: Nativo_SSP_Optimization_Roadmap.docx (black-and-white, no emojis,
Heading 1/2/3 styles, bullet lists, body text).
"""

import zipfile
from datetime import datetime
from xml.sax.saxutils import escape as xml_escape

# ----------------------------- helpers --------------------------------------

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def p(text="", style=None, bold=False):
    """Return a single <w:p> paragraph string."""
    style_xml = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    runs = ""
    if text:
        rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
        runs = (
            f'<w:r>{rpr}<w:t xml:space="preserve">{xml_escape(text)}</w:t></w:r>'
        )
    return f"<w:p>{style_xml}{runs}</w:p>"


def p_runs(parts, style=None):
    """Build a paragraph from a list of (text, bold_bool) tuples."""
    style_xml = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    runs = []
    for text, bold in parts:
        rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
        runs.append(
            f'<w:r>{rpr}<w:t xml:space="preserve">{xml_escape(text)}</w:t></w:r>'
        )
    return f"<w:p>{style_xml}{''.join(runs)}</w:p>"


def bullet(text):
    """Bulleted list item using the ListBullet style we define in styles.xml."""
    return (
        '<w:p><w:pPr><w:pStyle w:val="ListBullet"/>'
        '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
        '</w:pPr>'
        f'<w:r><w:t xml:space="preserve">{xml_escape(text)}</w:t></w:r></w:p>'
    )


def initiative(num, title, what_it_is, how_to_do_it, kpis):
    """Render one numbered initiative as a small block of paragraphs."""
    out = []
    # Title line: "1. Bid-Shading Countermeasures" as Heading 3
    out.append(p(f"{num}. {title}", style="Heading3"))
    out.append(
        p_runs(
            [("What it is: ", True), (what_it_is, False)],
            style="BodyText",
        )
    )
    out.append(
        p_runs(
            [("How to do it: ", True), (how_to_do_it, False)],
            style="BodyText",
        )
    )
    out.append(
        p_runs(
            [("KPIs to watch: ", True), (kpis, False)],
            style="BodyText",
        )
    )
    return "".join(out)


# ----------------------------- content --------------------------------------

# Each section is (Heading 1 title, intro paragraph, [initiatives...])
# Initiative tuple: (number, title, what, how, kpis)

SECTIONS = []

# ---- A. Auction & Floor Mechanics ----
SECTIONS.append((
    "A. Auction and Floor Mechanics",
    "Tighten the core auction so every winning bid clears at the highest "
    "defensible price. These levers are the foundation: every downstream "
    "demand and packaging initiative compounds on top of a clean auction.",
    [
        (1, "Bid-Shading Countermeasures",
         "DSPs shade aggressively against predictable dynamic floors. Add jitter and rotation to break their shading models.",
         "Inject plus or minus 5 to 10 percent stochastic noise around the dynamic floor; rotate floor logic variants per auction; monitor DSP bid-distribution histograms to detect when shading is winning.",
         "Bid-shading detection rate, post-jitter eCPM lift, win-price vs floor distance."),
        (2, "Net-Revenue Floor Optimization",
         "Optimize floors for net revenue after take rate, data costs, and infra, not gross CPM.",
         "Build a margin-aware floor objective function; feed in per-DSP take rate, per-publisher revshare, and per-request infra cost; A/B test against the gross-CPM-optimized baseline.",
         "Net margin per impression, gross-vs-net CPM divergence, contribution margin."),
        (3, "Real-Time Bid-Landscape Anchoring",
         "Anchor floors to live bid density rather than historical CPM averages. Historical anchoring always lags the market.",
         "Shift the floor model from a rolling-window historical baseline to in-flight bid density across the last N bids in the segment; reduce update latency from hours to minutes.",
         "Floor-to-clearing-price gap, market-shift response time, floor-driven no-bid percentage."),
        (4, "Automated Parameter Search",
         "Most SSPs A/B floor strategies, not parameters. Use Bayesian or multi-armed bandit search to tune parameters per segment continuously.",
         "Define the parameter space (jitter percentage, anchor window, segment granularity); run bandit allocation across cells; auto-promote winning parameters.",
         "Parameters tuned per quarter, lift per tuning cycle, exploration vs exploitation ratio."),
        (5, "First-Price Hygiene Audit",
         "Verify the auction is truly first-price end-to-end. Hidden second-price logic anywhere in the stack costs 5 to 15 percent.",
         "Trace one impression end-to-end through every fee and auction hop; check Prebid translators, DSP responses, and internal re-auction logic; remove any reduction steps.",
         "Win price equals bid price match percentage, fee-leakage dollars."),
        (6, "Late-Bid Acceptance and Bid Caching",
         "Accept bids that arrive slightly after timeout if they would win. Recovers 2 to 4 percent of lost revenue.",
         "Define a 50 to 100 millisecond grace window post-timeout; accept only if the bid clears the current top; measure end-user latency impact to avoid UX regression.",
         "Late-bid win rate, revenue recovered, page-load impact."),
        (7, "Quality-Weighted Tie Breaking",
         "When bids are within 1 percent, break ties by buyer quality (low IVT, low dispute rate), not alphabetic or random.",
         "Build a buyer quality score (IVT, dispute, brand-safety history); apply it within the tie-break threshold; document the rules transparently to buyers.",
         "Tie-break event count, complaint rate post-rule, buyer quality score distribution."),
        (8, "Reserve Price Personalization Per Buyer",
         "Buyer-specific floors based on each DSP's historical willingness-to-pay outperform a single global floor curve.",
         "Cluster DSPs by clearing-price elasticity; set per-buyer reserve prices conditioned on segment, geo, and time-of-day; bound deviations to avoid disclosure risk.",
         "Per-buyer eCPM lift, no-bid rate per buyer, blended margin."),
        (9, "Bid Duplication Detection",
         "The same impression hitting multiple paths inflates QPS, erodes DSP trust, and triggers throttling.",
         "Hash impression keys at the edge; flag duplicate request fan-out from publisher wrappers; alert on duplication ratios above a defined threshold; collapse where contractually permitted.",
         "Duplication ratio per publisher, DSP throttle events, infra cost per win."),
    ],
))

# ---- B. Supply Quality & Inventory Hygiene ----
SECTIONS.append((
    "B. Supply Quality and Inventory Hygiene",
    "Buy-side budgets follow clean supply. Continuous supply-quality controls "
    "are now table-stakes for preferred-path designation and for retaining "
    "performance budgets that route around MFA and IVT exposure.",
    [
        (10, "MFA Detection and Suppression",
         "Identify and downrank Made-for-Advertising inventory that the buy side is actively avoiding.",
         "Score domains using ad density, session duration, content originality, and refresh patterns; integrate Jounce and DV signals; suppress or floor-elevate high-MFA supply.",
         "Percent MFA traffic, post-suppression CPM lift, DSP MFA-block rate."),
        (11, "Real-Time IVT and GIVT Monitoring",
         "Continuous monitoring at impression level, not monthly batch reports.",
         "Real-time IVT scoring; geo, publisher, and device segmentation; integrate IAS, DV, and HUMAN signals; automated suppression rules; quarterly buyer-facing transparency reports.",
         "GIVT percent, SIVT percent, makegood and clawback dollars, buyer dispute rate."),
        (12, "Domain and Category Tiering",
         "Curated allow-list with category exclusions to protect auction quality.",
         "Tiered domain list (premium, standard, restricted); monthly review of low-engagement and high-complaint domains; expose category controls to buyers via deal IDs.",
         "CPM by tier, complaint rate, fill-rate impact post-removal."),
        (13, "Ads.txt and Seller.json Hygiene",
         "Ensure every reseller path is authorised and current. Automated, not quarterly manual.",
         "Automated weekly crawl of publisher ads.txt; reconcile against seller.json; alert on unauthorised entries; remove stale paths; publish a compliance dashboard.",
         "Percent authorised paths, DSP trust scores, blocked-path revenue."),
        (14, "Deal ID Hygiene",
         "Most SSPs accumulate hundreds of stale deal IDs. Auditing and culling builds DSP trust.",
         "Quarterly audit of all active deal IDs; flag zero-spend deals; expire after 90 days inactive; publish a deal-catalog dashboard to buyers.",
         "Active deal count, stale deal percent, buyer deal-catalog feedback."),
        (15, "Creative Quality and Brand Suitability Tiering",
         "Beyond IVT, buyers want creative-level quality and brand-suitability tiering at the bid-stream layer.",
         "Classify creatives via category, sensitive-content tags, and historical complaint rate; pass tier signals in bid responses; expose suitability filters to PMP buyers.",
         "Suitability-tier CPM premium, creative complaint rate, buyer brand-safety pass rate."),
        (16, "Automated Makegood and Clawback Workflow",
         "Manual makegood and clawback handling is operationally expensive and slows buyer trust recovery.",
         "Tie IVT and viewability monitoring directly to a workflow engine; auto-issue makegoods and process clawbacks within an SLA; expose status to buyers via portal.",
         "Time-to-resolution, manual ops hours saved, buyer dispute satisfaction."),
    ],
))

# ---- C. Demand Side & Path Optimization ----
SECTIONS.append((
    "C. Demand Side and Path Optimization",
    "Send fewer, smarter requests; route them to the buyer most likely to "
    "clear at the highest net price; and earn preferred-path status with "
    "the DSPs that drive incremental spend.",
    [
        (17, "DSP-Specific Traffic Shaping and QPS Capacity Alignment",
         "Global shaping leaves money on the table. Each DSP has different sweet spots and stated QPS capacity; wasted QPS equals lost spend and triggers throttling.",
         "Profile each top DSP on what they buy (geo, device, format, viewability, audience); train a per-DSP send-decision model; survey and enforce per-DSP QPS caps; allocate based on revenue-per-request rank.",
         "Per-DSP bid rate, per-DSP win rate, QPS efficiency ratio, infra cost per dollar of revenue."),
        (18, "Predicted-Revenue Ranking",
         "Do not just suppress no-bid traffic. Rank by expected revenue and send the top N percent to each DSP within their stated QPS cap.",
         "Train a predicted-revenue model (geo by placement by signal by DSP history); enforce per-DSP QPS budgets; send the highest-ranked requests first.",
         "Revenue per request sent, QPS utilization, DSP no-bid percent."),
        (19, "Win-Rate Feedback Loop",
         "When win rate drops on a segment, the DSP is likely throttling. Pull back proactively before they blacklist.",
         "Monitor win-rate per DSP and segment in 15-minute windows; auto-reduce send-rate when win-rate drops below the DSP-specific threshold; alert the account team.",
         "Win-rate stability, DSP throttle events avoided, segment-level send-rate."),
        (20, "SPO Positioning with Top DSPs",
         "Become a preferred or direct path for top DSPs by reducing duplication and fee layers.",
         "Pull DSP SPO scorecards; identify duplicate reseller chains; deprecate low-margin routes; pitch DSPs with a one-page SPO narrative covering fee stack, latency, quality, and signal richness.",
         "Preferred-path designation count, DSP-level spend share, fee transparency score."),
        (21, "Asymmetric Signal Distribution",
         "Send richer signals to preferred-path DSPs only. Creates a real incentive for SPO designation.",
         "Tier signal payloads by DSP relationship; preferred DSPs get full first-party, attention, and audience overlays; document the tiering as a buyer benefit.",
         "Preferred-DSP CPM uplift vs standard, SPO conversion rate."),
        (22, "Latency Optimization",
         "Reduce auction round-trip time. DSPs drop slow paths from the auction.",
         "Measure end-to-end latency per DSP and geo; trim payload size; co-locate with key DSPs; tune TCP and TLS reuse; right-size timeouts per DSP capacity.",
         "p50 and p95 latency, DSP drop-off rate, win-rate vs latency curve."),
        (23, "Bid Request Signal Enrichment",
         "Improve OpenRTB payload quality. DSPs bid more when signals are richer.",
         "Audit current payload (user, context, placement, supplychain); add GPID, content object, sua, and dooh signals where relevant; validate per DSP feedback loop.",
         "Bid rate by DSP, CPM uplift per added signal, DSP signal-completeness score."),
        (24, "Demand Partner Diversification",
         "Map current DSP coverage against the target universe and close gaps.",
         "Maintain a target list of 40-plus DSPs across retail-media, CTV, EMEA, and APAC; prioritise by estimated spend and integration cost; run a quarterly onboarding cadence.",
         "Number of active DSPs, top-5 concentration percent, fill from new DSPs."),
        (25, "Retail Media DSP Onboarding",
         "Onboard fast-growing retail media DSPs (Criteo RM, CitrusAd, Skai) actively seeking native supply.",
         "Native-format spec alignment; commerce-intent signal exposure; pilot with one anchor advertiser per DSP.",
         "Retail-media revenue percent, new-advertiser count."),
        (26, "International Demand Expansion",
         "Onboard strong EMEA and APAC DSPs to lift non-US CPMs.",
         "Audit non-US fill and CPM gaps; onboard 3 to 5 regional DSPs; localise floors and currency settings.",
         "Non-US eCPM, geo fill rate."),
        (27, "Buyer Whitelist and Performance Tiering",
         "Promote high-quality buyers; restrict low-CPM, low-quality bidders.",
         "Score buyers by CPM contribution, win-rate, dispute rate, and IVT exposure; tier into premium, standard, and restricted; gate inventory access by tier.",
         "Buyer-tier CPM, low-tier suppression dollars recovered."),
        (28, "Minimum Spend and Volume Commitments",
         "Contractual spend commitments from top DSPs and agencies.",
         "Quarterly commitments tied to first-look, preferred-path status, or PMP packages; monthly pacing reviews.",
         "Committed dollars, commitment-vs-delivery percent, churn risk score."),
        (29, "Bid Loss and Win Rate Analytics",
         "Operationalise bid-loss reasons (price, targeting, format, frequency) as a recurring optimisation input.",
         "Parse DSP nbr and loss codes; produce a weekly dashboard per DSP; route price-loss to the floor team and targeting-loss to the signal team.",
         "Loss reason mix shift, win-rate uplift after intervention."),
        (30, "Geo-Pricing Arbitrage and Routing",
         "The same impression can clear higher via different DSP routing. The same impression to a US DSP vs an EU DSP can differ 10 to 20 percent.",
         "Build routing logic that considers DSP geo-pricing curves; route to the highest-paying eligible DSP per impression.",
         "Geo-routed CPM uplift, routing-decision accuracy."),
        (31, "Bid-Stream A/B Testing for Buyers",
         "Let DSPs run controlled experiments on Nativo supply. Sticky differentiator that deepens buy-side integration.",
         "Expose a sandboxed experiment framework to top DSPs (signal variants, floor variants, payload variants); enforce statistical guardrails; share results in standardised reports.",
         "DSP experiments per quarter, post-experiment spend uplift, DSP NPS."),
    ],
))

# ---- D. Signals, Identity & Audience ----
SECTIONS.append((
    "D. Signals, Identity, and Audience Productization",
    "Treat the bid-stream as inventory. Audit what is being given away free, "
    "productize what differentiates Nativo (native adjacency, attention, "
    "first-party content behaviour), and stay current on identity and "
    "privacy frameworks the buy side is steering spend toward.",
    [
        (32, "Signal Productization and Monetization",
         "Audit what is currently given away free in the bid stream (first-party data, attention, sentiment, return-visitor) and productize as paid overlays.",
         "Inventory all signals currently exposed; classify as free, premium, or restricted; build pricing tiers; pitch as audience-uplift over base CPM.",
         "Signal-overlay revenue, premium-signal adoption percent."),
        (33, "Identity Integration (UID2, RampID, ID5)",
         "Support major privacy-safe identity frameworks. DSPs increasingly route spend to authenticated supply.",
         "Integrate UID2.0, LiveRamp RampID, and ID5; pass authenticated IDs where consent permits; publish authenticated-supply percent to buyers.",
         "Authenticated impressions percent, CPM authenticated vs anonymous, DSP match rate."),
        (34, "Cross-Device Identity Graph Integration",
         "A single ID framework caps match rates. Layering a cross-device graph lifts authenticated reach beyond any one provider.",
         "Integrate one or two cross-device graph providers (LiveRamp, Adstra, or equivalent); blend with UID2 and ID5; publish blended match-rate uplift to buyers.",
         "Blended match rate, cross-device CPM premium, authenticated reach uplift."),
        (35, "Contextual Targeting Productization",
         "Productise content adjacency, IAB taxonomy, sentiment, and brand-suitability as a cookieless signal layer.",
         "Classify content via NLP; output IAB v3 plus brand-suitability tiers and sentiment; expose as bid-stream signals and deal-ID overlays.",
         "Contextual deal CPM lift, buyer adoption rate."),
        (36, "First-Party Audience Packaging",
         "Turn content-consumption behaviour into buyable audience segments.",
         "Define segments by article-category affinity, engagement depth, and return frequency; package as deal-ID overlays; price as audience uplift over base CPM.",
         "Segment CPM uplift, segment deal revenue."),
        (37, "Seller-Defined Audiences (SDA)",
         "IAB Tech Lab cookieless standard. Buy side increasingly asking, low effort to support.",
         "Map first-party audience definitions to SDA taxonomy IDs; pass in the bid stream; document buyer-facing as a cookieless-ready signal.",
         "SDA-tagged impressions percent, SDA-deal CPM premium."),
        (38, "Attention and Viewability Measurement",
         "Integrate attention vendors (Adelaide, Lumen) to quantify native's engagement advantage.",
         "Wire vendor SDKs into native placements; surface attention scores to buyers; use as floor-justification in PMP pitches.",
         "Attention score, viewability percent, CPM lift on attention-verified deals."),
        (39, "Privacy Sandbox Readiness (Topics, Protected Audience API)",
         "Chrome's privacy roadmap continues to evolve. Being demonstrably ready protects spend from buyers steering away from non-compliant supply.",
         "Implement Topics API ingestion; pilot Protected Audience API auctions for native; publish a privacy-sandbox readiness scorecard for buyers.",
         "PAAPI win rate, Topics-tagged impressions percent, buyer privacy-readiness score."),
        (40, "Consent String Hygiene (TCF v2.2, GPP)",
         "Malformed or stale consent strings silently lose demand in EU and US state markets. Most SSPs do not actively monitor.",
         "Validate TCF v2.2 and GPP strings at the edge; reject or repair malformed strings; alert publishers to misconfigured CMPs; report consent-pass rate by geo.",
         "Consent-pass rate by geo, demand recovered from validation, CMP misconfiguration alerts resolved."),
    ],
))

# ---- E. PMP, Curation & Direct Deals ----
SECTIONS.append((
    "E. PMP, Curation, and Direct Deals",
    "Convert the cleaner foundation and richer signals into premium-priced "
    "deals and new revenue lines. PMP and curation are where Nativo's native "
    "adjacency story compounds into pricing power.",
    [
        (41, "PMP Packaging on Native Adjacency",
         "Standardised private marketplace packages built around native content adjacency. Nativo's unique signal.",
         "Bundle by vertical (commerce, finance, lifestyle), geo, and audience signal; set deal floors 20 to 40 percent above open auction; pitch agency trading desks directly.",
         "PMP revenue share percent, PMP eCPM vs open auction, deal win-rate."),
        (42, "Curation-as-a-Product",
         "Sell curated inventory bundles to OTHER SSPs and DSPs as a curation layer (OpenPath, ClearLine). New revenue line, no cannibalization.",
         "Define 3 to 5 hero bundles (Nativo Lifestyle Premium, Commerce Intent Native); integrate as curator on partner platforms; co-marketing with curation partners.",
         "Curation revenue, new buyer count, partner platform spend."),
        (43, "Buyer Self-Serve Curation",
         "Let buyers build deal IDs from your taxonomy without sales involvement. Reduces sales cycle friction.",
         "Build a self-serve curation UI exposing IAB taxonomy plus audience segments and supply tiers; auto-generate deal IDs; usage-based pricing.",
         "Self-serve deal count, time-to-deal, sales hours saved."),
        (44, "First-Look and Preferred Deals",
         "First-look deal IDs in exchange for spend commitments from top DSPs.",
         "Identify the top 10 DSPs by spend; structure first-look with floor plus volume commitment; track delivery weekly; auto-fallback to open auction if under-pacing.",
         "First-look fill, committed-vs-actual spend, incremental CPM."),
        (45, "Advertiser Category Demand Mapping",
         "Map which advertiser categories demand which inventory types. Feeds PMP packaging and sales.",
         "Cluster historical bids by advertiser vertical and supply attributes; surface high-demand-low-supply cells; use for PMP creation and sales targeting.",
         "Category fill rate, category CPM, PMP creation velocity."),
        (46, "Programmatic Guaranteed Automation",
         "Most PG deal execution is still manual. Automating it captures upfront-style commitments without the operational drag.",
         "Build PG deal execution into the same engine as PMP; automate forecasting, pacing, and delivery reporting; expose buyer-side and seller-side dashboards.",
         "PG revenue share, PG pacing accuracy, manual ops hours per PG deal."),
        (47, "Inventory Forecasting for PG and Upfronts",
         "PG and upfront commitments require credible forecasts. Without them, sales under-commits or oversells.",
         "Train an inventory-forecasting model (placement, geo, audience, seasonality); expose confidence intervals to sales; feed pacing engine.",
         "Forecast accuracy (MAPE), upfront commit dollars, oversell rate."),
    ],
))

# ---- F. Format Expansion ----
SECTIONS.append((
    "F. Format Expansion",
    "Extend Nativo's adjacency logic into higher-CPM formats and into "
    "formats the buy side is actively budgeting against.",
    [
        (48, "In-Stream Native Video",
         "Sponsored video in editorial context. Extends Nativo's adjacency logic into a higher-CPM format.",
         "Define an in-stream native spec; pilot with 3 to 5 publishers; align with video-buying DSPs; price as native-premium video.",
         "Video native CPM vs display native, publisher adoption count."),
        (49, "Outstream Optimization",
         "Most publishers run outstream suboptimally. Viewability is low, refresh policies are wrong.",
         "Audit outstream placements across top publishers; right-size refresh, viewability, and player size; A/B test optimal configurations.",
         "Outstream viewability percent, outstream CPM lift."),
        (50, "Shoppable and Commerce Native",
         "Adjacent to the retail-media DSP push. Shoppable native (product cards, embedded carts, deep-link CTAs) commands premium CPMs and unlocks performance budgets.",
         "Define shoppable native specs (product feed ingestion, deep-link template, attribution pixels); pilot with commerce publishers and retail-media DSPs; measure end-to-end conversion.",
         "Shoppable CPM premium, click-to-cart rate, retail-media spend share."),
        (51, "Generative AI for Native Creative Variants",
         "Auto-generate headline and image variants per placement to lift CTR. Native is uniquely well-suited because variants must respect editorial context.",
         "Integrate a controlled-generation pipeline (brand-safe prompts, advertiser approval workflow); A/B test variants per placement; cap generation per campaign.",
         "CTR lift per variant, advertiser approval rate, time-to-creative."),
    ],
))

# ---- G. Operational Levers & Publisher Ops ----
SECTIONS.append((
    "G. Operational Levers and Publisher Ops",
    "The unglamorous, high-leverage work. Tighten the plumbing, lock in "
    "supply, and protect margin while the demand-side initiatives mature.",
    [
        (52, "Cookie-Sync Optimization",
         "Drop low-match-rate sync partners; sync only on high-monetizable pages. Reduces page weight and improves match where it counts.",
         "Audit current syncs; rank by match rate and CPM uplift; drop the bottom quartile; trigger syncs based on a page-value score.",
         "Match rate, sync count, page load impact."),
        (53, "Geo-Specific Monetization Gap Analysis",
         "Identify regions where publisher yield trails demand benchmarks.",
         "Benchmark publisher CPM by geo against the network median; route gaps to the floor, DSP, or signal teams; quarterly review.",
         "Geo CPM vs benchmark, gap closure rate."),
        (54, "Publisher Churn Prediction",
         "Predict drop-off before it happens; intervene early.",
         "Train a churn-signal model on impression decline, CPM drop, and complaint spike; segment into high-value, at-risk, and dormant; trigger an account-team playbook.",
         "Churn rate, retention uplift, revenue saved."),
        (55, "Publisher Format Expansion and Upsell",
         "Publishers running only one Nativo format are an upsell opportunity.",
         "Tag publishers by formats live; identify single-format publishers; build a pitch with an incremental revenue estimate; track placement adoption.",
         "Average formats per publisher, incremental revenue per upsell."),
        (56, "Header Bidding Wrapper Audit",
         "Publisher-side Prebid hygiene. Adapter versions, timeouts, price granularity, bid caching.",
         "Quarterly audit of the top 20 publishers' Prebid configs; standardise price buckets; upgrade outdated adapters; align server-to-server vs client-to-server mix.",
         "Per-publisher CPM lift, Prebid win rate, latency p95."),
        (57, "Take Rate and Margin Cube",
         "Net revenue view after take rate, infra, and data costs. Not just gross.",
         "Build a margin cube (publisher by DSP by deal type); rank by net contribution; reallocate optimisation effort to the top-margin combinations.",
         "Gross margin, contribution margin per combination."),
        (58, "Prebid Server / Wrapper-as-a-Service for Mid-Tail",
         "Mid-tail publishers often run weak wrappers, leaking yield and locking Nativo out of the auction. Offering a managed wrapper locks in supply.",
         "Productize a Prebid Server wrapper with sensible defaults, Nativo demand pre-wired, and a self-serve config UI; bundle with onboarding support.",
         "Mid-tail publisher count on managed wrapper, mid-tail CPM lift, Nativo win rate on managed supply."),
        (59, "Yield Management Between Direct-Sold and Programmatic",
         "Most publishers waterfall direct-sold over programmatic, leaving money on the table when programmatic would have cleared higher.",
         "Implement unified decisioning (programmatic competes against direct-sold on a true-value basis); honour direct-sold delivery commitments via pacing constraints.",
         "Yield uplift on contested impressions, direct-sold delivery accuracy, publisher net revenue lift."),
        (60, "Carbon and Sustainability Reporting (Scope3, GreenBids)",
         "Buy-side RFPs increasingly require carbon reporting. Being measurably greener is becoming a procurement gate, not a nice-to-have.",
         "Integrate Scope3 or GreenBids; report grams of CO2 per thousand impressions per path; suppress high-carbon paths or surface as a buyer choice.",
         "gCO2e per kimp, percent of supply with carbon score, sustainability-gated revenue retained."),
    ],
))

# ---- H. Measurement & Attribution ----
SECTIONS.append((
    "H. Measurement and Attribution",
    "Without measurement maturity, every optimization above is a story, not "
    "a result. Measurement infrastructure is a prerequisite for declaring "
    "wins and for the buyer-facing transparency that earns preferred-path "
    "status.",
    [
        (61, "Lift Testing Infrastructure",
         "Proper holdouts for every optimization, not just before-and-after comparisons.",
         "Build a per-impression random-assignment framework; enforce holdout cells for every major optimization; gate rollout on statistical significance.",
         "Percent of optimizations launched via lift test, false-positive rate."),
        (62, "Buyer-Facing Transparency Reports",
         "Quarterly reports to DSPs showing their performance, supply quality, and benchmarks. Builds trust and surfaces fixable issues.",
         "Auto-generate per-DSP reports (CPM, win rate, IVT, latency, supply mix); review with the top 10 DSPs quarterly.",
         "DSP NPS, issue resolution rate, post-report spend growth."),
    ],
))


# ----------------------------- assemble document ----------------------------

body_parts = []

# Title
body_parts.append(p("Nativo SSP Optimization Roadmap", style="Title"))
body_parts.append(p(f"Prepared: {datetime.now().strftime('%B %Y')}", style="Subtitle"))

# Executive summary
body_parts.append(p("Executive Summary", style="Heading1"))
body_parts.append(p(
    "Nativo's structural advantage is native adjacency: editorial context that "
    "the rest of the SSP landscape cannot easily replicate. The roadmap below "
    "is built so that this advantage compounds rather than leaks. It is "
    "organised buy-side-first: tighten the auction, clean the supply, route "
    "demand intelligently, productize the signals, expand into the formats "
    "the buy side is actively budgeting against, harden operations, and "
    "measure everything.",
    style="BodyText"))
body_parts.append(p(
    "The document covers 62 initiatives across 8 sections. Each initiative "
    "follows the same structure: What it is, How to do it, KPIs to watch. "
    "A phased roadmap and a risks-and-dependencies summary close the document.",
    style="BodyText"))
body_parts.append(p(
    "The strategic narrative for leadership: Phase 1 protects and grows "
    "current revenue by removing leakage; Phase 2 converts the cleaner "
    "foundation into premium pricing and new revenue lines (curation, "
    "PG, signal productization); Phase 3 builds defensibility (attention, "
    "Privacy Sandbox readiness, lift-tested optimization, transparency).",
    style="BodyText"))
body_parts.append(p(
    "The tactical narrative for product and engineering: most items map to "
    "existing surfaces (floor engine, traffic shaper, Prebid wrapper, "
    "deal manager) and can be sequenced as parameter and signal changes "
    "before larger systems work (PAAPI, lift-testing infra, unified "
    "decisioning). Sequencing detail is in section I.",
    style="BodyText"))

# How to read
body_parts.append(p("How to Read This Document", style="Heading1"))
body_parts.append(p(
    "Sections A through H are the initiative catalogue, grouped buy-side-first.",
    style="BodyText"))
body_parts.append(bullet("A. Auction and Floor Mechanics"))
body_parts.append(bullet("B. Supply Quality and Inventory Hygiene"))
body_parts.append(bullet("C. Demand Side and Path Optimization"))
body_parts.append(bullet("D. Signals, Identity, and Audience Productization"))
body_parts.append(bullet("E. PMP, Curation, and Direct Deals"))
body_parts.append(bullet("F. Format Expansion"))
body_parts.append(bullet("G. Operational Levers and Publisher Ops"))
body_parts.append(bullet("H. Measurement and Attribution"))
body_parts.append(p(
    "Section I sequences the initiatives across three phases. Section J "
    "summarises risks and dependencies.",
    style="BodyText"))

# Sections A-H
for title, intro, items in SECTIONS:
    body_parts.append(p(title, style="Heading1"))
    body_parts.append(p(intro, style="BodyText"))
    for num, t, what, how, kpis in items:
        body_parts.append(initiative(num, t, what, how, kpis))

# I. Phased Roadmap
body_parts.append(p("I. Phased Roadmap", style="Heading1"))
body_parts.append(p(
    "Suggested phasing balances impact against execution capacity. Final "
    "ordering should be set with engineering and product capacity in mind, "
    "and with measurement infrastructure (61) built early enough to "
    "validate Phase 2 and Phase 3 wins.",
    style="BodyText"))

body_parts.append(p("Phase 1 - Foundation Tuning (Q1 to Q2)", style="Heading2"))
body_parts.append(p(
    "Build directly on the existing dynamic-floor and traffic-shaping "
    "foundation. Quick-win optimizations with measurable lift in the first "
    "one to two quarters.",
    style="BodyText"))
for item in [
    "Bid-shading countermeasures (1)",
    "First-price hygiene audit (5)",
    "Net-revenue floor optimization (2)",
    "DSP-specific traffic shaping and win-rate feedback loop (17, 19)",
    "Bid duplication detection (9)",
    "Ads.txt and seller.json hygiene (13)",
    "MFA suppression and real-time IVT monitoring (10, 11)",
    "Consent string hygiene (40)",
    "Lift testing infrastructure stood up in parallel (61)",
]:
    body_parts.append(bullet(item))

body_parts.append(p("Phase 2 - Demand and Packaging (Q2 to Q3)", style="Heading2"))
body_parts.append(p(
    "Convert the cleaner, better-tuned foundation into premium CPM and "
    "new revenue lines.",
    style="BodyText"))
for item in [
    "SPO positioning with top DSPs and asymmetric signal distribution (20, 21)",
    "PMP packaging on native adjacency (41)",
    "Curation-as-a-product launch (42)",
    "Signal productization and first-party audience packaging (32, 36)",
    "Identity integration (UID2, RampID, ID5) and SDA (33, 37)",
    "Retail media DSP onboarding (25)",
    "Programmatic guaranteed automation (46)",
    "Buyer-facing transparency reports v1 (62)",
]:
    body_parts.append(bullet(item))

body_parts.append(p("Phase 3 - Compound Levers and Expansion (Q3 to Q4)", style="Heading2"))
body_parts.append(p(
    "Slower-burn, defensibility-building work that compounds over multiple "
    "cycles.",
    style="BodyText"))
for item in [
    "Contextual targeting productization (35)",
    "Attention measurement integration (38)",
    "In-stream native video and shoppable native (48, 50)",
    "Generative AI for native creative variants (51)",
    "Publisher churn prediction (54)",
    "Privacy Sandbox readiness (39)",
    "Cross-device identity graph integration (34)",
    "Carbon and sustainability reporting (60)",
    "Yield management between direct-sold and programmatic (59)",
    "Bid-stream A/B testing for buyers (31)",
]:
    body_parts.append(bullet(item))

# J. Risks & Dependencies
body_parts.append(p("J. Risks and Dependencies", style="Heading1"))
body_parts.append(bullet(
    "Engineering capacity. Bid-shading countermeasures and per-DSP shaping "
    "require ML and data-engineering bandwidth; sequence with product."))
body_parts.append(bullet(
    "DSP relationships. SPO positioning and asymmetric signals require "
    "buy-side conversations; start outreach in parallel with technical work."))
body_parts.append(bullet(
    "Publisher buy-in. Wrapper audits, format expansion, and unified "
    "decisioning require publisher cooperation; lead with revenue-positive "
    "case studies."))
body_parts.append(bullet(
    "Measurement maturity. Proper lift testing is a prerequisite for "
    "declaring wins; build the infra before launching large optimizations."))
body_parts.append(bullet(
    "Privacy and consent. Identity, audience packaging, signal "
    "monetization, and consent-string hygiene all depend on a clean "
    "consent posture; align with legal and privacy early."))
body_parts.append(bullet(
    "Buyer trust. Transparency reports and creative-suitability tiering "
    "amplify wins from supply quality and IVT work; without them, the "
    "underlying improvements are invisible to the buy side."))

body = "".join(body_parts)


# ----------------------------- xml parts ------------------------------------

document_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:document {W_NS}>'
    f'<w:body>{body}'
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
    '</w:body></w:document>'
)

styles_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:styles {W_NS}>'
    # Default doc-wide font
    '<w:docDefaults><w:rPrDefault><w:rPr>'
    '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/>'
    '</w:rPr></w:rPrDefault></w:docDefaults>'
    # Normal
    '<w:style w:type="paragraph" w:styleId="Normal" w:default="1">'
    '<w:name w:val="Normal"/><w:pPr><w:spacing w:after="120"/></w:pPr>'
    '</w:style>'
    # Body Text
    '<w:style w:type="paragraph" w:styleId="BodyText">'
    '<w:name w:val="Body Text"/><w:basedOn w:val="Normal"/>'
    '<w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr>'
    '</w:style>'
    # Title
    '<w:style w:type="paragraph" w:styleId="Title">'
    '<w:name w:val="Title"/><w:basedOn w:val="Normal"/>'
    '<w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr>'
    '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
    '<w:b/><w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr>'
    '</w:style>'
    # Subtitle
    '<w:style w:type="paragraph" w:styleId="Subtitle">'
    '<w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/>'
    '<w:pPr><w:spacing w:before="0" w:after="360"/></w:pPr>'
    '<w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
    '</w:style>'
    # Heading 1
    '<w:style w:type="paragraph" w:styleId="Heading1">'
    '<w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>'
    '<w:next w:val="BodyText"/>'
    '<w:pPr><w:keepNext/><w:spacing w:before="360" w:after="120"/>'
    '<w:outlineLvl w:val="0"/></w:pPr>'
    '<w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'
    '</w:style>'
    # Heading 2
    '<w:style w:type="paragraph" w:styleId="Heading2">'
    '<w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>'
    '<w:next w:val="BodyText"/>'
    '<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/>'
    '<w:outlineLvl w:val="1"/></w:pPr>'
    '<w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>'
    '</w:style>'
    # Heading 3
    '<w:style w:type="paragraph" w:styleId="Heading3">'
    '<w:name w:val="heading 3"/><w:basedOn w:val="Normal"/>'
    '<w:next w:val="BodyText"/>'
    '<w:pPr><w:keepNext/><w:spacing w:before="200" w:after="60"/>'
    '<w:outlineLvl w:val="2"/></w:pPr>'
    '<w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
    '</w:style>'
    # ListBullet
    '<w:style w:type="paragraph" w:styleId="ListBullet">'
    '<w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/>'
    '<w:pPr><w:spacing w:after="60"/><w:ind w:left="360" w:hanging="360"/></w:pPr>'
    '</w:style>'
    '</w:styles>'
)

numbering_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:numbering {W_NS}>'
    '<w:abstractNum w:abstractNumId="0">'
    '<w:lvl w:ilvl="0">'
    '<w:start w:val="1"/><w:numFmt w:val="bullet"/>'
    '<w:lvlText w:val="-"/><w:lvlJc w:val="left"/>'
    '<w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr>'
    '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>'
    '</w:lvl>'
    '</w:abstractNum>'
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
    '</w:numbering>'
)

content_types_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/word/numbering.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
    '</Types>'
)

root_rels_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="word/document.xml"/>'
    '</Relationships>'
)

doc_rels_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
    'Target="styles.xml"/>'
    '<Relationship Id="rId2" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" '
    'Target="numbering.xml"/>'
    '</Relationships>'
)

# ----------------------------- write zip ------------------------------------

OUT_PATH = "Nativo_SSP_Optimization_Roadmap.docx"
with zipfile.ZipFile(OUT_PATH, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types_xml)
    z.writestr("_rels/.rels", root_rels_xml)
    z.writestr("word/_rels/document.xml.rels", doc_rels_xml)
    z.writestr("word/document.xml", document_xml)
    z.writestr("word/styles.xml", styles_xml)
    z.writestr("word/numbering.xml", numbering_xml)

print(f"Wrote {OUT_PATH}")
