"""
Builds Nativo_SSP_Working_Plan.docx using only the Python stdlib.

Tight, workstream-grouped optimization plan written from the perspective of
the Nativo SSP working team. Each initiative is presented compactly:
  - Action: what we will actually do
  - Audit signal: what we will measure

Output: Nativo_SSP_Working_Plan.docx
"""

import zipfile
from datetime import datetime
from xml.sax.saxutils import escape as xml_escape

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def p(text="", style=None):
    style_xml = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    runs = ""
    if text:
        runs = f'<w:r><w:t xml:space="preserve">{xml_escape(text)}</w:t></w:r>'
    return f"<w:p>{style_xml}{runs}</w:p>"


def p_runs(parts, style=None):
    style_xml = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    runs = []
    for text, bold in parts:
        rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
        runs.append(
            f'<w:r>{rpr}<w:t xml:space="preserve">{xml_escape(text)}</w:t></w:r>'
        )
    return f"<w:p>{style_xml}{''.join(runs)}</w:p>"


def bullet(text):
    return (
        '<w:p><w:pPr><w:pStyle w:val="ListBullet"/>'
        '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
        f'<w:r><w:t xml:space="preserve">{xml_escape(text)}</w:t></w:r></w:p>'
    )


def initiative(num, title, action, audit):
    """Compact 3-line initiative block: title + action + audit signal."""
    out = []
    out.append(p(f"{num}. {title}", style="Heading3"))
    out.append(p_runs([("Action: ", True), (action, False)], style="BodyText"))
    out.append(p_runs([("Audit signal: ", True), (audit, False)], style="BodyText"))
    return "".join(out)


# ----------------------------- content --------------------------------------
# Workstream tuple: (name, status_line, [initiatives])
# Initiative tuple: (num, title, action, audit_signal)

WORKSTREAMS = [
    (
        "Workstream 1: Demand Mismatch and Routing",
        "Current team focus. Reduce wasted QPS and route each impression to "
        "the DSP most likely to clear it at the highest price.",
        [
            (1, "DSP-Specific Traffic Shaping and QPS Alignment",
             "Profile each top DSP on what they actually buy (geo, device, format, viewability); enforce per-DSP QPS caps; rank requests by predicted revenue per DSP and send the top slice within their stated capacity.",
             "Per-DSP bid rate, win rate, no-bid rate, and infra cost per dollar of revenue."),
            (2, "Predicted-Revenue Ranking",
             "Score every request with a predicted-revenue model (geo by placement by signal by DSP history) and prioritise highest-ranked requests when QPS is constrained.",
             "Revenue per request sent, QPS utilization, lift vs unranked baseline."),
            (3, "Win-Rate Feedback Loop",
             "Monitor win rate per DSP and segment in 15-minute windows; auto-throttle send-rate when win rate falls below the DSP-specific threshold; alert account team before a manual blacklist hits.",
             "Win-rate stability, throttle events avoided, recovery time post-throttle."),
            (4, "Bid Loss and Win Rate Analytics",
             "Parse DSP nbr and loss codes weekly; route price losses to the floor team and targeting losses to the signal team; close the loop with a follow-up audit two weeks later.",
             "Loss reason mix shift, win-rate uplift after intervention."),
            (5, "Bid Duplication Detection",
             "Hash impression keys at the edge; flag duplicate fan-out from publisher wrappers; collapse where contractually allowed and escalate where not.",
             "Duplication ratio per publisher, DSP throttle events tied to duplication."),
            (6, "Geo-Pricing Routing",
             "Build routing logic that respects per-DSP geo-pricing curves; route each impression to the highest-paying eligible DSP rather than broadcasting.",
             "Geo-routed CPM uplift, routing-decision accuracy."),
        ],
    ),
    (
        "Workstream 2: Timeouts and Latency",
        "Already in progress. DSPs drop slow paths; small latency wins "
        "translate directly into win-rate gains.",
        [
            (7, "Per-DSP Latency Optimization",
             "Measure end-to-end latency per DSP and geo; trim payload size; tune TCP and TLS reuse; right-size timeouts per DSP capacity rather than using one global value.",
             "p50 and p95 latency per DSP, DSP drop-off rate, win rate vs latency curve."),
            (8, "Late-Bid Acceptance and Bid Caching",
             "Define a 50 to 100 millisecond grace window post-timeout; accept a late bid only if it clears the current top; verify no UX regression on page load.",
             "Late-bid win rate, revenue recovered, page-load impact."),
            (9, "Header Bidding Wrapper and Timeout Audit",
             "Audit top publishers' Prebid configs quarterly; standardise price buckets; upgrade outdated adapters; align timeouts and server-to-server vs client-to-server mix.",
             "Per-publisher CPM lift, Prebid win rate, latency p95."),
        ],
    ),
    (
        "Workstream 3: Ads.txt and Path Hygiene",
        "Already in progress. Cleaner paths build DSP trust and unlock "
        "preferred-path designation.",
        [
            (10, "Ads.txt and Seller.json Automated Hygiene",
             "Move from quarterly manual checks to weekly automated crawl; reconcile publisher ads.txt against seller.json; alert on unauthorised entries; remove stale paths; expose a compliance dashboard.",
             "Percent authorised paths, unauthorised entries detected and resolved, DSP trust scores."),
            (11, "Deal ID Hygiene",
             "Quarterly audit of all active deal IDs; flag zero-spend deals; auto-expire after 90 days inactive; publish a clean deal catalog to buyers.",
             "Active deal count, stale deal percent, buyer deal-catalog feedback."),
            (12, "First-Price Hygiene Audit",
             "Trace one impression end-to-end through every fee and auction hop; check Prebid translators, DSP responses, and any internal re-auction logic; remove any hidden second-price reduction steps.",
             "Win price equals bid price match percentage, fee-leakage dollars."),
        ],
    ),
    (
        "Workstream 4: Performance Monitoring and Audit",
        "Already in progress. Move from periodic batch reporting to "
        "continuous, impression-level monitoring with automated suppression.",
        [
            (13, "Real-Time IVT and GIVT Monitoring",
             "Score IVT at impression level in real time; segment by geo, publisher, and device; integrate IAS, DV, and HUMAN signals; trigger automated suppression when thresholds break.",
             "GIVT and SIVT percent, makegood and clawback dollars, buyer dispute rate."),
            (14, "MFA Detection and Suppression",
             "Score domains using ad density, session duration, content originality, and refresh patterns; integrate Jounce or DV signals; suppress or floor-elevate high-MFA supply.",
             "Percent MFA traffic, post-suppression CPM lift, DSP MFA-block rate."),
            (15, "Bid-Shading Detection and Countermeasures",
             "Monitor DSP bid-distribution histograms for shading patterns; inject 5 to 10 percent stochastic noise around dynamic floors; rotate floor logic variants per auction.",
             "Shading detection rate, post-jitter eCPM lift, win-price vs floor distance."),
            (16, "Consent String Monitoring (TCF v2.2 and GPP)",
             "Validate consent strings at the edge; reject or repair malformed strings; alert publishers on misconfigured CMPs; report consent-pass rate by geo.",
             "Consent-pass rate by geo, demand recovered from validation, CMP misconfigurations resolved."),
        ],
    ),
    (
        "Workstream 5: Whitelisting and Quality Tiering",
        "Already in progress. Tier buyers and supply so that high-quality "
        "demand sees high-quality supply, and the rest is gated.",
        [
            (17, "Buyer Whitelist and Performance Tiering",
             "Score buyers on CPM contribution, win rate, dispute rate, and IVT exposure; tier into premium, standard, and restricted; gate inventory access by tier.",
             "Buyer-tier CPM, low-tier suppression dollars recovered, dispute rate by tier."),
            (18, "Domain and Category Tiering",
             "Maintain a tiered domain list (premium, standard, restricted); monthly review of low-engagement and high-complaint domains; expose category controls to buyers via deal IDs.",
             "CPM by tier, complaint rate, fill-rate impact post-removal."),
            (19, "Quality-Weighted Tie Breaking",
             "When two bids are within 1 percent, break ties by buyer quality score (IVT, dispute, brand-safety history) rather than alphabetic or random.",
             "Tie-break event count, complaint rate post-rule, buyer quality score distribution."),
        ],
    ),
    (
        "Workstream 6: Floor and Signal Hygiene",
        "Adjacent to current work. Foundational levers that compound the "
        "value of every other workstream.",
        [
            (20, "Net-Revenue Floor Optimization",
             "Re-target the floor objective from gross CPM to net revenue after take rate, data costs, and infra; A/B test against the gross-CPM-optimized baseline.",
             "Net margin per impression, gross-vs-net CPM divergence, contribution margin."),
            (21, "Real-Time Bid-Landscape Anchoring",
             "Shift the floor model from a rolling-window historical baseline to in-flight bid density across the last N bids in the segment; cut update latency from hours to minutes.",
             "Floor-to-clearing-price gap, market-shift response time, floor-driven no-bid percent."),
            (22, "Bid Request Signal Enrichment",
             "Audit current OpenRTB payload (user, context, placement, supplychain); add GPID, content object, sua, and dooh signals where relevant; validate per DSP feedback loop.",
             "Bid rate by DSP, CPM uplift per added signal, DSP signal-completeness score."),
            (23, "Cookie-Sync Optimization",
             "Audit current sync partners; rank by match rate and CPM uplift; drop the bottom quartile; trigger syncs only on high-monetisable pages via a page-value score.",
             "Match rate, sync count, page-load impact."),
        ],
    ),
]


# ----------------------------- assemble -------------------------------------

body_parts = []

body_parts.append(p("Nativo SSP Optimization Working Plan", style="Title"))
body_parts.append(p(f"Prepared: {datetime.now().strftime('%B %Y')}", style="Subtitle"))

# Purpose
body_parts.append(p("Purpose", style="Heading1"))
body_parts.append(p(
    "This document is the working plan for the Nativo SSP optimization "
    "team. It captures what we are already running and what we will scale "
    "next, organised by the workstreams the team operates in: demand "
    "mismatch and routing, timeouts and latency, ads.txt and path "
    "hygiene, performance monitoring and audit, whitelisting and quality "
    "tiering, and floor and signal hygiene.",
    style="BodyText"))
body_parts.append(p(
    "Each initiative is written in a compact form: a single Action line "
    "describing what we will do, and a single Audit signal line "
    "describing what we will measure to validate it. The document is "
    "deliberately tight and execution-oriented; longer narrative and "
    "phased sequencing are kept out of scope here.",
    style="BodyText"))

# Workstreams
for name, status, items in WORKSTREAMS:
    body_parts.append(p(name, style="Heading1"))
    body_parts.append(p_runs([("Status: ", True), (status, False)], style="BodyText"))
    for num, title, action, audit in items:
        body_parts.append(initiative(num, title, action, audit))

# Closing audit note
body_parts.append(p("Audit and Validation", style="Heading1"))
body_parts.append(p(
    "Every optimization listed above will be audited at the end of its "
    "cycle against the audit signal stated for that item. The audit "
    "outcome decides whether the change is scaled across more supply or "
    "demand, iterated on, or rolled back. Items with insufficient signal "
    "(low traffic, short observation window, or confounded by parallel "
    "changes) will be re-run rather than declared a result.",
    style="BodyText"))


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
    '<w:docDefaults><w:rPrDefault><w:rPr>'
    '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/>'
    '</w:rPr></w:rPrDefault></w:docDefaults>'
    '<w:style w:type="paragraph" w:styleId="Normal" w:default="1">'
    '<w:name w:val="Normal"/><w:pPr><w:spacing w:after="120"/></w:pPr>'
    '</w:style>'
    '<w:style w:type="paragraph" w:styleId="BodyText">'
    '<w:name w:val="Body Text"/><w:basedOn w:val="Normal"/>'
    '<w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr>'
    '</w:style>'
    '<w:style w:type="paragraph" w:styleId="Title">'
    '<w:name w:val="Title"/><w:basedOn w:val="Normal"/>'
    '<w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr>'
    '<w:rPr><w:b/><w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr>'
    '</w:style>'
    '<w:style w:type="paragraph" w:styleId="Subtitle">'
    '<w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/>'
    '<w:pPr><w:spacing w:before="0" w:after="360"/></w:pPr>'
    '<w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
    '</w:style>'
    '<w:style w:type="paragraph" w:styleId="Heading1">'
    '<w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>'
    '<w:next w:val="BodyText"/>'
    '<w:pPr><w:keepNext/><w:spacing w:before="360" w:after="120"/>'
    '<w:outlineLvl w:val="0"/></w:pPr>'
    '<w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'
    '</w:style>'
    '<w:style w:type="paragraph" w:styleId="Heading2">'
    '<w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>'
    '<w:next w:val="BodyText"/>'
    '<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/>'
    '<w:outlineLvl w:val="1"/></w:pPr>'
    '<w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>'
    '</w:style>'
    '<w:style w:type="paragraph" w:styleId="Heading3">'
    '<w:name w:val="heading 3"/><w:basedOn w:val="Normal"/>'
    '<w:next w:val="BodyText"/>'
    '<w:pPr><w:keepNext/><w:spacing w:before="200" w:after="60"/>'
    '<w:outlineLvl w:val="2"/></w:pPr>'
    '<w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
    '</w:style>'
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


OUT_PATH = "Nativo_SSP_Working_Plan.docx"
with zipfile.ZipFile(OUT_PATH, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types_xml)
    z.writestr("_rels/.rels", root_rels_xml)
    z.writestr("word/_rels/document.xml.rels", doc_rels_xml)
    z.writestr("word/document.xml", document_xml)
    z.writestr("word/styles.xml", styles_xml)
    z.writestr("word/numbering.xml", numbering_xml)

print(f"Wrote {OUT_PATH}")
