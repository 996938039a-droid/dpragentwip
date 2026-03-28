"""
agents/handlers/review.py
══════════════════════════
Review screen with all intelligence upgrades:
  E19: Priority ordering — highest-impact assumptions first
  E20: Banker objection prediction
  E21: Ratio benchmarking (your value vs industry norm vs bank minimum)
  E22: Input sensitivity labels (HIGH/MEDIUM/LOW impact)
  G26: Correction tracking display
  G27: Provisional field re-confirmation
  G28: Section confidence scores
"""

import re
from core.session_store import SessionStore
from agents.extractor import extract_json, DEFAULT_MODEL

# ── E21: Ratio benchmarks by industry ─────────────────────────────────────────
RATIO_NORMS = {
    # code: {ratio: (your_label, bank_min, industry_typical, bank_note)}
    "manufacturing_general": {
        "current_ratio": (1.33, 1.5,  "Bank min: 1.33"),
        "dscr":          (1.25, 1.5,  "Bank min: 1.25"),
        "debt_equity":   (None, 3.0,  "Max D/E: 3:1"),
        "tol_tnw":       (None, 4.0,  "Max TOL/TNW: 4:1"),
    },
    "food_processing": {
        "current_ratio": (1.25, 1.5,  "Bank min: 1.25"),
        "dscr":          (1.25, 1.4,  "Bank min: 1.25"),
        "debt_equity":   (None, 3.0,  "Max D/E: 3:1"),
        "tol_tnw":       (None, 4.0,  "Max TOL/TNW: 4:1"),
    },
    "service_it": {
        "current_ratio": (1.2,  1.5,  "Bank min: 1.2"),
        "dscr":          (1.25, 1.5,  "Bank min: 1.25"),
        "debt_equity":   (None, 2.0,  "Service firms: max D/E 2:1"),
        "tol_tnw":       (None, 3.0,  "Max TOL/TNW: 3:1"),
    },
}
# Default if industry not in map
RATIO_NORMS_DEFAULT = {
    "current_ratio": (1.33, 1.5,  "Bank min: 1.33"),
    "dscr":          (1.25, 1.5,  "Bank min: 1.25"),
    "debt_equity":   (None, 3.0,  "Max D/E: 3:1"),
    "tol_tnw":       (None, 4.0,  "Max TOL/TNW: 4:1"),
}

# ── E22: Sensitivity tiers per benchmark field ─────────────────────────────────
# HIGH = 10% change in this field moves DSCR by >0.2 or IRR by >3pp
# MEDIUM = moves DSCR by 0.1-0.2
# LOW = moves DSCR by <0.1
SENSITIVITY = {
    "revenue.price_escalation":    "HIGH",
    "capacity.annual_increment":   "HIGH",
    "capacity.max_util":           "HIGH",
    "opex.power_pct_revenue":      "HIGH",
    "employee.increment":          "MEDIUM",
    "opex.sga_base":               "MEDIUM",
    "opex.marketing_pct_revenue":  "MEDIUM",
    "opex.transport_escalation":   "LOW",
    "opex.misc_escalation":        "LOW",
    "opex.rm_pct_fa":              "LOW",
    "opex.insurance_pct_fa":       "LOW",
    "opex.rm_escalation":          "LOW",
    "opex.insurance_escalation":   "LOW",
    "opex.sga_escalation":         "LOW",
    "material.escalation_generic": "MEDIUM",
    "material.escalation_agri":    "MEDIUM",
    "finance.od_rate":             "LOW",
    "wc.creditor_days_admin":      "LOW",
    "wc.stock_days_fg":            "LOW",
}

# ── E20: Banker objection rules ───────────────────────────────────────────────
def _banker_objections(store: SessionStore) -> list[str]:
    """E20: Predict questions/objections a bank appraiser will raise."""
    objections = []
    cm  = store.capital_means
    rv  = store.revenue_model
    pp  = store.project_profile
    wc  = store.working_capital

    if cm.total_project_cost > 0:
        promoter_pct = cm.promoter_equity / cm.total_project_cost
        if promoter_pct < 0.25:
            objections.append(
                f"📌 **Promoter contribution {promoter_pct*100:.0f}%** — bank will "
                f"ask why equity is below 25%. Prepare justification or increase equity."
            )

    if rv.year1_utilization >= 0.75:
        objections.append(
            f"📌 **Year 1 utilisation {rv.year1_utilization*100:.0f}%** is aggressive. "
            f"Appraiser will scrutinise market demand evidence."
        )

    if wc.implementation_months > 12:
        objections.append(
            f"📌 **{wc.implementation_months}mo implementation** — bank will ask for "
            f"a phased implementation schedule with milestone dates."
        )

    tls = cm.term_loans
    if tls:
        tl = tls[0]
        debt_to_equity = tl.amount_lakhs / cm.promoter_equity if cm.promoter_equity > 0 else 99
        if debt_to_equity > 3:
            objections.append(
                f"📌 **Debt/Equity = {debt_to_equity:.1f}:1** — above the typical 3:1 limit. "
                f"Bank may ask for additional collateral or higher promoter contribution."
            )

    if not pp.city or not pp.state:
        objections.append(
            "📌 **Location not specified** — appraiser will need site visit address "
            "and municipal/panchayat approvals confirmation."
        )

    products_missing_rm = [
        p.name for p in rv.products
        if not any(m.applies_to is None or p.name in (m.applies_to or [])
                   for m in store.cost_structure.raw_materials)
    ]
    if products_missing_rm and store.cost_structure.raw_materials:
        objections.append(
            f"📌 **No RM mapped to: {', '.join(products_missing_rm[:2])}** — "
            f"appraiser will ask for complete cost-of-production working."
        )

    return objections


# ── G28: Section confidence scores ───────────────────────────────────────────
def _compute_confidence(store: SessionStore) -> dict:
    """G28: Score 0-100 for each section based on how much was AI vs user-provided."""
    scores = {}
    pp, cm, rv, cs, mp, wc = (store.project_profile, store.capital_means,
                               store.revenue_model, store.cost_structure,
                               store.manpower, store.working_capital)

    # Profile: all T1 fields directly from user
    profile_fields = [pp.company_name, pp.promoter_name, pp.city, pp.state, pp.operation_start_date]
    scores["profile"] = int(sum(1 for f in profile_fields if f) / len(profile_fields) * 100)

    # Capital: assets + finance sources
    cap_score = 100 if (cm.assets and cm.finance_sources) else (
        50 if cm.assets else (30 if cm.finance_sources else 0))
    scores["capital"] = cap_score

    # Revenue: products + utilization
    scores["revenue"] = 100 if (rv.products and rv.year1_utilization > 0) else (
        60 if rv.products else 0)

    # Costs
    scores["costs"] = 100 if cs.raw_materials else (
        70 if not store.flags.rm_cogs else 0)  # service = no RM needed = 70

    # Manpower
    scores["manpower"] = 100 if mp.categories else 0

    # WC
    wc_fields = [wc.debtor_days > 0, wc.creditor_days_rm > 0,
                 wc.stock_days_rm > 0, wc.implementation_months > 0]
    scores["wc"] = int(sum(wc_fields) / len(wc_fields) * 100)

    # Store back
    for sec, score in scores.items():
        store.set_section_confidence(sec, score)

    return scores


OVERRIDE_PROMPT = """The user wants to change some assumption values.
Message: "{msg}"
Return JSON of fields to update (only fields mentioned, null for others):
{{
  "revenue.price_escalation": null,
  "opex.rm_pct_fa": null,
  "opex.power_pct_revenue": null,
  "opex.sga_base": null,
  "opex.transport_escalation": null,
  "opex.misc_escalation": null,
  "employee.increment": null,
  "finance.od_rate": null,
  "wc.creditor_days_admin": null,
  "wc.stock_days_fg": null,
  "capacity.annual_increment": null,
  "capacity.max_util": null
}}
Values as plain numbers (fractions for percentages: 6% → 0.06)."""


class ReviewHandler:
    def __init__(self, api_key: str, model: str = DEFAULT_MODEL):
        self.api_key = api_key
        self.model   = model

    def build_review_screen(self, store: SessionStore, benchmarks: dict) -> str:
        pp  = store.project_profile
        cm  = store.capital_means
        rv  = store.revenue_model
        cs  = store.cost_structure
        mp  = store.manpower
        wc  = store.working_capital

        # G28: Compute confidence scores
        conf = _compute_confidence(store)

        lines = [
            f"## 📋 DPR Assumption Review",
            f"**{pp.company_name}** · {pp.industry or 'Business'}",
            f"{pp.city}, {pp.state}  ·  Operations from {pp.operation_start_date}  ·  {pp.projection_years} year projection",
            "",
        ]

        # G28: Section confidence summary
        conf_bar = " ".join(
            f"{'✅' if v>=80 else '🟡' if v>=50 else '🔴'} {sec.title()} ({v}%)"
            for sec, v in conf.items()
        )
        lines += [f"**Input completeness:** {conf_bar}", "", "---", ""]

        # E20: Banker objections at the top
        objections = _banker_objections(store)
        if objections:
            lines += ["### ⚠️  Likely Banker Questions", ""]
            lines += objections
            lines += ["", "---", ""]

        # ── T1: User-provided ───────────────────────────────────────────────
        lines += ["### ✅ Your Inputs (T1 — Blue)", "", "**Project & Finance:**"]
        for a in cm.assets:
            lines.append(f"  - {a.name}: ₹{a.cost_lakhs:.1f}L  [{a.category.value}]")
        for src in cm.finance_sources:
            if src.is_term_loan:
                lines.append(
                    f"  - Term Loan: ₹{src.amount_lakhs:.1f}L @ {src.interest_rate*100:.1f}%, "
                    f"{src.tenor_months}mo, {src.moratorium_months}mo moratorium"
                )
            elif src.is_od:
                lines.append(f"  - OD Limit: ₹{src.amount_lakhs:.1f}L")
            elif src.is_equity:
                pct = src.amount_lakhs / cm.total_project_cost * 100 if cm.total_project_cost > 0 else 0
                lines.append(f"  - Promoter Equity: ₹{src.amount_lakhs:.1f}L ({pct:.0f}% of project cost)")

        lines += ["", "**Products:**"]
        for p in rv.products:
            lines.append(
                f"  - {p.name}: ₹{p.price_per_unit:,.0f}/{p.unit}, "
                f"{p.capacity_per_day}/day, {p.split_percent*100:.0f}% split"
            )
        lines.append(
            f"  - Year 1 utilisation: **{rv.year1_utilization*100:.0f}%**  |  "
            f"{rv.working_days_per_month} days/month"
        )

        if cs.raw_materials:
            lines += ["", "**Raw Materials:**"]
            for m in cs.raw_materials:
                scope = f" [applies to: {', '.join(m.applies_to)}]" if m.applies_to else ""
                lines.append(f"  - {m.name}: ₹{m.price_per_unit:.0f}/{m.unit}, {m.input_per_output}/unit{scope}")

        if mp.categories:
            lines += ["", "**Manpower:**"]
            for e in mp.categories:
                lines.append(f"  - {e.designation}: {e.count} × ₹{e.monthly_salary_lakhs:.2f}L/mo")

        lines += [
            "",
            f"**Working Capital:** Debtor {wc.debtor_days}d  |  Creditor {wc.creditor_days_rm}d  "
            f"|  RM Stock {wc.stock_days_rm}d  |  FG Stock {wc.stock_days_fg}d  "
            f"|  Impl {wc.implementation_months}mo",
        ]

        # G26: Show corrections if any
        if store.corrections:
            lines += ["", "**📝 Corrections made:**"]
            for field, data in store.corrections.items():
                lines.append(
                    f"  - {field}: ~~{data['original']}~~ → **{data['revised']}**"
                )

        # G27: Provisional fields needing re-confirmation
        if store.provisional_fields:
            lines += [
                "",
                f"**🔄 Please re-confirm (entered early):** {', '.join(store.provisional_fields)}"
            ]

        # ── T2: Benchmarks — E19 priority order, E22 sensitivity ──────────────
        lines += ["", "---", "", "### 📊 Benchmarked Values (T2 — Green)",
                  "*(Edit: e.g. \"make power 6%, salary increment 5%\")*", ""]

        # E19: Order by sensitivity (HIGH first, then MEDIUM, then LOW)
        ordered_keys = (
            [k for k in SENSITIVITY if SENSITIVITY[k] == "HIGH" and k in benchmarks] +
            [k for k in SENSITIVITY if SENSITIVITY[k] == "MEDIUM" and k in benchmarks] +
            [k for k in SENSITIVITY if SENSITIVITY[k] == "LOW" and k in benchmarks] +
            [k for k in benchmarks if k not in SENSITIVITY]
        )

        sens_label = {"HIGH": "🔴 HIGH impact", "MEDIUM": "🟡 MEDIUM", "LOW": "🟢 LOW"}
        current_tier = None
        for key in ordered_keys:
            bm  = benchmarks.get(key, {})
            val = bm.get("value", "—")
            reason = bm.get("reason", "")
            confidence = bm.get("confidence", "medium")
            rng  = bm.get("range", "")
            sens = SENSITIVITY.get(key, "LOW")

            # E19: Print tier header when tier changes
            if sens != current_tier:
                current_tier = sens
                lines.append(f"\n**{sens_label.get(sens, sens)}:**")

            # Format value
            if isinstance(val, float) and val < 1:
                disp = f"{val*100:.1f}%"
            elif isinstance(val, float):
                disp = f"₹{val:.1f}L"
            else:
                disp = str(val)

            conf_icon = "🔵" if confidence == "high" else "🟢" if confidence == "medium" else "⚪"
            label = key.split(".")[-1].replace("_", " ").title()
            line  = f"- **{label}**: {disp} {conf_icon}  — _{reason}_"
            if rng:
                line += f"\n  _Range: {rng}_"
            lines.append(line)

        # ── E21: Ratio benchmarking ────────────────────────────────────────────
            dscr_result = quick_dscr_estimate(store)
        norms = RATIO_NORMS.get(store.industry_code, RATIO_NORMS_DEFAULT)

        lines += ["", "---", "", "### 📐 Projected Ratios vs Industry Norms", ""]
        lines.append(
            f"| Ratio | Your Projection | Bank Minimum | Industry Typical | Status |"
        )
        lines.append(f"|---|---|---|---|---|")

        # DSCR estimate
        if dscr_result:
            bep  = dscr_result.get("bep_util", 0)
            nmin, ntyp, note = norms.get("dscr", (1.25, 1.5, "Bank min: 1.25"))
                    # rough DSCR from viability check
            bep_data  = estimate_bep_utilization(store)
            tls = cm.term_loans
            if tls and bep_data:
                tl = tls[0]
                mora = tl.moratorium_months
                rep  = tl.tenor_months - mora
                ann_p  = tl.amount_lakhs / (rep/12) if rep > 0 else 0
                ann_i  = tl.amount_lakhs * tl.interest_rate
                ds     = ann_p + ann_i
                y2u    = min(rv.year1_utilization + rv.annual_utilization_increment, rv.max_utilization)
                rev_y2 = bep_data["rev_full"] * y2u
                ncf    = rev_y2 * 0.15 + cm.total_project_cost * 0.12 + ann_i
                est_dscr = ncf / ds if ds > 0 else 0
                status = "✅" if est_dscr >= nmin else "⚠️"
                lines.append(
                    f"| DSCR (Year 2 est.) | ~{est_dscr:.2f} | ≥ {nmin} | ~{ntyp} | {status} |"
                )

        # Current ratio proxy
        wc_req = (wc.debtor_days + wc.stock_days_rm)
        cred   = wc.creditor_days_rm
        cr_proxy = (wc_req / cred) if cred > 0 else 1.5
        nmin, ntyp, note = norms.get("current_ratio", (1.33, 1.5, "Bank min: 1.33"))
        status = "✅" if cr_proxy >= nmin else "⚠️"
        lines.append(
            f"| Current Ratio (est.) | ~{cr_proxy:.1f} | ≥ {nmin} | ~{ntyp} | {status} |"
        )

        # D/E ratio
        if cm.term_loans and cm.promoter_equity > 0:
            de = cm.term_loans[0].amount_lakhs / cm.promoter_equity
            nmin, ntyp, note = norms.get("debt_equity", (None, 3.0, "Max 3:1"))
            status = "✅" if de <= ntyp else "⚠️"
            lines.append(f"| Debt / Equity | {de:.1f}:1 | — | ≤ {ntyp}:1 | {status} |")

        lines += [
            "",
            "_Note: These are pre-generation estimates. Actual ratios will be in the Excel._",
        ]

        # ── T3: Statutory ─────────────────────────────────────────────────────
        lines += [
            "", "---", "", "### ⚙️  Statutory (T3 — Grey — Fixed)",
            "- TL Interest: **9% fixed** · Tenor: **84 months** · Moratorium: **6 months**",
            "- Depreciation: P&M 15%, Civil 10%, Furniture 10% (IT Act WDV)",
            "- Company tax: 25% base + 7% surcharge (if PBT > ₹100L) + 4% HEC",
            "", "---", "",
            "Type **confirm** to generate the DPR Excel, or tell me what to change.",
            "_(e.g. 'make power 8%, set max utilisation 90%')_",
        ]

        return "\n".join(lines)

    async def handle(self, message: str, store: SessionStore,
                     benchmarks: dict) -> tuple[str, bool]:
        msg_lower = message.lower().strip()
        confirm_words = {"confirm", "yes", "ok", "okay", "generate",
                         "looks good", "proceed", "go ahead", "done", "generate dpr"}
        if any(w in msg_lower for w in confirm_words):
            return "✅ Confirmed! Generating your DPR Excel...", True

        d = await extract_json(
            OVERRIDE_PROMPT.format(msg=message),
            self.api_key, fallback={}, model=self.model
        )
        changed = _apply_overrides(store, d, benchmarks)

        if changed:
            # G26: Record corrections
            for field in changed:
                store.record_correction(field, "prev", "updated")
            return (
                f"Updated: **{', '.join(changed)}**.\n\n"
                "Type **confirm** to generate, or make more changes."
            ), False
        return (
            "I didn't catch specific changes. "
            "Type **confirm** to generate, or specify what to change "
            "(e.g. 'make power 6%, salary increment 5%')."
        ), False


def _apply_overrides(store: SessionStore, d: dict, benchmarks: dict = None) -> list[str]:
    """Apply override dict. Also updates benchmarks dict if supplied."""
    changed = []
    cs  = store.cost_structure
    rv  = store.revenue_model
    mp  = store.manpower
    wc  = store.working_capital

    mapping = {
        "revenue.price_escalation":   lambda v: [setattr(p, "price_escalation", v) for p in rv.products],
        "opex.rm_pct_fa":             lambda v: setattr(cs, "rm_pct_of_fa", v),
        "opex.power_pct_revenue":     lambda v: setattr(cs, "power_pct_revenue", v),
        "opex.sga_base":              lambda v: setattr(cs, "sga_base_lakhs", v),
        "opex.transport_escalation":  lambda v: setattr(cs, "transport_escalation", v),
        "opex.misc_escalation":       lambda v: setattr(cs, "misc_escalation", v),
        "employee.increment":         lambda v: [setattr(e, "annual_increment", v) for e in mp.categories],
        "finance.od_rate":            lambda v: [setattr(s, "interest_rate", v)
                                                  for s in store.capital_means.od_sources],
        "wc.creditor_days_admin":     lambda v: setattr(wc, "creditor_days_admin", int(v)),
        "wc.stock_days_fg":           lambda v: setattr(wc, "stock_days_fg", int(v)),
        "capacity.annual_increment":  lambda v: setattr(rv, "annual_utilization_increment", v),
        "capacity.max_util":          lambda v: setattr(rv, "max_utilization", v),
    }

    for key, fn in mapping.items():
        val = d.get(key)
        if val is not None:
            try:
                fn(float(val))
                label = key.split(".")[-1].replace("_", " ")
                changed.append(label)
                # Update benchmarks dict too
                if benchmarks and key in benchmarks:
                    benchmarks[key]["value"]  = float(val)
                    benchmarks[key]["reason"] = "Updated by user"
                    benchmarks[key]["confidence"] = "high"
            except Exception as e:
                print(f"[review override] {key}: {e}")

    return changed
