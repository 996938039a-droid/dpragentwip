"""
agents/validators.py
═════════════════════
A1  Cross-field sanity checks (funding balance, totals)
A2  Unit normalization (crore/thousand → lakhs, % → fraction)
A3  Per-field plausibility checks (utilization, debtor days, etc.)
A4  Contradiction detection (cash-and-carry vs debtor days, etc.)
A5  Implied value capture ("pure service", "already own land")

Returns structured warnings that handlers append to their responses.
"""

import re
from core.session_store import SessionStore


# ─── A2: Unit Normalizer ──────────────────────────────────────────────────────

def normalize_currency(text: str) -> str:
    """
    Pre-process user text: convert crore/thousand/lakh mentions to plain lakhs.
    Returns cleaned text that the LLM extractor can parse more reliably.
    Examples:
      "₹2 crore" → "200 lakhs"
      "₹40,000/month" → "0.40 lakhs/month"
      "40 thousand" → "0.40 lakhs"
    """
    t = text

    # Remove commas in numbers
    t = re.sub(r'(\d),(\d)', r'\1\2', t)

    # crore → lakhs  (1 crore = 100 lakhs)
    def crore_sub(m):
        val = float(m.group(1))
        return f"{val * 100:.2f} lakhs"
    t = re.sub(r'(\d+(?:\.\d+)?)\s*(?:crore|cr\.?)', crore_sub, t, flags=re.IGNORECASE)

    # lakh/lac → keep as is (already in lakhs)
    t = re.sub(r'(\d+(?:\.\d+)?)\s*(?:lakh|lac|l\.?)\b', r'\1 lakhs', t, flags=re.IGNORECASE)

    # thousand → divide by 100 to get lakhs
    def thousand_sub(m):
        val = float(m.group(1))
        return f"{val / 100:.4f} lakhs"
    t = re.sub(r'(\d+(?:\.\d+)?)\s*(?:thousand|k)\b', thousand_sub, t, flags=re.IGNORECASE)

    return t


def normalize_rate(value: float) -> float:
    """
    If a rate looks like it was entered as a percentage (e.g. 9.5 instead of 0.095),
    convert it. Heuristic: if value > 1.0, assume it's a percentage.
    """
    if value > 1.0:
        return value / 100.0
    return value


# ─── A1: Cross-field Sanity Checks ───────────────────────────────────────────

def check_funding_balance(store: SessionStore) -> list[str]:
    """
    Check that total finance sources = total project cost.
    Returns list of warning strings (empty = all OK).
    """
    warnings = []
    cm = store.capital_means
    if not cm.assets or not cm.finance_sources:
        return warnings

    project_cost = cm.total_project_cost
    # Exclude OD from balance check (OD is working capital, not project finance)
    project_finance = sum(s.amount_lakhs for s in cm.finance_sources if not s.is_od)
    gap = project_finance - project_cost

    if abs(gap) > 0.5:
        if gap > 0:
            warnings.append(
                f"⚠️  **Funding surplus of ₹{gap:.1f}L**: "
                f"Finance sources (₹{project_finance:.1f}L) exceed project cost (₹{project_cost:.1f}L). "
                f"Either reduce loan/equity or add more assets."
            )
        else:
            warnings.append(
                f"⚠️  **Funding gap of ₹{abs(gap):.1f}L**: "
                f"Project cost (₹{project_cost:.1f}L) exceeds finance sources (₹{project_finance:.1f}L). "
                f"Please increase term loan, OD limit, or promoter equity."
            )
    return warnings


# ─── A3: Per-field Plausibility Checks ────────────────────────────────────────

def check_plausibility(store: SessionStore) -> list[str]:
    """
    Check individual field values for plausibility.
    Returns list of warning strings.
    """
    warnings = []
    rv  = store.revenue_model
    wc  = store.working_capital
    cs  = store.cost_structure
    cm  = store.capital_means

    # Year 1 utilization
    if rv.year1_utilization > 0.90:
        warnings.append(
            f"⚠️  **Year 1 utilisation {rv.year1_utilization*100:.0f}% is very optimistic**. "
            f"Banks typically expect 50–60% in Year 1 for a new unit. "
            f"Consider revising to 60–70% to maintain credibility."
        )
    if rv.year1_utilization < 0.30:
        warnings.append(
            f"⚠️  **Year 1 utilisation {rv.year1_utilization*100:.0f}% is very low**. "
            f"This implies low revenue and likely a first-year loss. "
            f"Confirm this is intentional."
        )

    # Debtor days
    if wc.debtor_days > 120:
        warnings.append(
            f"⚠️  **Debtor days = {wc.debtor_days} is very high**. "
            f"This locks up significant cash in receivables. "
            f"Typical range for this industry is 30–60 days."
        )

    # Creditor days vs debtor days
    if wc.creditor_days_rm > 0 and wc.debtor_days > 0:
        if wc.debtor_days > wc.creditor_days_rm * 3:
            warnings.append(
                f"⚠️  **Debtors ({wc.debtor_days}d) >> Creditors ({wc.creditor_days_rm}d)**: "
                f"Large WC gap will require significant funding. Verify these numbers."
            )

    # Working days per month
    if rv.working_days_per_month > 30:
        warnings.append(
            f"⚠️  **Working days/month = {rv.working_days_per_month}**: "
            f"Maximum is 30. Please correct."
        )
    if rv.working_days_per_month < 20:
        warnings.append(
            f"ℹ️  **Working days/month = {rv.working_days_per_month}**: "
            f"That's less than 20 days. Is this intentional (e.g. seasonal business)?"
        )

    # Product split percentages
    if len(rv.products) > 1:
        total_split = sum(p.split_percent for p in rv.products)
        if abs(total_split - 1.0) > 0.05:
            warnings.append(
                f"⚠️  **Product splits sum to {total_split*100:.0f}%** (should be 100%). "
                f"Please adjust product capacity splits."
            )

    # Power cost anomaly
    if cs.power_pct_revenue > 0.20:
        warnings.append(
            f"⚠️  **Power cost = {cs.power_pct_revenue*100:.0f}% of revenue** seems very high. "
            f"Typical range is 3–12%. Please verify."
        )

    # Implementation period
    if wc.implementation_months > 24:
        warnings.append(
            f"⚠️  **Implementation period = {wc.implementation_months} months** is unusually long. "
            f"Typical range is 6–18 months. Confirm this is correct."
        )

    # Promoter equity check (A17 — equity adequacy)
    if cm.assets and cm.finance_sources:
        promoter_pct = cm.promoter_equity / cm.total_project_cost if cm.total_project_cost > 0 else 0
        if 0 < promoter_pct < 0.25:
            warnings.append(
                f"⚠️  **Promoter equity is only {promoter_pct*100:.0f}% of project cost** "
                f"(₹{cm.promoter_equity:.1f}L of ₹{cm.total_project_cost:.1f}L). "
                f"Most banks require at least 25–30%. Consider increasing contribution."
            )

    return warnings


# ─── A4: Contradiction Detection ─────────────────────────────────────────────

def check_contradictions(store: SessionStore) -> list[str]:
    """
    Detect logical contradictions across fields.
    Returns list of warning strings.
    """
    warnings = []
    wc  = store.working_capital
    rv  = store.revenue_model
    cs  = store.cost_structure
    f   = store.flags

    # Cash-and-carry but debtor days set
    if wc.debtor_days == 0 and f.trade_debtors == 1:
        pass  # consistent — no debtors, flag also off

    # Service industry but RM entered
    if not f.rm_cogs and cs.raw_materials:
        warnings.append(
            f"ℹ️  **{len(cs.raw_materials)} raw material(s) entered** but industry is "
            f"flagged as service/non-manufacturing. "
            f"If these are direct cost inputs, that's fine — otherwise remove them."
        )

    # High utilization + very long moratorium
    tls = store.capital_means.term_loans
    if tls and rv.year1_utilization >= 0.70:
        mora = tls[0].moratorium_months
        if mora >= 18:
            warnings.append(
                f"ℹ️  **Year 1 utilisation = {rv.year1_utilization*100:.0f}%** but moratorium "
                f"is {mora} months. If operations start early, moratorium may be shorter "
                f"than entered. Verify with the bank."
            )

    return warnings


# ─── A5: Implied Value Capture ────────────────────────────────────────────────

def detect_implied_flags(message: str, store: SessionStore) -> list[str]:
    """
    Scan for phrases that imply specific values — capture them automatically.
    Returns list of auto-set messages for display.
    """
    auto_set = []
    text = message.lower()

    # "pure service" / "no inventory" → zero all stock days
    if any(p in text for p in ["pure service", "no inventory", "no stock", "no raw material"]):
        if store.working_capital.stock_days_rm != 0:
            store.working_capital.stock_days_rm = 0
            store.working_capital.stock_days_fg = 0
            auto_set.append("stock days set to 0 (pure service)")

    # "cash business" / "cash and carry" → zero debtor days
    if any(p in text for p in ["cash and carry", "cash business", "cash only",
                                "no credit", "no debtors", "immediate payment"]):
        if store.working_capital.debtor_days != 0:
            store.working_capital.debtor_days = 0
            auto_set.append("debtor days set to 0 (cash business)")

    # "already own land" → note, but can't reduce civil works automatically
    if any(p in text for p in ["already own", "own the land", "land is mine",
                                "land already", "own land"]):
        auto_set.append("noted: land is owned — ensure Civil Works cost excludes land value")

    # "no transport" / "local delivery only"
    if any(p in text for p in ["no transport", "local only", "self pickup"]):
        store.set_flag("transport", 0)
        auto_set.append("transport cost flag disabled")

    return auto_set


# ─── C12: Working Capital Gap ────────────────────────────────────────────────

def check_wc_gap(store: SessionStore) -> list[str]:
    """
    Compare estimated WC requirement to OD limit.
    Rough estimate: (debtors + stock - creditors) as % of annual revenue.
    """
    warnings = []
    wc  = store.working_capital
    rv  = store.revenue_model
    cs  = store.cost_structure
    cm  = store.capital_means

    if not rv.products or rv.year1_utilization <= 0:
        return warnings

    # Rough revenue estimate
    rev = sum(
        p.price_per_unit * p.capacity_per_day * rv.working_days_per_month
        * 12 * rv.year1_utilization * p.split_percent
        for p in rv.products
    ) / 100000

    cogs = rev * 0.4  # rough estimate if no RM data yet
    if cs.raw_materials:
        cogs = sum(
            m.price_per_unit * m.input_per_output
            * (sum(p.capacity_per_day * rv.working_days_per_month * 12
                   * rv.year1_utilization * p.split_percent for p in rv.products
                   if m.applies_to is None or p.name in (m.applies_to or [])))
            for m in cs.raw_materials
        ) / 100000

    debtors   = rev / 365 * max(wc.debtor_days, 0)
    stock     = cogs / 365 * max(wc.stock_days_rm, 0)
    creditors = cogs / 365 * max(wc.creditor_days_rm, 0)
    wc_req    = debtors + stock - creditors

    od_limit = wc.wc_loan_amount
    if wc_req > 0 and od_limit > 0:
        gap = wc_req - od_limit
        if gap > od_limit * 0.25:
            warnings.append(
                f"⚠️  **WC gap detected**: Estimated WC requirement ≈ ₹{wc_req:.1f}L "
                f"but OD limit is ₹{od_limit:.1f}L (shortfall ≈ ₹{gap:.1f}L). "
                f"Either increase the OD limit or tighten debtor/stock days."
            )

    return warnings


# ─── C13: Break-even utilization ─────────────────────────────────────────────

def estimate_bep_utilization(store: SessionStore) -> dict | None:
    """
    Rough BEP utilization estimate before full generation.
    Returns dict with bep_util, year1_util, warning (or None if can't compute).
    """
    rv  = store.revenue_model
    cs  = store.cost_structure
    cm  = store.capital_means

    if not rv.products or not cm.finance_sources:
        return None

    # Revenue at 100% utilization
    rev_full = sum(
        p.price_per_unit * p.capacity_per_day * rv.working_days_per_month
        * 12 * p.split_percent
        for p in rv.products
    ) / 100000

    if rev_full <= 0:
        return None

    # Fixed costs (rough)
    nfa_est  = cm.total_project_cost
    manpower = sum(e.count * e.monthly_salary_lakhs * 12
                   for e in store.manpower.categories) if store.manpower.categories else 0
    sga      = cs.sga_base_lakhs
    rm_m     = cs.rm_pct_of_fa * nfa_est
    ins      = cs.insurance_pct_of_fa * nfa_est
    depr     = nfa_est * 0.12  # blended WDV estimate
    tl_int   = sum(s.amount_lakhs * s.interest_rate
                   for s in cm.finance_sources if s.is_term_loan)
    fixed    = manpower + sga + rm_m + ins + depr + tl_int + cs.transport_base_lakhs

    # Variable costs at full utilization
    cogs_full = 0
    if cs.raw_materials:
        cogs_full = sum(
            m.price_per_unit * m.input_per_output
            * sum(p.capacity_per_day * rv.working_days_per_month * 12 * p.split_percent
                  for p in rv.products
                  if m.applies_to is None or p.name in (m.applies_to or []))
            for m in cs.raw_materials
        ) / 100000
    pwr = cs.power_pct_revenue * rev_full
    mkt = cs.marketing_pct_revenue * rev_full
    var  = cogs_full + pwr + mkt

    # Contribution margin ratio
    cm_ratio = (rev_full - var) / rev_full if rev_full > var else 0.01

    # BEP = Fixed / Contribution per unit of revenue
    bep_util = fixed / (rev_full * cm_ratio) if cm_ratio > 0 else None

    if bep_util is None or bep_util <= 0:
        return None

    return {
        "bep_util":   min(bep_util, 2.0),
        "year1_util": rv.year1_utilization,
        "rev_full":   rev_full,
        "fixed":      fixed,
    }


def bep_warning(store: SessionStore) -> list[str]:
    """Return warnings about BEP vs Year 1 utilization."""
    warnings = []
    result = estimate_bep_utilization(store)
    if result is None:
        return warnings
    bep   = result["bep_util"]
    y1    = result["year1_util"]
    if bep > y1:
        warnings.append(
            f"⚠️  **Loss expected in Year 1**: Break-even utilisation ≈ {bep*100:.0f}% "
            f"but Year 1 is set at {y1*100:.0f}%. "
            f"This is fine if you've planned for it — the bank will want to see "
            f"when you turn cash-positive."
        )
    elif bep > 0.7:
        warnings.append(
            f"ℹ️  **BEP at {bep*100:.0f}% utilisation** — relatively high. "
            f"The project has limited buffer if revenue underperforms."
        )
    return warnings


# ─── C14: Moratorium vs Implementation ───────────────────────────────────────

def check_moratorium_vs_impl(store: SessionStore) -> list[str]:
    """Check that moratorium ≥ implementation period."""
    warnings = []
    wc  = store.working_capital
    tls = store.capital_means.term_loans
    if not tls or wc.implementation_months < 0:
        return warnings
    mora = tls[0].moratorium_months
    impl = wc.implementation_months
    if impl > mora:
        warnings.append(
            f"⚠️  **Moratorium ({mora}mo) < Implementation period ({impl}mo)**: "
            f"TL principal repayment will start {impl - mora} months before operations begin. "
            f"Consider requesting moratorium = {impl} months to avoid pre-operational cash strain."
        )
    return warnings


# ─── C11: Quick DSCR estimate ─────────────────────────────────────────────────

def quick_dscr_estimate(store: SessionStore) -> list[str]:
    """
    Back-of-envelope DSCR for Year 2 (first full operating year post-mora).
    Shows a warning if it looks unfeasible before generating the full model.
    """
    warnings = []
    bep = estimate_bep_utilization(store)
    if bep is None:
        return warnings

    rv  = store.revenue_model
    cm  = store.capital_means
    tls = cm.term_loans
    if not tls:
        return warnings

    tl   = tls[0]
    mora = tl.moratorium_months
    repayment_months = tl.tenor_months - mora
    if repayment_months <= 0:
        return warnings

    annual_principal = tl.amount_lakhs / (repayment_months / 12) if repayment_months > 0 else 0
    annual_interest  = tl.amount_lakhs * tl.interest_rate
    debt_service     = annual_principal + annual_interest

    # Estimate Year 2 revenue and PAT
    y2_util = min(rv.year1_utilization + rv.annual_utilization_increment, rv.max_utilization)
    rev_y2  = bep["rev_full"] * y2_util
    fixed   = bep["fixed"]
    cogs_y2 = (bep["rev_full"] * y2_util) * (1 - 0.40)  # rough
    pat_y2  = rev_y2 - fixed - cogs_y2
    depr    = cm.total_project_cost * 0.12
    ncf_y2  = pat_y2 + depr + annual_interest   # numerator

    if debt_service > 0:
        dscr = ncf_y2 / debt_service
        if dscr < 1.0:
            warnings.append(
                f"⚠️  **Estimated DSCR ≈ {dscr:.2f}** (target ≥ 1.25): "
                f"At current assumptions, the project may not service its debt comfortably. "
                f"Review revenue projections, loan amount, or cost structure before generating."
            )
        elif dscr < 1.25:
            warnings.append(
                f"ℹ️  **Estimated DSCR ≈ {dscr:.2f}** — just above the minimum (1.25). "
                f"The margin is thin. Banks may ask for additional security or higher equity."
            )
    return warnings


# ─── D16: Entity type inference ──────────────────────────────────────────────

def infer_entity_type(message: str, store: SessionStore) -> str | None:
    """
    Infer entity type from natural language cues.
    Returns the inferred EntityType value string, or None.
    """
    from core.session_store import EntityType
    text = message.lower()
    hints = {
        EntityType.PROPRIETORSHIP: ["proprietor", "sole owner", "self-employed",
                                     "individual", "own business", "my shop"],
        EntityType.PARTNERSHIP:    ["partner", "partnership", "two of us",
                                     "my brother and i", "joint"],
        EntityType.LLP:            ["llp", "limited liability partnership"],
        EntityType.COMPANY:        ["pvt ltd", "private limited", "company",
                                     "incorporated", "director"],
    }
    for entity_type, keywords in hints.items():
        if any(kw in text for kw in keywords):
            return entity_type.value
    return None


# ─── D18: Duplicate detection ────────────────────────────────────────────────

def find_duplicates(store: SessionStore) -> list[str]:
    """Check for near-duplicate product or material names."""
    warnings = []

    def similar(a: str, b: str) -> bool:
        a, b = a.lower().strip(), b.lower().strip()
        if a == b:
            return True
        # Check if one string is a prefix substring of the other
        shorter, longer = (a, b) if len(a) <= len(b) else (b, a)
        if len(shorter) >= 4 and shorter in longer:
            return True
        # Check word overlap: if >60% of words in shorter are in longer
        words_s = set(shorter.split())
        words_l = set(longer.split())
        if len(words_s) >= 2:
            overlap = len(words_s & words_l) / len(words_s)
            if overlap >= 0.6:
                return True
        return False

    prods = store.revenue_model.products
    for i in range(len(prods)):
        for j in range(i + 1, len(prods)):
            if similar(prods[i].name, prods[j].name):
                warnings.append(
                    f"⚠️  **Possible duplicate products**: '{prods[i].name}' and '{prods[j].name}'. "
                    f"If these are the same product, remove one."
                )

    mats = store.cost_structure.raw_materials
    for i in range(len(mats)):
        for j in range(i + 1, len(mats)):
            if similar(mats[i].name, mats[j].name):
                warnings.append(
                    f"⚠️  **Possible duplicate materials**: '{mats[i].name}' and '{mats[j].name}'. "
                    f"If these are the same input, remove one."
                )

    return warnings


# ─── Master validator ─────────────────────────────────────────────────────────

def run_all_validators(store: SessionStore, message: str = "") -> list[str]:
    """
    Run all relevant validators. Called before advancing to next section.
    Returns list of warning/info strings to append to handler response.
    """
    warnings = []
    if message:
        warnings += detect_implied_flags(message, store)
    warnings += check_funding_balance(store)
    warnings += check_plausibility(store)
    warnings += check_contradictions(store)
    warnings += find_duplicates(store)
    warnings += check_moratorium_vs_impl(store)
    return warnings


def format_warnings(warnings: list[str]) -> str:
    """Format warnings block for display."""
    if not warnings:
        return ""
    return "\n\n---\n**⚡ Quick checks:**\n" + "\n".join(warnings)
