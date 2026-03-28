"""
agents/benchmark_engine.py
═══════════════════════════
Single batch LLM call that fills ALL Tier 2 benchmark fields.
Returns values + reasoning for display on the review screen.
"""

import json
from agents.extractor import llm_call, clean_json, DEFAULT_MODEL
from core.session_store import SessionStore


# ── B6: State-level industrial power tariff table (₹/kWh, 2024 estimates) ─────
STATE_POWER_TARIFF = {
    "haryana": 7.0, "punjab": 7.5, "uttar pradesh": 6.5, "uttarakhand": 5.5,
    "rajasthan": 8.0, "gujarat": 7.5, "maharashtra": 9.0, "karnataka": 9.5,
    "tamil nadu": 8.5, "andhra pradesh": 7.8, "telangana": 8.0, "kerala": 6.5,
    "west bengal": 7.0, "odisha": 5.5, "jharkhand": 6.0, "chhattisgarh": 5.0,
    "madhya pradesh": 7.5, "himachal pradesh": 5.0, "jammu and kashmir": 4.5,
    "delhi": 8.5, "goa": 6.0, "assam": 6.5, "default": 7.5,
}

# ── B9: Benchmark ranges per field {low, typical, high} ──────────────────────
BENCHMARK_RANGES = {
    "opex.power_pct_revenue":    {"low": 0.03, "typical": 0.07, "high": 0.15,
                                   "label": "Power & Fuel % Revenue"},
    "opex.rm_pct_fa":            {"low": 0.01, "typical": 0.02, "high": 0.04,
                                   "label": "R&M % Net Fixed Assets"},
    "opex.marketing_pct_revenue":{"low": 0.01, "typical": 0.04, "high": 0.10,
                                   "label": "Marketing % Revenue"},
    "opex.sga_base":             {"low": 3.0,  "typical": 8.0,  "high": 25.0,
                                   "label": "SGA Base (₹ Lakhs)"},
    "employee.increment":        {"low": 0.05, "typical": 0.08, "high": 0.15,
                                   "label": "Salary Increment"},
    "revenue.price_escalation":  {"low": 0.03, "typical": 0.05, "high": 0.10,
                                   "label": "Price Escalation"},
    "capacity.max_util":         {"low": 0.70, "typical": 0.85, "high": 0.95,
                                   "label": "Max Utilisation"},
    "wc.stock_days_fg":          {"low": 0,    "typical": 7,    "high": 30,
                                   "label": "FG Stock Days"},
}


BENCHMARK_SYSTEM = (
    "You are a financial analyst specialising in Indian MSME project finance. "
    "Given a business description, industry, location, and project scale, "
    "provide realistic benchmark values for financial assumptions. "
    "Return ONLY valid JSON. No markdown, no preamble."
)

BENCHMARK_PROMPT_TEMPLATE = """
You are benchmarking financial assumptions for an Indian MSME project.

Business: {description}
Industry: {industry}
Location: {location}
Project Cost: ₹{project_cost:.1f} Lakhs
Products: {products}
Raw Materials: {materials}

Provide realistic benchmark values for ALL fields below.
For each field return: value (number) and reason (1 sentence, max 12 words).

Fields needed:
1.  capacity.annual_increment     — annual utilisation increment (fraction, e.g. 0.05)
2.  capacity.max_util              — max utilisation ceiling (fraction, e.g. 0.85)
3.  revenue.price_escalation       — annual price escalation per product (fraction, e.g. 0.05)
4.  material.escalation_generic    — annual RM cost escalation generic (fraction, e.g. 0.06)
5.  material.escalation_agri       — annual RM escalation for agri/wood inputs (fraction, e.g. 0.08)
6.  opex.rm_pct_fa                 — repair & maintenance as % of net fixed assets (fraction, e.g. 0.02)
7.  opex.rm_escalation             — R&M cost escalation (fraction p.a., e.g. 0.06)
8.  opex.insurance_pct_fa          — insurance as % of net fixed assets (fraction, e.g. 0.004)
9.  opex.insurance_escalation      — insurance escalation (fraction p.a., e.g. 0.05)
10. opex.power_pct_revenue         — power & fuel as % of revenue (fraction, e.g. 0.07)
11. opex.power_escalation          — power cost escalation (fraction p.a., e.g. 0.06)
12. opex.marketing_pct_revenue     — marketing as % of revenue (fraction, e.g. 0.04)
13. opex.transport_escalation      — transport cost escalation (fraction p.a., e.g. 0.10)
14. opex.misc_escalation           — misc expenses escalation (fraction p.a., e.g. 0.08)
15. opex.sga_base                  — selling, general & admin base amount year 1 (INR Lakhs)
16. opex.sga_escalation            — SGA escalation (fraction p.a., e.g. 0.10)
17. employee.increment             — annual salary increment (fraction, e.g. 0.08)
18. finance.od_rate                — OD/CC interest rate (fraction p.a., e.g. 0.09)
19. wc.creditor_days_admin         — creditor days for admin expenses (days, e.g. 30)
20. wc.stock_days_fg               — finished goods stock days (days, e.g. 7)

Return JSON exactly in this format:
{{
  "capacity.annual_increment":   {{"value": 0.05, "reason": "..."}},
  "capacity.max_util":           {{"value": 0.85, "reason": "..."}},
  "revenue.price_escalation":    {{"value": 0.05, "reason": "..."}},
  "material.escalation_generic": {{"value": 0.06, "reason": "..."}},
  "material.escalation_agri":    {{"value": 0.08, "reason": "..."}},
  "opex.rm_pct_fa":              {{"value": 0.02, "reason": "..."}},
  "opex.rm_escalation":          {{"value": 0.06, "reason": "..."}},
  "opex.insurance_pct_fa":       {{"value": 0.004,"reason": "..."}},
  "opex.insurance_escalation":   {{"value": 0.05, "reason": "..."}},
  "opex.power_pct_revenue":      {{"value": 0.07, "reason": "..."}},
  "opex.power_escalation":       {{"value": 0.06, "reason": "..."}},
  "opex.marketing_pct_revenue":  {{"value": 0.04, "reason": "..."}},
  "opex.transport_escalation":   {{"value": 0.10, "reason": "..."}},
  "opex.misc_escalation":        {{"value": 0.08, "reason": "..."}},
  "opex.sga_base":               {{"value": 10.0, "reason": "..."}},
  "opex.sga_escalation":         {{"value": 0.10, "reason": "..."}},
  "employee.increment":          {{"value": 0.08, "reason": "..."}},
  "finance.od_rate":             {{"value": 0.09, "reason": "..."}},
  "wc.creditor_days_admin":      {{"value": 30,   "reason": "..."}},
  "wc.stock_days_fg":            {{"value": 7,    "reason": "..."}}
}}
"""

# Fallback values if benchmark call fails
FALLBACK_BENCHMARKS = {
    "capacity.annual_increment":   {"value": 0.05,  "reason": "Standard ramp-up for new MSME unit"},
    "capacity.max_util":           {"value": 0.85,  "reason": "Industry standard ceiling"},
    "revenue.price_escalation":    {"value": 0.05,  "reason": "Typical product price inflation"},
    "material.escalation_generic": {"value": 0.06,  "reason": "General commodity inflation"},
    "material.escalation_agri":    {"value": 0.08,  "reason": "Agri/wood input price volatility"},
    "opex.rm_pct_fa":              {"value": 0.02,  "reason": "Standard R&M for manufacturing"},
    "opex.rm_escalation":          {"value": 0.06,  "reason": "Maintenance cost inflation"},
    "opex.insurance_pct_fa":       {"value": 0.004, "reason": "Standard industrial insurance rate"},
    "opex.insurance_escalation":   {"value": 0.05,  "reason": "Insurance premium increases"},
    "opex.power_pct_revenue":      {"value": 0.07,  "reason": "Average manufacturing power intensity"},
    "opex.power_escalation":       {"value": 0.06,  "reason": "Annual tariff revision"},
    "opex.marketing_pct_revenue":  {"value": 0.04,  "reason": "Standard MSME marketing spend"},
    "opex.transport_escalation":   {"value": 0.10,  "reason": "Fuel price volatility"},
    "opex.misc_escalation":        {"value": 0.08,  "reason": "General expense inflation"},
    "opex.sga_base":               {"value": 5.0,   "reason": "Administrative overhead estimate"},
    "opex.sga_escalation":         {"value": 0.10,  "reason": "Admin cost escalation"},
    "employee.increment":          {"value": 0.08,  "reason": "Labour market wage growth"},
    "finance.od_rate":             {"value": 0.09,  "reason": "Current PSU bank OD rates"},
    "wc.creditor_days_admin":      {"value": 30,    "reason": "Standard admin payment terms"},
    "wc.stock_days_fg":            {"value": 7,     "reason": "Minimal FG holding for MSME"},
}


class BenchmarkEngine:
    def __init__(self, api_key: str, model: str = DEFAULT_MODEL):
        self.api_key = api_key
        self.model   = model

    async def generate(self, store: SessionStore) -> dict:
        """
        Run one batch benchmark call for the given store.
        Returns dict: {key: {value, reason}} for all T2 fields.
        """
        pp = store.project_profile
        rv = store.revenue_model
        cs = store.cost_structure

        products_desc = ", ".join(
            f"{p.name} @ ₹{p.price_per_unit}/{p.unit}"
            for p in rv.products[:5]
        )
        materials_desc = ", ".join(
            f"{m.name} ₹{m.price_per_unit}/{m.unit}"
            for m in cs.raw_materials[:5]
        )

        prompt = BENCHMARK_PROMPT_TEMPLATE.format(
            description   = store.business_description or f"{pp.industry} in {pp.city}, {pp.state}",
            industry      = pp.industry,
            location      = f"{pp.city}, {pp.state}",
            project_cost  = store.capital_means.total_project_cost,
            products      = products_desc or "N/A",
            materials     = materials_desc or "N/A",
        )

        try:
            raw = await llm_call(prompt, self.api_key,
                                 system=BENCHMARK_SYSTEM,
                                 model=self.model,
                                 max_tokens=2000)
            result = json.loads(clean_json(raw))
            # Validate structure — every key must have value + reason
            validated = {}
            for k, v in result.items():
                if isinstance(v, dict) and "value" in v and "reason" in v:
                    validated[k] = v
                else:
                    validated[k] = FALLBACK_BENCHMARKS.get(k, {"value": 0, "reason": "default"})
            # Fill any missing keys from fallbacks
            for k, v in FALLBACK_BENCHMARKS.items():
                if k not in validated:
                    validated[k] = v
            validated = self._apply_industry_overrides(store, validated)
            validated = self._adjust_for_location_and_scale(store, validated)
            validated = self._add_confidence_and_ranges(validated)
            return validated
        except Exception as e:
            print(f"[benchmark] failed: {e}, using fallbacks")
            result = dict(FALLBACK_BENCHMARKS)
            result = self._apply_industry_overrides(store, result)
            result = self._adjust_for_location_and_scale(store, result)
            result = self._add_confidence_and_ranges(result)
            return result


    def _adjust_for_location_and_scale(self, store, benchmarks: dict) -> dict:
        """B6: Adjust power benchmark for state tariff. B7: Scale SGA to project size."""
        pp = store.project_profile

        # B6: State-level power tariff adjustment
        state = (pp.state or "").lower()
        tariff = STATE_POWER_TARIFF.get(state, STATE_POWER_TARIFF["default"])
        # Baseline tariff = 7.5, adjust power % proportionally
        base_tariff = 7.5
        if "opex.power_pct_revenue" in benchmarks:
            adj = benchmarks["opex.power_pct_revenue"]["value"] * (tariff / base_tariff)
            benchmarks["opex.power_pct_revenue"]["value"]  = round(adj, 4)
            benchmarks["opex.power_pct_revenue"]["reason"] = (
                f"{pp.state} industrial tariff ≈ ₹{tariff:.1f}/kWh (adjusted)"
            )

        # B7: Scale-aware SGA (project cost drives administrative overhead)
        pc = store.capital_means.total_project_cost
        if pc > 0 and "opex.sga_base" in benchmarks:
            # SGA scales roughly as: 3L for <50L project, 8L for 100L, 20L for 500L+
            sga = min(max(pc * 0.06, 3.0), 25.0)
            benchmarks["opex.sga_base"]["value"]  = round(sga, 1)
            benchmarks["opex.sga_base"]["reason"] = (
                f"Scale-adjusted for ₹{pc:.0f}L project (6% of project cost, capped)"
            )

        return benchmarks

    def _add_confidence_and_ranges(self, benchmarks: dict) -> dict:
        """B8: Add confidence level. B9: Add range context to each benchmark."""
        for key, data in benchmarks.items():
            # B8: Set confidence based on whether it came from LLM or fallback
            if "reason" in data and data["reason"] != "default":
                data["confidence"] = "medium"   # LLM-generated
            else:
                data["confidence"] = "low"       # fallback

            # B9: Add range if available
            if key in BENCHMARK_RANGES:
                r = BENCHMARK_RANGES[key]
                val = data.get("value", r["typical"])
                if val <= r["low"]:
                    position = "low end"
                elif val >= r["high"]:
                    position = "high end"
                else:
                    position = "typical range"
                data["range"] = (
                    f"Low: {r['low']} / Typical: {r['typical']} / High: {r['high']} "
                    f"— your value is at {position}"
                )
        return benchmarks

    def _apply_industry_overrides(self, store, benchmarks: dict) -> dict:
        """Merge industry-specific benchmark overrides on top of LLM results."""
        from core.industry_config import get_profile
        profile = get_profile(getattr(store, "industry_code", "manufacturing_general"))
        for k, v in profile.benchmarks.items():
            # Industry override takes precedence over LLM value for key benchmarks
            if k in benchmarks:
                benchmarks[k]["value"]  = v
                benchmarks[k]["reason"] = f"Industry standard for {profile.name}"
            else:
                benchmarks[k] = {"value": v, "reason": f"Industry standard for {profile.name}"}
        return benchmarks

    def check_rm_margin(self, store) -> list[str]:
        """B10: Cross-check if implied gross margin is within industry norms."""
        warnings = []
        rv  = store.revenue_model
        cs  = store.cost_structure
        from core.industry_config import get_profile
        profile = get_profile(getattr(store, "industry_code", "manufacturing_general"))

        if not rv.products or not cs.raw_materials:
            return warnings

        # Estimate Year 1 revenue
        rev = sum(
            p.price_per_unit * p.capacity_per_day * rv.working_days_per_month
            * 12 * rv.year1_utilization * p.split_percent
            for p in rv.products
        ) / 100000
        if rev <= 0:
            return warnings

        # Estimate COGS
        cogs = sum(
            m.price_per_unit * m.input_per_output
            * sum(p.capacity_per_day * rv.working_days_per_month * 12
                  * rv.year1_utilization * p.split_percent
                  for p in rv.products
                  if m.applies_to is None or p.name in (m.applies_to or []))
            for m in cs.raw_materials
        ) / 100000

        gm_pct = (rev - cogs) / rev * 100 if rev > 0 else 0
        rm_pct = cogs / rev * 100

        # Industry gross margin norms
        gm_norms = {
            "manufacturing_general": (25, 55),
            "food_processing":       (15, 40),
            "textile_garment":       (20, 45),
            "trading_retail":        (10, 30),
            "service_it":            (50, 85),
            "steel_metal":           (10, 30),
        }
        code = getattr(store, "industry_code", "manufacturing_general")
        lo, hi = gm_norms.get(code, (20, 60))

        if gm_pct < lo:
            warnings.append(
                f"⚠️  **Gross margin ≈ {gm_pct:.0f}%** is below typical range ({lo}–{hi}%) "
                f"for {profile.name}. RM cost is {rm_pct:.0f}% of revenue — check input prices "
                f"and quantities."
            )
        elif gm_pct > hi:
            warnings.append(
                f"ℹ️  **Gross margin ≈ {gm_pct:.0f}%** is above typical range ({lo}–{hi}%) "
                f"for {profile.name}. This may be optimistic — verify selling prices."
            )
        return warnings

    def apply_to_store(self, store: SessionStore, benchmarks: dict) -> SessionStore:
        """
        Apply benchmark values to a SessionStore.
        Only sets fields that weren't explicitly provided by the user.
        """
        def get(key: str) -> float:
            return benchmarks.get(key, FALLBACK_BENCHMARKS.get(key, {})).get("value", 0)

        rv = store.revenue_model
        cs = store.cost_structure
        mp = store.manpower

        # Capacity
        rv.annual_utilization_increment = get("capacity.annual_increment")
        rv.max_utilization              = get("capacity.max_util")

        # Revenue — apply same escalation to all products
        for p in rv.products:
            if p.price_escalation == 0.05:  # default — override with benchmark
                p.price_escalation = get("revenue.price_escalation")

        # Raw materials — use agri escalation for wood/agri inputs, generic for others
        agri_keywords = {"willow", "wood", "bamboo", "cane", "jute", "cotton",
                         "wool", "silk", "leather", "rubber", "oil", "resin"}
        for mat in cs.raw_materials:
            name_lower = mat.name.lower()
            is_agri = any(kw in name_lower for kw in agri_keywords)
            mat.price_escalation = (
                get("material.escalation_agri") if is_agri
                else get("material.escalation_generic")
            )

        # Opex
        cs.rm_pct_of_fa          = get("opex.rm_pct_fa")
        cs.rm_escalation         = get("opex.rm_escalation")
        cs.insurance_pct_of_fa   = get("opex.insurance_pct_fa")
        cs.insurance_escalation  = get("opex.insurance_escalation")
        cs.power_pct_revenue     = get("opex.power_pct_revenue")
        cs.power_escalation      = get("opex.power_escalation")
        cs.marketing_pct_revenue = get("opex.marketing_pct_revenue")
        cs.transport_escalation  = get("opex.transport_escalation")
        cs.misc_escalation       = get("opex.misc_escalation")
        cs.sga_base_lakhs        = get("opex.sga_base")
        cs.sga_escalation        = get("opex.sga_escalation")

        # Manpower
        for emp in mp.categories:
            emp.annual_increment = get("employee.increment")

        # Finance
        for src in store.capital_means.od_sources:
            src.interest_rate = get("finance.od_rate")

        # Working capital
        store.working_capital.creditor_days_admin = int(get("wc.creditor_days_admin"))
        store.working_capital.stock_days_fg       = int(get("wc.stock_days_fg"))
        store.working_capital.wc_loan_amount      = (
            store.capital_means.od_sources[0].amount_lakhs
            if store.capital_means.od_sources else 0
        )
        store.working_capital.wc_interest_rate = get("finance.od_rate")

        return store
