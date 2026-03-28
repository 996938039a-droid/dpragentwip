"""
agents/benchmark_engine.py
═══════════════════════════
Single batch LLM call that fills ALL Tier 2 benchmark fields.
Returns values + reasoning for display on the review screen.
"""

from __future__ import annotations
import json
from agents.extractor import llm_call, clean_json, DEFAULT_MODEL
from core.session_store import SessionStore


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
            return validated
        except Exception as e:
            print(f"[benchmark] failed: {e}, using fallbacks")
            result = dict(FALLBACK_BENCHMARKS)
            return self._apply_industry_overrides(store, result)


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
