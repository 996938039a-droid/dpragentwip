"""
agents/handlers/review.py
═══════════════════════════
Generates the single review screen showing all assumptions (T1+T2+T3),
handles user overrides, and triggers Excel generation on confirm.
"""

from __future__ import annotations
import re
from core.session_store import SessionStore
from agents.extractor import extract_json, DEFAULT_MODEL


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
        """Build the complete review screen text."""
        pp = store.project_profile
        cm = store.capital_means
        rv = store.revenue_model
        cs = store.cost_structure
        mp = store.manpower
        wc = store.working_capital

        lines = [
            "## 📋 Complete Assumption Review",
            f"**{pp.company_name}** — {pp.industry}",
            f"{pp.city}, {pp.state}  |  Operations from {pp.operation_start_date}  |  {pp.projection_years} year projection",
            "",
            "---",
            "",
            "### ✅ You provided (T1 — Blue)",
            "",
            "**Project & Finance:**",
            f"- Entity: {pp.entity_type.value}",
            f"- Total project cost: ₹{cm.total_project_cost:.1f}L",
        ]

        for a in cm.assets:
            lines.append(f"  - {a.name}: ₹{a.cost_lakhs:.1f}L")

        for src in cm.finance_sources:
            if src.is_term_loan:
                lines.append(f"- Term Loan: ₹{src.amount_lakhs:.1f}L @ {src.interest_rate*100:.1f}%, "
                             f"{src.tenor_months}mo tenor, {src.moratorium_months}mo mora")
            elif src.is_od:
                lines.append(f"- OD Limit: ₹{src.amount_lakhs:.1f}L")
            elif src.is_equity:
                lines.append(f"- Promoter Equity: ₹{src.amount_lakhs:.1f}L")

        lines += ["", "**Products:**"]
        for p in rv.products:
            lines.append(f"- {p.name}: ₹{p.price_per_unit:,.0f}/{p.unit}, "
                        f"{p.capacity_per_day}/day, {p.split_percent*100:.0f}% of production")
        lines.append(f"- Year 1 utilisation: {rv.year1_utilization*100:.0f}%  |  "
                    f"{rv.working_days_per_month} working days/month")

        lines += ["", "**Raw Materials:**"]
        for m in cs.raw_materials:
            lines.append(f"- {m.name}: ₹{m.price_per_unit:.0f}/{m.unit}, "
                        f"{m.input_per_output} per output unit")
        if cs.transport_base_lakhs > 0:
            lines.append(f"- Transport: ₹{cs.transport_base_lakhs:.1f}L/yr  |  "
                        f"Misc: ₹{cs.misc_base_lakhs:.1f}L/yr")

        lines += ["", "**Manpower:**"]
        for e in mp.categories:
            lines.append(f"- {e.designation}: {e.count} × ₹{e.monthly_salary_lakhs:.2f}L/mo")

        lines += ["", f"**Working Capital:** Debtor {wc.debtor_days}d  |  "
                     f"Creditor {wc.creditor_days_rm}d  |  Stock {wc.stock_days_rm}d  |  "
                     f"Impl {wc.implementation_months}mo"]

        # T2 benchmarks
        lines += ["", "---", "", "### 📊 Benchmarked estimates (T2 — Green)",
                  "*(Change anything by typing e.g. \"make power 6%, salary increment 5%\")*", ""]

        def b(key: str) -> str:
            bm = benchmarks.get(key, {})
            val = bm.get("value", "—")
            reason = bm.get("reason", "")
            if isinstance(val, float) and val < 1:
                display = f"{val*100:.1f}%"
            elif isinstance(val, float):
                display = f"₹{val:.1f}L"
            else:
                display = str(val)
            return f"**{display}** — {reason}"

        lines += [
            f"- Annual utilisation increment: {b('capacity.annual_increment')}",
            f"- Max utilisation ceiling: {b('capacity.max_util')}",
            f"- Price escalation (all products): {b('revenue.price_escalation')}",
            f"- RM cost escalation (generic): {b('material.escalation_generic')}",
            f"- RM cost escalation (agri/wood): {b('material.escalation_agri')}",
            f"- R&M rate: {b('opex.rm_pct_fa')}",
            f"- Insurance rate: {b('opex.insurance_pct_fa')}",
            f"- Power & Fuel % revenue: {b('opex.power_pct_revenue')}",
            f"- Marketing % revenue: {b('opex.marketing_pct_revenue')}",
            f"- SGA base: {b('opex.sga_base')}",
            f"- Transport escalation: {b('opex.transport_escalation')}",
            f"- Misc escalation: {b('opex.misc_escalation')}",
            f"- Salary increment: {b('employee.increment')}",
            f"- OD interest rate: {b('finance.od_rate')}",
            f"- Creditor days (admin): {b('wc.creditor_days_admin')}",
            f"- FG stock days: {b('wc.stock_days_fg')}",
        ]

        # T3 statutory
        lines += ["", "---", "", "### ⚙️ Statutory (T3 — Grey)", "",
                  "- Depreciation: P&M 15%, Civil 10%, Furniture 10% (IT Act WDV)",
                  "- Company tax: 25% + 7% surcharge (if PBT > ₹100L) + 4% HEC",
                  "- Months in year: 12",
                  "", "---", "",
                  "Type **confirm** to generate your DPR Excel, or tell me what to change."]

        return "\n".join(lines)

    async def handle(self, message: str, store: SessionStore,
                     benchmarks: dict) -> tuple[str, bool]:
        """
        Handle user response on review screen.
        Returns (reply_text, should_generate_excel).
        """
        msg_lower = message.lower().strip()
        confirm_words = {"confirm", "yes", "ok", "okay", "generate",
                         "looks good", "proceed", "go ahead", "done"}

        if any(w in msg_lower for w in confirm_words):
            return "✅ Confirmed! Generating your DPR Excel...", True

        # Try to parse overrides
        d = await extract_json(
            OVERRIDE_PROMPT.format(msg=message),
            self.api_key, fallback={}, model=self.model
        )
        changed = _apply_overrides(store, d)

        if changed:
            return (f"Updated: {', '.join(changed)}.\n\n"
                    "Type **confirm** to generate, or make more changes."), False
        else:
            return ("I didn't catch any specific changes. "
                    "Type **confirm** to generate, or specify what to change "
                    "(e.g. 'make power 6%, set salary increment to 5%')."), False


def _apply_overrides(store: SessionStore, d: dict) -> list[str]:
    """Apply override dict to store. Returns list of changed field names."""
    changed = []
    cs = store.cost_structure
    rv = store.revenue_model
    mp = store.manpower
    wc = store.working_capital

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
                changed.append(key.split(".")[-1].replace("_", " "))
            except Exception as e:
                print(f"[review override] {key}: {e}")

    return changed
