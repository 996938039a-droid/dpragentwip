"""
agents/handlers/intake.py
═══════════════════════════
Runs 6 parallel LLM extractions on the first user message,
applies everything found, marks complete sections, returns
the first question for whatever is still missing.
"""

from __future__ import annotations
import asyncio
from core.session_store import (
    SessionStore, Asset, AssetCategory, FinanceSource,
    Product, RawMaterial, EmployeeCategory, EntityType
)
from agents.extractor import extract_all_parallel, DEFAULT_MODEL
from agents.handlers.flag_detector import update_flags_from_message
from core.industry_config import detect_industry, default_wc_params, questions_to_skip

PROFILE_PROMPT = """Extract company/project profile from this message.
Message: "{msg}"
Return JSON (null for missing):
{{"company_name":null,"promoter_name":null,"entity_type":null,
  "industry":null,"city":null,"state":null,
  "operation_start_date":null,"projection_years":7}}
entity_type must be one of: Proprietorship, Partnership, LLP, Company
operation_start_date format: YYYY-MM"""

CAPITAL_PROMPT = """Extract capital and finance details from this message.
Message: "{msg}"
Return JSON (null/empty for missing):
{{"assets":[{{"name":"...","cost_lakhs":0,"category":"Civil Works|Plant & Machinery|Furniture & Fixture|Vehicle|Electrical & Fittings|Other"}}],
  "term_loans":[{{"amount_lakhs":0,"rate_pa":0,"tenor_months":0,"moratorium_months":0}}],
  "od_limit_lakhs":null,"promoter_equity_lakhs":null}}
All numbers plain (no ₹, no commas). Rate as fraction (9.5% → 0.095)."""

REVENUE_PROMPT = """Extract revenue/product details from this message.
Message: "{msg}"
Return JSON (null for missing):
{{"products":[{{"name":"...","unit":"...","price_per_unit":0,"capacity_per_day":0,"output_ratio":1.0,"split_percent":0}}],
  "year1_utilization":null,"working_days_per_month":null}}
split_percent as fraction (30% → 0.3). price_per_unit plain number."""

COSTS_PROMPT = """Extract raw material costs from this message.
Message: "{msg}"
Return JSON (empty lists for missing):
{{"raw_materials":[{{"name":"...","unit":"...","price_per_unit":0,"input_per_output":0}}],
  "transport_base_lakhs":null,"misc_base_lakhs":null}}
price_per_unit and input_per_output must be plain numbers."""

MANPOWER_PROMPT = """Extract manpower/staffing from this message.
Message: "{msg}"
Return JSON (empty list for missing):
{{"categories":[{{"designation":"...","count":0,"monthly_salary_lakhs":0}}]}}
monthly_salary_lakhs as number (₹40,000/month → 0.40)."""

WC_PROMPT = """Extract working capital parameters from this message.
Message: "{msg}"
Return JSON (null for missing):
{{"debtor_days":null,"creditor_days_rm":null,"stock_days_rm":null,"implementation_months":null}}
All plain integers."""


class IntakeHandler:
    def __init__(self, api_key: str, model: str = DEFAULT_MODEL):
        self.api_key = api_key
        self.model   = model

    async def handle(self, message: str, store: SessionStore) -> str:
        """
        Extract all sections in parallel, apply to store,
        return greeting + list of what was captured + first missing question.
        """
        prompts = [
            (PROFILE_PROMPT.format(msg=message),   {}),
            (CAPITAL_PROMPT.format(msg=message),   {"assets":[],"term_loans":[],"od_limit_lakhs":None,"promoter_equity_lakhs":None}),
            (REVENUE_PROMPT.format(msg=message),   {"products":[],"year1_utilization":None,"working_days_per_month":None}),
            (COSTS_PROMPT.format(msg=message),     {"raw_materials":[],"transport_base_lakhs":None,"misc_base_lakhs":None}),
            (MANPOWER_PROMPT.format(msg=message),  {"categories":[]}),
            (WC_PROMPT.format(msg=message),        {"debtor_days":None,"creditor_days_rm":None,"stock_days_rm":None,"implementation_months":None}),
        ]

        results = await extract_all_parallel(prompts, self.api_key, self.model)
        profile_d, capital_d, revenue_d, costs_d, manpower_d, wc_d = results

        # Apply profile
        _apply_profile(store, profile_d)

        # Detect industry from description, set profile and initialise field flags
        industry_text = (store.project_profile.industry or store.business_description or message)
        profile = detect_industry(industry_text)
        store.industry_code = profile.code

        # Initialise field_flags_dict from the detected industry profile
        # Handlers can call store.set_flag("wip", 1) to override during conversation
        if store.field_flags_dict is None:
            store.field_flags_dict = profile.flags.as_dict()

        # Set business type for backward compatibility
        if "trading" in profile.code or "retail" in profile.code:
            store.business_type = "TRADING"
        elif "service" in profile.code or "it" in profile.code or "education" in profile.code:
            store.business_type = "SERVICE"
        else:
            store.business_type = "MANUFACTURING"

        # Apply industry-default WC parameters if not yet set by user
        wc = store.working_capital
        wc_defaults = default_wc_params(profile)
        if wc.debtor_days         < 0: wc.debtor_days         = wc_defaults["debtor_days"]
        if wc.creditor_days_rm    < 0: wc.creditor_days_rm    = wc_defaults["creditor_days_rm"]
        if wc.stock_days_rm       < 0: wc.stock_days_rm       = wc_defaults["stock_days_rm"]
        if wc.stock_days_fg       == 7: wc.stock_days_fg      = wc_defaults["stock_days_fg"]
        if wc.wip_days            == 0: wc.wip_days           = wc_defaults["wip_days"]
        if wc.cold_store_days     == 0: wc.cold_store_days    = wc_defaults["cold_store_days"]
        if wc.implementation_months < 0: wc.implementation_months = wc_defaults["implementation_months"]

        _apply_capital(store, capital_d)
        _apply_revenue(store, revenue_d)
        _apply_costs(store, costs_d)
        _apply_manpower(store, manpower_d)
        _apply_wc(store, wc_d)

        # Scan full first message for flag signals
        update_flags_from_message(message, store)

        # Build response
        captured  = _captured_summary(store)
        next_sec  = store.next_incomplete_section()
        from agents.handlers import FIRST_QUESTIONS
        first_q   = FIRST_QUESTIONS.get(next_sec, "")

        from core.industry_config import get_profile
        profile_obj  = get_profile(store.industry_code)
        industry_lbl = store.project_profile.industry or profile_obj.name
        notes        = profile_obj.industry_notes

        f = store.flags
        applicable = []
        if f.raw_materials:   applicable.append("Raw Materials")
        if f.wip:             applicable.append("WIP Stock")
        if f.finished_goods:  applicable.append("Finished Goods")
        if f.cold_store:      applicable.append("Cold Store")
        if f.intangibles:     applicable.append("Intangibles")
        applicable_str = ", ".join(applicable) if applicable else "Standard"

        reply = (
            f"Got it. Detected **{profile_obj.name}** industry.\n\n"
            f"_{notes}_\n\n"
            f"**Applicable items for your industry:** {applicable_str}\n"
            f"**Drawings:** Always included ✓\n\n"
            "I've pre-filled industry-specific benchmark values — "
            "you'll see those as green cells on the review screen.\n\n"
        )

        if captured:
            reply += "**Already captured from your message:**\n" + captured + "\n\n"

        if next_sec == "review":
            reply += "✅ All inputs captured! Running benchmarks...\n\n"
        else:
            sec_name = {
                "profile":  "project details",
                "capital":  "capital & finance",
                "revenue":  "revenue model",
                "costs":    "input costs",
                "manpower": "manpower",
                "wc":       "working capital",
            }.get(next_sec, next_sec)
            reply += f"Still need a few details about your **{sec_name}**:\n\n---\n\n{first_q}"

        store.current_section = next_sec
        return reply


# ── Apply helpers ─────────────────────────────────────────────────────────────

def _apply_profile(store: SessionStore, d: dict):
    pp = store.project_profile
    if d.get("company_name"):    pp.company_name  = d["company_name"]
    if d.get("promoter_name"):   pp.promoter_name = d["promoter_name"]
    if d.get("industry"):        pp.industry      = d["industry"]
    if d.get("city"):            pp.city          = d["city"]
    if d.get("state"):           pp.state         = d["state"]
    if d.get("operation_start_date"):
        pp.operation_start_date = d["operation_start_date"]
    if d.get("projection_years"):
        pp.projection_years = int(d["projection_years"])
    if d.get("entity_type"):
        try:    pp.entity_type = EntityType(d["entity_type"])
        except: pass


def _apply_capital(store: SessionStore, d: dict):
    cm = store.capital_means
    cat_map = {
        "civil":        AssetCategory.CIVIL_WORKS,
        "plant":        AssetCategory.PLANT_MACHINERY,
        "machinery":    AssetCategory.PLANT_MACHINERY,
        "furniture":    AssetCategory.FURNITURE,
        "vehicle":      AssetCategory.VEHICLE,
        "electrical":   AssetCategory.ELECTRICAL,
        "pre-operative":AssetCategory.PRE_OPERATIVE,
    }
    for a in (d.get("assets") or []):
        name = a.get("name","")
        cost = float(a.get("cost_lakhs") or 0)
        if cost <= 0: continue
        cat_hint = (a.get("category","") + " " + name).lower()
        category = AssetCategory.OTHER
        for k, v in cat_map.items():
            if k in cat_hint:
                category = v
                break
        # Don't duplicate
        if not any(x.name.lower() == name.lower() for x in cm.assets):
            cm.assets.append(Asset(name=name, category=category, cost_lakhs=cost))

    for tl in (d.get("term_loans") or []):
        amt = float(tl.get("amount_lakhs") or 0)
        if amt <= 0: continue
        if not any(s.is_term_loan for s in cm.finance_sources):
            cm.finance_sources.append(FinanceSource(
                label="Term Loan",
                amount_lakhs=amt,
                is_term_loan=True,
                interest_rate=float(tl.get("rate_pa") or 0.095),
                tenor_months=int(tl.get("tenor_months") or 84),
                moratorium_months=int(tl.get("moratorium_months") or 0),
            ))

    if d.get("od_limit_lakhs"):
        od = float(d["od_limit_lakhs"])
        if od > 0 and not any(s.is_od for s in cm.finance_sources):
            cm.finance_sources.append(FinanceSource(
                label="OD Limit", amount_lakhs=od, is_od=True, interest_rate=0.09))

    if d.get("promoter_equity_lakhs"):
        eq = float(d["promoter_equity_lakhs"])
        if eq > 0 and not any(s.is_equity for s in cm.finance_sources):
            cm.finance_sources.append(FinanceSource(
                label="Promoter Equity", amount_lakhs=eq, is_equity=True))


def _apply_revenue(store: SessionStore, d: dict):
    rv = store.revenue_model
    for p in (d.get("products") or []):
        name = p.get("name","")
        if not name: continue
        if any(x.name.lower() == name.lower() for x in rv.products): continue
        rv.products.append(Product(
            name=name,
            unit=p.get("unit","units"),
            price_per_unit=float(p.get("price_per_unit") or 0),
            capacity_per_day=float(p.get("capacity_per_day") or 0),
            output_ratio=float(p.get("output_ratio") or 1.0),
            split_percent=float(p.get("split_percent") or 1.0),
        ))
    if d.get("year1_utilization") is not None:
        rv.year1_utilization = float(d["year1_utilization"])
    if d.get("working_days_per_month") is not None:
        rv.working_days_per_month = int(d["working_days_per_month"])


def _apply_costs(store: SessionStore, d: dict):
    cs = store.cost_structure
    for m in (d.get("raw_materials") or []):
        name = m.get("name","")
        if not name: continue
        if any(x.name.lower() == name.lower() for x in cs.raw_materials): continue
        cs.raw_materials.append(RawMaterial(
            name=name,
            unit=m.get("unit","kg"),
            price_per_unit=float(m.get("price_per_unit") or 0),
            input_per_output=float(m.get("input_per_output") or 1),
        ))
    if d.get("transport_base_lakhs") is not None:
        cs.transport_base_lakhs = float(d["transport_base_lakhs"])
    if d.get("misc_base_lakhs") is not None:
        cs.misc_base_lakhs = float(d["misc_base_lakhs"])


def _apply_manpower(store: SessionStore, d: dict):
    mp = store.manpower
    for c in (d.get("categories") or []):
        desig = c.get("designation","")
        if not desig: continue
        if any(x.designation.lower() == desig.lower() for x in mp.categories): continue
        mp.categories.append(EmployeeCategory(
            designation=desig,
            count=int(c.get("count") or 1),
            monthly_salary_lakhs=float(c.get("monthly_salary_lakhs") or 0),
        ))


def _apply_wc(store: SessionStore, d: dict):
    wc = store.working_capital
    if d.get("debtor_days") is not None:       wc.debtor_days = int(d["debtor_days"])
    if d.get("creditor_days_rm") is not None:  wc.creditor_days_rm = int(d["creditor_days_rm"])
    if d.get("stock_days_rm") is not None:     wc.stock_days_rm = int(d["stock_days_rm"])
    if d.get("implementation_months") is not None:
        wc.implementation_months = int(d["implementation_months"])


def _captured_summary(store: SessionStore) -> str:
    lines = []
    pp = store.project_profile
    cm = store.capital_means
    rv = store.revenue_model
    cs = store.cost_structure
    mp = store.manpower
    wc = store.working_capital

    if pp.company_name:
        lines.append(f"✅ Profile: {pp.company_name}, {pp.city or 'location TBD'}")
    if cm.assets:
        lines.append(f"✅ Capital: {len(cm.assets)} assets, ₹{cm.total_project_cost:.0f}L project cost")
    if rv.products:
        lines.append(f"✅ Revenue: {len(rv.products)} product(s)")
    if cs.raw_materials:
        lines.append(f"✅ Costs: {len(cs.raw_materials)} raw material(s)")
    if mp.categories:
        lines.append(f"✅ Manpower: {len(mp.categories)} category/ies")
    if wc.is_complete:
        lines.append(f"✅ Working capital: debtor {wc.debtor_days}d, creditor {wc.creditor_days_rm}d")
    return "\n".join(lines)
