"""agents/handlers/revenue.py — collect missing revenue T1 fields"""
from agents.handlers.flag_detector import update_flags_from_message, flag_change_summary
from __future__ import annotations
from core.session_store import SessionStore
from agents.extractor import extract_json, DEFAULT_MODEL
from agents.handlers import FIRST_QUESTIONS
from agents.handlers.intake import _apply_revenue

PROMPT = """Extract revenue/product details from this message.
Message: "{msg}"
Current products captured: {n_products}
Return JSON (null/[] for missing):
{{"products":[{{"name":"...","unit":"...","price_per_unit":0,"capacity_per_day":0,"output_ratio":1.0,"split_percent":0}}],
  "year1_utilization":null,"working_days_per_month":null}}
split_percent as fraction (30% → 0.3). All numbers plain."""

class RevenueHandler:
    def __init__(self, api_key, model=DEFAULT_MODEL):
        self.api_key = api_key; self.model = model

    async def handle(self, message: str, store: SessionStore) -> str:

        # Scan message for applicability signals and update flags
        _flag_changes = update_flags_from_message(message, store)
        rv = store.revenue_model
        d = await extract_json(
            PROMPT.format(msg=message, n_products=len(rv.products)),
            self.api_key, fallback={"products":[]}, model=self.model)
        _apply_revenue(store, d)
        if rv.products and rv.year1_utilization > 0:
            store.current_section = store.next_incomplete_section()
            return (f"✅ Revenue captured: {len(rv.products)} product(s).\n\n---\n\n"
                    + FIRST_QUESTIONS.get(store.current_section, ""))
        missing = []
        if not rv.products: missing.append("product details (name, price, capacity)")
        if not rv.year1_utilization: missing.append("Year 1 utilisation %")
        return f"Still need: **{', '.join(missing)}**.\n\n" + FIRST_QUESTIONS["revenue"]
