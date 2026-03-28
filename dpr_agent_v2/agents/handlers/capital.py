"""agents/handlers/capital.py — collect missing capital T1 fields"""
from agents.handlers.flag_detector import update_flags_from_message, flag_change_summary
from __future__ import annotations
from core.session_store import SessionStore, Asset, AssetCategory, FinanceSource
from agents.extractor import extract_json, DEFAULT_MODEL
from agents.handlers import FIRST_QUESTIONS
from agents.handlers.intake import _apply_capital

PROMPT = """Extract capital/finance details from this message.
Message: "{msg}"
Return JSON (null/[] for missing):
{{"assets":[{{"name":"...","cost_lakhs":0,"category":"Civil Works|Plant & Machinery|Furniture & Fixture|Vehicle|Other"}}],
  "term_loans":[{{"amount_lakhs":0,"rate_pa":0,"tenor_months":0,"moratorium_months":0}}],
  "od_limit_lakhs":null,"promoter_equity_lakhs":null}}
Numbers plain (no ₹). Rate as fraction (9.5% → 0.095)."""

class CapitalHandler:
    def __init__(self, api_key, model=DEFAULT_MODEL):
        self.api_key = api_key; self.model = model

    async def handle(self, message: str, store: SessionStore) -> str:

        # Scan message for applicability signals and update flags
        _flag_changes = update_flags_from_message(message, store)
        d = await extract_json(PROMPT.format(msg=message), self.api_key,
                               fallback={"assets":[],"term_loans":[]}, model=self.model)
        _apply_capital(store, d)
        cm = store.capital_means
        if cm.assets and cm.finance_sources:
            store.current_section = store.next_incomplete_section()
            summary = (f"✅ Capital captured: {len(cm.assets)} assets, "
                       f"₹{cm.total_project_cost:.0f}L total cost.\n\n---\n\n")
            return summary + FIRST_QUESTIONS.get(store.current_section, "")
        missing = []
        if not cm.assets: missing.append("asset list with costs")
        if not any(s.is_term_loan for s in cm.finance_sources): missing.append("term loan details")
        if not any(s.is_equity for s in cm.finance_sources): missing.append("promoter equity amount")
        return f"Still need: **{', '.join(missing)}**.\n\n" + FIRST_QUESTIONS["capital"]
