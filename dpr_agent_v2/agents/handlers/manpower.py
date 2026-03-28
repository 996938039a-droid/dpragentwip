"""agents/handlers/manpower.py — collect missing manpower T1 fields"""
from agents.handlers.flag_detector import update_flags_from_message, flag_change_summary
from __future__ import annotations
from core.session_store import SessionStore
from agents.extractor import extract_json, DEFAULT_MODEL
from agents.handlers import FIRST_QUESTIONS
from agents.handlers.intake import _apply_manpower

PROMPT = """Extract staffing/manpower from this message.
Message: "{msg}"
Return JSON:
{{"categories":[{{"designation":"...","count":0,"monthly_salary_lakhs":0}}]}}
monthly_salary_lakhs as decimal (₹40,000 → 0.40, ₹18,000 → 0.18).
If user says "8 workers at ₹18,000 each" → count=8, monthly_salary_lakhs=0.18."""

class ManpowerHandler:
    def __init__(self, api_key, model=DEFAULT_MODEL):
        self.api_key = api_key; self.model = model

    async def handle(self, message: str, store: SessionStore) -> str:

        # Scan message for applicability signals and update flags
        _flag_changes = update_flags_from_message(message, store)
        d = await extract_json(PROMPT.format(msg=message), self.api_key,
                               fallback={"categories":[]}, model=self.model)
        _apply_manpower(store, d)
        mp = store.manpower
        if mp.categories:
            total = sum(e.count * e.monthly_salary_lakhs * 12 for e in mp.categories)
            store.current_section = store.next_incomplete_section()
            return (f"✅ Manpower captured: {len(mp.categories)} categories, "
                    f"₹{total:.1f}L annual base salary.\n\n---\n\n"
                    + FIRST_QUESTIONS.get(store.current_section, ""))
        return "Still need employee details.\n\n" + FIRST_QUESTIONS["manpower"]
