"""agents/handlers/wc.py — collect missing working capital T1 fields"""
from agents.handlers.flag_detector import update_flags_from_message, flag_change_summary
from __future__ import annotations
from core.session_store import SessionStore
from agents.extractor import extract_json, DEFAULT_MODEL
from agents.handlers import FIRST_QUESTIONS, get_first_question
from agents.handlers.intake import _apply_wc

PROMPT = """Extract working capital parameters from this message.
Message: "{msg}"
Current: debtor_days={dd}, creditor_days_rm={cd}, stock_days={sd}, impl_months={im}
Return JSON with only fields PRESENT in message:
{{"debtor_days":null,"creditor_days_rm":null,"stock_days_rm":null,"implementation_months":null}}
All plain integers."""

class WCHandler:
    def __init__(self, api_key, model=DEFAULT_MODEL):
        self.api_key = api_key; self.model = model

    async def handle(self, message: str, store: SessionStore) -> str:

        # Scan message for applicability signals and update flags
        _flag_changes = update_flags_from_message(message, store)
        wc = store.working_capital
        d = await extract_json(
            PROMPT.format(msg=message, dd=wc.debtor_days, cd=wc.creditor_days_rm,
                          sd=wc.stock_days_rm, im=wc.implementation_months),
            self.api_key, fallback={}, model=self.model)
        _apply_wc(store, d)
        if wc.is_complete:
            store.current_section = "review"
            return (f"✅ Working capital captured.\n\n---\n\n"
                    "Generating your complete assumptions review...")
        missing = []
        if wc.debtor_days    < 0: missing.append("debtor days")
        if wc.creditor_days_rm < 0: missing.append("creditor days (RM)")
        if wc.stock_days_rm  < 0: missing.append("stock days")
        if wc.implementation_months < 0: missing.append("implementation months")
        return f"Still need: **{', '.join(missing)}**.\n\n" + FIRST_QUESTIONS["wc"]
