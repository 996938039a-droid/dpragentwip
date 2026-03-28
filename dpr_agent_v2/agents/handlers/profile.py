"""
agents/handlers/profile.py — collect missing profile T1 fields
"""
from agents.handlers.flag_detector import update_flags_from_message, flag_change_summary
from __future__ import annotations
from core.session_store import SessionStore, EntityType
from agents.extractor import extract_json, DEFAULT_MODEL
from agents.handlers import FIRST_QUESTIONS

PROMPT = """Extract project profile from this message.
Message: "{msg}"
Current: company="{company}", promoter="{promoter}", entity="{entity}",
         city="{city}", state="{state}", start="{start}"
Return JSON with only the fields PRESENT in the message (omit others):
{{"company_name":null,"promoter_name":null,"entity_type":null,
  "city":null,"state":null,"operation_start_date":null,"projection_years":null}}
entity_type: Proprietorship|Partnership|LLP|Company
operation_start_date: YYYY-MM format"""


class ProfileHandler:
    def __init__(self, api_key: str, model: str = DEFAULT_MODEL):
        self.api_key = api_key
        self.model   = model

    async def handle(self, message: str, store: SessionStore) -> str:

        # Scan message for applicability signals and update flags
        _flag_changes = update_flags_from_message(message, store)
        pp = store.project_profile
        d = await extract_json(
            PROMPT.format(msg=message, company=pp.company_name,
                          promoter=pp.promoter_name, entity=pp.entity_type.value,
                          city=pp.city, state=pp.state, start=pp.operation_start_date),
            self.api_key, fallback={}, model=self.model
        )
        if d.get("company_name"):  pp.company_name  = d["company_name"]
        if d.get("promoter_name"): pp.promoter_name = d["promoter_name"]
        if d.get("city"):          pp.city          = d["city"]
        if d.get("state"):         pp.state         = d["state"]
        if d.get("operation_start_date"): pp.operation_start_date = d["operation_start_date"]
        if d.get("projection_years"):     pp.projection_years = int(d["projection_years"])
        if d.get("entity_type"):
            try: pp.entity_type = EntityType(d["entity_type"])
            except: pass

        missing = _missing(pp)
        if not missing:
            store.current_section = store.next_incomplete_section()
            return f"✅ Profile captured.\n\n---\n\n{FIRST_QUESTIONS.get(store.current_section,'')}"
        return _ask(missing)


def _missing(pp) -> list[str]:
    m = []
    if not pp.company_name:          m.append("company name")
    if not pp.promoter_name:         m.append("promoter name")
    if not pp.city:                  m.append("city")
    if not pp.state:                 m.append("state")
    if not pp.operation_start_date:  m.append("operation start date")
    return m

def _ask(missing: list[str]) -> str:
    return (f"Still need: **{', '.join(missing)}**.\n\n"
            "Please provide these for your project profile.")
