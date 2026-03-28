"""
agents/handlers/flag_detector.py
══════════════════════════════════
Scans user messages during the conversation and updates store.field_flags_dict
when the user mentions something that implies a flag should be 1.

Called from every handler after extracting data — it reads the raw message
and sets flags accordingly. This is the "based on the conversation" part.

Examples:
  "we have a cold room" → set cold_store=1
  "WIP is about 5 days" → set wip=1, ask_wip_days=1
  "we have intangible IP" → set intangibles=1
  "no transport needed" → set transport=0
  "this is a partnership, we draw salaries" → drawings already=1 (always)
"""

from core.session_store import SessionStore


# Keyword triggers — each entry: (keywords, flag_name, value)
# Positive signals (set flag=1 when keyword found in message)
_POSITIVE = [
    # Inventory
    (["wip","work in progress","work-in-progress","in-process","semi-finished"],
     "wip", 1),
    (["cold stor","cold room","cold chain","refrigerat","freezer","chiller"],
     "cold_store", 1),
    (["finished good","fg stock","fg days","packed good"],
     "finished_goods", 1),
    # Assets
    (["intangible","patent","trademark","goodwill","brand value","ip asset","licence asset"],
     "intangibles", 1),
    (["investment","fdr","fixed deposit","mutual fund","non-current invest"],
     "nc_investments", 1),
    # Finance
    (["vehicle loan","car loan","truck loan","two-wheeler loan","auto loan"],
     "vehicle_loan", 1),
    (["unsecured loan","friends and family loan","director loan","personal loan to business"],
     "unsecured_loans", 1),
    # Cold chain
    (["cold chain","reefer","temperature control","refrigerated transport"],
     "cold_store", 1),
]

# Negative signals (set flag=0 when explicitly stated)
_NEGATIVE = [
    (["no transport","no delivery","not applicable transport","transport not"],
     "transport", 0),
    (["no marketing","no advertisement","no ad spend"],
     "marketing", 0),
    (["no raw material","no material","no rm","no inputs","no inventory",
      "no stock","pure service","zero inventory"],
     "raw_materials", 0),
    (["no wip","no work in progress","instant production"],
     "wip", 0),
    (["no cold","no refrigerat","ambient storage only"],
     "cold_store", 0),
]


def update_flags_from_message(message: str, store: SessionStore) -> list[str]:
    """
    Scan message, update store.field_flags_dict for any triggered keywords.
    Returns list of changes made (for logging).
    """
    # Ensure flag dict is initialised
    if store.field_flags_dict is None:
        store.field_flags_dict = get_profile(store.industry_code).flags.as_dict()

    text    = message.lower()
    changes = []

    for keywords, flag_name, value in _POSITIVE + _NEGATIVE:
        if any(kw in text for kw in keywords):
            old = store.field_flags_dict.get(flag_name)
            if old != value:
                store.set_flag(flag_name, value)
                action = "enabled" if value == 1 else "disabled"
                changes.append(f"{flag_name} → {value} ({action})")

    return changes


def flag_change_summary(changes: list[str]) -> str:
    """Format flag changes for display in agent response."""
    if not changes:
        return ""
    lines = ["_(Automatically updated applicability flags based on your input:_"]
    for c in changes:
        lines.append(f"  • {c}")
    lines.append("_)_")
    return "\n".join(lines)
