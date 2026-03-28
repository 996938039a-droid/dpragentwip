"""
agents/handlers/costs.py
══════════════════════════
Collect raw material costs AND ask whether each material
is shared across all products or specific to certain products.
This fixes the fundamental multi-product COGS calculation.
"""
from agents.handlers.flag_detector import update_flags_from_message, flag_change_summary
from __future__ import annotations
from core.session_store import SessionStore
from agents.extractor import extract_json, DEFAULT_MODEL
from agents.handlers import FIRST_QUESTIONS
from agents.handlers.intake import _apply_costs

MATERIALS_PROMPT = """Extract raw material costs from this message.
Message: "{msg}"
Return JSON ([] for missing):
{{"raw_materials":[{{"name":"...","unit":"...","price_per_unit":0,"input_per_output":0}}],
  "transport_base_lakhs":null,"misc_base_lakhs":null}}
price_per_unit and input_per_output must be plain numbers."""

MAPPING_PROMPT = """A manufacturing business makes these products:
{products}

It uses these raw materials:
{materials}

The user says: "{msg}"

For each material, determine which products it goes into:
- If it goes into ALL products equally → applies_to: null
- If it goes into specific products only → applies_to: [list of exact product names]

Return ONLY JSON:
{{"mappings": [{{"material_name": "...", "applies_to": null_or_list}}]}}

Examples of "all products":
  "packaging goes into all", "oil is used for every product", "this is for everything"
Examples of "specific products":
  "EW Cleft only for English Willow bats", "this resin is only for grade A"
  "Kashmir willow is for KW bats only"

If the user hasn't addressed product mapping at all, return all as null (shared)."""


class CostsHandler:
    def __init__(self, api_key, model=DEFAULT_MODEL):
        self.api_key = api_key
        self.model   = model
        self._awaiting_mapping = False

    async def handle(self, message: str, store: SessionStore) -> str:

        # Scan message for applicability signals and update flags
        _flag_changes = update_flags_from_message(message, store)
        from core.industry_config import get_profile
        profile = get_profile(store.industry_code)

        # Service/non-RM industries: skip raw material questions
        if not profile.applicability.has_rm_cost:
            store.current_section = store.next_incomplete_section()
            from agents.handlers import FIRST_QUESTIONS
            return (f"✅ No raw material costs for **{profile.name}** businesses — skipping COGS section.\n\n"
                    f"---\n\n{FIRST_QUESTIONS.get(store.current_section, '')}")

        cs = store.cost_structure
        rv = store.revenue_model

        # Step 1: Extract raw materials if not yet done
        if not cs.raw_materials:
            d = await extract_json(
                MATERIALS_PROMPT.format(msg=message),
                self.api_key, fallback={"raw_materials": []}, model=self.model)
            _apply_costs(store, d)

        if not cs.raw_materials:
            return "Still need raw material details.\n\n" + FIRST_QUESTIONS["costs"]

        # Step 2: If multiple products exist, ask about product-material mapping
        # Only ask if we haven't already resolved mapping and there are 2+ products
        needs_mapping = (
            len(rv.products) > 1
            and any(m.applies_to is None for m in cs.raw_materials)
            and not self._awaiting_mapping
        )

        if needs_mapping:
            self._awaiting_mapping = True
            prod_list = "\n".join(f"  - {p.name}" for p in rv.products)
            mat_list  = "\n".join(f"  - {m.name}" for m in cs.raw_materials)
            return (
                f"✅ Captured {len(cs.raw_materials)} raw material(s).\n\n"
                f"You have **{len(rv.products)} products**. Quick question — "
                f"do all raw materials go into every product, or are some materials "
                f"specific to certain products?\n\n"
                f"**Your products:**\n{prod_list}\n\n"
                f"**Your materials:**\n{mat_list}\n\n"
                f"Example answer: *\"EW Cleft only for English Willow Grade A, "
                f"KW Cleft only for Kashmir bats, everything else goes into all\"*\n\n"
                f"Or just say **\"all shared\"** if every material goes into every product."
            )

        # Step 3: If awaiting mapping, parse the user's answer
        if self._awaiting_mapping:
            self._awaiting_mapping = False

            msg_lower = message.lower()
            all_shared = any(w in msg_lower for w in
                             ["all shared", "same for all", "all products",
                              "every product", "goes into all", "shared"])

            if not all_shared and rv.products:
                prod_list_str = "\n".join(f"  - {p.name}" for p in rv.products)
                mat_list_str  = "\n".join(f"  - {m.name}" for m in cs.raw_materials)
                d = await extract_json(
                    MAPPING_PROMPT.format(
                        products=prod_list_str,
                        materials=mat_list_str,
                        msg=message,
                    ),
                    self.api_key,
                    fallback={"mappings": []},
                    model=self.model,
                )
                _apply_mapping(store, d.get("mappings", []))

            # Also try to extract transport/misc from this message if not yet set
            if cs.transport_base_lakhs == 0 or cs.misc_base_lakhs == 0:
                d2 = await extract_json(
                    MATERIALS_PROMPT.format(msg=message),
                    self.api_key, fallback={}, model=self.model)
                if d2.get("transport_base_lakhs"):
                    cs.transport_base_lakhs = float(d2["transport_base_lakhs"])
                if d2.get("misc_base_lakhs"):
                    cs.misc_base_lakhs = float(d2["misc_base_lakhs"])

        # Step 4: Done
        store.current_section = store.next_incomplete_section()
        mapping_summary = _mapping_summary(store)
        return (
            f"✅ Cost structure complete.\n\n{mapping_summary}\n\n---\n\n"
            + FIRST_QUESTIONS.get(store.current_section, "")
        )


def _apply_mapping(store: SessionStore, mappings: list):
    """Apply product-material mapping from LLM response."""
    for entry in mappings:
        mat_name   = entry.get("material_name", "")
        applies_to = entry.get("applies_to")   # None or list
        for mat in store.cost_structure.raw_materials:
            if mat.name.lower() == mat_name.lower():
                mat.applies_to = applies_to if applies_to else None
                break


def _mapping_summary(store: SessionStore) -> str:
    """Show how materials are mapped for user confirmation."""
    lines = ["**Material → Product mapping:**"]
    for mat in store.cost_structure.raw_materials:
        if mat.applies_to is None:
            lines.append(f"  - {mat.name}: **all products**")
        else:
            lines.append(f"  - {mat.name}: **{', '.join(mat.applies_to)}** only")
    return "\n".join(lines)
