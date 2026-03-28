"""agents/handlers/__init__.py — industry-aware first questions"""

from core.industry_config import get_profile

# Base questions (industry-agnostic)
FIRST_QUESTIONS = {
    "profile":  (
        "Let\'s start with your project details.\n\n"
        "What is the name of your company, and who is the promoter? "
        "Also tell me your location (city & state), entity type "
        "(Proprietorship / Partnership / LLP / Company), and when you "
        "expect to start operations."
    ),
    "capital":  (
        "Now let\'s cover your project costs.\n\n"
        "What assets will you be creating or buying? Give me name and cost "
        "for each (e.g. \'Civil Works ₹45L, Plant & Machinery ₹35L\'). "
        "Also: term loan amount, OD limit, and promoter contribution. "
        "(Interest rate 9%, tenor 84 months, moratorium 6 months are fixed.)"
    ),
    "revenue":  (
        "Let\'s model your revenue.\n\n"
        "What do you sell? For each product/service: name, unit, selling price, "
        "and daily capacity (or monthly capacity for services). "
        "Also: Year 1 utilisation % and working days per month."
    ),
    "costs":    (
        "Now let\'s cover your input costs.\n\n"
        "What raw materials or inputs do you buy? For each: name, unit, price per unit, "
        "and quantity needed per unit of output. "
        "Also: annual transport and miscellaneous costs."
    ),
    "manpower": (
        "Tell me about your team.\n\n"
        "List each employee category with count and monthly salary. "
        "Example: \'1 Manager at ₹40,000, 8 Skilled Workers at ₹18,000 each\'"
    ),
    "wc": (
        "Working capital details:\n\n"
        "Debtor days (how long customers take to pay), "
        "creditor days (how long before you pay suppliers), "
        "raw material stock days, and "
        "implementation period (months before operations start)."
    ),
    "review": "✅ All inputs captured. Generating your review screen...",
}


def get_first_question(section: str, industry_code: str = "manufacturing_general") -> str:
    """
    Return the first question for a section, tailored to the industry.
    """
    profile = get_profile(industry_code)
    app     = profile.applicability

    if section == "costs":
        if not app.has_rm_cost:
            return (
                "Let\'s cover your direct operating costs.\n\n"
                f"For a **{profile.name}** business, what are your main direct cost inputs? "
                f"({profile.rm_hint}). "
                "Also: annual transport and miscellaneous costs if applicable."
            )
        return FIRST_QUESTIONS["costs"]

    if section == "revenue":
        unit_hint = profile.revenue_unit_hint
        if unit_hint != "units":
            return (
                f"Let\'s model your revenue.\n\n"
                f"For a **{profile.name}** business, what do you offer? "
                f"For each service/product: name, unit ({unit_hint}), price per {unit_hint}, "
                f"and capacity per month. Also: Year 1 utilisation %."
            )
        return FIRST_QUESTIONS["revenue"]

    if section == "wc":
        if not app.has_raw_materials:
            return (
                "Working capital details:\n\n"
                "Debtor days (how long customers take to pay), "
                "and implementation period (months before operations start). "
                f"Industry default stock days: RM={profile.typical_stock_rm}, "
                f"FG={profile.typical_stock_fg} — confirm or change."
            )
        extra = ""
        if app.has_wip:
            extra += f" WIP days (default {profile.typical_wip_days})."
        if app.has_cold_store:
            extra += f" Cold store days (default {profile.typical_cold_days})."
        if extra:
            return FIRST_QUESTIONS["wc"] + extra
        return FIRST_QUESTIONS["wc"]

    return FIRST_QUESTIONS.get(section, "")
