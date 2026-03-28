"""
core/industry_config.py
════════════════════════
Intelligent industry layer — flag-driven approach.

Every field has a 1/0 flag per industry:
  1 = applicable  → ask user in conversation + show row in Excel
  0 = not applicable → skip question + hide row in Excel

Flags are set at industry detection time and stored in SessionStore.
Handlers update flags live as the conversation reveals more about the business.
Drawings (1) for EVERY industry — every business owner may draw from the firm.

Usage:
    profile = detect_industry("cricket bat manufacturer, Haryana")
    store.field_flags = FieldFlags.from_profile(profile)
    store.industry_code = profile.code
    # later in handlers:
    store.field_flags.set("wip", 1)   # user mentioned WIP
"""

from __future__ import annotations
from dataclasses import dataclass, field


# ─── FieldFlags ───────────────────────────────────────────────────────────────
# Single source of truth for what to ask and what to show.
# 1 = applicable,  0 = not applicable

@dataclass
class FieldFlags:
    """
    Each flag is an int: 1 (applicable) or 0 (not applicable).

    These drive TWO things simultaneously:
      1. Conversation — whether to ask the user for this input
      2. Excel        — whether to show or hide the row/section

    Drawings is ALWAYS 1 — every business (company, LLP, proprietorship)
    may have owner drawings / dividends, so it is never suppressed.
    """

    # ── Inventory / W Cap ────────────────────────────────────────────────────
    raw_materials:    int = 1   # RM stock, RM creditors, COGS section
    finished_goods:   int = 1   # FG stock row
    wip:              int = 0   # Work-in-progress row
    cold_store:       int = 0   # Cold store / other stores row
    trade_debtors:    int = 1   # Trade receivables (almost always 1)

    # ── P&L cost lines ────────────────────────────────────────────────────────
    drawings:         int = 1   # ALWAYS 1 — every industry has drawings/dividends
    transport:        int = 1   # Transport cost line
    power:            int = 1   # Power & fuel
    marketing:        int = 1   # Marketing expenses
    rm_cogs:          int = 1   # Raw material COGS section in Expenses

    # ── BS non-current items ─────────────────────────────────────────────────
    intangibles:      int = 0   # Intangible assets
    security_deposits:int = 1   # Security deposits
    nc_investments:   int = 0   # Non-current investments
    fd_investments:   int = 1   # Fixed deposits / bank investments
    vehicle_loan:     int = 1   # Vehicle loan
    unsecured_loans:  int = 1   # Unsecured loans
    other_term_liab:  int = 1   # Other term liabilities

    # ── Finance ───────────────────────────────────────────────────────────────
    od_facility:      int = 1   # OD / WC loan

    # ── Revenue ───────────────────────────────────────────────────────────────
    capacity_revenue: int = 1   # Capacity/day based (manufacturing)
    service_revenue:  int = 0   # Project/contract/service based

    # ── Questions to skip in conversation ─────────────────────────────────────
    ask_rm_details:   int = 1   # Ask detailed RM cost questions
    ask_wip_days:     int = 0   # Ask WIP stock days
    ask_cold_days:    int = 0   # Ask cold store days
    ask_cold_chain:   int = 0   # Ask about cold chain logistics

    def set(self, key: str, value: int):
        """Update a single flag. Called by handlers as conversation reveals more."""
        if hasattr(self, key):
            setattr(self, key, value)

    def get(self, key: str, default: int = 1) -> int:
        return getattr(self, key, default)

    def as_dict(self) -> dict:
        """Return all flags as a plain dict for logging/review."""
        import dataclasses
        return dataclasses.asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "FieldFlags":
        valid = {k for k in cls.__dataclass_fields__}
        return cls(**{k: v for k, v in d.items() if k in valid})


# ─── Industry Profile ─────────────────────────────────────────────────────────

@dataclass
class IndustryProfile:
    code:       str
    name:       str
    keywords:   list[str]

    # Default flags for this industry (1/0 per field)
    flags:      FieldFlags = field(default_factory=FieldFlags)

    # T2 benchmark overrides
    benchmarks: dict = field(default_factory=dict)

    # Default WC parameters
    typical_debtor_days:    int = 30
    typical_creditor_days:  int = 15
    typical_stock_rm:       int = 20
    typical_stock_fg:       int = 7
    typical_wip_days:       int = 0
    typical_cold_days:      int = 0
    typical_impl_months:    int = 6

    # Conversation hints
    revenue_unit_hint: str = "units"
    rm_hint:           str = "raw materials per unit of output"
    industry_notes:    str = ""


# ─── Industry Registry ────────────────────────────────────────────────────────

PROFILES: list[IndustryProfile] = [

    IndustryProfile(
        code="manufacturing_general", name="General Manufacturing",
        keywords=["manufactur","fabricat","assembl","produc","factory","workshop",
                  "plant","casting","forging","machining","tooling","bat","leather",
                  "rubber","plastic","metal","wood","paper","glass","ceramic"],
        flags=FieldFlags(
            raw_materials=1, finished_goods=1, wip=0, cold_store=0, trade_debtors=1,
            drawings=1, transport=1, power=1, marketing=1, rm_cogs=1,
            intangibles=0, security_deposits=1, nc_investments=0, fd_investments=1,
            vehicle_loan=1, unsecured_loans=1, other_term_liab=1, od_facility=1,
            capacity_revenue=1, service_revenue=0,
            ask_rm_details=1, ask_wip_days=0, ask_cold_days=0, ask_cold_chain=0,
        ),
        benchmarks={"opex.power_pct_revenue":0.07,"opex.marketing_pct_revenue":0.04,"opex.sga_base":8.0,"wc.stock_days_fg":7},
        typical_debtor_days=30, typical_creditor_days=15, typical_stock_rm=20, typical_stock_fg=7,
        industry_notes="General manufacturing: raw material costs dominate. Ramp-up 50%→85% over 4 years.",
    ),

    IndustryProfile(
        code="food_processing", name="Food Processing / FMCG",
        keywords=["food","biscuit","snack","pickle","sauce","jam","dairy","milk",
                  "bakery","confectionery","fmcg","beverage","juice","flour","rice",
                  "edible","masala","spice","oil","ghee","namkeen","chips","noodle"],
        flags=FieldFlags(
            raw_materials=1, finished_goods=1, wip=1, cold_store=1, trade_debtors=1,
            drawings=1, transport=1, power=1, marketing=1, rm_cogs=1,
            intangibles=0, security_deposits=1, nc_investments=0, fd_investments=1,
            vehicle_loan=1, unsecured_loans=1, other_term_liab=1, od_facility=1,
            capacity_revenue=1, service_revenue=0,
            ask_rm_details=1, ask_wip_days=1, ask_cold_days=1, ask_cold_chain=1,
        ),
        benchmarks={"opex.power_pct_revenue":0.06,"opex.marketing_pct_revenue":0.06,"opex.sga_base":10.0,"wc.stock_days_fg":10,"material.escalation_agri":0.10},
        typical_debtor_days=21, typical_creditor_days=10, typical_stock_rm=15, typical_stock_fg=10,
        typical_wip_days=2, typical_cold_days=15,
        industry_notes="Food processing: agri inputs volatile. Cold store and WIP significant. Marketing spend higher.",
    ),

    IndustryProfile(
        code="textile_garment", name="Textile / Garment",
        keywords=["textile","garment","cloth","fabric","weav","knit","stitch",
                  "apparel","fashion","yarn","thread","dyeing","spinning",
                  "readymade","shirt","trouser","saree","kurti","hosiery"],
        flags=FieldFlags(
            raw_materials=1, finished_goods=1, wip=1, cold_store=0, trade_debtors=1,
            drawings=1, transport=1, power=1, marketing=1, rm_cogs=1,
            intangibles=0, security_deposits=1, nc_investments=0, fd_investments=1,
            vehicle_loan=1, unsecured_loans=1, other_term_liab=1, od_facility=1,
            capacity_revenue=1, service_revenue=0,
            ask_rm_details=1, ask_wip_days=1, ask_cold_days=0, ask_cold_chain=0,
        ),
        benchmarks={"opex.power_pct_revenue":0.08,"opex.marketing_pct_revenue":0.05,"opex.sga_base":6.0,"wc.stock_days_fg":14,"material.escalation_agri":0.09},
        typical_debtor_days=45, typical_creditor_days=30, typical_stock_rm=30, typical_stock_fg=14, typical_wip_days=5,
        industry_notes="Textile/garment: long WC cycles. Debtors 45 days common. WIP and FG stock both significant.",
    ),

    IndustryProfile(
        code="pharma_chemical", name="Pharma / Chemical",
        keywords=["pharma","chemical","drug","medicine","formulation","api",
                  "tablet","capsule","syrup","injection","resin","polymer",
                  "dye","pigment","fertiliser","pesticide","paint","adhesive"],
        flags=FieldFlags(
            raw_materials=1, finished_goods=1, wip=1, cold_store=0, trade_debtors=1,
            drawings=1, transport=1, power=1, marketing=1, rm_cogs=1,
            intangibles=1, security_deposits=1, nc_investments=0, fd_investments=1,
            vehicle_loan=1, unsecured_loans=1, other_term_liab=1, od_facility=1,
            capacity_revenue=1, service_revenue=0,
            ask_rm_details=1, ask_wip_days=1, ask_cold_days=0, ask_cold_chain=0,
        ),
        benchmarks={"opex.power_pct_revenue":0.05,"opex.marketing_pct_revenue":0.08,"opex.sga_base":15.0,"wc.stock_days_fg":21},
        typical_debtor_days=60, typical_creditor_days=30, typical_stock_rm=30, typical_stock_fg=21, typical_wip_days=7,
        industry_notes="Pharma/chemical: longer WC cycle, regulatory compliance costs, higher marketing. WIP significant.",
    ),

    IndustryProfile(
        code="agro_processing", name="Agro Processing / Cold Chain",
        keywords=["agro","agricult","cold chain","cold storage","potato","onion",
                  "mango","fruit","vegetable","pulses","rice mill","flour mill",
                  "dal","cereal","seed","cashew","groundnut","cotton gin"],
        flags=FieldFlags(
            raw_materials=1, finished_goods=1, wip=0, cold_store=1, trade_debtors=1,
            drawings=1, transport=1, power=1, marketing=1, rm_cogs=1,
            intangibles=0, security_deposits=1, nc_investments=0, fd_investments=1,
            vehicle_loan=1, unsecured_loans=1, other_term_liab=1, od_facility=1,
            capacity_revenue=1, service_revenue=0,
            ask_rm_details=1, ask_wip_days=0, ask_cold_days=1, ask_cold_chain=1,
        ),
        benchmarks={"opex.power_pct_revenue":0.09,"opex.marketing_pct_revenue":0.03,"opex.sga_base":5.0,"wc.stock_days_fg":30,"material.escalation_agri":0.12},
        typical_debtor_days=21, typical_creditor_days=7, typical_stock_rm=30, typical_stock_fg=30, typical_cold_days=45,
        typical_impl_months=9,
        industry_notes="Agro processing: seasonal procurement, high cold storage costs, commodity price volatility.",
    ),

    IndustryProfile(
        code="trading_retail", name="Trading / Retail / Distribution",
        keywords=["trading","wholesale","retail","distribut","dealer","stockist",
                  "merchant","import","export","shop","store","supermarket","e-commerce"],
        flags=FieldFlags(
            raw_materials=1, finished_goods=1, wip=0, cold_store=0, trade_debtors=1,
            drawings=1, transport=1, power=1, marketing=1, rm_cogs=1,
            intangibles=0, security_deposits=1, nc_investments=0, fd_investments=1,
            vehicle_loan=1, unsecured_loans=1, other_term_liab=1, od_facility=1,
            capacity_revenue=1, service_revenue=0,
            ask_rm_details=1, ask_wip_days=0, ask_cold_days=0, ask_cold_chain=0,
        ),
        benchmarks={"opex.power_pct_revenue":0.01,"opex.marketing_pct_revenue":0.02,"opex.sga_base":5.0,"wc.stock_days_fg":30},
        typical_debtor_days=30, typical_creditor_days=30, typical_stock_rm=0, typical_stock_fg=30,
        industry_notes="Trading: COGS = purchases for resale. Inventory days longer.",
        rm_hint="goods purchased for resale per unit sold",
    ),

    IndustryProfile(
        code="service_it", name="Service / IT / Consulting",
        keywords=["software","it service","it company","technology","consult","advisory",
                  "bpo","kpo","outsourc","saas","app","digital","data",
                  "engineering service","design service","legal","accounting"],
        flags=FieldFlags(
            raw_materials=0, finished_goods=0, wip=0, cold_store=0, trade_debtors=1,
            drawings=1, transport=0, power=1, marketing=1, rm_cogs=0,
            intangibles=1, security_deposits=1, nc_investments=0, fd_investments=1,
            vehicle_loan=1, unsecured_loans=1, other_term_liab=1, od_facility=1,
            capacity_revenue=0, service_revenue=1,
            ask_rm_details=0, ask_wip_days=0, ask_cold_days=0, ask_cold_chain=0,
        ),
        benchmarks={"opex.power_pct_revenue":0.02,"opex.marketing_pct_revenue":0.05,"opex.sga_base":12.0,"wc.stock_days_fg":0},
        typical_debtor_days=45, typical_creditor_days=30, typical_stock_rm=0, typical_stock_fg=0, typical_impl_months=3,
        industry_notes="Service: manpower is primary cost. No physical inventory. Debtors typically 45+ days.",
        rm_hint="direct service delivery costs (hosting, licences, subcontract)",
        revenue_unit_hint="projects / contracts / hours",
    ),

    IndustryProfile(
        code="hospitality", name="Hotel / Restaurant / Hospitality",
        keywords=["hotel","restaurant","resort","motel","guesthouse","dhaba",
                  "cafeteria","canteen","catering","bar","pub","cloud kitchen","hospitality"],
        flags=FieldFlags(
            raw_materials=1, finished_goods=0, wip=0, cold_store=1, trade_debtors=1,
            drawings=1, transport=1, power=1, marketing=1, rm_cogs=1,
            intangibles=0, security_deposits=1, nc_investments=0, fd_investments=1,
            vehicle_loan=1, unsecured_loans=1, other_term_liab=1, od_facility=1,
            capacity_revenue=1, service_revenue=0,
            ask_rm_details=1, ask_wip_days=0, ask_cold_days=1, ask_cold_chain=0,
        ),
        benchmarks={"opex.power_pct_revenue":0.08,"opex.marketing_pct_revenue":0.04,"opex.sga_base":10.0,"wc.stock_days_fg":0,"material.escalation_agri":0.10},
        typical_debtor_days=7, typical_creditor_days=15, typical_stock_rm=15, typical_stock_fg=0, typical_cold_days=15,
        typical_impl_months=12,
        industry_notes="Hospitality: high fixed costs, seasonal revenue. Power and staff dominate.",
        revenue_unit_hint="rooms / covers / seats",
    ),

    IndustryProfile(
        code="healthcare", name="Healthcare / Clinic / Hospital",
        keywords=["hospital","clinic","nursing home","diagnostic","path lab",
                  "ayurveda","dental","eye","maternity","health center",
                  "physiotherapy","healthcare"],
        flags=FieldFlags(
            raw_materials=1, finished_goods=0, wip=0, cold_store=0, trade_debtors=1,
            drawings=1, transport=1, power=1, marketing=1, rm_cogs=1,
            intangibles=1, security_deposits=1, nc_investments=0, fd_investments=1,
            vehicle_loan=1, unsecured_loans=1, other_term_liab=1, od_facility=1,
            capacity_revenue=1, service_revenue=0,
            ask_rm_details=1, ask_wip_days=0, ask_cold_days=0, ask_cold_chain=0,
        ),
        benchmarks={"opex.power_pct_revenue":0.04,"opex.marketing_pct_revenue":0.03,"opex.sga_base":8.0,"wc.stock_days_fg":0},
        typical_debtor_days=30, typical_creditor_days=30, typical_stock_rm=20, typical_stock_fg=0, typical_impl_months=9,
        industry_notes="Healthcare: high asset intensity (equipment), significant staff costs, compliance overhead.",
        revenue_unit_hint="patients / beds / procedures",
    ),

    IndustryProfile(
        code="construction", name="Construction / Real Estate / Infrastructure",
        keywords=["construct","build","real estate","developer","contractor",
                  "civil work","infrastructure","housing","apartment","road",
                  "bridge","tunnel","interior","fit-out"],
        flags=FieldFlags(
            raw_materials=1, finished_goods=0, wip=1, cold_store=0, trade_debtors=1,
            drawings=1, transport=1, power=1, marketing=1, rm_cogs=1,
            intangibles=0, security_deposits=1, nc_investments=0, fd_investments=1,
            vehicle_loan=1, unsecured_loans=1, other_term_liab=1, od_facility=1,
            capacity_revenue=0, service_revenue=1,
            ask_rm_details=1, ask_wip_days=1, ask_cold_days=0, ask_cold_chain=0,
        ),
        benchmarks={"opex.power_pct_revenue":0.03,"opex.marketing_pct_revenue":0.02,"opex.sga_base":8.0,"wc.stock_days_fg":0,"material.escalation_generic":0.09},
        typical_debtor_days=60, typical_creditor_days=30, typical_stock_rm=30, typical_stock_fg=0, typical_wip_days=60,
        typical_impl_months=3,
        industry_notes="Construction: high WIP, long debtor cycle, advance payments common.",
    ),

    IndustryProfile(
        code="steel_metal", name="Steel / Metal Fabrication",
        keywords=["steel","metal","iron","aluminium","copper","brass","zinc",
                  "roll","sheet","pipe","tube","wire","struct","girder","beam",
                  "fabricat","weld","galvaniz"],
        flags=FieldFlags(
            raw_materials=1, finished_goods=1, wip=1, cold_store=0, trade_debtors=1,
            drawings=1, transport=1, power=1, marketing=1, rm_cogs=1,
            intangibles=0, security_deposits=1, nc_investments=0, fd_investments=1,
            vehicle_loan=1, unsecured_loans=1, other_term_liab=1, od_facility=1,
            capacity_revenue=1, service_revenue=0,
            ask_rm_details=1, ask_wip_days=1, ask_cold_days=0, ask_cold_chain=0,
        ),
        benchmarks={"opex.power_pct_revenue":0.10,"opex.marketing_pct_revenue":0.02,"opex.sga_base":6.0,"wc.stock_days_fg":10,"material.escalation_generic":0.08},
        typical_debtor_days=45, typical_creditor_days=30, typical_stock_rm=30, typical_stock_fg=10, typical_wip_days=5,
        industry_notes="Steel/metal: power-intensive, RM is majority cost. Commodity price volatility high.",
    ),

    IndustryProfile(
        code="education", name="Education / Training Institute",
        keywords=["school","college","coaching","training","institute","academy",
                  "tutorial","education","e-learning","skill","vocational"],
        flags=FieldFlags(
            raw_materials=0, finished_goods=0, wip=0, cold_store=0, trade_debtors=1,
            drawings=1, transport=0, power=1, marketing=1, rm_cogs=0,
            intangibles=1, security_deposits=1, nc_investments=0, fd_investments=1,
            vehicle_loan=1, unsecured_loans=1, other_term_liab=1, od_facility=0,
            capacity_revenue=1, service_revenue=0,
            ask_rm_details=0, ask_wip_days=0, ask_cold_days=0, ask_cold_chain=0,
        ),
        benchmarks={"opex.power_pct_revenue":0.03,"opex.marketing_pct_revenue":0.05,"opex.sga_base":10.0,"wc.stock_days_fg":0},
        typical_debtor_days=0, typical_creditor_days=30, typical_stock_rm=0, typical_stock_fg=0, typical_impl_months=6,
        industry_notes="Education: staff costs primary. Fee income upfront (low debtors). No inventory.",
    ),
]

_DEFAULT = PROFILES[0]
_BY_CODE  = {p.code: p for p in PROFILES}


# ─── Public API ───────────────────────────────────────────────────────────────

def get_profile(code: str) -> IndustryProfile:
    return _BY_CODE.get(code, _DEFAULT)


def detect_industry(description: str) -> IndustryProfile:
    """Score-based detection — highest keyword match wins."""
    text  = description.lower()
    best  = _DEFAULT
    score = 0
    for p in PROFILES:
        s = sum(1 for kw in p.keywords if kw in text)
        if s > score:
            score = s
            best  = p
    return best


def default_wc_params(profile: IndustryProfile) -> dict:
    return {
        "debtor_days":           profile.typical_debtor_days,
        "creditor_days_rm":      profile.typical_creditor_days,
        "stock_days_rm":         profile.typical_stock_rm,
        "stock_days_fg":         profile.typical_stock_fg,
        "wip_days":              profile.typical_wip_days,
        "cold_store_days":       profile.typical_cold_days,
        "implementation_months": profile.typical_impl_months,
    }


def questions_to_skip(profile: IndustryProfile) -> list[str]:
    """Return sections to skip in conversation based on flags."""
    skip = []
    if not profile.flags.ask_rm_details:
        skip.append("costs")
    return skip


def applicable_wc_items(profile: IndustryProfile) -> dict:
    """Helper for workbook_builder — returns flag dict for W Cap rows."""
    f = profile.flags
    return {
        "raw_materials":  f.raw_materials,
        "finished_goods": f.finished_goods,
        "wip":            f.wip,
        "cold_store":     f.cold_store,
    }


def applicable_bs_items(profile: IndustryProfile) -> dict:
    """Helper for workbook_builder — returns flag dict for BS rows."""
    f = profile.flags
    return {
        "intangibles":    f.intangibles,
        "security_dep":   f.security_deposits,
        "nc_investments": f.nc_investments,
    }
