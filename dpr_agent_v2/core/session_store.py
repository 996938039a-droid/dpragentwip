"""
core/session_store.py
══════════════════════
All dataclasses that represent the DPR session state.

This is the single shared data object that all handlers read from
and write to. The orchestrator, handlers, and Excel generator all
work exclusively through this store.

Design principles:
  - Pure data — no logic, no LLM calls, no Excel code here
  - All fields have safe defaults
  - Designed to be JSON-serialisable for session persistence
"""

from dataclasses import dataclass, field
from enum import Enum
from typing import Optional


# ─── Enums ────────────────────────────────────────────────────────────────────

class EntityType(str, Enum):
    PROPRIETORSHIP = "Proprietorship"
    PARTNERSHIP    = "Partnership"
    LLP            = "LLP"
    COMPANY        = "Company"


class AssetCategory(str, Enum):
    CIVIL_WORKS     = "Civil Works"
    PLANT_MACHINERY = "Plant & Machinery"
    FURNITURE       = "Furniture & Fixture"
    VEHICLE         = "Vehicle"
    ELECTRICAL      = "Electrical & Fittings"
    PRE_OPERATIVE   = "Pre-operative Expenses"
    OTHER           = "Other"


class SectionStatus(str, Enum):
    PENDING     = "pending"
    IN_PROGRESS = "in_progress"
    COMPLETE    = "complete"


# ─── Sub-dataclasses ─────────────────────────────────────────────────────────

@dataclass
class Asset:
    name:       str
    category:   AssetCategory
    cost_lakhs: float


@dataclass
class FinanceSource:
    """A single source of project finance."""
    label:              str           # e.g. "Term Loan 1", "OD Limit"
    amount_lakhs:       float
    is_term_loan:       bool  = False
    is_od:              bool  = False
    is_equity:          bool  = False
    is_subsidy:         bool  = False
    interest_rate:      float = 0.09  # p.a. fraction (fixed at 9%)
    tenor_months:       int   = 84   # fixed 84 months
    moratorium_months:  int   = 6    # fixed 6 months moratorium
    is_vehicle_loan:    bool  = False
    is_unsecured:       bool  = False
    is_other_term:      bool  = False


@dataclass
class Product:
    name:              str
    unit:              str
    price_per_unit:    float
    capacity_per_day:  float
    output_ratio:      float = 1.0    # fraction of input that becomes output
    split_percent:     float = 1.0    # fraction of total output this product represents
    price_escalation:  float = 0.05   # T2


@dataclass
class RawMaterial:
    name:                 str
    unit:                 str
    price_per_unit:       float
    input_per_output:     float              # units of this input per 1 unit of applicable output
    price_escalation:     float = 0.06       # T2
    applies_to:           list = None        # None = ALL products; list of product names = specific products only
    # Examples:
    #   applies_to = None                          → shared input, used in all products equally
    #   applies_to = ["English Willow Grade A"]    → only this product uses this material
    #   applies_to = ["KW Standard", "KW Junior"]  → two products share this material


@dataclass
class EmployeeCategory:
    designation:          str
    count:                int
    monthly_salary_lakhs: float
    annual_increment:     float = 0.07  # T2


# ─── Section dataclasses ──────────────────────────────────────────────────────

@dataclass
class ProjectProfile:
    company_name:         str           = ""
    promoter_name:        str           = ""
    entity_type:          EntityType    = EntityType.COMPANY
    industry:             str           = ""
    city:                 str           = ""
    state:                str           = ""
    operation_start_date: str           = ""   # YYYY-MM
    projection_years:     int           = 7
    status:               SectionStatus = SectionStatus.PENDING


@dataclass
class CapitalMeans:
    assets:         list[Asset]         = field(default_factory=list)
    finance_sources:list[FinanceSource] = field(default_factory=list)
    status:         SectionStatus       = SectionStatus.PENDING

    @property
    def total_project_cost(self) -> float:
        return sum(a.cost_lakhs for a in self.assets)

    @property
    def total_finance(self) -> float:
        return sum(f.amount_lakhs for f in self.finance_sources)

    @property
    def promoter_equity(self) -> float:
        return sum(f.amount_lakhs for f in self.finance_sources if f.is_equity)

    @property
    def term_loans(self) -> list[FinanceSource]:
        return [f for f in self.finance_sources if f.is_term_loan]

    @property
    def od_sources(self) -> list[FinanceSource]:
        return [f for f in self.finance_sources if f.is_od]

    @property
    def vehicle_loans(self) -> list[FinanceSource]:
        return [f for f in self.finance_sources if f.is_vehicle_loan]

    @property
    def unsecured_loans(self) -> list[FinanceSource]:
        return [f for f in self.finance_sources if f.is_unsecured]

    @property
    def other_term_liabs(self) -> list[FinanceSource]:
        return [f for f in self.finance_sources if f.is_other_term]

    investment_deposits_lakhs: float = 0.0    # Fixed deposits / investments (non-project)
    other_non_current_lakhs:  float = 0.0    # Other non-current assets
    intangible_assets_lakhs:  float = 0.0    # Intangible assets
    nc_investments_lakhs:     float = 0.0    # Non-current investments
    security_deposits_lakhs:  float = 0.0    # Security deposits

    @property
    def is_balanced(self) -> bool:
        return abs(self.total_project_cost - self.total_finance) < 0.5


@dataclass
class RevenueModel:
    products:                  list[Product]  = field(default_factory=list)
    year1_utilization:         float          = 0.5
    annual_utilization_increment: float       = 0.05   # T2
    max_utilization:           float          = 0.85   # T2
    working_days_per_month:    int            = 26
    non_operating_income_lakhs:float          = 0.0    # Annual non-operating income (FD interest etc.)
    status:                    SectionStatus  = SectionStatus.PENDING


@dataclass
class CostStructure:
    raw_materials:            list[RawMaterial] = field(default_factory=list)
    transport_base_lakhs:     float             = 0.0
    misc_base_lakhs:          float             = 0.0
    # T2 fields — filled by benchmark engine
    rm_pct_of_fa:             float             = 0.02
    rm_escalation:            float             = 0.06
    insurance_pct_of_fa:      float             = 0.004
    insurance_escalation:     float             = 0.05
    power_pct_revenue:        float             = 0.07
    power_escalation:         float             = 0.06
    marketing_pct_revenue:    float             = 0.04
    marketing_escalation:     float             = 0.0
    transport_escalation:     float             = 0.10
    misc_escalation:          float             = 0.08
    sga_base_lakhs:           float             = 5.0
    sga_escalation:           float             = 0.10
    drawings_base_lakhs:      float             = 0.0   # Proprietor/partner drawings per year
    drawings_escalation:      float             = 0.0
    status:                   SectionStatus     = SectionStatus.PENDING


@dataclass
class ManpowerStructure:
    categories:    list[EmployeeCategory] = field(default_factory=list)
    status:        SectionStatus          = SectionStatus.PENDING


@dataclass
class WorkingCapital:
    debtor_days:          int           = -1    # -1 = not set
    creditor_days_rm:     int           = -1
    creditor_days_admin:  int           = 30    # T2
    stock_days_rm:        int           = -1
    stock_days_fg:        int           = 7     # T2
    wc_loan_amount:       float         = 0.0   # T3 = OD limit
    wc_interest_rate:     float         = 0.0   # T3 = OD rate
    wip_days:             int           = 0     # Work-in-progress days
    cold_store_days:      int           = 0     # Cold store / other store days
    implementation_months:int           = -1
    status:               SectionStatus = SectionStatus.PENDING

    @property
    def is_complete(self) -> bool:
        return (self.debtor_days > 0
                and self.creditor_days_rm > 0
                and self.stock_days_rm > 0
                and self.implementation_months > 0)


@dataclass
class BenchmarkValues:
    """
    Stores all T2 benchmark values after the benchmark engine runs.
    Also stores the reasoning text for display on the review screen.
    """
    values:   dict[str, float]  = field(default_factory=dict)  # key → value
    reasons:  dict[str, str]    = field(default_factory=dict)  # key → reasoning text
    source:   dict[str, str]    = field(default_factory=dict)  # key → "llm"|"fallback"
    complete: bool              = False


# ─── Root session store ───────────────────────────────────────────────────────

@dataclass
class SessionStore:
    """
    The single root object for an entire DPR session.
    All handlers read from and write to this object.
    The Excel generator reads from this object exclusively.
    """

    # Business metadata
    business_type:       str  = "MANUFACTURING"
    industry:            str  = ""
    business_description:str  = ""

    # Section data
    project_profile: ProjectProfile   = field(default_factory=ProjectProfile)
    capital_means:   CapitalMeans     = field(default_factory=CapitalMeans)
    revenue_model:   RevenueModel     = field(default_factory=RevenueModel)
    cost_structure:  CostStructure    = field(default_factory=CostStructure)
    manpower:        ManpowerStructure= field(default_factory=ManpowerStructure)
    working_capital: WorkingCapital   = field(default_factory=WorkingCapital)

    # Benchmark values (populated after benchmark engine runs)
    benchmarks:      BenchmarkValues  = field(default_factory=BenchmarkValues)

    # ── G: Intelligence tracking (G26/G27/G28) ───────────────────────────────
    corrections:       dict  = None   # G26: {field: [original, revised]} tracking
    provisional_fields:list  = None   # G27: fields entered early, needs re-confirm
    section_confidence:dict  = None   # G28: {section: 0-100} confidence score

    # ── Industry intelligence layer
    industry_code:   str = "manufacturing_general"  # set by detect_industry()
    # field_flags is populated on first message via detect_industry()
    # Stored as dict so it survives JSON serialisation; access via FieldFlags.from_dict()
    field_flags_dict: dict = None   # populated from industry profile at intake

    # Conversation state
    current_section: str = "intake"   # intake → profile → capital → revenue → costs → manpower → wc → review → done

    # ── Computed helpers ──────────────────────────────────────────────────────

    def record_correction(self, field: str, old_val, new_val):
        """G26: Track when user corrects a previously entered value."""
        if self.corrections is None:
            self.corrections = {}
        self.corrections[field] = {"original": old_val, "revised": new_val}

    def mark_provisional(self, field: str):
        """G27: Mark a field as provisional — entered early, needs re-confirmation."""
        if self.provisional_fields is None:
            self.provisional_fields = []
        if field not in self.provisional_fields:
            self.provisional_fields.append(field)

    def set_section_confidence(self, section: str, score: int):
        """G28: Record how much was user-confirmed vs AI-inferred for a section."""
        if self.section_confidence is None:
            self.section_confidence = {}
        self.section_confidence[section] = max(0, min(100, score))

    def get_section_confidence(self, section: str) -> int:
        if self.section_confidence:
            return self.section_confidence.get(section, 50)
        return 50

    @property
    def flags(self):
        """Live FieldFlags object. Handlers call store.flags.set('wip',1) etc."""
        from core.industry_config import FieldFlags, get_profile
        if self.field_flags_dict:
            return FieldFlags.from_dict(self.field_flags_dict)
        return get_profile(self.industry_code).flags

    def set_flag(self, key: str, value: int):
        """Set a single flag and persist it."""
        from core.industry_config import FieldFlags, get_profile
        if self.field_flags_dict is None:
            self.field_flags_dict = get_profile(self.industry_code).flags.as_dict()
        self.field_flags_dict[key] = value

    @property
    def n_products(self) -> int:
        return len(self.revenue_model.products)

    @property
    def n_materials(self) -> int:
        return len(self.cost_structure.raw_materials)

    @property
    def n_employees(self) -> int:
        return len(self.manpower.categories)

    @property
    def projection_years(self) -> int:
        return self.project_profile.projection_years

    @property
    def all_t1_complete(self) -> bool:
        """True when all required T1 fields have been collected."""
        pp = self.project_profile
        cm = self.capital_means
        rv = self.revenue_model
        cs = self.cost_structure
        mp = self.manpower
        wc = self.working_capital

        profile_ok  = bool(pp.company_name and pp.promoter_name
                          and pp.city and pp.state and pp.operation_start_date)
        capital_ok  = bool(cm.assets and cm.finance_sources)
        revenue_ok  = bool(rv.products and rv.year1_utilization > 0)
        costs_ok    = bool(cs.raw_materials)
        manpower_ok = bool(mp.categories)
        wc_ok       = wc.is_complete

        return all([profile_ok, capital_ok, revenue_ok, costs_ok, manpower_ok, wc_ok])

    def section_complete(self, section: str) -> bool:
        """Check if a specific section has all required T1 fields."""
        if section == "profile":
            pp = self.project_profile
            return bool(pp.company_name and pp.promoter_name
                       and pp.city and pp.state and pp.operation_start_date)
        if section == "capital":
            return bool(self.capital_means.assets and self.capital_means.finance_sources)
        if section == "revenue":
            return bool(self.revenue_model.products and self.revenue_model.year1_utilization > 0)
        if section == "costs":
            return bool(self.cost_structure.raw_materials)
        if section == "manpower":
            return bool(self.manpower.categories)
        if section == "wc":
            return self.working_capital.is_complete
        return False

    def next_incomplete_section(self) -> str:
        """Return the next section that needs T1 data."""
        for section in ["profile", "capital", "revenue", "costs", "manpower", "wc"]:
            if not self.section_complete(section):
                return section
        return "review"  # all T1 complete → go to review screen

    def to_layout_engine(self):
        """Create a LayoutEngine from the current store state."""
        from core.layout_engine import LayoutEngine
        return LayoutEngine(
            n_products=max(1, self.n_products),
            n_materials=max(1, self.n_materials),
            n_employees=max(1, self.n_employees),
            n_years=self.projection_years,
        )
