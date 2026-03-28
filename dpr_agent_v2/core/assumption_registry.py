"""
core/assumption_registry.py
════════════════════════════
THE single source of truth for every field in the DPR model.

Every field that appears in the Assumption sheet is registered here with:
  - Its section (A–J)
  - Its tier (1=user, 2=benchmark, 3=statutory)
  - Its anchor row (the section's first data row — NEVER changes)
  - Its max_count (how many items the section reserves space for)
  - Its rows_per_item (how many rows each product/material/employee takes)
  - Its field_offset (which row within the item block this field sits on)

The Layout Engine reads ONLY from this registry. No row numbers exist
anywhere else in the codebase.

Section Anchor Map (derived from MAX counts):
  A  Capacity          anchor=4   (5 rows, fixed)
  B  Revenue/Products  anchor=13  (MAX_PRODUCTS=10, 4 rows each → 40 rows)
  C  Raw Materials     anchor=56  (MAX_MATERIALS=20, 3 rows each → 60 rows)
  D  Opex              anchor=119 (14 rows, fixed)
  E  Manpower          anchor=136 (MAX_EMPLOYEES=15, 2 rows each → 30 rows)
  F  Finance/TL        anchor=169 (10 rows, fixed)
  G  Working Capital   anchor=182 (9 rows, fixed)
  H  Depreciation      anchor=194 (6 rows, fixed)
  I  Implementation    anchor=203 (2 rows, fixed)
  J  Tax               anchor=208 (5 rows, fixed)
"""

from dataclasses import dataclass, field
from enum import Enum
from typing import Optional


# ─── Tier enum ───────────────────────────────────────────────────────────────

class FieldTier(int, Enum):
    USER       = 1   # must be collected from user
    BENCHMARK  = 2   # filled by benchmark engine if not provided
    STATUTORY  = 3   # regulatory / auto-filled, never asked


# ─── Section enum ────────────────────────────────────────────────────────────

class Section(str, Enum):
    A = "A"   # Capacity
    B = "B"   # Revenue / Products
    C = "C"   # Raw Materials
    D = "D"   # Operating Expenses
    E = "E"   # Manpower
    F = "F"   # Finance / Term Loan
    G = "G"   # Working Capital
    H = "H"   # Depreciation
    I = "I"   # Implementation
    J = "J"   # Tax
    K = "K"   # Balance Sheet Non-Current Items


# ─── Field descriptor ────────────────────────────────────────────────────────

@dataclass
class FieldDef:
    key:          str           # unique dotted key e.g. "capacity.year1_util"
    label:        str           # human label for review screen
    section:      Section
    tier:         FieldTier
    unit:         str           # display unit e.g. "fraction", "INR Lakhs", "days"
    default:      float | str | None = None   # T3 default or T2 fallback
    # For fixed fields (non-repeating):
    row_offset:   int = 0       # rows below section anchor (0-based)
    # For repeating blocks (products/materials/employees):
    is_repeating: bool = False
    rows_per_item:int = 0       # how many rows each item occupies
    item_field_offset: int = 0  # which row within the item is THIS field


# ─── Section capacity constants ───────────────────────────────────────────────

MAX_PRODUCTS   = 10   # max products/services in Revenue section
MAX_MATERIALS  = 20   # max raw materials in Cost section
MAX_EMPLOYEES  = 15   # max employee categories in Manpower section

ROWS_PER_PRODUCT  = 4   # name, price, capacity+split, escalation
ROWS_PER_MATERIAL = 3   # name+unit, price, escalation
ROWS_PER_EMPLOYEE = 2   # designation+count, salary+increment

# ─── Section anchor rows (first data row of each section) ────────────────────
# These are computed from the maxes above and NEVER change.
# They are the load-bearing constants of the entire model.

ANCHOR_A = 4    # Capacity — rows 4–8
ANCHOR_B = 13   # Revenue  — rows 13–52  (header row 13, data 14+)
ANCHOR_C = 56   # Raw mats — rows 56–115 (header row 56, data 57+)
                # = ANCHOR_B + 1 + MAX_PRODUCTS * ROWS_PER_PRODUCT + 2
ANCHOR_D = 119  # Opex     — rows 119–132
                # = ANCHOR_C + 1 + MAX_MATERIALS * ROWS_PER_MATERIAL + 2
ANCHOR_E = 136  # Manpower — rows 136–165 (header row 136, data 137+)
                # = ANCHOR_D + 14 + 3
ANCHOR_F = 169  # Finance  — rows 169–178
                # = ANCHOR_E + 1 + MAX_EMPLOYEES * ROWS_PER_EMPLOYEE + 2
ANCHOR_G = 182  # WC       — rows 182–190
ANCHOR_H = 196  # Depr — shifted +2 to clear G rows 192-193
ANCHOR_I = 205  # Impl
ANCHOR_J = 210  # Tax
ANCHOR_K = 220  # BS Non-Current Items


# ─── Complete field registry ──────────────────────────────────────────────────

FIELDS: list[FieldDef] = [

    # ── A: Capacity ──────────────────────────────────────────────────────────
    FieldDef("capacity.year1_util",       "Year 1 Capacity Utilisation",
             Section.A, FieldTier.USER,       "fraction (0-1)", 0.5,  row_offset=0),
    FieldDef("capacity.annual_increment", "Annual Utilisation Increment",
             Section.A, FieldTier.BENCHMARK,  "fraction p.a.",  0.05, row_offset=1),
    FieldDef("capacity.max_util",         "Maximum Utilisation Ceiling",
             Section.A, FieldTier.BENCHMARK,  "fraction (0-1)", 0.85, row_offset=2),
    FieldDef("capacity.working_days",     "Working Days per Month",
             Section.A, FieldTier.USER,       "days",           26,   row_offset=3),
    FieldDef("capacity.months_in_year",   "Months in a Year",
             Section.A, FieldTier.STATUTORY,  "months",         12,   row_offset=4),

    # ── B: Revenue — repeating block per product ──────────────────────────────
    # Each product occupies 4 rows:
    #   offset 0: product name (text, col B) + unit (col D)
    #   offset 1: base price year 1 (col E)
    #   offset 2: capacity per day (col E) + split % (col F)
    #   offset 3: annual price escalation (col E)
    FieldDef("revenue.product.name",       "Product / Service Name",
             Section.B, FieldTier.USER,       "text",  None,
             is_repeating=True, rows_per_item=ROWS_PER_PRODUCT, item_field_offset=0),
    FieldDef("revenue.product.price",      "Base Selling Price (Yr 1)",
             Section.B, FieldTier.USER,       "INR/unit", None,
             is_repeating=True, rows_per_item=ROWS_PER_PRODUCT, item_field_offset=1),
    FieldDef("revenue.product.capacity",   "Production Capacity per Day",
             Section.B, FieldTier.USER,       "units/day", None,
             is_repeating=True, rows_per_item=ROWS_PER_PRODUCT, item_field_offset=2),
    FieldDef("revenue.product.escalation", "Annual Price Escalation",
             Section.B, FieldTier.BENCHMARK,  "fraction p.a.", 0.05,
             is_repeating=True, rows_per_item=ROWS_PER_PRODUCT, item_field_offset=3),

    # ── C: Raw Materials — repeating block per material ───────────────────────
    # Each material occupies 3 rows:
    #   offset 0: material name (col B) + unit (col D)
    #   offset 1: base price year 1 (col E)
    #   offset 2: annual cost escalation (col E) + input_per_output (col F)
    FieldDef("material.name",              "Material Name",
             Section.C, FieldTier.USER,       "text",  None,
             is_repeating=True, rows_per_item=ROWS_PER_MATERIAL, item_field_offset=0),
    FieldDef("material.price",             "Base Price per Unit (Yr 1)",
             Section.C, FieldTier.USER,       "INR/unit", None,
             is_repeating=True, rows_per_item=ROWS_PER_MATERIAL, item_field_offset=1),
    FieldDef("material.escalation",        "Annual Cost Escalation",
             Section.C, FieldTier.BENCHMARK,  "fraction p.a.", 0.06,
             is_repeating=True, rows_per_item=ROWS_PER_MATERIAL, item_field_offset=2),

    # ── D: Operating Expenses — fixed rows ────────────────────────────────────
    FieldDef("opex.rm_pct_fa",             "Repair & Maintenance (% of Net Fixed Assets)",
             Section.D, FieldTier.BENCHMARK,  "fraction", 0.02, row_offset=0),
    FieldDef("opex.rm_escalation",         "R&M Cost Annual Escalation",
             Section.D, FieldTier.BENCHMARK,  "fraction p.a.", 0.06, row_offset=1),
    FieldDef("opex.insurance_pct_fa",      "Insurance (% of Net Fixed Assets)",
             Section.D, FieldTier.BENCHMARK,  "fraction", 0.004, row_offset=2),
    FieldDef("opex.insurance_escalation",  "Insurance Annual Escalation",
             Section.D, FieldTier.BENCHMARK,  "fraction p.a.", 0.05, row_offset=3),
    FieldDef("opex.power_pct_revenue",     "Power & Fuel (% of Revenue)",
             Section.D, FieldTier.BENCHMARK,  "fraction", 0.07, row_offset=4),
    FieldDef("opex.power_escalation",      "Power Cost Annual Escalation",
             Section.D, FieldTier.BENCHMARK,  "fraction p.a.", 0.06, row_offset=5),
    FieldDef("opex.marketing_pct_revenue", "Marketing Expenses (% of Revenue)",
             Section.D, FieldTier.BENCHMARK,  "fraction", 0.04, row_offset=6),
    FieldDef("opex.marketing_escalation",  "Marketing Annual Escalation",
             Section.D, FieldTier.STATUTORY,  "fraction p.a.", 0.0, row_offset=7),
    FieldDef("opex.transport_base",        "Transportation Cost — Base Year 1",
             Section.D, FieldTier.USER,       "INR Lakhs", 0.0, row_offset=8),
    FieldDef("opex.transport_escalation",  "Transport Cost Annual Escalation",
             Section.D, FieldTier.BENCHMARK,  "fraction p.a.", 0.10, row_offset=9),
    FieldDef("opex.misc_base",             "Miscellaneous Expenses — Base Year 1",
             Section.D, FieldTier.USER,       "INR Lakhs", 0.0, row_offset=10),
    FieldDef("opex.misc_escalation",       "Miscellaneous Annual Escalation",
             Section.D, FieldTier.BENCHMARK,  "fraction p.a.", 0.08, row_offset=11),
    FieldDef("opex.sga_base",              "Selling, General & Admin — Base Year 1",
             Section.D, FieldTier.BENCHMARK,  "INR Lakhs", 5.0, row_offset=12),
    FieldDef("opex.sga_escalation",        "SGA Annual Escalation",
             Section.D, FieldTier.BENCHMARK,  "fraction p.a.", 0.10, row_offset=13),
    FieldDef("opex.drawings_base",         "Proprietor Drawings / Dividends (Base Year 1)",
             Section.D, FieldTier.USER,       "INR Lakhs", 0.0, row_offset=14),
    FieldDef("opex.drawings_escalation",   "Drawings Annual Escalation",
             Section.D, FieldTier.BENCHMARK,  "fraction p.a.", 0.0, row_offset=15),

    # ── E: Manpower — repeating block per category ────────────────────────────
    # Each category occupies 2 rows:
    #   offset 0: designation (col B) + headcount (col D)
    #   offset 1: monthly salary (col E) + annual increment (col F)
    FieldDef("employee.designation",       "Designation",
             Section.E, FieldTier.USER,       "text", None,
             is_repeating=True, rows_per_item=ROWS_PER_EMPLOYEE, item_field_offset=0),
    FieldDef("employee.salary",            "Monthly Salary",
             Section.E, FieldTier.USER,       "INR Lakhs/month", None,
             is_repeating=True, rows_per_item=ROWS_PER_EMPLOYEE, item_field_offset=1),
    FieldDef("employee.increment",         "Annual Salary Increment",
             Section.E, FieldTier.BENCHMARK,  "fraction p.a.", 0.07,
             is_repeating=True, rows_per_item=ROWS_PER_EMPLOYEE, item_field_offset=1),

    # ── F: Finance — fixed rows ───────────────────────────────────────────────
    FieldDef("finance.tl_amount",          "Term Loan Amount",
             Section.F, FieldTier.USER,       "INR Lakhs", 0.0, row_offset=1),
    FieldDef("finance.tl_rate",            "Term Loan Interest Rate  [Fixed: 9%]",
             Section.F, FieldTier.STATUTORY,  "fraction p.a.", 0.09, row_offset=2),
    FieldDef("finance.tl_tenor",           "Term Loan Total Tenor  [Fixed: 84 months]",
             Section.F, FieldTier.STATUTORY,  "months", 84, row_offset=3),
    FieldDef("finance.tl_moratorium",      "Moratorium Period  [Fixed: 6 months]",
             Section.F, FieldTier.STATUTORY,  "months", 6, row_offset=4),
    FieldDef("finance.tl_repayment",       "Repayment Months (derived)",
             Section.F, FieldTier.STATUTORY,  "months", None, row_offset=5),
    FieldDef("finance.od_limit",           "OD / CC Limit",
             Section.F, FieldTier.USER,       "INR Lakhs", 0.0, row_offset=7),
    FieldDef("finance.od_rate",            "OD Interest Rate",
             Section.F, FieldTier.BENCHMARK,  "fraction p.a.", 0.09, row_offset=8),
    FieldDef("finance.vehicle_loan_amount","Vehicle Loan — Amount",
             Section.F, FieldTier.USER,       "INR Lakhs", 0.0, row_offset=9),
    FieldDef("finance.unsecured_loan_amount","Unsecured Loans — Amount",
             Section.F, FieldTier.USER,       "INR Lakhs", 0.0, row_offset=10),
    FieldDef("finance.other_term_liab",    "Other Term Liabilities — Amount",
             Section.F, FieldTier.USER,       "INR Lakhs", 0.0, row_offset=11),

    # ── G: Working Capital — fixed rows ──────────────────────────────────────
    FieldDef("wc.debtor_days",             "Debtor Days (Receivables)",
             Section.G, FieldTier.USER,       "days", 30, row_offset=0),
    FieldDef("wc.creditor_days_rm",        "Creditor Days — Raw Materials",
             Section.G, FieldTier.USER,       "days", 15, row_offset=1),
    FieldDef("wc.creditor_days_admin",     "Creditor Days — Admin Expenses",
             Section.G, FieldTier.BENCHMARK,  "days", 30, row_offset=2),
    FieldDef("wc.stock_days_rm",           "Raw Material Stock Days",
             Section.G, FieldTier.USER,       "days", 15, row_offset=3),
    FieldDef("wc.stock_days_fg",           "Finished Goods Stock Days",
             Section.G, FieldTier.BENCHMARK,  "days", 7, row_offset=4),
    FieldDef("wc.wc_loan_amount",          "Working Capital Loan Amount",
             Section.G, FieldTier.STATUTORY,  "INR Lakhs", 0.0, row_offset=5),
    FieldDef("wc.wc_interest_rate",        "Working Capital Interest Rate",
             Section.G, FieldTier.STATUTORY,  "fraction p.a.", 0.0, row_offset=6),
    FieldDef("wc.investment_deposits",     "Investment & Deposits (FD / Investments)",
             Section.G, FieldTier.USER,       "INR Lakhs", 0.0, row_offset=7),
    FieldDef("wc.other_non_current",       "Other Non-Current Assets",
             Section.G, FieldTier.USER,       "INR Lakhs", 0.0, row_offset=8),
    FieldDef("wc.non_operating_income",    "Non-Operating Income (Annual Base)",
             Section.G, FieldTier.USER,       "INR Lakhs", 0.0, row_offset=9),
    FieldDef("wc.wip_days",               "Work-in-Progress Days",
             Section.G, FieldTier.USER,       "days", 0, row_offset=10),
    FieldDef("wc.cold_store_days",        "Cold Store / Other Store Days",
             Section.G, FieldTier.USER,       "days", 0, row_offset=11),

    # ── H: Depreciation — statutory rates ────────────────────────────────────
    FieldDef("depr.plant_machinery",       "Plant & Machinery (WDV)",
             Section.H, FieldTier.STATUTORY,  "fraction p.a.", 0.15, row_offset=0),
    FieldDef("depr.civil_works",           "Civil Works / Building (WDV)",
             Section.H, FieldTier.STATUTORY,  "fraction p.a.", 0.10, row_offset=1),
    FieldDef("depr.furniture",             "Furniture & Fixtures (WDV)",
             Section.H, FieldTier.STATUTORY,  "fraction p.a.", 0.10, row_offset=2),
    FieldDef("depr.vehicle",               "Vehicles (WDV)",
             Section.H, FieldTier.STATUTORY,  "fraction p.a.", 0.15, row_offset=3),
    FieldDef("depr.electrical",            "Electrical & Fittings (WDV)",
             Section.H, FieldTier.STATUTORY,  "fraction p.a.", 0.10, row_offset=4),
    FieldDef("depr.other",                 "Other Assets (WDV)",
             Section.H, FieldTier.STATUTORY,  "fraction p.a.", 0.15, row_offset=5),

    # ── I: Implementation ─────────────────────────────────────────────────────
    FieldDef("impl.months",                "Implementation Period (months before COD)",
             Section.I, FieldTier.USER,       "months", 6, row_offset=0),

    # ── J: Tax ───────────────────────────────────────────────────────────────
    FieldDef("tax.entity_type",            "Entity Type",
             Section.J, FieldTier.USER,       "type", "Company", row_offset=0),
    FieldDef("tax.company_rate",           "Company Tax Rate (base)",
             Section.J, FieldTier.STATUTORY,  "fraction", 0.25, row_offset=1),
    FieldDef("tax.surcharge",              "Surcharge",
             Section.J, FieldTier.STATUTORY,  "fraction", 0.07, row_offset=2),
    FieldDef("tax.hec",                    "Health & Education Cess",
             Section.J, FieldTier.STATUTORY,  "fraction", 0.04, row_offset=3),
    FieldDef("tax.llp_rate",              "LLP / Proprietorship Tax Rate",
             Section.J, FieldTier.STATUTORY,  "fraction", 0.30, row_offset=4),
    FieldDef("tax.surcharge_threshold",   "Surcharge Threshold",
             Section.J, FieldTier.STATUTORY,  "INR Lakhs", 100.0, row_offset=5),

    # ── K: Balance Sheet Non-Current Items ───────────────────────────────────
    FieldDef("balance.intangible_assets",  "Intangible Assets (Year 1 Book Value)",
             Section.K, FieldTier.USER,       "INR Lakhs", 0.0, row_offset=0),
    FieldDef("balance.nc_investments",     "Non-Current Investments",
             Section.K, FieldTier.USER,       "INR Lakhs", 0.0, row_offset=1),
    FieldDef("balance.security_deposits",  "Security Deposits",
             Section.K, FieldTier.USER,       "INR Lakhs", 0.0, row_offset=2),
    FieldDef("balance.share_capital_add",  "Additional Share Capital (beyond equity contribution)",
             Section.K, FieldTier.USER,       "INR Lakhs", 0.0, row_offset=3),
]


# ─── Registry lookup helpers ──────────────────────────────────────────────────

_BY_KEY: dict[str, FieldDef] = {f.key: f for f in FIELDS}

# Anchor map — the row each section's data starts on
SECTION_ANCHOR: dict[Section, int] = {
    Section.A: ANCHOR_A,
    Section.B: ANCHOR_B,
    Section.C: ANCHOR_C,
    Section.D: ANCHOR_D,
    Section.E: ANCHOR_E,
    Section.F: ANCHOR_F,
    Section.G: ANCHOR_G,
    Section.H: ANCHOR_H,
    Section.I: ANCHOR_I,
    Section.J: ANCHOR_J,
    Section.K: ANCHOR_K,
}

SECTION_LABEL: dict[Section, str] = {
    Section.A: "A  |  CAPACITY PARAMETERS",
    Section.B: "B  |  REVENUE PARAMETERS  (one block per product / service)",
    Section.C: "C  |  RAW MATERIAL PARAMETERS  (one block per input material)",
    Section.D: "D  |  OPERATING EXPENSE PARAMETERS",
    Section.E: "E  |  MANPOWER PARAMETERS",
    Section.F: "F  |  FINANCE PARAMETERS",
    Section.G: "G  |  WORKING CAPITAL PARAMETERS",
    Section.H: "H  |  DEPRECIATION RATES  (Income Tax Act — WDV Method)",
    Section.I: "I  |  IMPLEMENTATION SCHEDULE",
    Section.J: "J  |  TAX PARAMETERS",
    Section.K: "K  |  BALANCE SHEET — NON-CURRENT ITEMS",
}


def get_field(key: str) -> FieldDef:
    """Return field definition by key. Raises KeyError if not found."""
    return _BY_KEY[key]


def fields_by_tier(tier: FieldTier) -> list[FieldDef]:
    """Return all fields of a given tier."""
    return [f for f in FIELDS if f.tier == tier]


def fields_by_section(section: Section) -> list[FieldDef]:
    """Return all fields in a given section."""
    return [f for f in FIELDS if f.section == section]


def tier2_fields() -> list[FieldDef]:
    """Return all T2 (benchmark) fields."""
    return fields_by_tier(FieldTier.BENCHMARK)


def tier1_fields() -> list[FieldDef]:
    """Return all T1 (user-provided) fields."""
    return fields_by_tier(FieldTier.USER)
