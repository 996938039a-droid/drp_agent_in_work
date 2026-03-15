"""
field_registry.py
─────────────────
The complete registry of every input variable in the DPR model.

Each field is classified into one of three tiers:

  TIER_1 — Business-specific. Agent MUST ask. No benchmark possible.
            Examples: product price, loan amount, asset cost, staff salary.

  TIER_2 — Industry-benchmarkable. Benchmark Engine generates a smart
            default. Agent shows default + asks user to accept or override.
            Examples: power % of revenue, debtor days, R&M rate.

  TIER_3 — Regulatory/statutory. Used silently. Logged in Assumption sheet.
            Examples: IT Act depreciation rates, tax slabs, HEC rate.

This registry is the single source of truth for:
  - What the conversation agent asks
  - In what order
  - With what prompt text
  - How the extracted value maps to the session store
"""

from __future__ import annotations
from dataclasses import dataclass, field
from enum import Enum
from typing import Optional, Any, Callable


class FieldTier(str, Enum):
    TIER_1 = "TIER_1"   # Always ask — user must provide
    TIER_2 = "TIER_2"   # Benchmark default — user confirms or overrides
    TIER_3 = "TIER_3"   # Statutory/regulatory — used silently


class FieldType(str, Enum):
    NUMBER      = "number"      # float
    PERCENT     = "percent"     # float 0.0–1.0 stored as fraction
    INTEGER     = "integer"     # int
    TEXT        = "text"        # str
    DATE        = "date"        # "YYYY-MM" string
    BOOLEAN     = "boolean"     # bool
    ENUM        = "enum"        # one of allowed values


@dataclass
class FieldDef:
    """
    Definition of one input field in the DPR model.
    """
    key: str                        # matches session store path (dot notation)
    tier: FieldTier
    field_type: FieldType
    label: str                      # plain English name
    unit: str                       # unit description for display
    question: str                   # what the agent asks (Tier 1)
    tier2_prompt: str = ""          # what agent says when showing benchmark (Tier 2)
    default: Any = None             # Tier 3 statutory default
    allowed_values: list = field(default_factory=list)  # for ENUM fields
    min_val: Optional[float] = None
    max_val: Optional[float] = None
    section: str = ""               # which conversation section owns this
    depends_on: str = ""            # only ask if this other field is set
    per_product: bool = False       # True = one entry per product
    per_material: bool = False      # True = one entry per raw material
    per_employee: bool = False      # True = one entry per employee category
    per_loan: bool = False          # True = one entry per loan


# ─── Shorthand constructors ──────────────────────────────────────────────────

def t1(key, label, unit, question, field_type=FieldType.NUMBER,
       section="", min_val=None, max_val=None,
       per_product=False, per_material=False, per_employee=False, per_loan=False,
       allowed_values=None):
    return FieldDef(
        key=key, tier=FieldTier.TIER_1, field_type=field_type,
        label=label, unit=unit, question=question, section=section,
        min_val=min_val, max_val=max_val,
        per_product=per_product, per_material=per_material,
        per_employee=per_employee, per_loan=per_loan,
        allowed_values=allowed_values or [],
    )

def t2(key, label, unit, tier2_prompt, field_type=FieldType.PERCENT,
       section="", min_val=None, max_val=None,
       per_product=False, per_material=False, per_employee=False, per_loan=False):
    return FieldDef(
        key=key, tier=FieldTier.TIER_2, field_type=field_type,
        label=label, unit=unit, question="", tier2_prompt=tier2_prompt,
        section=section, min_val=min_val, max_val=max_val,
        per_product=per_product, per_material=per_material,
        per_employee=per_employee, per_loan=per_loan,
    )

def t3(key, label, unit, default, field_type=FieldType.PERCENT, section=""):
    return FieldDef(
        key=key, tier=FieldTier.TIER_3, field_type=field_type,
        label=label, unit=unit, question="", default=default,
        section=section,
    )


# ═══════════════════════════════════════════════════════════════════════════════
#  COMPLETE FIELD REGISTRY
# ═══════════════════════════════════════════════════════════════════════════════

FIELDS: list[FieldDef] = [

    # ── SECTION 0: Intake ─────────────────────────────────────────────────────
    t1("session.business_type",
       "Business Type", "type",
       "To get started, tell me about your business. What does it do, what does it produce or sell?",
       field_type=FieldType.TEXT, section="intake"),

    # ── SECTION 1: Project Profile ────────────────────────────────────────────
    t1("project_profile.company_name",
       "Company Name", "text",
       "What is the name of your company or business?",
       field_type=FieldType.TEXT, section="profile"),

    t1("project_profile.promoter_name",
       "Promoter / Owner Name", "text",
       "Who is the promoter or owner of the business?",
       field_type=FieldType.TEXT, section="profile"),

    t1("project_profile.entity_type",
       "Legal Entity Type", "type",
       "Is the business a Proprietorship, Partnership, LLP, or Private Limited Company?",
       field_type=FieldType.ENUM,
       allowed_values=["Proprietorship","Partnership","LLP","Company"],
       section="profile"),

    t1("project_profile.industry",
       "Industry / Sector", "text",
       "What industry or sector does this project belong to? (e.g. edible oil processing, cold storage, textile weaving)",
       field_type=FieldType.TEXT, section="profile"),

    t1("project_profile.city",
       "City", "text",
       "Which city is the project located in?",
       field_type=FieldType.TEXT, section="profile"),

    t1("project_profile.state",
       "State", "text",
       "Which state?",
       field_type=FieldType.TEXT, section="profile"),

    t1("project_profile.operation_start_date",
       "Commercial Operation Date", "YYYY-MM",
       "When do you expect the business to start commercial operations? (month and year)",
       field_type=FieldType.DATE, section="profile"),

    t1("project_profile.projection_years",
       "Projection Years", "years",
       "How many years should the financial projections cover? (typically 7 years for bank loans)",
       field_type=FieldType.INTEGER, min_val=3, max_val=15, section="profile"),

    # ── SECTION 2: Capital & Means ────────────────────────────────────────────
    t1("capital_means.assets[].name",
       "Asset Name", "text",
       "What assets will be created or purchased for this project? Please list each one (e.g. Civil Works, Plant & Machinery, Furniture).",
       field_type=FieldType.TEXT, section="capital"),

    t1("capital_means.assets[].cost_lakhs",
       "Asset Cost", "INR Lakhs",
       "What is the cost of {asset_name} in lakhs?",
       field_type=FieldType.NUMBER, min_val=0, section="capital", per_product=False),

    t1("capital_means.finance_sources[].term_loan_amount",
       "Term Loan Amount", "INR Lakhs",
       "How much term loan are you seeking from the bank? (in lakhs)",
       field_type=FieldType.NUMBER, min_val=0, section="capital", per_loan=True),

    t1("capital_means.finance_sources[].rate_pa",
       "Term Loan Interest Rate", "% per annum",
       "What is the expected interest rate on the term loan?",
       field_type=FieldType.PERCENT, section="capital", per_loan=True),

    t1("capital_means.finance_sources[].tenor_months",
       "Loan Tenor", "months",
       "What is the total repayment period for the loan? (in months, e.g. 84 for 7 years)",
       field_type=FieldType.INTEGER, min_val=12, max_val=240, section="capital",
       per_loan=True),

    t1("capital_means.finance_sources[].moratorium_months",
       "Moratorium Period", "months",
       "Is there a moratorium period (time before repayments start)? How many months?",
       field_type=FieldType.INTEGER, min_val=0, max_val=36, section="capital",
       per_loan=True),

    t1("capital_means.od_limit",
       "OD / Working Capital Limit", "INR Lakhs",
       "Is there an overdraft or working capital loan from the bank? If yes, what is the limit in lakhs? (Enter 0 if none)",
       field_type=FieldType.NUMBER, min_val=0, section="capital"),

    # ── SECTION 3: Revenue Model ──────────────────────────────────────────────
    t1("revenue_model.products[].name",
       "Product / Service Name", "text",
       "What do you sell? Please name each product or service.",
       field_type=FieldType.TEXT, section="revenue", per_product=True),

    t1("revenue_model.products[].price_per_unit",
       "Selling Price per Unit", "INR per unit",
       "What is the selling price per unit for {product_name}?",
       field_type=FieldType.NUMBER, min_val=0, section="revenue", per_product=True),

    t1("revenue_model.products[].capacity_per_day",
       "Capacity per Day", "units/day",
       "What is the total input or production capacity per day for {product_name}?",
       field_type=FieldType.NUMBER, min_val=0, section="revenue", per_product=True),

    t1("revenue_model.products[].output_ratio",
       "Yield / Output Ratio", "fraction",
       "What fraction of input becomes finished output? (e.g. 0.36 means 36% yield — for {product_name})",
       field_type=FieldType.PERCENT, min_val=0.01, max_val=1.0,
       section="revenue", per_product=True),

    t1("revenue_model.year1_utilization",
       "Year 1 Capacity Utilisation", "fraction",
       "What percentage of full capacity do you expect to use in the first year of operations? (e.g. 0.50 for 50%)",
       field_type=FieldType.PERCENT, min_val=0.1, max_val=1.0, section="revenue"),

    t1("revenue_model.annual_utilization_increment",
       "Annual Utilisation Ramp-up", "fraction per year",
       "By how much will utilisation increase each year? (e.g. 0.05 means 5% additional capacity each year)",
       field_type=FieldType.PERCENT, min_val=0.0, max_val=0.20, section="revenue"),

    t1("revenue_model.max_utilization",
       "Maximum Utilisation Ceiling", "fraction",
       "What is the maximum practical capacity you can ever reach? (e.g. 0.85 for 85% — leaving room for maintenance)",
       field_type=FieldType.PERCENT, min_val=0.5, max_val=1.0, section="revenue"),

    t1("revenue_model.working_days_per_month",
       "Working Days per Month", "days",
       "How many days per month will the business operate?",
       field_type=FieldType.INTEGER, min_val=15, max_val=31, section="revenue"),

    # ── SECTION 4: Cost Structure ─────────────────────────────────────────────
    t1("cost_structure.raw_materials[].name",
       "Raw Material Name", "text",
       "What raw materials or inputs do you purchase? Please list each one.",
       field_type=FieldType.TEXT, section="costs", per_material=True),

    t1("cost_structure.raw_materials[].price_per_unit",
       "Raw Material Price", "INR per unit",
       "What is the current purchase price per unit for {material_name}?",
       field_type=FieldType.NUMBER, min_val=0, section="costs", per_material=True),

    t1("cost_structure.raw_materials[].input_per_output_unit",
       "Input Required per Output Unit", "input units / output unit",
       "How many units of {material_name} are needed to produce one unit of output?",
       field_type=FieldType.NUMBER, min_val=0, section="costs", per_material=True),

    t1("cost_structure.raw_materials[].price_escalation_pa",
       "Raw Material Annual Cost Escalation", "fraction per year",
       "By what percentage do you expect {material_name} costs to rise each year?",
       field_type=FieldType.PERCENT, min_val=0.0, max_val=0.25,
       section="costs", per_material=True),

    t1("cost_structure.transport_base_lakhs",
       "Transportation Cost (Year 1)", "INR Lakhs",
       "What is your estimated annual transportation and logistics cost in the first year? (in lakhs)",
       field_type=FieldType.NUMBER, min_val=0, section="costs"),

    t1("cost_structure.misc_base_lakhs",
       "Miscellaneous Expenses (Year 1)", "INR Lakhs",
       "What is your estimated annual miscellaneous expense in the first year? (in lakhs)",
       field_type=FieldType.NUMBER, min_val=0, section="costs"),

    # ── SECTION 5: Manpower ───────────────────────────────────────────────────
    t1("manpower.categories[].designation",
       "Employee Designation", "text",
       "What employee categories or designations will you have? (e.g. Manager, Operator, Admin, Guard)",
       field_type=FieldType.TEXT, section="manpower", per_employee=True),

    t1("manpower.categories[].count",
       "Number of Employees", "headcount",
       "How many {designation} employees will you have?",
       field_type=FieldType.INTEGER, min_val=1, section="manpower", per_employee=True),

    t1("manpower.categories[].monthly_salary_lakhs",
       "Monthly Salary per Person", "INR Lakhs per month",
       "What is the monthly salary per person for {designation}? (in lakhs, e.g. 0.25 for ₹25,000)",
       field_type=FieldType.NUMBER, min_val=0, section="manpower", per_employee=True),

    # ── SECTION 6: Finance & Working Capital ──────────────────────────────────
    t1("finance_wc.debtor_days",
       "Debtor Days", "days",
       "How many days after a sale do your customers typically pay you? (e.g. 10 for cash-near business, 45 for credit-heavy)",
       field_type=FieldType.INTEGER, min_val=0, max_val=120, section="finance"),

    t1("finance_wc.creditor_days_rm",
       "Creditor Days — Raw Materials", "days",
       "How many days do you take to pay your raw material suppliers?",
       field_type=FieldType.INTEGER, min_val=0, max_val=90, section="finance"),

    t1("finance_wc.implementation_months",
       "Implementation Period", "months",
       "How many months will it take to complete construction and installation before the business starts?",
       field_type=FieldType.INTEGER, min_val=1, max_val=36, section="finance"),

    # ── TIER 2: Benchmarkable fields ─────────────────────────────────────────
    t2("cost_structure.rm_pct_of_fa",
       "Repair & Maintenance Rate", "% of net fixed assets",
       "For repair & maintenance, {benchmark_midpoint}% of net fixed assets is typical for {business_description}. "
       "This covers routine servicing of machinery and premises.",
       section="costs"),

    t2("cost_structure.rm_escalation_pa",
       "R&M Cost Escalation", "% per year",
       "R&M costs typically escalate at {benchmark_midpoint}% per year for {business_description}.",
       field_type=FieldType.PERCENT, section="costs"),

    t2("cost_structure.insurance_pct_of_fa",
       "Insurance Rate", "% of fixed assets",
       "Insurance premiums are typically {benchmark_midpoint}% of fixed assets annually for {business_description}.",
       section="costs"),

    t2("cost_structure.power_pct_revenue",
       "Power & Fuel as % of Revenue", "% of revenue",
       "Power & fuel costs for {business_description} typically run {benchmark_midpoint}% of revenue "
       "(range: {benchmark_range}). This varies significantly by energy intensity.",
       section="costs"),

    t2("cost_structure.power_escalation_pa",
       "Power Cost Annual Escalation", "% per year",
       "Power costs in India typically escalate at {benchmark_midpoint}% per year.",
       section="costs"),

    t2("cost_structure.marketing_pct_revenue",
       "Marketing Expenses as % of Revenue", "% of revenue",
       "For {business_description}, marketing & sales expenses are typically {benchmark_midpoint}% of revenue. "
       "({benchmark_range}) — B2B businesses tend to spend less than B2C.",
       section="costs"),

    t2("cost_structure.sga_base_lakhs",
       "Selling, General & Admin Expenses (Base Year)", "INR Lakhs",
       "General admin overheads for a business of this scale are typically ₹{benchmark_midpoint}L in year 1.",
       field_type=FieldType.NUMBER, section="costs"),

    t2("cost_structure.transport_escalation_pa",
       "Transport Cost Annual Escalation", "% per year",
       "Transportation costs typically escalate at {benchmark_midpoint}% per year (fuel price-linked).",
       section="costs"),

    t2("cost_structure.misc_escalation_pa",
       "Miscellaneous Expenses Escalation", "% per year",
       "Miscellaneous expenses typically grow at {benchmark_midpoint}% per year for {business_description}.",
       section="costs"),

    t2("manpower.categories[].annual_increment_pa",
       "Annual Salary Increment", "% per year",
       "Annual salary increments in {business_description} are typically {benchmark_midpoint}% per year.",
       field_type=FieldType.PERCENT, section="manpower", per_employee=True),

    t2("finance_wc.creditor_days_admin",
       "Creditor Days — Admin Expenses", "days",
       "Admin expense payment cycles are typically {benchmark_midpoint} days for {business_description}.",
       field_type=FieldType.INTEGER, section="finance"),

    t2("finance_wc.stock_days_rm",
       "Raw Material Stock Days", "days",
       "{business_description} typically holds {benchmark_midpoint} days of raw material inventory. "
       "({benchmark_range}) — depends on supply chain reliability.",
       field_type=FieldType.INTEGER, section="finance"),

    t2("finance_wc.od_rate",
       "OD / CC Interest Rate", "% per annum",
       "Overdraft / CC rates in India are typically {benchmark_midpoint}% for MSME accounts.",
       field_type=FieldType.PERCENT, section="finance"),

    t2("revenue_model.products[].price_escalation_pa",
       "Annual Selling Price Escalation", "% per year",
       "Selling prices for {product_name} in {business_description} typically rise {benchmark_midpoint}% per year. "
       "({benchmark_range})",
       field_type=FieldType.PERCENT, section="revenue", per_product=True),

    # ── TIER 3: Statutory / regulatory — used silently ────────────────────────
    t3("depreciation_rates.plant_machinery",
       "P&M Depreciation Rate (IT Act)", "% WDV per annum",
       default=0.15, section="system"),

    t3("depreciation_rates.civil_works",
       "Civil Works Depreciation Rate (IT Act)", "% WDV per annum",
       default=0.10, section="system"),

    t3("depreciation_rates.furniture",
       "Furniture Depreciation Rate (IT Act)", "% WDV per annum",
       default=0.10, section="system"),

    t3("depreciation_rates.vehicle",
       "Vehicle Depreciation Rate (IT Act)", "% WDV per annum",
       default=0.15, section="system"),

    t3("depreciation_rates.electrical",
       "Electrical Fittings Depreciation Rate (IT Act)", "% WDV per annum",
       default=0.10, section="system"),

    t3("depreciation_rates.pre_operative",
       "Pre-operative Expenses Depreciation Rate (IT Act)", "% WDV per annum",
       default=0.20, section="system"),

    t3("tax_config.company_basic_rate",
       "Company Tax Rate", "fraction",
       default=0.30, section="system"),

    t3("tax_config.hec_rate",
       "Health & Education Cess", "fraction",
       default=0.04, section="system"),

    t3("tax_config.surcharge_rate_1cr_10cr",
       "Surcharge Rate (₹1Cr–10Cr income)", "fraction",
       default=0.07, section="system"),

    t3("tax_config.surcharge_rate_above_10cr",
       "Surcharge Rate (above ₹10Cr income)", "fraction",
       default=0.12, section="system"),

    t3("tax_config.partnership_rate",
       "Partnership Firm Tax Rate", "fraction",
       default=0.30, section="system"),

    t3("tax_config.partnership_surcharge_rate",
       "Partnership Surcharge Rate", "fraction",
       default=0.12, section="system"),

    t3("finance_wc.wc_interest_rate",
       "Working Capital Loan Interest Rate", "fraction",
       default=0.0, section="system"),    # 0 until OD is used
]


# ─── Index helpers ────────────────────────────────────────────────────────────

FIELD_MAP: dict[str, FieldDef] = {f.key: f for f in FIELDS}

TIER_1_FIELDS = [f for f in FIELDS if f.tier == FieldTier.TIER_1]
TIER_2_FIELDS = [f for f in FIELDS if f.tier == FieldTier.TIER_2]
TIER_3_FIELDS = [f for f in FIELDS if f.tier == FieldTier.TIER_3]

def fields_for_section(section: str) -> list[FieldDef]:
    return [f for f in FIELDS if f.section == section]

def tier1_for_section(section: str) -> list[FieldDef]:
    return [f for f in FIELDS
            if f.section == section and f.tier == FieldTier.TIER_1]

def tier2_for_section(section: str) -> list[FieldDef]:
    return [f for f in FIELDS
            if f.section == section and f.tier == FieldTier.TIER_2]
