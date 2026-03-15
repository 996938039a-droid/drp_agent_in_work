"""
session_store.py
────────────────
Canonical data contract for the DPR Agent system.
Every agent reads from and writes to a SessionStore instance.
No agent communicates directly with another — only through this store.

Units: All monetary values in INR Lakhs unless explicitly noted.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any
from enum import Enum
from datetime import datetime
import json
import uuid


# ═══════════════════════════════════════════════════════════════════════════════
#  ENUMS
# ═══════════════════════════════════════════════════════════════════════════════

class BusinessType(str, Enum):
    MANUFACTURING  = "MANUFACTURING"
    TRADING        = "TRADING"
    SERVICE        = "SERVICE"
    MIXED          = "MIXED"
    UNKNOWN        = "UNKNOWN"


class RevenueModelType(str, Enum):
    """Six canonical revenue model patterns the system supports."""
    MANUFACTURING       = "MANUFACTURING"        # Capacity × Days × Yield × Util × Price
    TRADING             = "TRADING"              # Purchase × (1 + Margin%) × Volume
    SERVICE_CAPACITY    = "SERVICE_CAPACITY"     # Units × Occupancy% × Rate × Days
    SERVICE_TIME        = "SERVICE_TIME"         # Staff × Hours × Util% × Rate
    SERVICE_SUBSCRIPTION= "SERVICE_SUBSCRIPTION" # Customers × Fee × 12
    MIXED               = "MIXED"                # Sum of sub-models


class EntityType(str, Enum):
    PROPRIETORSHIP  = "Proprietorship"
    PARTNERSHIP     = "Partnership"
    LLP             = "LLP"
    COMPANY         = "Company"


class FinanceSourceType(str, Enum):
    PROMOTER_EQUITY = "Promoter Equity"
    TERM_LOAN       = "Term Loan"
    OD_LIMIT        = "OD / CC Limit"
    SUBSIDY         = "Subsidy / Grant"
    UNSECURED_LOAN  = "Unsecured Loan"


class AssetCategory(str, Enum):
    CIVIL_WORKS      = "Civil Works"
    PLANT_MACHINERY  = "Plant & Machinery"
    FURNITURE        = "Furniture & Fixture"
    VEHICLE          = "Vehicle"
    ELECTRICAL       = "Electrical & Fittings"
    PRE_OPERATIVE    = "Pre-operative Expenses"
    OTHER            = "Other"


class SectionStatus(str, Enum):
    PENDING    = "PENDING"
    IN_PROGRESS= "IN_PROGRESS"
    COMPLETE   = "COMPLETE"


# ═══════════════════════════════════════════════════════════════════════════════
#  SUB-DATACLASSES
# ═══════════════════════════════════════════════════════════════════════════════

@dataclass
class Asset:
    name: str
    category: AssetCategory
    cost_lakhs: float
    disbursement_schedule: Dict[int, float] = field(default_factory=lambda: {1: 1.0})
    # month_number -> fraction (must sum to 1.0)

    def validate(self):
        assert self.cost_lakhs > 0, f"Asset cost must be > 0: {self.name}"
        total = sum(self.disbursement_schedule.values())
        assert abs(total - 1.0) < 0.001, f"Disbursement schedule must sum to 1.0 for {self.name}, got {total}"


@dataclass
class FinanceSource:
    source_type: FinanceSourceType
    amount_lakhs: float
    rate_pa: float = 0.0           # annual interest rate (0.09 = 9%)
    tenor_months: int = 0          # total repayment period
    moratorium_months: int = 0     # interest-only period before repayment
    label: str = ""                # e.g. "SBI Term Loan"

    @property
    def repayment_months(self) -> int:
        return max(self.tenor_months - self.moratorium_months, 1)

    @property
    def monthly_emi_principal(self) -> float:
        """Flat principal repayment per month after moratorium."""
        return self.amount_lakhs / self.repayment_months


@dataclass
class Product:
    """One revenue-generating product or service stream."""
    name: str
    unit: str                          # e.g. "litres", "kg", "rooms", "hours"

    # MANUFACTURING fields
    capacity_per_day: float = 0.0     # total input capacity per day (in input units)
    output_ratio: float = 1.0         # fraction of input that becomes output
    split_percent: float = 1.0        # fraction of total output that is this product

    # TRADING fields
    purchase_volume_per_year: float = 0.0
    purchase_price: float = 0.0
    margin_percent: float = 0.0

    # SERVICE_CAPACITY fields
    capacity_units: float = 0.0       # e.g. number of beds, rooms, seats
    occupancy_rate_year1: float = 0.0

    # SERVICE_TIME fields
    staff_count: float = 0.0
    hours_per_day: float = 0.0
    utilization_rate: float = 0.0

    # SERVICE_SUBSCRIPTION fields
    subscribers_year1: float = 0.0
    monthly_fee: float = 0.0
    subscriber_growth_rate: float = 0.0

    # COMMON
    price_per_unit: float = 0.0       # selling price per unit of output
    price_escalation_pa: float = 0.04 # annual price increase


@dataclass
class RawMaterial:
    name: str
    unit: str
    price_per_unit: float
    input_per_output_unit: float      # how many input units per output unit
    price_escalation_pa: float = 0.05


@dataclass
class EmployeeCategory:
    designation: str
    count: int
    monthly_salary_lakhs: float
    is_fixed: bool = True             # True = fixed cost, False = variable
    annual_increment_pa: float = 0.05


@dataclass
class DepreciationRates:
    plant_machinery: float  = 0.15
    civil_works: float      = 0.10
    furniture: float        = 0.10
    vehicle: float          = 0.15
    electrical: float       = 0.10
    pre_operative: float    = 0.20
    other: float            = 0.15

    def rate_for(self, category: AssetCategory) -> float:
        mapping = {
            AssetCategory.CIVIL_WORKS:     self.civil_works,
            AssetCategory.PLANT_MACHINERY: self.plant_machinery,
            AssetCategory.FURNITURE:       self.furniture,
            AssetCategory.VEHICLE:         self.vehicle,
            AssetCategory.ELECTRICAL:      self.electrical,
            AssetCategory.PRE_OPERATIVE:   self.pre_operative,
            AssetCategory.OTHER:           self.other,
        }
        return mapping.get(category, self.other)


# ═══════════════════════════════════════════════════════════════════════════════
#  SECTION STORES
# ═══════════════════════════════════════════════════════════════════════════════

@dataclass
class ProjectProfile:
    company_name: str           = ""
    promoter_name: str          = ""
    entity_type: EntityType     = EntityType.PROPRIETORSHIP
    industry: str               = ""
    city: str                   = ""
    state: str                  = ""
    operation_start_date: str   = ""   # "YYYY-MM" format
    projection_years: int       = 7
    status: SectionStatus       = SectionStatus.PENDING


@dataclass
class CapitalMeans:
    assets: List[Asset]               = field(default_factory=list)
    finance_sources: List[FinanceSource] = field(default_factory=list)
    status: SectionStatus             = SectionStatus.PENDING

    @property
    def total_project_cost(self) -> float:
        return sum(a.cost_lakhs for a in self.assets)

    @property
    def total_finance(self) -> float:
        return sum(f.amount_lakhs for f in self.finance_sources)

    @property
    def is_balanced(self) -> bool:
        return abs(self.total_project_cost - self.total_finance) < 0.01

    @property
    def term_loans(self) -> List[FinanceSource]:
        return [f for f in self.finance_sources
                if f.source_type == FinanceSourceType.TERM_LOAN]

    @property
    def od_sources(self) -> List[FinanceSource]:
        return [f for f in self.finance_sources
                if f.source_type == FinanceSourceType.OD_LIMIT]

    @property
    def promoter_contribution(self) -> float:
        return sum(f.amount_lakhs for f in self.finance_sources
                   if f.source_type == FinanceSourceType.PROMOTER_EQUITY)


@dataclass
class RevenueModel:
    model_type: RevenueModelType      = RevenueModelType.MANUFACTURING
    working_days_per_month: int       = 26
    year1_utilization: float          = 0.50
    annual_utilization_increment: float = 0.05
    max_utilization: float            = 0.85
    products: List[Product]           = field(default_factory=list)
    status: SectionStatus             = SectionStatus.PENDING

    def utilization_for_year(self, year: int) -> float:
        """1-indexed year. Returns capacity utilisation fraction."""
        util = self.year1_utilization
        for _ in range(year - 1):
            if util < self.max_utilization:
                util = min(util + self.annual_utilization_increment,
                           self.max_utilization)
        return util


@dataclass
class CostStructure:
    raw_materials: List[RawMaterial]  = field(default_factory=list)
    rm_pct_of_fa: float               = 0.02    # R&M as % of net fixed assets
    rm_escalation_pa: float           = 0.10
    insurance_pct_of_fa: float        = 0.01
    insurance_escalation_pa: float    = 0.10
    power_pct_revenue: float          = 0.06
    power_escalation_pa: float        = 0.05
    marketing_pct_revenue: float      = 0.05
    marketing_escalation_pa: float    = 0.00    # usually kept constant
    transport_base_lakhs: float       = 0.0
    transport_escalation_pa: float    = 0.15
    misc_base_lakhs: float            = 0.0
    misc_escalation_pa: float         = 0.20
    sga_base_lakhs: float             = 0.0
    sga_escalation_pa: float          = 0.10
    status: SectionStatus             = SectionStatus.PENDING


@dataclass
class ManpowerStructure:
    categories: List[EmployeeCategory] = field(default_factory=list)
    status: SectionStatus              = SectionStatus.PENDING

    @property
    def total_annual_fixed_salary_lakhs(self) -> float:
        return sum(e.count * e.monthly_salary_lakhs * 12
                   for e in self.categories if e.is_fixed)

    @property
    def total_annual_variable_salary_lakhs(self) -> float:
        return sum(e.count * e.monthly_salary_lakhs * 12
                   for e in self.categories if not e.is_fixed)


@dataclass
class FinanceWorkingCapital:
    # Working capital days
    debtor_days: int            = -1
    creditor_days_rm: int       = -1
    creditor_days_admin: int    = 30
    stock_days_rm: int          = 7
    stock_days_fg: int          = 0
    wc_loan_amount: float       = 0.0
    wc_interest_rate: float     = 0.0
    # Implementation schedule (months before COD)
    implementation_months: int  = -1
    status: SectionStatus       = SectionStatus.PENDING


@dataclass
class TaxConfig:
    entity_type: EntityType     = EntityType.COMPANY
    # These are standard rates — only change if regulatory update
    company_basic_rate: float   = 0.30
    company_surcharge_threshold_cr: float = 1.0   # 1 crore
    surcharge_rate_1cr_10cr: float  = 0.07
    surcharge_rate_above_10cr: float = 0.12
    hec_rate: float             = 0.04
    individual_slabs: List[tuple] = field(default_factory=lambda: [
        (250000,  0.00),
        (500000,  0.05),
        (750000,  0.10),
        (1000000, 0.15),
        (1250000, 0.20),
        (1500000, 0.25),
        (float('inf'), 0.30),
    ])
    partnership_rate: float     = 0.30
    partnership_surcharge_threshold: float = 10000000  # 1 crore
    partnership_surcharge_rate: float = 0.12
    status: SectionStatus       = SectionStatus.PENDING


# ═══════════════════════════════════════════════════════════════════════════════
#  MASTER SESSION STORE
# ═══════════════════════════════════════════════════════════════════════════════

@dataclass
class SessionStore:
    """
    The single shared data contract for the entire DPR Agent system.
    All agents read from and write to an instance of this class.
    """
    session_id: str                         = field(default_factory=lambda: str(uuid.uuid4()))
    created_at: str                         = field(default_factory=lambda: datetime.now().isoformat())
    business_type: BusinessType             = BusinessType.UNKNOWN
    revenue_model_type: RevenueModelType    = RevenueModelType.MANUFACTURING
    conversation_language: str              = "english"

    # Section stores
    project_profile:    ProjectProfile         = field(default_factory=ProjectProfile)
    capital_means:      CapitalMeans           = field(default_factory=CapitalMeans)
    revenue_model:      RevenueModel           = field(default_factory=RevenueModel)
    cost_structure:     CostStructure          = field(default_factory=CostStructure)
    manpower:           ManpowerStructure      = field(default_factory=ManpowerStructure)
    finance_wc:         FinanceWorkingCapital  = field(default_factory=FinanceWorkingCapital)
    depreciation_rates: DepreciationRates      = field(default_factory=DepreciationRates)
    tax_config:         TaxConfig              = field(default_factory=TaxConfig)

    # Generation outputs
    excel_path: Optional[str]   = None
    document_path: Optional[str]= None
    validation_results: Dict    = field(default_factory=dict)
    generation_errors: List[str]= field(default_factory=list)

    # ── Convenience Properties ──────────────────────────────────────────────

    @property
    def n_products(self) -> int:
        return len(self.revenue_model.products)

    @property
    def n_materials(self) -> int:
        return len(self.cost_structure.raw_materials)

    @property
    def n_employee_categories(self) -> int:
        return len(self.manpower.categories)

    @property
    def n_term_loans(self) -> int:
        return len(self.capital_means.term_loans)

    @property
    def n_years(self) -> int:
        return self.project_profile.projection_years

    @property
    def n_assets_by_category(self) -> Dict[AssetCategory, List[Asset]]:
        result: Dict[AssetCategory, List[Asset]] = {}
        for a in self.capital_means.assets:
            result.setdefault(a.category, []).append(a)
        return result

    @property
    def completed_sections(self) -> List[str]:
        sections = {
            "project_profile":  self.project_profile.status,
            "capital_means":    self.capital_means.status,
            "revenue_model":    self.revenue_model.status,
            "cost_structure":   self.cost_structure.status,
            "manpower":         self.manpower.status,
            "finance_wc":       self.finance_wc.status,
            "tax_config":       self.tax_config.status,
        }
        return [k for k, v in sections.items() if v == SectionStatus.COMPLETE]

    @property
    def is_ready_for_generation(self) -> bool:
        required = ["project_profile", "capital_means", "revenue_model",
                    "cost_structure", "manpower", "finance_wc"]
        return all(s in self.completed_sections for s in required)

    # ── Serialization ────────────────────────────────────────────────────────

    def to_dict(self) -> Dict[str, Any]:
        """Serialize for storage / API transmission."""
        import dataclasses
        def _convert(obj):
            if isinstance(obj, (BusinessType, RevenueModelType, EntityType,
                                FinanceSourceType, AssetCategory, SectionStatus)):
                return obj.value
            if dataclasses.is_dataclass(obj):
                return {k: _convert(v) for k, v in dataclasses.asdict(obj).items()}
            if isinstance(obj, list):
                return [_convert(i) for i in obj]
            if isinstance(obj, dict):
                return {k: _convert(v) for k, v in obj.items()}
            return obj
        return _convert(self)

    def to_json(self, indent: int = 2) -> str:
        return json.dumps(self.to_dict(), indent=indent)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'SessionStore':
        """Deserialize from storage."""
        # Simplified loader — full implementation would recursively reconstruct
        store = cls()
        store.session_id = data.get("session_id", store.session_id)
        store.business_type = BusinessType(data.get("business_type", "UNKNOWN"))
        store.revenue_model_type = RevenueModelType(
            data.get("revenue_model_type", "MANUFACTURING"))
        return store

    # ── Validation ────────────────────────────────────────────────────────────

    def validate_completeness(self) -> List[str]:
        """Returns list of missing required fields."""
        errors = []
        pp = self.project_profile
        if not pp.company_name:   errors.append("Missing: company_name")
        if not pp.promoter_name:  errors.append("Missing: promoter_name")
        if not pp.operation_start_date: errors.append("Missing: operation_start_date")
        if pp.projection_years < 1: errors.append("projection_years must be >= 1")

        cm = self.capital_means
        if not cm.assets:         errors.append("Missing: no assets defined")
        if not cm.finance_sources: errors.append("Missing: no finance sources defined")
        if not cm.is_balanced:
            gap = cm.total_project_cost - cm.total_finance
            errors.append(f"Capital imbalance: cost-means gap = {gap:.2f} Lakhs")

        rm = self.revenue_model
        if not rm.products:       errors.append("Missing: no products/services defined")
        if rm.year1_utilization <= 0: errors.append("year1_utilization must be > 0")

        return errors
