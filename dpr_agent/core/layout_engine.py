"""
layout_engine.py
────────────────
The Layout Engine is the single most critical module in the Excel Generator.

It computes the EXACT row and column number of every logical cell in every
sheet, given the dynamic input counts (products, materials, employees, loans).

Every formula written to the workbook is generated via:
    layout.ref("Revenue", "total_revenue", year=1)
    → returns "=D47" or whatever the actual address is

This guarantees formula correctness regardless of how many rows expand
or contract due to dynamic content.

Design principles:
  - No hardcoded cell addresses anywhere outside this file
  - Every sheet has a "layout map": logical_name → (row, col)
  - Column letters for year columns: col_for_year(year) → "D", "E", ...
  - All row computations are pure functions of (n_products, n_materials, etc.)
"""

from __future__ import annotations
from typing import Dict, Tuple, Optional, List
from openpyxl.utils import get_column_letter
from core.session_store import (
    SessionStore, AssetCategory, RevenueModelType
)


# ═══════════════════════════════════════════════════════════════════════════════
#  CONSTANTS — Sheet column layout
# ═══════════════════════════════════════════════════════════════════════════════

# Standard column assignments (1-indexed, openpyxl convention)
COL_LABEL       = 2   # B  — row labels / descriptions
COL_BASIS       = 3   # C  — basis / rate column
COL_YEAR_START  = 4   # D  — Year 1 data starts here

# Maximum allowed counts (unused rows are hidden with height=0, value=0)
MAX_PRODUCTS        = 15
MAX_MATERIALS       = 15
MAX_EMPLOYEES       = 10
MAX_TERM_LOANS      = 5
MAX_ASSET_CATEGORIES= 8    # per the AssetCategory enum

# Rows per product block in Revenue sheet
ROWS_PER_PRODUCT    = 6    # sub-header(1) + volume(1) + liters(1) + split(1) + price(1) + revenue(1)

# Rows per material block in Expenses sheet
ROWS_PER_MATERIAL   = 3    # quantity, price, cost

# Rows per employee category in ManPower sheet
ROWS_PER_EMPLOYEE   = 2    # count+salary row, annual cost row

# Rows per asset class in Depreciation sheet
ROWS_PER_ASSET_CLASS= 5    # opening, addition, depreciation, closing, blank


# ═══════════════════════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

def col_letter(col: int) -> str:
    """Convert 1-indexed column number to Excel letter. 1→A, 2→B, 28→AB."""
    return get_column_letter(col)


def year_col(year: int) -> int:
    """1-indexed year → column number. Year 1 → COL_YEAR_START (4=D)."""
    return COL_YEAR_START + (year - 1)


def cell_addr(row: int, col: int) -> str:
    """Return Excel address string like 'D17'."""
    return f"{col_letter(col)}{row}"


def cell_range(row: int, col_start: int, col_end: int) -> str:
    """Return range like 'D17:J17'."""
    return f"{col_letter(col_start)}{row}:{col_letter(col_end)}{row}"


def col_range(row_start: int, row_end: int, col: int) -> str:
    """Return range like 'D5:D17'."""
    return f"{col_letter(col)}{row_start}:{col_letter(col)}{row_end}"


def xref(sheet: str, row: int, col: int, abs_row=False, abs_col=False) -> str:
    """Cross-sheet reference. Returns e.g. =Revenue!D23 or =Revenue!$D$23."""
    c = f"${col_letter(col)}$" if abs_col else col_letter(col)
    r = f"${row}" if abs_row else str(row)
    return f"'{sheet}'!{c}{r}"


# ═══════════════════════════════════════════════════════════════════════════════
#  PER-SHEET LAYOUT DEFINITIONS
# ═══════════════════════════════════════════════════════════════════════════════

class AssumptionLayout:
    """
    Assumption sheet layout.
    The sheet is divided into 8 sections, each a labeled block.
    ALL parameters from every other sheet reference back to here.
    """
    SHEET = "Assumption"

    # Section header rows — spaced to accommodate dynamic content above
    # Expenses: 7 types × 2 rows each = 14 rows (34–53)
    # Manpower header must come AFTER expenses end
    HDR_CAPACITY        = 3
    HDR_REVENUE         = 12
    HDR_RAWMATERIAL     = 22
    HDR_EXPENSES        = 33
    HDR_MANPOWER        = 56    # after expenses (row 53) + 2 blank rows
    HDR_FINANCE         = 68    # after manpower (up to 10 cats × 1 row = 10 rows → 56+10+2)
    HDR_WORKING_CAPITAL = 82    # after finance (up to 5 loans × 5 rows = 25 → 68+12+2)
    HDR_DEPRECIATION    = 93    # after WC (8 rows → 82+8+2)
    HDR_IMPL_SCHEDULE   = 102   # after depreciation (6 rows → 93+6+2)

    # ── Section A: Capacity ─────────────────────────────────────────────────
    CAP_YEAR1_UTIL      = (4,  5)   # E4
    CAP_ANNUAL_INCREMENT= (5,  5)   # E5
    CAP_MAX_UTIL        = (6,  5)   # E6
    CAP_WORKING_DAYS    = (7,  5)   # E7
    CAP_MONTHS_YEAR     = (8,  5)   # E8  (always 12)

    # ── Section B: Revenue (per product, dynamic) ───────────────────────────
    # Product i (0-indexed) lives at row REV_BASE_ROW + i*3
    REV_BASE_ROW        = 14
    REV_ROWS_PER_PRODUCT= 3
    REV_COL_NAME        = 2   # B — product name
    REV_COL_PRICE       = 5   # E — base price
    REV_COL_ESCALATION  = 6   # F — price escalation

    @classmethod
    def rev_price_row(cls, product_idx: int) -> int:
        return cls.REV_BASE_ROW + product_idx * cls.REV_ROWS_PER_PRODUCT

    @classmethod
    def rev_escalation_row(cls, product_idx: int) -> int:
        return cls.REV_BASE_ROW + product_idx * cls.REV_ROWS_PER_PRODUCT + 1

    # ── Section C: Raw Material (per material, dynamic) ─────────────────────
    # Computed dynamically based on n_products
    @classmethod
    def rm_base_row(cls, n_products: int) -> int:
        return cls.HDR_RAWMATERIAL + 1 + n_products * cls.REV_ROWS_PER_PRODUCT

    RM_ROWS_PER_MATERIAL = 2
    RM_COL_NAME          = 2
    RM_COL_PRICE         = 5
    RM_COL_ESCALATION    = 6

    @classmethod
    def rm_price_row(cls, material_idx: int, n_products: int) -> int:
        base = cls.rm_base_row(n_products)
        return base + material_idx * cls.RM_ROWS_PER_MATERIAL

    # ── Section D: Expenses ─────────────────────────────────────────────────
    # Rate row + escalation row pairs (each pair occupies 2 consecutive rows)
    # Layout: row N = rate/base value, row N+1 = escalation rate
    EXP_RM_PCT_FA       = (34, 5)
    EXP_RM_ESCALATION   = (35, 5)   # row 35 = one below rate
    EXP_INS_PCT_FA      = (37, 5)
    EXP_INS_ESCALATION  = (38, 5)
    EXP_POWER_PCT_REV   = (40, 5)
    EXP_POWER_ESCALATION= (41, 5)
    EXP_MKT_PCT_REV     = (43, 5)
    EXP_MKT_ESCALATION  = (44, 5)
    EXP_TRANSPORT_BASE  = (46, 5)
    EXP_TRANSPORT_ESC   = (47, 5)
    EXP_MISC_BASE       = (49, 5)
    EXP_MISC_ESC        = (50, 5)
    EXP_SGA_BASE        = (52, 5)
    EXP_SGA_ESC         = (53, 5)

    # ── Section E: Manpower ──────────────────────────────────────────────────
    MP_BASE_ROW         = 58    # HDR_MANPOWER + 2 (sub-header row + first data)
    MP_ROWS_PER_CAT     = 1     # one row per employee category
    MP_COL_DESIG        = 2
    MP_COL_COUNT        = 4
    MP_COL_SALARY       = 5
    MP_COL_INCREMENT    = 6

    @classmethod
    def mp_count_row(cls, cat_idx: int) -> int:
        return cls.MP_BASE_ROW + cat_idx * cls.MP_ROWS_PER_CAT

    # ── Section F: Finance ───────────────────────────────────────────────────
    FIN_BASE_ROW        = 70    # HDR_FINANCE + 2
    FIN_ROWS_PER_LOAN   = 5     # amount, rate, tenor, moratorium, repayment_months
    FIN_COL_LABEL       = 2
    FIN_COL_VALUE       = 5

    @classmethod
    def fin_amount_row(cls, loan_idx: int) -> int:
        return cls.FIN_BASE_ROW + loan_idx * cls.FIN_ROWS_PER_LOAN

    @classmethod
    def fin_rate_row(cls, loan_idx: int) -> int:
        return cls.FIN_BASE_ROW + loan_idx * cls.FIN_ROWS_PER_LOAN + 1

    @classmethod
    def fin_tenor_row(cls, loan_idx: int) -> int:
        return cls.FIN_BASE_ROW + loan_idx * cls.FIN_ROWS_PER_LOAN + 2

    @classmethod
    def fin_moratorium_row(cls, loan_idx: int) -> int:
        return cls.FIN_BASE_ROW + loan_idx * cls.FIN_ROWS_PER_LOAN + 3

    FIN_OD_LIMIT_ROW    = 83    # OD limit at end of finance section
    FIN_OD_RATE_ROW     = 84

    # ── Section G: Working Capital ───────────────────────────────────────────
    # HDR_WORKING_CAPITAL = 82, data starts at 84
    WC_DEBTOR_DAYS      = (87, 5)
    WC_CREDITOR_RM      = (88, 5)
    WC_CREDITOR_ADMIN   = (89, 5)
    WC_STOCK_RM         = (90, 5)
    WC_STOCK_FG         = (91, 5)
    WC_LOAN_AMOUNT      = (92, 5)
    WC_INTEREST_RATE    = (93, 5)

    # ── Section H: Depreciation ──────────────────────────────────────────────
    DEP_PM_RATE         = (97, 5)
    DEP_CIVIL_RATE      = (98, 5)
    DEP_FURN_RATE       = (99, 5)
    DEP_VEH_RATE        = (100, 5)
    DEP_ELEC_RATE       = (101, 5)
    DEP_PREOP_RATE      = (102, 5)

    # ── Section I: Implementation Schedule ───────────────────────────────────
    IMPL_MONTHS         = (106, 5)


class RevenueLayout:
    """
    Revenue sheet layout.
    Dynamic: each product gets ROWS_PER_PRODUCT rows.
    """
    SHEET = "Revenue"

    HDR_ROW         = 2
    SUBHDR_ROW      = 5
    MONTHS_ROW      = 7
    DAYS_ROW        = 8
    CAPACITY_ROW    = 9    # total capacity per day
    UTIL_ROW        = 12   # operational utilization %

    # Production blocks (1 per product)
    PROD_BASE_ROW   = 13   # first product starts here
    # Each product block:
    #   +0: Total production in tons (= months × days × capacity × output_ratio × util)
    #   +1: Total production in liters (tons × 1000)
    #   +2: Product liters (liters × split%)
    #   +3: Price per unit
    #   +4: Revenue (product_liters × price) / 10^5

    @classmethod
    def prod_volume_row(cls, idx: int) -> int:
        return cls.PROD_BASE_ROW + idx * ROWS_PER_PRODUCT

    @classmethod
    def prod_liters_row(cls, idx: int) -> int:
        return cls.PROD_BASE_ROW + idx * ROWS_PER_PRODUCT + 1

    @classmethod
    def prod_split_row(cls, idx: int) -> int:
        return cls.PROD_BASE_ROW + idx * ROWS_PER_PRODUCT + 2

    @classmethod
    def prod_price_row(cls, idx: int) -> int:
        return cls.PROD_BASE_ROW + idx * ROWS_PER_PRODUCT + 3

    @classmethod
    def prod_revenue_row(cls, idx: int) -> int:
        return cls.PROD_BASE_ROW + idx * ROWS_PER_PRODUCT + 4

    @classmethod
    def total_revenue_row(cls, n_products: int) -> int:
        """Total Revenue row comes after all product blocks."""
        return cls.PROD_BASE_ROW + n_products * ROWS_PER_PRODUCT + 1


class ManPowerLayout:
    """ManPower sheet layout."""
    SHEET = "ManPower"
    HDR_ROW         = 2
    SUBHDR_ROW      = 5
    BASE_ROW        = 7

    @classmethod
    def emp_row(cls, idx: int) -> int:
        return cls.BASE_ROW + idx

    @classmethod
    def annual_fixed_row(cls, n_employees: int) -> int:
        return cls.BASE_ROW + n_employees + 2

    @classmethod
    def annual_variable_row(cls, n_employees: int) -> int:
        return cls.BASE_ROW + n_employees + 3

    @classmethod
    def pl_fixed_row(cls, n_employees: int) -> int:
        """Row that feeds into P&L (after months-in-operation scaling)."""
        return cls.BASE_ROW + n_employees + 6

    @classmethod
    def pl_variable_row(cls, n_employees: int) -> int:
        return cls.BASE_ROW + n_employees + 7


class DepreciationLayout:
    """Depreciation sheet layout. One block per asset category."""
    SHEET = "Depreciation"
    HDR_ROW     = 2
    SUBHDR_ROW  = 5
    MONTHS_ROW  = 7
    BASE_ROW    = 9

    # Each asset class block:
    #   +0: header row
    #   +1: Opening Balance
    #   +2: Additions
    #   +3: Depreciation charge
    #   +4: Closing Balance

    @classmethod
    def opening_row(cls, asset_idx: int) -> int:
        return cls.BASE_ROW + asset_idx * ROWS_PER_ASSET_CLASS + 1

    @classmethod
    def addition_row(cls, asset_idx: int) -> int:
        return cls.BASE_ROW + asset_idx * ROWS_PER_ASSET_CLASS + 2

    @classmethod
    def charge_row(cls, asset_idx: int) -> int:
        return cls.BASE_ROW + asset_idx * ROWS_PER_ASSET_CLASS + 3

    @classmethod
    def closing_row(cls, asset_idx: int) -> int:
        return cls.BASE_ROW + asset_idx * ROWS_PER_ASSET_CLASS + 4

    @classmethod
    def summary_base_row(cls, n_asset_classes: int) -> int:
        return cls.BASE_ROW + n_asset_classes * ROWS_PER_ASSET_CLASS + 2

    @classmethod
    def gross_block_row(cls, n_asset_classes: int) -> int:
        return cls.summary_base_row(n_asset_classes)

    @classmethod
    def cumul_depr_row(cls, n_asset_classes: int) -> int:
        return cls.summary_base_row(n_asset_classes) + 1

    @classmethod
    def net_block_row(cls, n_asset_classes: int) -> int:
        return cls.summary_base_row(n_asset_classes) + 2

    @classmethod
    def check_row(cls, n_asset_classes: int) -> int:
        return cls.summary_base_row(n_asset_classes) + 4


class ExpensesLayout:
    """Expenses sheet layout."""
    SHEET = "Expenses"
    HDR_ROW         = 2
    SUBHDR_ROW      = 4

    # Raw material blocks
    RM_BASE_ROW     = 6

    @classmethod
    def rm_qty_row(cls, idx: int) -> int:
        return cls.RM_BASE_ROW + idx * ROWS_PER_MATERIAL

    @classmethod
    def rm_price_row(cls, idx: int) -> int:
        return cls.RM_BASE_ROW + idx * ROWS_PER_MATERIAL + 1

    @classmethod
    def rm_cost_row(cls, idx: int) -> int:
        return cls.RM_BASE_ROW + idx * ROWS_PER_MATERIAL + 2

    @classmethod
    def total_cogs_row(cls, n_materials: int) -> int:
        return cls.RM_BASE_ROW + n_materials * ROWS_PER_MATERIAL + 1

    @classmethod
    def overhead_base_row(cls, n_materials: int) -> int:
        return cls.total_cogs_row(n_materials) + 3

    @classmethod
    def revenue_ref_row(cls, n_materials: int) -> int:
        """Revenue pull row (needed for % calculations)."""
        return cls.overhead_base_row(n_materials)

    @classmethod
    def net_block_ref_row(cls, n_materials: int) -> int:
        return cls.overhead_base_row(n_materials) + 1

    @classmethod
    def rm_rate_row(cls, n_materials: int) -> int:
        return cls.overhead_base_row(n_materials) + 3

    @classmethod
    def rm_amount_row(cls, n_materials: int) -> int:
        return cls.overhead_base_row(n_materials) + 4

    @classmethod
    def ins_rate_row(cls, n_materials: int) -> int:
        return cls.overhead_base_row(n_materials) + 5

    @classmethod
    def ins_amount_row(cls, n_materials: int) -> int:
        return cls.overhead_base_row(n_materials) + 6

    @classmethod
    def mkt_rate_row(cls, n_materials: int) -> int:
        return cls.overhead_base_row(n_materials) + 7

    @classmethod
    def mkt_amount_row(cls, n_materials: int) -> int:
        return cls.overhead_base_row(n_materials) + 8

    @classmethod
    def power_rate_row(cls, n_materials: int) -> int:
        return cls.overhead_base_row(n_materials) + 9

    @classmethod
    def power_amount_row(cls, n_materials: int) -> int:
        return cls.overhead_base_row(n_materials) + 10

    @classmethod
    def sga_base_row(cls, n_materials: int) -> int:
        return cls.overhead_base_row(n_materials) + 11

    @classmethod
    def sga_amount_row(cls, n_materials: int) -> int:
        return cls.overhead_base_row(n_materials) + 12

    @classmethod
    def transport_row(cls, n_materials: int) -> int:
        return cls.overhead_base_row(n_materials) + 13

    @classmethod
    def misc_row(cls, n_materials: int) -> int:
        return cls.overhead_base_row(n_materials) + 14


class TermLoanLayout:
    """Term Loan sheet layout. One monthly schedule per loan."""
    SHEET = "Term Loan"
    HDR_ROW         = 2
    RATE_ROW        = 4
    SUBHDR_ROW      = 5
    MONTHLY_BASE    = 6
    MAX_MONTHS      = 120   # supports up to 10-year loans

    # Right-side annual summary
    SUMMARY_YEAR_ROW    = 5   # year labels
    SUMMARY_MONTHS_ROW  = 7
    SUMMARY_REPAY_ROW   = 8   # repayment months per year
    SUMMARY_PRINCIPAL_ROW= 9
    SUMMARY_OUTSTANDING_ROW = 10
    SUMMARY_INTEREST_ROW= 11
    SUMMARY_COL_START   = 10  # J column = col 10

    # OD Interest (below monthly schedule)
    @classmethod
    def od_base_row(cls) -> int:
        return cls.MONTHLY_BASE + cls.MAX_MONTHS + 3


class WCapLayout:
    """Working Capital sheet layout."""
    SHEET    = "W Cap"
    HDR_ROW  = 2
    SUBHDR   = 4
    DAYS_ROW = 6

    CL_HDR_ROW       = 9
    CREDITORS_ROW    = 10
    CREDITOR_DAYS_ROW= 11
    ADMIN_CRED_ROW   = 14
    ADMIN_DAYS_ROW   = 15
    TOTAL_CL_ROW     = 17

    CA_HDR_ROW       = 19
    STOCK_ROW        = 20
    STOCK_DAYS_ROW   = 21
    DEBTORS_ROW      = 23
    DEBTOR_DAYS_ROW  = 24
    CASH_ROW         = 26
    TOTAL_CA_ROW     = 28

    WC_REQ_ROW       = 31
    BANK_LOAN_ROW    = 33
    WC_INTEREST_ROW  = 35


class TaxLayout:
    """Tax sheet layout."""
    SHEET       = "Tax"
    STATUS_ROW  = 3   # entity type cell
    STATUS_COL  = 3   # C3
    HDR_ROW     = 5

    # Individual slabs: rows 33-35
    IND_BASIC_ROW   = 33
    IND_HEC_ROW     = 34
    IND_TOTAL_ROW   = 35

    # Company: rows 39-41
    CO_BASIC_ROW    = 39
    CO_HEC_ROW      = 40
    CO_TOTAL_ROW    = 41

    # Partnership: rows 45-47
    PART_BASIC_ROW  = 45
    PART_HEC_ROW    = 46
    PART_TOTAL_ROW  = 47

    TAXABLE_INC_ROW = 30
    YEAR_ROW        = 29
    DATA_COL_START  = 3  # C


class PLLayout:
    """Profit & Loss sheet layout."""
    SHEET       = "PL"
    HDR_ROW     = 2
    YEAR_ROW    = 4
    DATA_COL_START = 6  # F — Year 1

    # Income
    SALES_ROW        = 9
    TOTAL_REV_ROW    = 10

    # Cost of Sales
    COGS_ROW         = 13
    TOTAL_COGS_ROW   = 15

    # Gross Profit
    GROSS_PROFIT_ROW = 17

    # Operating Expenses (dynamic — based on n expense line items)
    OE_BASE_ROW      = 20

    # Fixed expense rows within operating expenses (relative to OE_BASE_ROW)
    OE_RM_OFFSET     = 0
    OE_INSURANCE_OFFSET = 1
    OE_MARKETING_OFFSET = 2
    OE_SGA_OFFSET    = 3
    OE_TRANSPORT_OFFSET = 4
    OE_MISC_OFFSET   = 5
    OE_SALARIES_OFFSET  = 6
    OE_DEPRECIATION_OFFSET = 7

    N_OPEX_LINES     = 8   # fixed count of opex line items

    @classmethod
    def total_opex_row(cls) -> int:
        return cls.OE_BASE_ROW + cls.N_OPEX_LINES + 1

    @classmethod
    def ebit_row(cls) -> int:
        return cls.total_opex_row() + 2

    # Interest
    @classmethod
    def interest_tl_row(cls) -> int:
        return cls.ebit_row() + 3

    @classmethod
    def interest_wc_row(cls) -> int:
        return cls.ebit_row() + 4

    @classmethod
    def total_interest_row(cls) -> int:
        return cls.ebit_row() + 5

    @classmethod
    def ebat_row(cls) -> int:
        return cls.ebit_row() + 7

    @classmethod
    def pbt_row(cls) -> int:
        return cls.ebat_row() + 4

    @classmethod
    def current_tax_row(cls) -> int:
        return cls.pbt_row() + 3

    @classmethod
    def pat_row(cls) -> int:
        return cls.pbt_row() + 5

    @classmethod
    def retained_profit_row(cls) -> int:
        return cls.pat_row() + 2

    @classmethod
    def ebitda_row(cls) -> int:
        return cls.retained_profit_row() + 3

    @classmethod
    def ebitda_margin_row(cls) -> int:
        return cls.ebitda_row() + 1

    @classmethod
    def interest_coverage_row(cls) -> int:
        return cls.ebitda_row() + 2


class BSLayout:
    """Balance Sheet layout."""
    SHEET       = "BS"
    HDR_ROW     = 2
    YEAR_ROW    = 5
    DATA_COL_START = 3  # C — Year 1

    # LIABILITIES
    CURR_LIAB_HDR    = 7
    ST_BORROWINGS_ROW= 8
    CREDITORS_ROW    = 9
    TAX_PROVISION_ROW= 10
    TL_CURRENT_ROW   = 11
    VEHICLE_LOAN_ROW = 12
    DTL_ROW          = 13
    OTHER_CL_ROW     = 14
    TOTAL_CL_ROW     = 15

    TERM_LIAB_HDR    = 17
    TERM_LOAN_ROW    = 18
    OTHER_TL_ROW     = 21
    TOTAL_TL_ROW     = 22

    TOTAL_OUTSIDE_LIAB_ROW = 23

    SHAREHOLDER_HDR  = 25
    SHARE_CAPITAL_ROW= 26
    RESERVES_ROW     = 27
    SHAREHOLDERS_FUND_ROW = 28
    TOTAL_LIAB_ROW   = 29

    # ASSETS
    CURR_ASSETS_HDR  = 31
    CASH_ROW         = 32
    FD_ROW           = 33
    RECEIVABLES_ROW  = 34
    STOCK_CONS_ROW   = 36
    WIP_ROW          = 37
    FG_ROW           = 38
    TOTAL_CA_ROW     = 40

    FA_HDR_ROW       = 42
    GROSS_BLOCK_ROW  = 43
    ACCUM_DEPR_ROW   = 44
    NET_BLOCK_ROW    = 45

    INTANGIBLES_ROW  = 47
    NON_CURR_INV_ROW = 48
    SECURITY_DEP_ROW = 49
    OTHER_NCA_ROW    = 50

    TOTAL_ASSETS_ROW = 52

    # Below totals
    TNW_ROW          = 54
    NWC_ROW          = 55
    CURR_RATIO_ROW   = 56
    TOL_TNW_ROW      = 58
    BALANCE_CHECK_ROW= 60


class CFSLayout:
    """Cash Flow Statement layout."""
    SHEET       = "CFS"
    HDR_ROW     = 2
    YEAR_ROW    = 4
    DATA_COL_START = 3  # C — Year 1

    # Operating
    OA_HDR_ROW      = 5
    PBT_ROW         = 6
    ADD_DEPR_ROW    = 7
    ADD_INT_ROW     = 8
    LESS_NONOP_ROW  = 9
    DTL_CHANGE_ROW  = 10

    WC_CHG_HDR_ROW  = 12
    CHG_INVESTMENTS_ROW = 14
    CHG_RECEIVABLES_ROW = 15
    CHG_INVENTORY_ROW   = 16
    CHG_OTHER_CA_ROW    = 17
    CHG_STB_ROW         = 20
    CHG_CREDITORS_ROW   = 21
    TAXES_PAID_ROW      = 23
    NET_OPERATING_ROW   = 25

    # Investing
    INVEST_HDR_ROW  = 27
    CAPEX_ROW       = 28
    INVESTMENTS_ROW = 29
    OTHER_NCA_ROW   = 30
    NONOP_INC_ROW   = 31
    NET_INVESTING_ROW = 33

    # Financing
    FIN_HDR_ROW     = 35
    SHARE_CAP_ROW   = 36
    TL_ROW          = 37
    INT_PAID_ROW    = 38
    NET_FINANCING_ROW = 40

    NET_CHANGE_ROW  = 42
    OPENING_CASH_ROW= 43
    CLOSING_CASH_ROW= 44


class RatioLayout:
    """Ratios / Dashboard sheet layout."""
    SHEET       = "Ratio"
    HDR_ROW     = 2
    YEAR_ROW    = 4
    DATA_COL_START = 3  # C — Year 1

    DSCR_HDR_ROW    = 6
    PAT_ROW         = 8
    DEPR_ROW        = 9
    INT_ROW         = 10
    NUMERATOR_ROW   = 11

    PRINC_ROW       = 13
    INT_DENOM_ROW   = 14
    DENOMINATOR_ROW = 15
    DSCR_ROW        = 17
    AVG_DSCR_ROW    = 18

    ROCE_HDR_ROW    = 20
    ROCE_PAT_ROW    = 22
    ROCE_TAX_ROW    = 23
    ROCE_INT_ROW    = 24
    EBIT_ROW        = 25
    TOTAL_ASSETS_ROW= 27
    CURR_LIAB_ROW   = 28
    CAP_EMPLOYED_ROW= 29
    ROCE_ROW        = 31
    AVG_ROCE_ROW    = 32

    IRR_HDR_ROW     = 34
    CAP_CHG_ROW     = 36
    EBITDA_ROW      = 37
    NCF_ROW         = 38
    IRR_ROW         = 40

    BEP_HDR_ROW     = 43
    BEP_YEAR_ROW    = 45
    BEP_INT_ROW     = 48
    BEP_DEPR_ROW    = 49
    BEP_SGA_ROW     = 50
    BEP_TOTAL_FC_ROW= 52
    BEP_CONTRIBUTION_ROW = 54
    BEP_PCT_ROW     = 56

    PROF_HDR_ROW    = 59
    PROF_BEP_ROW    = 62
    PROF_OPR_MARGIN_ROW = 63
    PROF_NPM_ROW    = 64
    PROF_DE_ROW     = 65
    PROF_AT_ROW     = 66


# ═══════════════════════════════════════════════════════════════════════════════
#  MASTER LAYOUT ENGINE
# ═══════════════════════════════════════════════════════════════════════════════

class LayoutEngine:
    """
    Master cell address resolver for the entire workbook.

    Usage:
        engine = LayoutEngine(store)
        addr = engine.addr("Revenue", "total_revenue_row")  → 47
        ref  = engine.ref("Revenue", "total_revenue", year=1) → "D47"
        xref = engine.xref("PL", "Revenue", "total_revenue", year=1)
               → "=Revenue!D47"
    """

    def __init__(self, store: SessionStore):
        self.store      = store
        self.n_products = store.n_products
        self.n_materials= store.n_materials
        self.n_employees= store.n_employee_categories
        self.n_loans    = store.n_term_loans
        self.n_years    = store.n_years
        self.n_assets   = len(store.capital_means.assets)
        self.n_asset_classes = len(set(
            a.category for a in store.capital_means.assets
        )) or 1

        # Pre-build the complete address map
        self._map: Dict[str, Dict[str, int]] = {}
        self._build_all_layouts()

    # ── Public API ────────────────────────────────────────────────────────────

    def row(self, sheet: str, logical: str) -> int:
        """Return row number for a logical cell name in a sheet."""
        try:
            val = self._map[sheet][logical]
            return val[0] if isinstance(val, tuple) else val
        except KeyError:
            raise KeyError(f"Layout key not found: sheet='{sheet}' logical='{logical}'\n"
                           f"Available keys: {sorted(self._map.get(sheet, {}).keys())}")

    def ref(self, sheet: str, logical: str, year: Optional[int] = None,
            col: Optional[int] = None) -> str:
        """
        Return cell address string like 'D17'.
        Provide year (1-indexed) OR explicit col (1-indexed).
        """
        r = self.row(sheet, logical)
        c = year_col(year) if year is not None else (col or COL_LABEL)
        return cell_addr(r, c)

    def xref(self, from_sheet: str, to_sheet: str, logical: str,
             year: Optional[int] = None, col: Optional[int] = None,
             abs_row=False, abs_col=False) -> str:
        """
        Return cross-sheet formula reference like =Revenue!D47.
        Automatically wraps in single quotes if sheet name has spaces.
        """
        r = self.row(to_sheet, logical)
        c = year_col(year) if year is not None else (col or COL_LABEL)
        col_str = f"${col_letter(c)}" if abs_col else col_letter(c)
        row_str = f"${r}" if abs_row else str(r)
        sheet_str = f"'{to_sheet}'" if ' ' in to_sheet else to_sheet
        return f"={sheet_str}!{col_str}{row_str}"

    def assumption_ref(self, logical: str, abs_both=True) -> str:
        """Return absolute Assumption sheet reference like =Assumption!$E$4."""
        val = self._map["Assumption"][logical]
        r, c = val if isinstance(val, tuple) else (val, COL_LABEL)
        col_str = f"${col_letter(c)}" if abs_both else col_letter(c)
        row_str = f"${r}" if abs_both else str(r)
        return f"=Assumption!{col_str}{row_str}"

    def year_col(self, year: int) -> int:
        return year_col(year)

    def year_col_letter(self, year: int) -> str:
        return col_letter(year_col(year))

    def year_range(self, logical: str, sheet: str) -> str:
        """Return range across all projection years like D17:J17."""
        r = self.row(sheet, logical)
        c_start = year_col(1)
        c_end   = year_col(self.n_years)
        return f"{col_letter(c_start)}{r}:{col_letter(c_end)}{r}"

    # ── Layout Builders ───────────────────────────────────────────────────────

    def _build_all_layouts(self):
        self._build_assumption_layout()
        self._build_revenue_layout()
        self._build_manpower_layout()
        self._build_depreciation_layout()
        self._build_expenses_layout()
        self._build_term_loan_layout()
        self._build_wcap_layout()
        self._build_tax_layout()
        self._build_pl_layout()
        self._build_bs_layout()
        self._build_cfs_layout()
        self._build_ratio_layout()

    def _set(self, sheet: str, logical: str, row_or_tuple, col: int = None):
        """Register a (row, col) or row in the map."""
        if sheet not in self._map:
            self._map[sheet] = {}
        if isinstance(row_or_tuple, tuple):
            self._map[sheet][logical] = row_or_tuple  # (row, col) stored as tuple
        else:
            self._map[sheet][logical] = (row_or_tuple, col or COL_LABEL)

    def _build_assumption_layout(self):
        AL = AssumptionLayout
        s = AL.SHEET
        # Capacity
        self._set(s, "cap_year1_util",       AL.CAP_YEAR1_UTIL)
        self._set(s, "cap_annual_increment",  AL.CAP_ANNUAL_INCREMENT)
        self._set(s, "cap_max_util",          AL.CAP_MAX_UTIL)
        self._set(s, "cap_working_days",      AL.CAP_WORKING_DAYS)
        self._set(s, "cap_months_year",       AL.CAP_MONTHS_YEAR)
        # Revenue: per product
        for i in range(self.n_products):
            self._set(s, f"rev_price_p{i}",      (AL.rev_price_row(i),      AL.REV_COL_PRICE))
            self._set(s, f"rev_escalation_p{i}",  (AL.rev_escalation_row(i), AL.REV_COL_ESCALATION))
        # Raw Materials: per material
        for i in range(self.n_materials):
            self._set(s, f"rm_price_m{i}",
                      (AL.rm_price_row(i, self.n_products), AL.RM_COL_PRICE))
            self._set(s, f"rm_escalation_m{i}",
                      (AL.rm_price_row(i, self.n_products) + 1, AL.RM_COL_ESCALATION))
        # Expenses
        for attr in ["EXP_RM_PCT_FA","EXP_RM_ESCALATION","EXP_INS_PCT_FA",
                     "EXP_INS_ESCALATION","EXP_POWER_PCT_REV","EXP_POWER_ESCALATION",
                     "EXP_MKT_PCT_REV","EXP_MKT_ESCALATION","EXP_TRANSPORT_BASE",
                     "EXP_TRANSPORT_ESC","EXP_MISC_BASE","EXP_MISC_ESC",
                     "EXP_SGA_BASE","EXP_SGA_ESC"]:
            self._set(s, attr.lower(), getattr(AL, attr))
        # Manpower
        for i in range(self.n_employees):
            self._set(s, f"mp_count_cat{i}",  (AL.mp_count_row(i), AL.MP_COL_COUNT))
            self._set(s, f"mp_salary_cat{i}", (AL.mp_count_row(i), AL.MP_COL_SALARY))
        # Finance: per loan
        for i in range(self.n_loans):
            self._set(s, f"fin_amount_l{i}",     (AL.fin_amount_row(i),     AL.FIN_COL_VALUE))
            self._set(s, f"fin_rate_l{i}",        (AL.fin_rate_row(i),       AL.FIN_COL_VALUE))
            self._set(s, f"fin_tenor_l{i}",       (AL.fin_tenor_row(i),      AL.FIN_COL_VALUE))
            self._set(s, f"fin_moratorium_l{i}",  (AL.fin_moratorium_row(i), AL.FIN_COL_VALUE))
        self._set(s, "fin_od_limit",  (AL.FIN_OD_LIMIT_ROW, AL.FIN_COL_VALUE))
        self._set(s, "fin_od_rate",   (AL.FIN_OD_RATE_ROW,  AL.FIN_COL_VALUE))
        # Working Capital
        for attr in ["WC_DEBTOR_DAYS","WC_CREDITOR_RM","WC_CREDITOR_ADMIN",
                     "WC_STOCK_RM","WC_STOCK_FG","WC_LOAN_AMOUNT","WC_INTEREST_RATE"]:
            self._set(s, attr.lower(), getattr(AL, attr))
        # Depreciation
        for attr in ["DEP_PM_RATE","DEP_CIVIL_RATE","DEP_FURN_RATE",
                     "DEP_VEH_RATE","DEP_ELEC_RATE","DEP_PREOP_RATE"]:
            self._set(s, attr.lower(), getattr(AL, attr))
        # Implementation
        self._set(s, "impl_months", AL.IMPL_MONTHS)

    def _build_revenue_layout(self):
        RL = RevenueLayout
        s  = RL.SHEET
        self._set(s, "months_row",      RL.MONTHS_ROW,   COL_LABEL)
        self._set(s, "days_row",        RL.DAYS_ROW,     COL_LABEL)
        self._set(s, "capacity_row",    RL.CAPACITY_ROW, COL_LABEL)
        self._set(s, "util_row",        RL.UTIL_ROW,     COL_LABEL)
        for i in range(self.n_products):
            self._set(s, f"prod_volume_p{i}",  RL.prod_volume_row(i),  COL_LABEL)
            self._set(s, f"prod_liters_p{i}",  RL.prod_liters_row(i),  COL_LABEL)
            self._set(s, f"prod_split_p{i}",   RL.prod_split_row(i),   COL_LABEL)
            self._set(s, f"prod_price_p{i}",   RL.prod_price_row(i),   COL_LABEL)
            self._set(s, f"prod_revenue_p{i}", RL.prod_revenue_row(i), COL_LABEL)
        self._set(s, "total_revenue",   RL.total_revenue_row(self.n_products), COL_LABEL)

    def _build_manpower_layout(self):
        ML = ManPowerLayout
        s  = ML.SHEET
        for i in range(self.n_employees):
            self._set(s, f"emp_row_{i}", ML.emp_row(i), COL_LABEL)
        self._set(s, "annual_fixed",    ML.annual_fixed_row(self.n_employees),    COL_LABEL)
        self._set(s, "annual_variable", ML.annual_variable_row(self.n_employees), COL_LABEL)
        self._set(s, "pl_fixed",        ML.pl_fixed_row(self.n_employees),        COL_LABEL)
        self._set(s, "pl_variable",     ML.pl_variable_row(self.n_employees),     COL_LABEL)

    def _build_depreciation_layout(self):
        DL = DepreciationLayout
        s  = DL.SHEET
        self._set(s, "months_row", DL.MONTHS_ROW, COL_LABEL)
        # Get ordered unique asset categories from store
        seen = []
        for a in self.store.capital_means.assets:
            if a.category not in seen:
                seen.append(a.category)
        self._asset_class_order = seen
        for i, cat in enumerate(seen):
            key = cat.value.lower().replace(" ", "_").replace("&", "and")
            self._set(s, f"dep_opening_{key}",    DL.opening_row(i),  COL_LABEL)
            self._set(s, f"dep_addition_{key}",   DL.addition_row(i), COL_LABEL)
            self._set(s, f"dep_charge_{key}",     DL.charge_row(i),   COL_LABEL)
            self._set(s, f"dep_closing_{key}",    DL.closing_row(i),  COL_LABEL)
        n = len(seen)
        self._set(s, "gross_block",  DL.gross_block_row(n),  COL_LABEL)
        self._set(s, "cumul_depr",   DL.cumul_depr_row(n),   COL_LABEL)
        self._set(s, "net_block",    DL.net_block_row(n),    COL_LABEL)
        self._set(s, "check_row",    DL.check_row(n),        COL_LABEL)

    def _build_expenses_layout(self):
        EL = ExpensesLayout
        s  = EL.SHEET
        n  = self.n_materials
        for i in range(n):
            self._set(s, f"rm_qty_m{i}",   EL.rm_qty_row(i),   COL_LABEL)
            self._set(s, f"rm_price_m{i}",  EL.rm_price_row(i), COL_LABEL)
            self._set(s, f"rm_cost_m{i}",   EL.rm_cost_row(i),  COL_LABEL)
        self._set(s, "total_cogs",    EL.total_cogs_row(n),     COL_LABEL)
        self._set(s, "revenue_ref",   EL.revenue_ref_row(n),    COL_LABEL)
        self._set(s, "net_block_ref", EL.net_block_ref_row(n),  COL_LABEL)
        self._set(s, "rm_rate",       EL.rm_rate_row(n),        COL_LABEL)
        self._set(s, "rm_amount",     EL.rm_amount_row(n),      COL_LABEL)
        self._set(s, "ins_rate",      EL.ins_rate_row(n),       COL_LABEL)
        self._set(s, "ins_amount",    EL.ins_amount_row(n),     COL_LABEL)
        self._set(s, "mkt_rate",      EL.mkt_rate_row(n),       COL_LABEL)
        self._set(s, "mkt_amount",    EL.mkt_amount_row(n),     COL_LABEL)
        self._set(s, "power_rate",    EL.power_rate_row(n),     COL_LABEL)
        self._set(s, "power_amount",  EL.power_amount_row(n),   COL_LABEL)
        self._set(s, "sga_base",      EL.sga_base_row(n),       COL_LABEL)
        self._set(s, "sga_amount",    EL.sga_amount_row(n),     COL_LABEL)
        self._set(s, "transport",     EL.transport_row(n),      COL_LABEL)
        self._set(s, "misc",          EL.misc_row(n),           COL_LABEL)

    def _build_term_loan_layout(self):
        TL = TermLoanLayout
        s  = TL.SHEET
        self._set(s, "rate_row",          TL.RATE_ROW,              COL_LABEL)
        self._set(s, "monthly_base",      TL.MONTHLY_BASE,          COL_LABEL)
        self._set(s, "summary_year_row",  TL.SUMMARY_YEAR_ROW,      TL.SUMMARY_COL_START)
        self._set(s, "summary_months",    TL.SUMMARY_MONTHS_ROW,    TL.SUMMARY_COL_START)
        self._set(s, "summary_principal", TL.SUMMARY_PRINCIPAL_ROW, TL.SUMMARY_COL_START)
        self._set(s, "summary_outstanding",TL.SUMMARY_OUTSTANDING_ROW,TL.SUMMARY_COL_START)
        self._set(s, "summary_interest",  TL.SUMMARY_INTEREST_ROW,  TL.SUMMARY_COL_START)
        self._set(s, "od_base",           TL.od_base_row(),         COL_LABEL)

    def _build_wcap_layout(self):
        WL = WCapLayout
        s  = WL.SHEET
        for attr in ["CREDITORS_ROW","CREDITOR_DAYS_ROW","ADMIN_CRED_ROW",
                     "ADMIN_DAYS_ROW","TOTAL_CL_ROW","STOCK_ROW","STOCK_DAYS_ROW",
                     "DEBTORS_ROW","DEBTOR_DAYS_ROW","CASH_ROW","TOTAL_CA_ROW",
                     "WC_REQ_ROW","BANK_LOAN_ROW","WC_INTEREST_ROW"]:
            self._set(s, attr.lower(), getattr(WL, attr), COL_BASIS)

    def _build_tax_layout(self):
        TxL = TaxLayout
        s   = TxL.SHEET
        for attr in ["IND_BASIC_ROW","IND_HEC_ROW","IND_TOTAL_ROW",
                     "CO_BASIC_ROW","CO_HEC_ROW","CO_TOTAL_ROW",
                     "PART_BASIC_ROW","PART_HEC_ROW","PART_TOTAL_ROW",
                     "TAXABLE_INC_ROW","YEAR_ROW"]:
            self._set(s, attr.lower(), getattr(TxL, attr), TxL.DATA_COL_START)

    def _build_pl_layout(self):
        PL = PLLayout
        s  = PL.SHEET
        dc = PL.DATA_COL_START
        for attr in ["SALES_ROW","TOTAL_REV_ROW","COGS_ROW","TOTAL_COGS_ROW",
                     "GROSS_PROFIT_ROW","OE_BASE_ROW"]:
            self._set(s, attr.lower(), getattr(PL, attr), dc)
        # Operating expense individual lines
        self._set(s, "opex_rm",         PL.OE_BASE_ROW + PL.OE_RM_OFFSET,          dc)
        self._set(s, "opex_insurance",  PL.OE_BASE_ROW + PL.OE_INSURANCE_OFFSET,   dc)
        self._set(s, "opex_marketing",  PL.OE_BASE_ROW + PL.OE_MARKETING_OFFSET,   dc)
        self._set(s, "opex_sga",        PL.OE_BASE_ROW + PL.OE_SGA_OFFSET,         dc)
        self._set(s, "opex_transport",  PL.OE_BASE_ROW + PL.OE_TRANSPORT_OFFSET,   dc)
        self._set(s, "opex_misc",       PL.OE_BASE_ROW + PL.OE_MISC_OFFSET,        dc)
        self._set(s, "opex_salaries",   PL.OE_BASE_ROW + PL.OE_SALARIES_OFFSET,    dc)
        self._set(s, "opex_depreciation",PL.OE_BASE_ROW + PL.OE_DEPRECIATION_OFFSET, dc)
        self._set(s, "total_opex_row",  PL.total_opex_row(), dc)
        self._set(s, "ebit_row",        PL.ebit_row(),       dc)
        self._set(s, "interest_tl_row", PL.interest_tl_row(), dc)
        self._set(s, "interest_wc_row", PL.interest_wc_row(), dc)
        self._set(s, "total_interest_row", PL.total_interest_row(), dc)
        self._set(s, "ebat_row",        PL.ebat_row(),       dc)
        self._set(s, "pbt_row",         PL.pbt_row(),        dc)
        self._set(s, "current_tax_row", PL.current_tax_row(), dc)
        self._set(s, "pat_row",         PL.pat_row(),        dc)
        self._set(s, "retained_profit_row", PL.retained_profit_row(), dc)
        self._set(s, "ebitda_row",      PL.ebitda_row(),     dc)
        self._set(s, "ebitda_margin_row", PL.ebitda_margin_row(), dc)
        self._set(s, "int_coverage_row", PL.interest_coverage_row(), dc)

    def _build_bs_layout(self):
        BS = BSLayout
        s  = BS.SHEET
        dc = BS.DATA_COL_START
        for attr in dir(BS):
            if attr.endswith("_ROW") and not attr.startswith("_"):
                self._set(s, attr.lower(), getattr(BS, attr), dc)

    def _build_cfs_layout(self):
        CFS = CFSLayout
        s   = CFS.SHEET
        dc  = CFS.DATA_COL_START
        for attr in dir(CFS):
            if attr.endswith("_ROW") and not attr.startswith("_"):
                self._set(s, attr.lower(), getattr(CFS, attr), dc)

    def _build_ratio_layout(self):
        RL = RatioLayout
        s  = RL.SHEET
        dc = RL.DATA_COL_START
        for attr in dir(RL):
            if attr.endswith("_ROW") and not attr.startswith("_"):
                self._set(s, attr.lower(), getattr(RL, attr), dc)

    # ── Diagnostic ────────────────────────────────────────────────────────────

    def dump_sheet(self, sheet: str) -> str:
        """Pretty print all logical → (row, col) mappings for a sheet."""
        lines = [f"\n{'─'*60}", f"  Layout: {sheet}", f"{'─'*60}"]
        for k, v in sorted(self._map.get(sheet, {}).items()):
            r, c = v
            lines.append(f"  {k:<40} → {col_letter(c)}{r}")
        return "\n".join(lines)

    def dump_all(self) -> str:
        return "\n".join(self.dump_sheet(s) for s in self._map)
