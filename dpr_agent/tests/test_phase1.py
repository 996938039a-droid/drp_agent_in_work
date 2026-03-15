"""
test_phase1.py
──────────────
Unit tests for Phase 1: SessionStore + LayoutEngine

Tests cover:
  1. SessionStore: creation, property computations, validation
  2. LayoutEngine: cell address correctness for 1 product, 5 products, 10 products
  3. Dynamic row shifts: verify downstream rows shift correctly
  4. Cross-sheet reference format
  5. Assumption ref format
  6. Edge cases: 1 loan, 3 loans; 1 employee, 8 employees
"""

import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from core.session_store import (
    SessionStore, Asset, FinanceSource, Product, RawMaterial,
    EmployeeCategory, BusinessType, RevenueModelType,
    EntityType, FinanceSourceType, AssetCategory, SectionStatus
)
from core.layout_engine import LayoutEngine, col_letter, year_col, cell_addr

# ─── ANSI colours for output ─────────────────────────────────────────────────
GREEN  = "\033[92m"
RED    = "\033[91m"
YELLOW = "\033[93m"
CYAN   = "\033[96m"
RESET  = "\033[0m"
BOLD   = "\033[1m"

passed = 0
failed = 0
errors = []

def ok(name):
    global passed
    passed += 1
    print(f"  {GREEN}✓{RESET}  {name}")

def fail(name, reason):
    global failed
    failed += 1
    errors.append((name, reason))
    print(f"  {RED}✗{RESET}  {name}")
    print(f"      {RED}{reason}{RESET}")

def section(title):
    print(f"\n{BOLD}{CYAN}{'─'*60}{RESET}")
    print(f"{BOLD}{CYAN}  {title}{RESET}")
    print(f"{BOLD}{CYAN}{'─'*60}{RESET}")

def assert_eq(test_name, actual, expected):
    if actual == expected:
        ok(test_name)
    else:
        fail(test_name, f"expected {repr(expected)}, got {repr(actual)}")

def assert_true(test_name, condition, msg=""):
    if condition:
        ok(test_name)
    else:
        fail(test_name, msg or "condition was False")

def assert_raises(test_name, exc_type, fn):
    try:
        fn()
        fail(test_name, f"Expected {exc_type.__name__} but no exception raised")
    except exc_type:
        ok(test_name)
    except Exception as e:
        fail(test_name, f"Expected {exc_type.__name__}, got {type(e).__name__}: {e}")


# ═══════════════════════════════════════════════════════════════════════════════
#  FIXTURES
# ═══════════════════════════════════════════════════════════════════════════════

def make_store(n_products=2, n_materials=2, n_employees=4, n_loans=1,
               n_years=7, add_od=True) -> SessionStore:
    store = SessionStore()
    store.project_profile.company_name = "Test Co"
    store.project_profile.promoter_name = "Test Promoter"
    store.project_profile.operation_start_date = "2026-04"
    store.project_profile.projection_years = n_years
    store.project_profile.entity_type = EntityType.COMPANY
    store.project_profile.status = SectionStatus.COMPLETE

    # Assets
    store.capital_means.assets = [
        Asset("Civil Works",   AssetCategory.CIVIL_WORKS,     200.0, {1: 0.8, 2: 0.2}),
        Asset("Plant & Mach",  AssetCategory.PLANT_MACHINERY, 110.0, {3: 1.0}),
        Asset("Furniture",     AssetCategory.FURNITURE,         5.0, {3: 1.0}),
    ]

    # Finance
    sources = []
    for i in range(n_loans):
        sources.append(FinanceSource(
            source_type=FinanceSourceType.TERM_LOAN,
            amount_lakhs=200.0 / n_loans,
            rate_pa=0.09,
            tenor_months=84,
            moratorium_months=18,
            label=f"TL {i+1}",
        ))
    if add_od:
        sources.append(FinanceSource(
            source_type=FinanceSourceType.OD_LIMIT,
            amount_lakhs=70.0,
            rate_pa=0.09,
        ))
    # Promoter equity = total cost - loans - od
    total_loans = sum(f.amount_lakhs for f in sources)
    promoter = max(0, store.capital_means.total_project_cost - total_loans)
    if promoter > 0:
        sources.append(FinanceSource(
            source_type=FinanceSourceType.PROMOTER_EQUITY,
            amount_lakhs=promoter
        ))
    store.capital_means.finance_sources = sources
    store.capital_means.status = SectionStatus.COMPLETE

    # Products
    products = []
    for i in range(n_products):
        products.append(Product(
            name=f"Product {i+1}",
            unit="litres",
            capacity_per_day=8.0,
            output_ratio=0.36,
            split_percent=1.0 / n_products,
            price_per_unit=175.0 + i * 20,
            price_escalation_pa=0.04,
        ))
    store.revenue_model.products = products
    store.revenue_model.year1_utilization = 0.50
    store.revenue_model.annual_utilization_increment = 0.05
    store.revenue_model.max_utilization = 0.85
    store.revenue_model.working_days_per_month = 28
    store.revenue_model.status = SectionStatus.COMPLETE

    # Materials
    materials = []
    for i in range(n_materials):
        materials.append(RawMaterial(
            name=f"Material {i+1}",
            unit="kg",
            price_per_unit=50.0 + i * 10,
            input_per_output_unit=2.78,
            price_escalation_pa=0.05,
        ))
    store.cost_structure.raw_materials = materials
    store.cost_structure.status = SectionStatus.COMPLETE

    # Employees
    desig = ["Manager", "Operator", "Admin", "Guard", "Helper",
             "Supervisor", "Technician", "Driver"]
    cats = []
    for i in range(n_employees):
        cats.append(EmployeeCategory(
            designation=desig[i % len(desig)],
            count=1 + i % 3,
            monthly_salary_lakhs=0.40 - i * 0.03,
            is_fixed=True,
            annual_increment_pa=0.05,
        ))
    store.manpower.categories = cats
    store.manpower.status = SectionStatus.COMPLETE

    store.finance_wc.debtor_days = 10
    store.finance_wc.creditor_days_rm = 10
    store.finance_wc.creditor_days_admin = 30
    store.finance_wc.stock_days_rm = 7
    store.finance_wc.status = SectionStatus.COMPLETE

    store.tax_config.entity_type = EntityType.COMPANY
    store.tax_config.status = SectionStatus.COMPLETE

    return store


# ═══════════════════════════════════════════════════════════════════════════════
#  TEST SUITE
# ═══════════════════════════════════════════════════════════════════════════════

section("1. SessionStore — Basic Properties")

s2 = make_store(n_products=2)
assert_eq("n_products (2)",  s2.n_products,  2)
assert_eq("n_materials (2)", s2.n_materials, 2)
assert_eq("n_employees (4)", s2.n_employee_categories, 4)
assert_eq("n_loans (1)",     s2.n_term_loans, 1)
assert_eq("n_years (7)",     s2.n_years,     7)

# ── Finance properties
assert_eq("total_project_cost", round(s2.capital_means.total_project_cost, 1), 315.0)
assert_true("is_balanced (cost=means)", s2.capital_means.is_balanced,
            f"gap={s2.capital_means.total_project_cost - s2.capital_means.total_finance:.2f}")

# ── EMI calculation
tl = s2.capital_means.term_loans[0]
assert_eq("TL repayment_months", tl.repayment_months, 66)  # 84 - 18
expected_emi = round(200.0 / 66, 6)
assert_eq("TL monthly EMI principal",
          round(tl.monthly_emi_principal, 6), expected_emi)

# ── Capacity utilization ramp
rm = s2.revenue_model
assert_eq("util year 1", rm.utilization_for_year(1), 0.50)
assert_eq("util year 2", round(rm.utilization_for_year(2), 2), 0.55)
assert_eq("util year 7", round(rm.utilization_for_year(7), 2), 0.80)

# ── Readiness check
assert_true("is_ready_for_generation", s2.is_ready_for_generation)

# ── Serialization roundtrip
j = s2.to_json()
assert_true("JSON serializable", isinstance(j, str) and len(j) > 100)
assert_true("JSON contains company_name", "Test Co" in j)


section("2. SessionStore — Validation")

empty = SessionStore()
errs = empty.validate_completeness()
assert_true("empty store has errors", len(errs) > 3)

partial = make_store()
partial.project_profile.company_name = ""
errs = partial.validate_completeness()
assert_true("missing company_name flagged", any("company_name" in e for e in errs))

imbalanced = make_store()
imbalanced.capital_means.finance_sources = [
    FinanceSource(FinanceSourceType.TERM_LOAN, 100.0, 0.09, 84, 18)
]
errs = imbalanced.validate_completeness()
assert_true("imbalance flagged", any("imbalance" in e.lower() or "gap" in e.lower()
                                      for e in errs), f"errors: {errs}")


section("3. LayoutEngine — Column Helpers")

assert_eq("col_letter(1)",  col_letter(1),  "A")
assert_eq("col_letter(2)",  col_letter(2),  "B")
assert_eq("col_letter(4)",  col_letter(4),  "D")
assert_eq("col_letter(26)", col_letter(26), "Z")
assert_eq("col_letter(27)", col_letter(27), "AA")

assert_eq("year_col(1)", year_col(1), 4)   # D
assert_eq("year_col(2)", year_col(2), 5)   # E
assert_eq("year_col(7)", year_col(7), 10)  # J

assert_eq("cell_addr(17,4)", cell_addr(17, 4), "D17")
assert_eq("cell_addr(1,1)",  cell_addr(1,  1), "A1")


section("4. LayoutEngine — 1 Product Configuration")

store1 = make_store(n_products=1)
eng1   = LayoutEngine(store1)

# Revenue rows — PROD_BASE_ROW=13, ROWS_PER_PRODUCT=6
# Product 0: sub_hdr=12, vol=13, lit=14, spl=15, prc=16, rev=17
assert_eq("revenue: prod_volume_p0",   eng1.row("Revenue", "prod_volume_p0"),  13)
assert_eq("revenue: prod_liters_p0",   eng1.row("Revenue", "prod_liters_p0"),  14)
assert_eq("revenue: prod_split_p0",    eng1.row("Revenue", "prod_split_p0"),   15)
assert_eq("revenue: prod_price_p0",    eng1.row("Revenue", "prod_price_p0"),   16)
assert_eq("revenue: prod_revenue_p0",  eng1.row("Revenue", "prod_revenue_p0"), 17)
# total_revenue_row = PROD_BASE_ROW + 1*6 + 1 = 13 + 6 + 1 = 20
assert_eq("revenue: total_revenue (1 product)", eng1.row("Revenue", "total_revenue"), 20)


section("5. LayoutEngine — 2 Product Configuration")

store2 = make_store(n_products=2)
eng2   = LayoutEngine(store2)

# Product 0: vol=13, rev=17 (unchanged)
assert_eq("p2: prod_volume_p0",  eng2.row("Revenue", "prod_volume_p0"),  13)
assert_eq("p2: prod_revenue_p0", eng2.row("Revenue", "prod_revenue_p0"), 17)

# Product 1 starts after product 0's 6 rows: 13 + 1*6 = 19
assert_eq("p2: prod_volume_p1",  eng2.row("Revenue", "prod_volume_p1"),  19)
assert_eq("p2: prod_price_p1",   eng2.row("Revenue", "prod_price_p1"),   22)
assert_eq("p2: prod_revenue_p1", eng2.row("Revenue", "prod_revenue_p1"), 23)

# total_revenue = 13 + 2*6 + 1 = 26
assert_eq("p2: total_revenue (2 products)", eng2.row("Revenue", "total_revenue"), 26)


section("6. LayoutEngine — 5 Product Configuration")

store5 = make_store(n_products=5)
eng5   = LayoutEngine(store5)

# Each product block = 6 rows (sub_hdr + 5 data rows)
for i in range(5):
    expected_vol = 13 + i * 6
    expected_rev = 13 + i * 6 + 4
    assert_eq(f"p5: prod_volume_p{i}",   eng5.row("Revenue", f"prod_volume_p{i}"),  expected_vol)
    assert_eq(f"p5: prod_revenue_p{i}",  eng5.row("Revenue", f"prod_revenue_p{i}"), expected_rev)

# total_revenue = 13 + 5*6 + 1 = 44
assert_eq("p5: total_revenue (5 products)", eng5.row("Revenue", "total_revenue"), 44)


section("7. LayoutEngine — 10 Product Configuration")

store10 = make_store(n_products=10)
eng10   = LayoutEngine(store10)

assert_eq("p10: prod_volume_p9",  eng10.row("Revenue", "prod_volume_p9"),  13 + 9*6)
assert_eq("p10: prod_revenue_p9", eng10.row("Revenue", "prod_revenue_p9"), 13 + 9*6 + 4)
# total_revenue = 13 + 10*6 + 1 = 74
assert_eq("p10: total_revenue (10 products)", eng10.row("Revenue", "total_revenue"), 74)


section("8. LayoutEngine — Dynamic row shift verification")

# When products increase, downstream row references in Expenses must also shift
# The key invariant: row addresses computed via the engine are ALWAYS consistent
# regardless of n_products

eng_1p = LayoutEngine(make_store(n_products=1, n_materials=2))
eng_5p = LayoutEngine(make_store(n_products=5, n_materials=2))

tr_1p = eng_1p.row("Revenue", "total_revenue")
tr_5p = eng_5p.row("Revenue", "total_revenue")
assert_true("total_revenue shifts with more products",
            tr_5p > tr_1p,
            f"1p row={tr_1p}, 5p row={tr_5p}")

# ManPower rows shift with more employees
eng_4e = LayoutEngine(make_store(n_employees=4))
eng_8e = LayoutEngine(make_store(n_employees=8))
fixed_4e = eng_4e.row("ManPower", "pl_fixed")
fixed_8e = eng_8e.row("ManPower", "pl_fixed")
assert_true("manpower pl_fixed shifts with more employees",
            fixed_8e > fixed_4e,
            f"4e row={fixed_4e}, 8e row={fixed_8e}")


section("9. LayoutEngine — ref() and xref() output format")

eng = LayoutEngine(make_store(n_products=2))

# ref() returns "D17" format
rev_row = eng.row("Revenue", "total_revenue")
ref_y1  = eng.ref("Revenue", "total_revenue", year=1)
assert_eq("ref() year=1 uses D column", ref_y1[0], "D")
assert_true("ref() contains row number", str(rev_row) in ref_y1)

ref_y7  = eng.ref("Revenue", "total_revenue", year=7)
assert_eq("ref() year=7 uses J column", ref_y7[0], "J")

# xref() starts with = and contains sheet name
xr = eng.xref("PL", "Revenue", "total_revenue", year=1)
assert_true("xref() starts with =",       xr.startswith("="))
assert_true("xref() contains sheet name", "Revenue" in xr)
assert_true("xref() contains D column",   "D" in xr)

# Sheet with space → wrapped in single quotes
xr_wc = eng.xref("PL", "W Cap", "wc_req_row", year=1)
assert_true("xref() wraps sheet-with-space in quotes", "'W Cap'" in xr_wc)


section("10. LayoutEngine — Assumption ref() format")

eng = LayoutEngine(make_store(n_products=2))
aref = eng.assumption_ref("cap_year1_util")
assert_true("assumption_ref starts with =",          aref.startswith("="))
assert_true("assumption_ref contains Assumption",     "Assumption" in aref)
assert_true("assumption_ref uses absolute refs ($)",  "$" in aref)


section("11. LayoutEngine — Multiple term loans")

s1loan  = make_store(n_loans=1)
s3loans = make_store(n_loans=3)
eng1l   = LayoutEngine(s1loan)
eng3l   = LayoutEngine(s3loans)

assert_eq("1 loan: n_term_loans",  s1loan.n_term_loans,  1)
assert_eq("3 loans: n_term_loans", s3loans.n_term_loans, 3)

# Each loan has 4 rows in Assumption
row_l0 = eng3l._map["Assumption"]["fin_amount_l0"][0]
row_l1 = eng3l._map["Assumption"]["fin_amount_l1"][0]
row_l2 = eng3l._map["Assumption"]["fin_amount_l2"][0]
assert_eq("loan 1 amount row offset from loan 0",   row_l1 - row_l0, 5)
assert_eq("loan 2 amount row offset from loan 0",   row_l2 - row_l0, 10)


section("12. LayoutEngine — Depreciation layout with multiple asset classes")

eng = LayoutEngine(make_store())
# 3 asset categories: CIVIL_WORKS(0), PLANT_MACHINERY(1), FURNITURE(2)
n = 3  # n_asset_classes

# Opening balances
open_cw  = eng.row("Depreciation", "dep_opening_civil_works")
open_pm  = eng.row("Depreciation", "dep_opening_plant_and_machinery")
open_furn= eng.row("Depreciation", "dep_opening_furniture_and_fixture")

assert_true("civil_works opening before P&M opening", open_cw < open_pm)
assert_true("P&M opening before furniture opening",   open_pm < open_furn)
assert_eq("gap between asset class openings = 5 rows", open_pm - open_cw, 5)

# Net block should come after all asset class blocks
net_block_row = eng.row("Depreciation", "net_block")
assert_true("net_block after last asset class",
            net_block_row > open_furn + 4)


section("13. LayoutEngine — P&L layout consistency")

eng = LayoutEngine(make_store())

# Key invariant: rows must be in ascending order
rows = [
    eng.row("PL", "sales_row"),
    eng.row("PL", "total_rev_row"),
    eng.row("PL", "cogs_row"),
    eng.row("PL", "total_cogs_row"),
    eng.row("PL", "gross_profit_row"),
    eng.row("PL", "oe_base_row"),
    eng.row("PL", "total_opex_row"),
    eng.row("PL", "ebit_row"),
    eng.row("PL", "interest_tl_row"),
    eng.row("PL", "total_interest_row"),
    eng.row("PL", "ebat_row"),
    eng.row("PL", "pbt_row"),
    eng.row("PL", "current_tax_row"),
    eng.row("PL", "pat_row"),
    eng.row("PL", "retained_profit_row"),
    eng.row("PL", "ebitda_row"),
]
for i in range(len(rows) - 1):
    assert_true(f"PL row ordering: row[{i}] < row[{i+1}]",
                rows[i] < rows[i+1],
                f"rows[{i}]={rows[i]}, rows[{i+1}]={rows[i+1]}")


section("14. LayoutEngine — BS balance sheet layout consistency")

eng = LayoutEngine(make_store())

# Liabilities before assets
total_liab = eng.row("BS", "total_liab_row")
total_assets = eng.row("BS", "total_assets_row")
balance_check = eng.row("BS", "balance_check_row")

assert_true("total_liab < total_assets", total_liab < total_assets)
assert_true("balance_check after total_assets", balance_check > total_assets)


section("15. LayoutEngine — year_range() helper")

eng = LayoutEngine(make_store(n_years=7))

rng = eng.year_range("total_revenue", "Revenue")
assert_true("year_range starts with D",     rng.startswith("D"))
assert_true("year_range ends with J (yr7)", rng.endswith("J" + str(eng.row("Revenue", "total_revenue"))))

eng5yr = LayoutEngine(make_store(n_years=5))
rng5 = eng5yr.year_range("total_revenue", "Revenue")
assert_true("5yr range ends with H", "H" in rng5.split(":")[1])


section("16. LayoutEngine — dump_sheet() works without error")

eng = LayoutEngine(make_store())
for sheet in ["Assumption", "Revenue", "ManPower", "Depreciation",
              "Expenses", "Term Loan", "W Cap", "Tax", "PL", "BS", "CFS", "Ratio"]:
    try:
        output = eng.dump_sheet(sheet)
        assert_true(f"dump_sheet({sheet}) returns non-empty string",
                    isinstance(output, str) and len(output) > 10)
    except Exception as e:
        fail(f"dump_sheet({sheet})", str(e))


section("17. RevenueModel.utilization_for_year() — edge cases")

rm = make_store().revenue_model
assert_eq("util year 1 = 50%",  rm.utilization_for_year(1), 0.50)

# Ramp: 0.50, 0.55, 0.60, 0.65, 0.70, 0.75, 0.80
expected = [0.50, 0.55, 0.60, 0.65, 0.70, 0.75, 0.80]
for yr, exp in enumerate(expected, 1):
    actual = round(rm.utilization_for_year(yr), 2)
    assert_eq(f"utilization year {yr}", actual, exp)

# Test ceiling: max_utilization = 0.85, should not exceed
rm2 = make_store().revenue_model
rm2.year1_utilization = 0.80
rm2.annual_utilization_increment = 0.10
rm2.max_utilization = 0.85
assert_eq("ceiling enforced year 1", rm2.utilization_for_year(1), 0.80)
assert_eq("ceiling enforced year 2", round(rm2.utilization_for_year(2), 2), 0.85)
assert_eq("ceiling enforced year 3", round(rm2.utilization_for_year(3), 2), 0.85)


# ═══════════════════════════════════════════════════════════════════════════════
#  SUMMARY
# ═══════════════════════════════════════════════════════════════════════════════

total = passed + failed
print(f"\n{'═'*60}")
print(f"{BOLD}  PHASE 1 TEST RESULTS{RESET}")
print(f"{'═'*60}")
print(f"  {GREEN}Passed: {passed}/{total}{RESET}")
if failed:
    print(f"  {RED}Failed: {failed}/{total}{RESET}")
    print(f"\n{RED}  Failed tests:{RESET}")
    for name, reason in errors:
        print(f"    {RED}✗ {name}{RESET}")
        print(f"      {reason}")
else:
    print(f"\n  {GREEN}{BOLD}ALL TESTS PASSED ✓{RESET}")
print(f"{'═'*60}\n")

# Exit code for CI
sys.exit(0 if failed == 0 else 1)
