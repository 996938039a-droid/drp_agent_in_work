"""
test_phase_intelligence.py
──────────────────────────
Tests the intelligence layer:
  1. Field registry — completeness, tier counts, section assignments
  2. Benchmark engine — fallback path, validation, formatting
  3. Orchestrator — Tier 3 defaults applied, store field appliers,
     missing field detection, extraction logic (without API calls)
"""

import sys, os, json, asyncio
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from agents.field_registry import (
    FIELDS, TIER_1_FIELDS, TIER_2_FIELDS, TIER_3_FIELDS,
    FieldTier, FieldType, fields_for_section, tier1_for_section, tier2_for_section
)
from agents.benchmark_engine import (
    BenchmarkEngine, Benchmark, FALLBACK_BENCHMARKS,
    VALIDATION_BOUNDS, AssumptionLogBuilder
)
from agents.orchestrator import Orchestrator, OrchestratorResponse
from core.session_store import (
    SessionStore, SectionStatus, EntityType, AssetCategory,
    FinanceSourceType, Asset, FinanceSource, Product, RawMaterial,
    EmployeeCategory
)

GREEN = "\033[92m"; RED = "\033[91m"; CYAN = "\033[96m"
RESET = "\033[0m";  BOLD = "\033[1m"

passed = 0; failed = 0; errors = []

def ok(name):
    global passed; passed += 1
    print(f"  {GREEN}✓{RESET}  {name}")

def fail(name, reason):
    global failed; failed += 1; errors.append((name, reason))
    print(f"  {RED}✗{RESET}  {name}")
    print(f"      {RED}{reason}{RESET}")

def section(title):
    print(f"\n{BOLD}{CYAN}{'─'*60}{RESET}\n{BOLD}{CYAN}  {title}{RESET}\n{BOLD}{CYAN}{'─'*60}{RESET}")

def assert_eq(name, actual, expected):
    if actual == expected: ok(name)
    else: fail(name, f"expected {repr(expected)}, got {repr(actual)}")

def assert_true(name, cond, msg=""):
    if cond: ok(name)
    else: fail(name, msg or "False")

def assert_in(name, val, container):
    if val in container: ok(name)
    else: fail(name, f"{repr(val)} not in container")


# ═══════════════════════════════════════════════════════════════════════════════
section("1. Field Registry — completeness")
# ═══════════════════════════════════════════════════════════════════════════════

assert_true("Total fields >= 50", len(FIELDS) >= 50, f"got {len(FIELDS)}")
assert_true("Tier 1 fields >= 15", len(TIER_1_FIELDS) >= 15,
            f"got {len(TIER_1_FIELDS)}")
assert_true("Tier 2 fields >= 10", len(TIER_2_FIELDS) >= 10,
            f"got {len(TIER_2_FIELDS)}")
assert_true("Tier 3 fields >= 8",  len(TIER_3_FIELDS) >= 8,
            f"got {len(TIER_3_FIELDS)}")

# All fields have required attributes
for f in FIELDS:
    assert_true(f"Field '{f.key}' has label",   bool(f.label))
    assert_true(f"Field '{f.key}' has unit",    bool(f.unit))
    assert_true(f"Field '{f.key}' has section", bool(f.section))
    if f.tier == FieldTier.TIER_1:
        assert_true(f"Tier1 '{f.key}' has question", bool(f.question))
    if f.tier == FieldTier.TIER_3:
        assert_true(f"Tier3 '{f.key}' has default",
                    f.default is not None, f"default is None")

# All sections covered
for sec in ["intake","profile","capital","revenue","costs","manpower","finance","system"]:
    fields_in_sec = fields_for_section(sec)
    assert_true(f"Section '{sec}' has fields",
                len(fields_in_sec) > 0, f"0 fields in {sec}")

# Tier 2 fields have tier2_prompt
for f in TIER_2_FIELDS:
    assert_true(f"Tier2 '{f.key}' has tier2_prompt", bool(f.tier2_prompt))


# ═══════════════════════════════════════════════════════════════════════════════
section("2. Field Registry — tier counts by section")
# ═══════════════════════════════════════════════════════════════════════════════

section_t1 = {}
section_t2 = {}
for sec in ["profile","capital","revenue","costs","manpower","finance"]:
    section_t1[sec] = len(tier1_for_section(sec))
    section_t2[sec] = len(tier2_for_section(sec))
    print(f"  {sec}: {section_t1[sec]} Tier1, {section_t2[sec]} Tier2")

assert_true("Profile has Tier1 fields",   section_t1["profile"] >= 5)
assert_true("Capital has Tier1 fields",   section_t1["capital"] >= 4)
assert_true("Revenue has Tier1 fields",   section_t1["revenue"] >= 5)
assert_true("Costs has Tier1 fields",     section_t1["costs"] >= 3)
assert_true("Manpower has Tier1 fields",  section_t1["manpower"] >= 2)

assert_true("Costs has Tier2 fields",     section_t2["costs"] >= 3,
            f"got {section_t2['costs']}")
assert_true("Finance has Tier2 fields",   section_t2["finance"] >= 2)


# ═══════════════════════════════════════════════════════════════════════════════
section("3. BenchmarkEngine — fallback path")
# ═══════════════════════════════════════════════════════════════════════════════

be = BenchmarkEngine()
fallbacks = be._build_fallbacks()

assert_true("Fallbacks built for all Tier2 fields",
            len(fallbacks) == len(TIER_2_FIELDS),
            f"got {len(fallbacks)}, expected {len(TIER_2_FIELDS)}")

for key, bm in fallbacks.items():
    assert_true(f"Fallback '{key}' has value",
                isinstance(bm.value, (int, float)))
    assert_true(f"Fallback '{key}' has range",
                bm.range_low <= bm.value <= bm.range_high or
                bm.range_low == bm.value,
                f"value={bm.value}, range=[{bm.range_low},{bm.range_high}]")
    assert_true(f"Fallback '{key}' source = fallback",
                bm.source == "fallback")
    assert_true(f"Fallback '{key}' has reason",
                bool(bm.reason))


# ═══════════════════════════════════════════════════════════════════════════════
section("4. BenchmarkEngine — validation")
# ═══════════════════════════════════════════════════════════════════════════════

# Test _validate with a well-formed response
good_response = {}
for f in TIER_2_FIELDS:
    fb = FALLBACK_BENCHMARKS.get(f.key, {})
    good_response[f.key] = {
        "value": fb.get("value", 0.05),
        "range_low": fb.get("low", 0.01),
        "range_high": fb.get("high", 0.10),
        "reason": "Test reason for this benchmark"
    }

validated = be._validate(good_response)
assert_eq("Validated count matches Tier2 count",
          len(validated), len(TIER_2_FIELDS))

for key, bm in validated.items():
    assert_true(f"Validated '{key}' source = llm",
                bm.source == "llm",
                f"got {bm.source}")

# Test _validate with out-of-bounds value → should fall back
bad_response = {
    "cost_structure.power_pct_revenue": {
        "value": 99.0,       # way out of bounds (max is 0.50)
        "range_low": 90.0,
        "range_high": 100.0,
        "reason": "Nonsense value"
    }
}
partial_validated = be._validate(bad_response)
if "cost_structure.power_pct_revenue" in partial_validated:
    pv_bm = partial_validated["cost_structure.power_pct_revenue"]
    assert_true("Out-of-bounds value falls back to fallback source",
                pv_bm.source == "fallback",
                f"got source={pv_bm.source}, value={pv_bm.value}")

# Test _validate with missing fields → all fall back
validated_empty = be._validate({})
assert_eq("Empty response produces fallback for all Tier2 fields",
          len(validated_empty), len(TIER_2_FIELDS))
for bm in validated_empty.values():
    assert_true("Empty response: all fallbacks",
                bm.source == "fallback")


# ═══════════════════════════════════════════════════════════════════════════════
section("5. BenchmarkEngine — JSON parsing")
# ═══════════════════════════════════════════════════════════════════════════════

# Valid JSON
valid_json = '{"cost_structure.rm_pct_of_fa": {"value": 0.025, "range_low": 0.01, "range_high": 0.05, "reason": "Test"}}'
parsed = be._parse_response(valid_json)
assert_true("Valid JSON parsed correctly",
            parsed["cost_structure.rm_pct_of_fa"]["value"] == 0.025)

# JSON with code fences
fenced = "```json\n" + valid_json + "\n```"
parsed_fenced = be._parse_response(fenced)
assert_true("Code-fenced JSON parsed correctly",
            "cost_structure.rm_pct_of_fa" in parsed_fenced)

# JSON embedded in prose
embedded = "Here are the benchmarks:\n\n" + valid_json + "\n\nEnd."
parsed_embedded = be._parse_response(embedded)
assert_true("JSON embedded in prose extracted correctly",
            "cost_structure.rm_pct_of_fa" in parsed_embedded)


# ═══════════════════════════════════════════════════════════════════════════════
section("6. BenchmarkEngine — display formatting")
# ═══════════════════════════════════════════════════════════════════════════════

from agents.field_registry import FIELD_MAP

test_bm = Benchmark(
    key="cost_structure.power_pct_revenue",
    value=0.22, range_low=0.18, range_high=0.26,
    reason="Refrigeration is the dominant cost in cold storage",
    source="llm"
)
test_field = FIELD_MAP.get("cost_structure.power_pct_revenue")
if test_field:
    display = BenchmarkEngine.format_for_display(test_bm, test_field, {})
    assert_true("Display string contains field label",
                test_field.label in display)
    assert_true("Display string contains formatted value",
                "22.0%" in display)
    assert_true("Display string contains range",
                "18–26%" in display or "18" in display)
    assert_true("Display string contains reason",
                "Refrigeration" in display)
    assert_true("Display string has acceptance prompt",
                "enter" in display.lower() or "accept" in display.lower())


# ═══════════════════════════════════════════════════════════════════════════════
section("7. BenchmarkEngine — batch display formatting")
# ═══════════════════════════════════════════════════════════════════════════════

fallbacks = be._build_fallbacks()
t2_costs  = tier2_for_section("costs")
batch_msg = BenchmarkEngine.format_batch_for_display(
    benchmarks=fallbacks,
    fields=t2_costs,
    context={"business": "cold storage unit", "industry": "Cold Storage"},
)
assert_true("Batch message is non-empty", len(batch_msg) > 100)
assert_true("Batch message contains enter prompt",
            "enter" in batch_msg.lower() or "accept" in batch_msg.lower())
assert_true("Batch message contains field labels",
            any(f.label in batch_msg for f in t2_costs))


# ═══════════════════════════════════════════════════════════════════════════════
section("8. AssumptionLogBuilder")
# ═══════════════════════════════════════════════════════════════════════════════

log = AssumptionLogBuilder()

# Add one of each tier
log.add_tier1("project_profile.company_name", "Company Name",
              "Test Co", "text", "profile")
test_bm2 = Benchmark("cost_structure.power_pct_revenue", 0.22,
                      0.18, 0.26, "Test reason", "llm")
log.add_tier2("cost_structure.power_pct_revenue", "Power & Fuel",
              0.22, "% of revenue", test_bm2,
              user_overrode=False, section="costs")
log.add_tier3("depreciation_rates.plant_machinery",
              "P&M Depreciation Rate", 0.15, "WDV % p.a.", "system")

rows = log.to_rows()
assert_eq("Log has 3 entries", len(rows), 3)

t1_row = next(r for r in rows if r["tier"].startswith("1"))
t2_row = next(r for r in rows if r["tier"].startswith("2"))
t3_row = next(r for r in rows if r["tier"].startswith("3"))

assert_eq("Tier1 source", t1_row["source"], "Conversation input")
assert_true("Tier2 source contains 'benchmark'",
            "benchmark" in t2_row["source"].lower() or
            "override" in t2_row["source"].lower())
assert_eq("Tier3 source", t3_row["source"], "Regulatory standard")

# Test risk flagging
bm_risky = Benchmark("k", 0.35, 0.18, 0.26, "reason", "llm")
log2 = AssumptionLogBuilder()
log2.add_tier2("k", "lbl", 0.35, "unit", bm_risky, True, "costs")
risky_row = log2.to_rows()[0]
assert_true("Above-range value gets risk flag",
            "above" in risky_row["risk_flag"].lower() or
            risky_row["risk_flag"] != "")


# ═══════════════════════════════════════════════════════════════════════════════
section("9. Orchestrator — Tier 3 defaults applied at startup")
# ═══════════════════════════════════════════════════════════════════════════════

orch = Orchestrator()

# All depreciation rates set
dr = orch.store.depreciation_rates
assert_eq("P&M depr rate = 0.15",    dr.plant_machinery, 0.15)
assert_eq("Civil depr rate = 0.10",  dr.civil_works,     0.10)
assert_eq("Furniture rate = 0.10",   dr.furniture,       0.10)
assert_eq("Vehicle rate = 0.15",     dr.vehicle,         0.15)

# Tax config set
tc = orch.store.tax_config
assert_eq("Company tax = 30%",       tc.company_basic_rate,    0.30)
assert_eq("HEC rate = 4%",           tc.hec_rate,              0.04)
assert_eq("Surcharge 1-10Cr = 7%",   tc.surcharge_rate_1cr_10cr, 0.07)

# Tier 3 values logged in assumption log
t3_logged = [e for e in orch.assumption_log.to_rows()
             if e["tier"].startswith("3")]
assert_true("Tier3 values logged at startup",
            len(t3_logged) >= 6, f"only {len(t3_logged)} logged")


# ═══════════════════════════════════════════════════════════════════════════════
section("10. Orchestrator — _apply_profile_fields")
# ═══════════════════════════════════════════════════════════════════════════════

orch2 = Orchestrator()
orch2._apply_profile_fields({
    "company_name": "Videhanutra India Pvt Ltd",
    "promoter_name": "Rajesh Kumar",
    "entity_type": "Company",
    "city": "Indore",
    "state": "Madhya Pradesh",
    "operation_start_date": "2026-04",
    "projection_years": 7,
})

pp = orch2.store.project_profile
assert_eq("company_name set",          pp.company_name,          "Videhanutra India Pvt Ltd")
assert_eq("promoter_name set",         pp.promoter_name,         "Rajesh Kumar")
assert_eq("entity_type set",           pp.entity_type.value,     "Company")
assert_eq("city set",                  pp.city,                  "Indore")
assert_eq("state set",                 pp.state,                 "Madhya Pradesh")
assert_eq("operation_start_date set",  pp.operation_start_date,  "2026-04")
assert_eq("projection_years set",      pp.projection_years,      7)


# ═══════════════════════════════════════════════════════════════════════════════
section("11. Orchestrator — _apply_capital_fields")
# ═══════════════════════════════════════════════════════════════════════════════

orch3 = Orchestrator()
orch3._apply_capital_fields({
    "assets": [
        {"name": "Civil Works",       "cost_lakhs": 202.0, "category": "Civil Works"},
        {"name": "Plant & Machinery", "cost_lakhs": 110.0, "category": "Plant & Machinery"},
    ],
    "term_loans": [
        {"amount_lakhs": 200, "rate_pa": 0.09,
         "tenor_months": 84, "moratorium_months": 18, "label": "SBI TL"},
    ],
    "od_limit_lakhs": 70,
    "od_rate_pa": 0.09,
})

cm = orch3.store.capital_means
assert_eq("2 assets added",             len(cm.assets), 2)
assert_eq("Asset 1 name",               cm.assets[0].name, "Civil Works")
assert_eq("Asset 1 cost",               cm.assets[0].cost_lakhs, 202.0)
assert_eq("Asset 1 category",           cm.assets[0].category, AssetCategory.CIVIL_WORKS)
assert_eq("1 term loan added",          len(cm.term_loans), 1)
assert_eq("TL amount",                  cm.term_loans[0].amount_lakhs, 200.0)
assert_eq("TL repayment months",        cm.term_loans[0].repayment_months, 66)
assert_eq("1 OD source added",          len(cm.od_sources), 1)
assert_eq("OD amount",                  cm.od_sources[0].amount_lakhs, 70.0)

# Compute promoter and check balance
orch3._compute_promoter_contribution()
assert_true("Balance after promoter contribution",
            cm.is_balanced,
            f"gap={cm.total_project_cost - cm.total_finance:.2f}")


# ═══════════════════════════════════════════════════════════════════════════════
section("12. Orchestrator — _apply_revenue_fields")
# ═══════════════════════════════════════════════════════════════════════════════

orch4 = Orchestrator()
orch4._apply_revenue_fields({
    "products": [
        {"name": "Mustard Oil", "unit": "litres",
         "price_per_unit": 175, "capacity_per_day": 8,
         "output_ratio": 0.36, "split_percent": 0.5},
        {"name": "Groundnut Oil", "unit": "litres",
         "price_per_unit": 195, "capacity_per_day": 8,
         "output_ratio": 0.36, "split_percent": 0.5},
    ],
    "year1_utilization": 0.50,
    "annual_utilization_increment": 0.05,
    "max_utilization": 0.85,
    "working_days_per_month": 28,
})

rm = orch4.store.revenue_model
assert_eq("2 products added",          len(rm.products), 2)
assert_eq("Product 1 name",            rm.products[0].name, "Mustard Oil")
assert_eq("Product 1 price",           rm.products[0].price_per_unit, 175.0)
assert_eq("Year1 utilization",         rm.year1_utilization, 0.50)
assert_eq("Working days",              rm.working_days_per_month, 28)

# No duplicate products
orch4._apply_revenue_fields({"products": [
    {"name": "Mustard Oil", "unit": "litres",
     "price_per_unit": 180, "capacity_per_day": 8,
     "output_ratio": 0.36, "split_percent": 0.5},
]})
assert_eq("No duplicate products added",  len(rm.products), 2)


# ═══════════════════════════════════════════════════════════════════════════════
section("13. Orchestrator — _write_tier2_to_store")
# ═══════════════════════════════════════════════════════════════════════════════

orch5 = Orchestrator()
orch5._write_tier2_to_store("cost_structure.rm_pct_of_fa",     0.025)
orch5._write_tier2_to_store("cost_structure.power_pct_revenue", 0.08)
orch5._write_tier2_to_store("finance_wc.stock_days_rm",         14)
orch5._write_tier2_to_store("finance_wc.creditor_days_admin",   30)

cs = orch5.store.cost_structure
fw = orch5.store.finance_wc
assert_eq("rm_pct_of_fa written",        cs.rm_pct_of_fa,         0.025)
assert_eq("power_pct_revenue written",   cs.power_pct_revenue,    0.08)
assert_eq("stock_days_rm written",       fw.stock_days_rm,        14)
assert_eq("creditor_days_admin written", fw.creditor_days_admin,  30)


# ═══════════════════════════════════════════════════════════════════════════════
section("14. Orchestrator — missing field detection")
# ═══════════════════════════════════════════════════════════════════════════════

orch6 = Orchestrator()

# Empty store
assert_true("Empty: profile fields missing",    len(orch6._missing_profile_fields()) > 0)
assert_true("Empty: capital fields missing",    len(orch6._missing_capital_fields()) > 0)
assert_true("Empty: revenue fields missing",    len(orch6._missing_revenue_fields()) > 0)
assert_true("Empty: costs fields missing",      len(orch6._missing_costs_fields()) > 0)
assert_true("Empty: finance fields missing",    len(orch6._missing_finance_fields()) > 0)

# Fill profile
orch6._apply_profile_fields({
    "company_name":"Test Co","promoter_name":"Test","entity_type":"Company",
    "city":"Mumbai","state":"Maharashtra","operation_start_date":"2026-04",
    "projection_years":7
})
assert_eq("Profile complete: no missing fields",
          len(orch6._missing_profile_fields()), 0)

# Fill capital
orch6._apply_capital_fields({
    "assets": [{"name":"Plant","cost_lakhs":100,"category":"Plant & Machinery"}],
    "term_loans":[{"amount_lakhs":80,"rate_pa":0.09,"tenor_months":84,"moratorium_months":18}],
    "od_limit_lakhs":0,
})
assert_eq("Capital complete: no missing fields",
          len(orch6._missing_capital_fields()), 0)


# ═══════════════════════════════════════════════════════════════════════════════
section("15. Orchestrator — JSON cleaning utility")
# ═══════════════════════════════════════════════════════════════════════════════

cases = [
    ('{"a":1}',                    '{"a":1}'),
    ('```json\n{"a":1}\n```',      '{"a":1}'),
    ('```\n{"a":1}\n```',          '{"a":1}'),
    ('  {"a":1}  ',                '{"a":1}'),
]
for raw, expected in cases:
    cleaned = Orchestrator._clean_json(raw)
    assert_eq(f"JSON cleaning: {repr(raw[:30])}", cleaned, expected)


# ═══════════════════════════════════════════════════════════════════════════════
#  SUMMARY
# ═══════════════════════════════════════════════════════════════════════════════

total = passed + failed
print(f"\n{'═'*60}")
print(f"{BOLD}  INTELLIGENCE LAYER TEST RESULTS{RESET}")
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

sys.exit(0 if failed == 0 else 1)
