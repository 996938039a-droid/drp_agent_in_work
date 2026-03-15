"""
benchmark_engine.py
────────────────────
The Benchmark Engine makes a single Claude API call once the business
type is known, and returns intelligent, business-specific benchmark
values for all Tier 2 fields.

Key design decisions:
  1. One API call per DPR session — not one per field. We ask for all
     Tier 2 benchmarks in a single structured prompt and parse the
     JSON response.

  2. The prompt includes: business description, industry, location,
     scale (project cost), and the exact list of fields needed.

  3. Response is strict JSON — no markdown, no prose. We parse it
     and validate each value falls within a sensible range.

  4. On failure (malformed JSON, API error), falls back to hardcoded
     safe defaults so generation always proceeds.

  5. Each benchmark carries: value, range_low, range_high, reason.
     The reason is shown to the user when presenting the default.
"""

from __future__ import annotations
import json
import re
from dataclasses import dataclass
from typing import Optional
from agents.field_registry import TIER_2_FIELDS, FieldDef, FieldType


# ─── Benchmark result dataclass ──────────────────────────────────────────────

@dataclass
class Benchmark:
    key: str
    value: float              # midpoint — what gets shown as default
    range_low: float
    range_high: float
    reason: str               # shown to user when presenting default
    source: str = "llm"       # "llm" | "fallback"


# ─── Fallback defaults (used when API call fails) ────────────────────────────

FALLBACK_BENCHMARKS: dict[str, dict] = {
    "cost_structure.rm_pct_of_fa":          {"value": 0.02, "low": 0.01, "high": 0.04,
        "reason": "Standard R&M provision for manufacturing assets"},
    "cost_structure.rm_escalation_pa":       {"value": 0.10, "low": 0.05, "high": 0.15,
        "reason": "Typical maintenance cost escalation"},
    "cost_structure.insurance_pct_of_fa":    {"value": 0.01, "low": 0.005, "high": 0.02,
        "reason": "Standard commercial insurance rate on fixed assets"},
    "cost_structure.power_pct_revenue":      {"value": 0.07, "low": 0.03, "high": 0.20,
        "reason": "Varies widely by industry energy intensity"},
    "cost_structure.power_escalation_pa":    {"value": 0.05, "low": 0.03, "high": 0.08,
        "reason": "Indian electricity tariff escalation trend"},
    "cost_structure.marketing_pct_revenue":  {"value": 0.04, "low": 0.01, "high": 0.10,
        "reason": "Typical MSME marketing spend"},
    "cost_structure.sga_base_lakhs":         {"value": 5.0,  "low": 2.0,  "high": 15.0,
        "reason": "General admin overhead for small MSME"},
    "cost_structure.transport_escalation_pa":{"value": 0.12, "low": 0.08, "high": 0.18,
        "reason": "Fuel price-linked transport cost escalation"},
    "cost_structure.misc_escalation_pa":     {"value": 0.10, "low": 0.05, "high": 0.20,
        "reason": "General miscellaneous expense growth"},
    "manpower.categories[].annual_increment_pa": {"value": 0.05, "low": 0.03, "high": 0.10,
        "reason": "Standard annual increment in Indian MSME sector"},
    "finance_wc.creditor_days_admin":        {"value": 30,   "low": 15,   "high": 60,
        "reason": "Typical admin expense payment cycle"},
    "finance_wc.stock_days_rm":              {"value": 10,   "low": 5,    "high": 30,
        "reason": "Raw material holding based on supply reliability"},
    "finance_wc.od_rate":                    {"value": 0.09, "low": 0.085,"high": 0.115,
        "reason": "Current MSME OD rates at Indian PSU banks"},
    "revenue_model.products[].price_escalation_pa": {"value": 0.04, "low": 0.02, "high": 0.08,
        "reason": "Typical selling price escalation in Indian markets"},
}

# ─── Validation bounds (LLM output must stay within these) ───────────────────

VALIDATION_BOUNDS: dict[str, tuple] = {
    "cost_structure.rm_pct_of_fa":           (0.005, 0.10),
    "cost_structure.rm_escalation_pa":        (0.0,   0.30),
    "cost_structure.insurance_pct_of_fa":     (0.001, 0.05),
    "cost_structure.power_pct_revenue":       (0.01,  0.50),
    "cost_structure.power_escalation_pa":     (0.0,   0.15),
    "cost_structure.marketing_pct_revenue":   (0.0,   0.25),
    "cost_structure.sga_base_lakhs":          (0.5,   100.0),
    "cost_structure.transport_escalation_pa": (0.0,   0.30),
    "cost_structure.misc_escalation_pa":      (0.0,   0.40),
    "manpower.categories[].annual_increment_pa": (0.0, 0.20),
    "finance_wc.creditor_days_admin":         (0,     120),
    "finance_wc.stock_days_rm":               (0,     90),
    "finance_wc.od_rate":                     (0.06,  0.18),
    "revenue_model.products[].price_escalation_pa": (0.0, 0.20),
}


# ═══════════════════════════════════════════════════════════════════════════════
#  BENCHMARK ENGINE
# ═══════════════════════════════════════════════════════════════════════════════

class BenchmarkEngine:
    """
    Generates intelligent, business-specific Tier 2 benchmarks via
    a single Claude API call.

    Usage:
        engine = BenchmarkEngine()
        benchmarks = await engine.generate(
            business_description="Cold storage unit in Nagpur, 500-ton capacity",
            industry="Cold Storage",
            location="Nagpur, Maharashtra",
            project_cost_lakhs=350.0,
        )
        # benchmarks: dict[str, Benchmark]
    """

    def __init__(self, model: str = "claude-sonnet-4-20250514",
                 api_key: str = ""):
        self.model   = model
        self.api_key = api_key

    # ── Public API ────────────────────────────────────────────────────────────

    async def generate(
        self,
        business_description: str,
        industry: str,
        location: str,
        project_cost_lakhs: float = 0.0,
    ) -> dict[str, Benchmark]:
        """
        Make the API call and return a dict of key → Benchmark.
        Falls back to safe defaults on any error.
        """
        try:
            raw = await self._call_api(
                business_description, industry, location, project_cost_lakhs
            )
            parsed = self._parse_response(raw)
            validated = self._validate(parsed)
            return validated
        except Exception as e:
            print(f"[BenchmarkEngine] API call failed: {e}. Using fallbacks.")
            return self._build_fallbacks()

    def generate_sync(
        self,
        business_description: str,
        industry: str,
        location: str,
        project_cost_lakhs: float = 0.0,
    ) -> dict[str, Benchmark]:
        """Synchronous version for non-async contexts."""
        import asyncio
        try:
            loop = asyncio.get_event_loop()
            if loop.is_running():
                import concurrent.futures
                with concurrent.futures.ThreadPoolExecutor() as pool:
                    future = pool.submit(
                        asyncio.run,
                        self.generate(business_description, industry,
                                      location, project_cost_lakhs)
                    )
                    return future.result(timeout=30)
            else:
                return loop.run_until_complete(
                    self.generate(business_description, industry,
                                  location, project_cost_lakhs)
                )
        except Exception as e:
            print(f"[BenchmarkEngine] Sync wrapper failed: {e}. Using fallbacks.")
            return self._build_fallbacks()

    # ── API call ──────────────────────────────────────────────────────────────

    async def _call_api(
        self,
        business_description: str,
        industry: str,
        location: str,
        project_cost_lakhs: float,
    ) -> str:
        """Make the Claude API call. Returns raw response text."""
        import aiohttp

        system_prompt = self._build_system_prompt()
        user_message  = self._build_user_prompt(
            business_description, industry, location, project_cost_lakhs
        )

        payload = {
            "model": self.model,
            "max_tokens": 2000,
            "system": system_prompt,
            "messages": [{"role": "user", "content": user_message}]
        }

        headers = {
            "Content-Type":      "application/json",
            "anthropic-version": "2023-06-01",
        }
        if self.api_key:
            headers["x-api-key"] = self.api_key

        async with aiohttp.ClientSession() as session:
            async with session.post(
                "https://api.anthropic.com/v1/messages",
                json=payload,
                headers=headers,
                timeout=aiohttp.ClientTimeout(total=30)
            ) as resp:
                if resp.status != 200:
                    text = await resp.text()
                    raise RuntimeError(f"API error {resp.status}: {text[:200]}")
                data = await resp.json()
                return data["content"][0]["text"]

    # ── Prompt builders ───────────────────────────────────────────────────────

    def _build_system_prompt(self) -> str:
        return """You are an expert Indian MSME financial analyst with deep knowledge of
sector-specific financial ratios, operating cost structures, and working capital
patterns across Indian industries.

You will be given a business description and asked to provide realistic benchmark
values for financial modelling parameters. Your benchmarks must:

1. Be specific to the exact business described — not generic
2. Reflect Indian market conditions (power costs, labour, logistics, etc.)
3. Account for the business scale (project cost / revenue)
4. Be defensible to a bank loan officer reviewing the DPR

You must respond ONLY with a valid JSON object. No prose, no markdown, no
explanation outside the JSON. The JSON must contain exactly the keys requested."""

    def _build_user_prompt(
        self,
        business_description: str,
        industry: str,
        location: str,
        project_cost_lakhs: float,
    ) -> str:
        # Build the list of fields to benchmark
        fields_section = "\n".join([
            f'  "{f.key}": {{"value": <number>, "range_low": <number>, '
            f'"range_high": <number>, "reason": "<one sentence, max 12 words>"}}'
            for f in TIER_2_FIELDS
        ])

        return f"""Business description: {business_description}
Industry: {industry}
Location: {location}
Estimated project cost: ₹{project_cost_lakhs:.0f} lakhs

Generate benchmark values for the following financial modelling parameters.
All percentage values must be expressed as decimals (0.06 = 6%, not 6).
All day values must be integers. All Lakh values must be floats.

Return ONLY this JSON object with all values filled in:
{{
{fields_section}
}}

Important context for each field:
- cost_structure.rm_pct_of_fa: Repair & maintenance as % of net fixed assets per year
- cost_structure.insurance_pct_of_fa: Insurance premium as % of fixed assets per year
- cost_structure.power_pct_revenue: Total power/electricity/fuel cost as % of annual revenue
- cost_structure.marketing_pct_revenue: Sales, distribution & marketing as % of revenue
- cost_structure.sga_base_lakhs: General selling, admin & overhead in INR Lakhs for Year 1
- finance_wc.stock_days_rm: How many days of raw material stock this business typically holds
- finance_wc.creditor_days_admin: Typical payment cycle for admin/overhead expenses
- revenue_model.products[].price_escalation_pa: Expected annual % increase in selling prices
"""

    # ── Response parsing ──────────────────────────────────────────────────────

    def _parse_response(self, raw: str) -> dict[str, dict]:
        """Extract and parse JSON from the LLM response."""
        # Strip any accidental markdown code fences
        text = raw.strip()
        text = re.sub(r'^```json\s*', '', text)
        text = re.sub(r'^```\s*',     '', text)
        text = re.sub(r'\s*```$',     '', text)

        try:
            return json.loads(text)
        except json.JSONDecodeError:
            # Try to extract JSON object from anywhere in the response
            match = re.search(r'\{[\s\S]+\}', text)
            if match:
                return json.loads(match.group())
            raise ValueError(f"Could not parse JSON from response: {text[:200]}")

    def _validate(self, parsed: dict[str, dict]) -> dict[str, Benchmark]:
        """
        Validate each benchmark value.
        - Must have required keys: value, range_low, range_high, reason
        - Value must be within VALIDATION_BOUNDS
        - Out-of-range values fall back to the hardcoded default
        """
        result: dict[str, Benchmark] = {}

        for f in TIER_2_FIELDS:
            key = f.key
            fb  = FALLBACK_BENCHMARKS.get(key, {})

            if key not in parsed:
                # Field missing from response — use fallback
                result[key] = self._fallback_for(key)
                continue

            raw_bm = parsed[key]

            # Extract fields with fallback
            try:
                value     = float(raw_bm.get("value",      fb.get("value", 0)))
                range_low = float(raw_bm.get("range_low",  fb.get("low",   value * 0.7)))
                range_high= float(raw_bm.get("range_high", fb.get("high",  value * 1.3)))
                reason    = str(raw_bm.get("reason", fb.get("reason", "Industry benchmark")))
            except (TypeError, ValueError):
                result[key] = self._fallback_for(key)
                continue

            # Bounds check
            bounds = VALIDATION_BOUNDS.get(key)
            if bounds:
                lo, hi = bounds
                if not (lo <= value <= hi):
                    print(f"[BenchmarkEngine] {key} value {value} out of bounds "
                          f"[{lo},{hi}] — using fallback")
                    result[key] = self._fallback_for(key)
                    continue

            result[key] = Benchmark(
                key=key,
                value=value,
                range_low=range_low,
                range_high=range_high,
                reason=reason,
                source="llm",
            )

        return result

    def _fallback_for(self, key: str) -> Benchmark:
        fb = FALLBACK_BENCHMARKS.get(key, {
            "value": 0.05, "low": 0.0, "high": 0.10,
            "reason": "Default estimate"
        })
        return Benchmark(
            key=key,
            value=fb["value"],
            range_low=fb["low"],
            range_high=fb["high"],
            reason=fb["reason"],
            source="fallback",
        )

    def _build_fallbacks(self) -> dict[str, Benchmark]:
        return {f.key: self._fallback_for(f.key) for f in TIER_2_FIELDS}

    # ── Formatting helpers (used by section agents) ───────────────────────────

    @staticmethod
    def format_for_display(bm: Benchmark, field: FieldDef,
                           context: dict) -> str:
        """
        Format a benchmark for display to the user in the chat.
        Returns a string like:
          "Power & fuel: 22% of revenue (range 18–26%). Reason. [Press enter to accept or type a value]"
        """
        # Format the value nicely
        if field.field_type == FieldType.PERCENT:
            val_str   = f"{bm.value * 100:.1f}%"
            range_str = f"{bm.range_low*100:.0f}–{bm.range_high*100:.0f}%"
        elif field.field_type == FieldType.INTEGER:
            val_str   = str(int(bm.value))
            range_str = f"{int(bm.range_low)}–{int(bm.range_high)}"
        else:
            val_str   = f"₹{bm.value:.1f}L"
            range_str = f"₹{bm.range_low:.1f}–{bm.range_high:.1f}L"

        source_note = "" if bm.source == "llm" else " *(industry default)*"

        return (
            f"**{field.label}**: {val_str}{source_note}\n"
            f"  *{bm.reason}* (typical range: {range_str})\n"
            f"  → Press enter to accept, or type a different value"
        )

    @staticmethod
    def format_batch_for_display(
        benchmarks: dict[str, Benchmark],
        fields: list[FieldDef],
        context: dict,
    ) -> str:
        """
        Format multiple benchmarks for display as a batch message.
        Used when showing all Tier 2 fields for a section at once.
        """
        lines = [
            "Here are the suggested values for your business. "
            "I've tailored these to your specific industry and scale.\n",
            "**Press enter to accept each, or type a different value:**\n",
        ]
        for i, f in enumerate(fields, 1):
            bm = benchmarks.get(f.key)
            if not bm:
                continue
            if f.field_type == FieldType.PERCENT:
                val_str   = f"{bm.value * 100:.1f}%"
                range_str = f"{bm.range_low*100:.0f}–{bm.range_high*100:.0f}%"
            elif f.field_type == FieldType.INTEGER:
                val_str   = str(int(bm.value))
                range_str = f"{int(bm.range_low)}–{int(bm.range_high)}"
            else:
                val_str   = f"₹{bm.value:.1f}L"
                range_str = f"₹{bm.range_low:.1f}–{bm.range_high:.1f}L"

            source_note = " *(default)*" if bm.source == "fallback" else ""
            lines.append(
                f"{i}. **{f.label}**: `{val_str}`{source_note}\n"
                f"   {bm.reason} *(range: {range_str})*"
            )

        return "\n".join(lines)


# ═══════════════════════════════════════════════════════════════════════════════
#  ASSUMPTION LOG BUILDER
# ═══════════════════════════════════════════════════════════════════════════════

class AssumptionLogBuilder:
    """
    Builds the complete record of every assumption used in the model.
    This populates the Assumption Log sheet in the Excel workbook.

    Each entry records:
      - The parameter name
      - The value used
      - How it was determined (user-provided / LLM benchmark / statutory)
      - The benchmark range (if Tier 2)
      - Any risk flags
    """

    def __init__(self):
        self.entries: list[dict] = []

    def add_tier1(self, key: str, label: str, value, unit: str,
                  section: str):
        """Record a user-provided Tier 1 value."""
        self.entries.append({
            "key": key, "label": label, "value": value, "unit": unit,
            "tier": "1 — User provided", "source": "Conversation input",
            "benchmark_range": "", "risk_flag": "", "section": section,
        })

    def add_tier2(self, key: str, label: str, value, unit: str,
                  bm: Benchmark, user_overrode: bool, section: str):
        """Record a Tier 2 value (benchmark default or user override)."""
        source = "User override" if user_overrode else f"LLM benchmark (midpoint of {bm.range_low}–{bm.range_high})"
        risk   = self._compute_risk(key, value, bm)
        self.entries.append({
            "key": key, "label": label, "value": value, "unit": unit,
            "tier": "2 — Benchmarked", "source": source,
            "benchmark_range": f"{bm.range_low}–{bm.range_high}",
            "risk_flag": risk, "section": section,
        })

    def add_tier3(self, key: str, label: str, value, unit: str,
                  section: str):
        """Record a Tier 3 statutory value used silently."""
        self.entries.append({
            "key": key, "label": label, "value": value, "unit": unit,
            "tier": "3 — Statutory / IT Act", "source": "Regulatory standard",
            "benchmark_range": "", "risk_flag": "", "section": section,
        })

    def _compute_risk(self, key: str, value: float,
                      bm: Benchmark) -> str:
        """Flag if a value is outside the benchmark range."""
        if value < bm.range_low:
            return f"Below typical range ({bm.range_low}–{bm.range_high})"
        if value > bm.range_high:
            return f"Above typical range ({bm.range_low}–{bm.range_high})"
        return ""

    def to_rows(self) -> list[dict]:
        """Return all entries as a list of dicts for Excel writing."""
        return self.entries
