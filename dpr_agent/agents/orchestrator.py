"""
orchestrator.py
───────────────
The Orchestrator is the master conversation controller.

It:
  1. Receives every user message
  2. Decides which Section Agent should handle it
  3. Calls the Benchmark Engine after Intake completes
  4. Tracks what's collected and what's missing
  5. Presents the confirmation summary
  6. Triggers generation when everything is confirmed

The Orchestrator talks to ONE LLM at a time.
All agents share the same session store.
"""

from __future__ import annotations
import json
import re
from typing import Optional
from dataclasses import dataclass

from core.session_store import (
    SessionStore, SectionStatus, BusinessType, RevenueModelType,
    EntityType, Asset, FinanceSource, Product, RawMaterial,
    EmployeeCategory, AssetCategory, FinanceSourceType
)
from agents.field_registry import (
    FieldTier, FieldDef, TIER_3_FIELDS,
    tier1_for_section, tier2_for_section, FIELD_MAP
)
from agents.benchmark_engine import BenchmarkEngine, Benchmark, AssumptionLogBuilder


# ─── Conversation state ───────────────────────────────────────────────────────

SECTIONS_IN_ORDER = [
    "intake",
    "profile",
    "capital",
    "revenue",
    "costs",
    "manpower",
    "finance",
    "confirm",
]

SECTION_DISPLAY_NAMES = {
    "intake":   "Business Description",
    "profile":  "Project Profile",
    "capital":  "Capital & Finance",
    "revenue":  "Revenue Model",
    "costs":    "Cost Structure",
    "manpower": "Manpower",
    "finance":  "Working Capital",
    "confirm":  "Confirmation",
}


@dataclass
class ConversationTurn:
    role: str    # "user" | "assistant"
    content: str


# ═══════════════════════════════════════════════════════════════════════════════
#  ORCHESTRATOR
# ═══════════════════════════════════════════════════════════════════════════════

class Orchestrator:
    """
    Drives the complete DPR conversation from first message to generation.

    Usage:
        orch = Orchestrator(store)
        response = await orch.process_message("I want a DPR for my oil mill")
        print(response.message)          # what to show the user
        print(response.ready_to_generate) # True when store is complete
    """

    def __init__(self, store: Optional[SessionStore] = None,
                 model: str = "claude-sonnet-4-20250514",
                 api_key: str = ""):
        self.store     = store or SessionStore()
        self.model     = model
        self.api_key   = api_key   # injected from Streamlit session_state
        self.history:  list[ConversationTurn] = []
        self.current_section: str = "intake"
        self.benchmarks: dict[str, Benchmark] = {}
        self.assumption_log = AssumptionLogBuilder()
        self._apply_tier3_defaults()

    # ── Apply Tier 3 defaults immediately ────────────────────────────────────

    def _apply_tier3_defaults(self):
        """Write all statutory values to store silently at startup."""
        store = self.store
        dr = store.depreciation_rates
        dr.plant_machinery  = 0.15
        dr.civil_works      = 0.10
        dr.furniture        = 0.10
        dr.vehicle          = 0.15
        dr.electrical       = 0.10
        dr.pre_operative    = 0.20

        tc = store.tax_config
        tc.company_basic_rate          = 0.30
        tc.hec_rate                    = 0.04
        tc.surcharge_rate_1cr_10cr     = 0.07
        tc.surcharge_rate_above_10cr   = 0.12
        tc.partnership_rate            = 0.30
        tc.partnership_surcharge_rate  = 0.12

        # Log them
        for f in TIER_3_FIELDS:
            self.assumption_log.add_tier3(
                f.key, f.label, f.default, f.unit, f.section
            )

    # ── Main entry point ──────────────────────────────────────────────────────

    async def process_message(self, user_message: str) -> "OrchestratorResponse":
        """
        Process one user message. Returns the next assistant message
        and any state changes.
        """
        self.history.append(ConversationTurn("user", user_message))

        # ── Tier 2 interception ───────────────────────────────────────────────
        # If we are awaiting a tier2 response, handle it first before
        # dispatching to section handlers — otherwise the section handler
        # loops and re-presents the same tier2 question indefinitely.
        if hasattr(self, "_awaiting_tier2") and self._awaiting_tier2:
            section = self._awaiting_tier2
            self._awaiting_tier2 = ""
            await self._apply_tier2_responses(section, user_message)

            # Advance to next section
            next_section_map = {
                "revenue":  "costs",
                "costs":    "manpower",
                "manpower": "finance",
                "finance":  "confirm",
            }
            next_sec = next_section_map.get(section, "confirm")

            # Mark current section complete
            section_store_map = {
                "revenue":  self.store.revenue_model,
                "costs":    self.store.cost_structure,
                "manpower": self.store.manpower,
                "finance":  self.store.finance_wc,
            }
            sec_store = section_store_map.get(section)
            if sec_store:
                sec_store.status = SectionStatus.COMPLETE

            self.current_section = next_sec
            first_q = self._first_question_for(next_sec)
            response = OrchestratorResponse(
                message=f"✅ {section.title()} complete.\n\n---\n\n{first_q}",
                section_completed=section,
                next_section=next_sec,
            )
            self.history.append(ConversationTurn("assistant", response.message))
            return response

        # ── Normal section dispatch ───────────────────────────────────────────
        handler = {
            "intake":   self._handle_intake,
            "profile":  self._handle_profile,
            "capital":  self._handle_capital,
            "revenue":  self._handle_revenue,
            "costs":    self._handle_costs,
            "manpower": self._handle_manpower,
            "finance":  self._handle_finance,
            "confirm":  self._handle_confirm,
        }.get(self.current_section, self._handle_intake)

        response = await handler(user_message)

        # Track tier2 state on the orchestrator instance
        if response.awaiting_tier2:
            self._awaiting_tier2 = response.awaiting_tier2

        self.history.append(ConversationTurn("assistant", response.message))
        return response

    # ── Section handlers ──────────────────────────────────────────────────────

    async def _handle_intake(self, user_message: str) -> "OrchestratorResponse":
        """
        Intake: detect business type, run benchmarks, then greedily extract
        ALL sections from the first message. Only ask for what is missing.
        """
        # ── Step 1: Detect business type + extract everything we can ─────────
        extraction_prompt = (
            "The user has sent their first message describing a business project. "
            "Extract as much information as possible across ALL sections.\n\n"
            f'User message: "{user_message}"\n\n'
            "Extract:\n"
            "1. business_description: 1-2 sentence clean description\n"
            "2. industry: specific sector (e.g. Cricket Bat Manufacturing)\n"
            "3. business_type: MANUFACTURING | TRADING | SERVICE | MIXED\n"
            "4. revenue_model_type: MANUFACTURING | TRADING | SERVICE_CAPACITY | SERVICE_TIME | SERVICE_SUBSCRIPTION | MIXED\n"
            "5. location_hint: any location mentioned\n"
            "6. profile: {company_name, promoter_name, entity_type (Proprietorship/Partnership/LLP/Company), city, state, operation_start_date (YYYY-MM), projection_years}\n"
            "7. capital: {assets: [{name, cost_lakhs, category (Civil Works/Plant & Machinery/Furniture & Fixture/Vehicle/Other)}], "
            "term_loan_amount, term_loan_rate (fraction), term_loan_tenor_months, moratorium_months, od_limit, promoter_equity}\n"
            "8. revenue: {products: [{name, unit, price_per_unit, capacity_per_day, output_ratio (default 1.0), split_percent}], "
            "year1_utilization (fraction), annual_utilization_increment (fraction), max_utilization (fraction), working_days_per_month}\n"
            "9. costs: {raw_materials: [{name, unit, price_per_unit, input_per_output_unit, price_escalation_pa (default 0.05)}], "
            "transport_base_lakhs, misc_base_lakhs}\n"
            "10. manpower: {categories: [{designation, count, monthly_salary_lakhs, annual_increment_pa (default 0.05)}]}\n"
            "11. finance: {debtor_days, creditor_days_rm, stock_days_rm, implementation_months}\n\n"
            "Return ONLY JSON. Use null for fields not mentioned. Numbers must be plain numbers (no ₹ symbol, no units in values).\n"
            "{\"business_description\":\"...\",\"industry\":\"...\",\"business_type\":\"MANUFACTURING\","
            "\"revenue_model_type\":\"MANUFACTURING\",\"location_hint\":\"...\","
            "\"profile\":{\"company_name\":\"...\",\"promoter_name\":\"...\",\"entity_type\":\"...\","
            "\"city\":\"...\",\"state\":\"...\",\"operation_start_date\":\"...\",\"projection_years\":7},"
            "\"capital\":{\"assets\":[],\"term_loan_amount\":null,\"term_loan_rate\":null,"
            "\"term_loan_tenor_months\":null,\"moratorium_months\":null,\"od_limit\":null,\"promoter_equity\":null},"
            "\"revenue\":{\"products\":[],\"year1_utilization\":null,\"annual_utilization_increment\":null,"
            "\"max_utilization\":null,\"working_days_per_month\":null},"
            "\"costs\":{\"raw_materials\":[],\"transport_base_lakhs\":null,\"misc_base_lakhs\":null},"
            "\"manpower\":{\"categories\":[]},"
            "\"finance\":{\"debtor_days\":null,\"creditor_days_rm\":null,\"stock_days_rm\":null,\"implementation_months\":null}}"
        )

        raw = await self._llm_call(extraction_prompt, system=self._extraction_system(), max_tokens=2000)
        try:
            extracted = json.loads(self._clean_json(raw))
        except Exception:
            extracted = {}

        # ── Step 2: Write business type to store ─────────────────────────────
        bt = extracted.get("business_type", "MANUFACTURING")
        rm = extracted.get("revenue_model_type", "MANUFACTURING")
        try:
            self.store.business_type = BusinessType(bt)
        except Exception:
            self.store.business_type = BusinessType.MANUFACTURING
        try:
            self.store.revenue_model_type = RevenueModelType(rm)
        except Exception:
            self.store.revenue_model_type = RevenueModelType.MANUFACTURING

        self.store.project_profile.industry = extracted.get("industry", "")
        self._intake_description = extracted.get("business_description", user_message)
        self._intake_location    = extracted.get("location_hint", "India")

        # ── Step 3: Apply all sections that were found ────────────────────────
        if extracted.get("profile"):
            try:
                self._apply_profile_fields(extracted["profile"])
            except Exception as e:
                print(f"[intake] profile extract error: {e}")

        if extracted.get("capital"):
            try:
                self._apply_capital_fields_from_intake(extracted["capital"])
            except Exception as e:
                print(f"[intake] capital extract error: {e}")

        if extracted.get("revenue"):
            try:
                self._apply_revenue_fields(extracted["revenue"])
            except Exception as e:
                print(f"[intake] revenue extract error: {e}")

        if extracted.get("costs"):
            try:
                self._apply_costs_fields(extracted["costs"])
            except Exception as e:
                print(f"[intake] costs extract error: {e}")

        if extracted.get("manpower"):
            try:
                self._apply_manpower_fields(extracted["manpower"])
            except Exception as e:
                print(f"[intake] manpower extract error: {e}")

        if extracted.get("finance"):
            try:
                self._apply_finance_fields(extracted["finance"])
            except Exception as e:
                print(f"[intake] finance extract error: {e}")

        # ── Step 4: Run benchmark engine ──────────────────────────────────────
        be = BenchmarkEngine(self.model, api_key=self.api_key)
        self.benchmarks = await be.generate(
            business_description=self._intake_description,
            industry=self.store.project_profile.industry,
            location=self._intake_location,
            project_cost_lakhs=self.store.capital_means.total_project_cost,
        )

        # ── Step 5: Mark sections complete if all fields present, advance ─────
        self._auto_complete_sections()

        # ── Step 6: Build response — jump to first incomplete section ─────────
        reply = self._build_intake_confirmation(extracted)
        return OrchestratorResponse(
            message=reply,
            section_completed="intake",
            next_section=self.current_section,
        )

    def _apply_capital_fields_from_intake(self, d: dict):
        """Apply capital fields extracted during intake."""
        from core.session_store import Asset, FinanceSource, AssetCategory, FinanceSourceType
        cm = self.store.capital_means

        # Assets
        cat_map = {
            "civil": AssetCategory.CIVIL_WORKS,
            "plant": AssetCategory.PLANT_MACHINERY,
            "machinery": AssetCategory.PLANT_MACHINERY,
            "furniture": AssetCategory.FURNITURE,
            "vehicle": AssetCategory.VEHICLE,
            "electrical": AssetCategory.ELECTRICAL,
        }
        for a in (d.get("assets") or []):
            name = a.get("name", "")
            cost = float(a.get("cost_lakhs") or 0)
            if cost <= 0:
                continue
            cat_hint = a.get("category", name).lower()
            category = AssetCategory.OTHER
            for k, v in cat_map.items():
                if k in cat_hint:
                    category = v
                    break
            cm.assets.append(Asset(name=name, category=category, cost_lakhs=cost))

        # Finance sources
        if d.get("term_loan_amount"):
            cm.finance_sources.append(FinanceSource(
                source_type=FinanceSourceType.TERM_LOAN,
                amount_lakhs=float(d["term_loan_amount"]),
                interest_rate=float(d.get("term_loan_rate") or 0.095),
                tenor_months=int(d.get("term_loan_tenor_months") or 84),
                moratorium_months=int(d.get("moratorium_months") or 0),
            ))
        if d.get("od_limit"):
            cm.finance_sources.append(FinanceSource(
                source_type=FinanceSourceType.OD_LIMIT,
                amount_lakhs=float(d["od_limit"]),
                interest_rate=float(d.get("term_loan_rate") or 0.095),
            ))
        if d.get("promoter_equity"):
            cm.finance_sources.append(FinanceSource(
                source_type=FinanceSourceType.PROMOTER_EQUITY,
                amount_lakhs=float(d["promoter_equity"]),
            ))

    def _apply_manpower_fields(self, d: dict):
        """Apply manpower fields."""
        from core.session_store import EmployeeCategory
        mp = self.store.manpower
        for c in (d.get("categories") or []):
            desig = c.get("designation", "")
            if not desig:
                continue
            mp.categories.append(EmployeeCategory(
                designation=desig,
                count=int(c.get("count") or 1),
                monthly_salary_lakhs=float(c.get("monthly_salary_lakhs") or 0.15),
                is_fixed=True,
                annual_increment_pa=float(c.get("annual_increment_pa") or 0.05),
            ))

    def _apply_finance_fields(self, d: dict):
        """Apply working capital fields."""
        fw = self.store.finance_wc
        if d.get("debtor_days") is not None:
            fw.debtor_days = int(d["debtor_days"])
        if d.get("creditor_days_rm") is not None:
            fw.creditor_days_rm = int(d["creditor_days_rm"])
        if d.get("stock_days_rm") is not None:
            fw.stock_days_rm = int(d["stock_days_rm"])
        if d.get("implementation_months") is not None:
            fw.implementation_months = int(d["implementation_months"])

    def _auto_complete_sections(self):
        """Mark sections complete if all required fields are present, and advance current_section."""
        sections_in_order = ["profile", "capital", "revenue", "costs", "manpower", "finance", "confirm"]

        # Profile
        if not self._missing_profile_fields():
            self.store.project_profile.status = SectionStatus.COMPLETE

        # Capital
        if self.store.capital_means.assets and self.store.capital_means.finance_sources:
            self.store.capital_means.status = SectionStatus.COMPLETE

        # Revenue
        if not self._missing_revenue_fields():
            self.store.revenue_model.status = SectionStatus.COMPLETE

        # Costs
        if not self._missing_costs_fields():
            self.store.cost_structure.status = SectionStatus.COMPLETE

        # Manpower
        if self.store.manpower.categories:
            self.store.manpower.status = SectionStatus.COMPLETE

        # Finance/WC
        fw = self.store.finance_wc
        if fw.debtor_days != -1 and fw.creditor_days_rm != -1 and fw.implementation_months != -1:
            self.store.finance_wc.status = SectionStatus.COMPLETE

        # Set current_section to first incomplete section
        status_map = {
            "profile":  self.store.project_profile.status,
            "capital":  self.store.capital_means.status,
            "revenue":  self.store.revenue_model.status,
            "costs":    self.store.cost_structure.status,
            "manpower": self.store.manpower.status,
            "finance":  self.store.finance_wc.status,
        }
        for sec in sections_in_order[:-1]:
            if status_map.get(sec) != SectionStatus.COMPLETE:
                self.current_section = sec
                return
        # All complete
        self.current_section = "confirm"

    async def _handle_profile(self, user_message: str) -> "OrchestratorResponse":
        """Extract project profile fields from conversation."""
        missing = self._missing_tier1_fields("profile")

        extraction_prompt = f"""Extract project profile information from the user's message.
Current store state:
  company_name:    "{self.store.project_profile.company_name}"
  promoter_name:   "{self.store.project_profile.promoter_name}"
  entity_type:     "{self.store.project_profile.entity_type.value}"
  city:            "{self.store.project_profile.city}"
  state:           "{self.store.project_profile.state}"
  operation_start: "{self.store.project_profile.operation_start_date}"
  projection_years: {self.store.project_profile.projection_years}

User message: "{user_message}"

Extract any of these fields present in the message. For entity_type, map to one of: Proprietorship, Partnership, LLP, Company.
For operation_start_date, format as YYYY-MM.

Respond ONLY with JSON of fields found (omit fields not mentioned):
{{"company_name":"...", "promoter_name":"...", "entity_type":"...", "city":"...", "state":"...", "operation_start_date":"...", "projection_years": 7}}"""

        raw = await self._llm_call(extraction_prompt, system=self._extraction_system())
        try:
            extracted = json.loads(self._clean_json(raw))
            self._apply_profile_fields(extracted)
        except Exception:
            pass

        # Check what's still missing
        missing_after = self._missing_profile_fields()

        if not missing_after:
            self.store.project_profile.status = SectionStatus.COMPLETE
            self.current_section = "capital"
            summary = self._profile_summary()
            next_q  = self._first_question_for("capital")
            return OrchestratorResponse(
                message=f"{summary}\n\n✅ Project profile complete.\n\n---\n\n{next_q}",
                section_completed="profile",
                next_section="capital",
            )
        else:
            next_q = self._ask_for_missing_profile(missing_after)
            return OrchestratorResponse(message=next_q)

    async def _handle_capital(self, user_message: str) -> "OrchestratorResponse":
        """Extract capital and finance fields."""
        extraction_prompt = f"""Extract capital cost and finance information from the user message.

Current assets: {json.dumps([{"name":a.name,"cost":a.cost_lakhs,"category":a.category.value} for a in self.store.capital_means.assets], indent=2)}
Current loans:  {json.dumps([{"type":f.source_type.value,"amount":f.amount_lakhs,"rate":f.rate_pa,"tenor":f.tenor_months,"moratorium":f.moratorium_months} for f in self.store.capital_means.finance_sources], indent=2)}

User message: "{user_message}"

Extract:
- assets: list of {{"name":"...", "cost_lakhs": <number>, "category": "Civil Works|Plant & Machinery|Furniture & Fixture|Vehicle|Electrical & Fittings|Other"}}
- term_loans: list of {{"amount_lakhs":<n>, "rate_pa":<fraction>, "tenor_months":<int>, "moratorium_months":<int>, "label":"..."}}
- od_limit_lakhs: <number or 0>
- od_rate_pa: <fraction>

Respond ONLY with JSON. Include ONLY fields explicitly mentioned by user.
{{"assets":[], "term_loans":[], "od_limit_lakhs":0, "od_rate_pa":0.09}}"""

        raw = await self._llm_call(extraction_prompt, system=self._extraction_system())
        try:
            extracted = json.loads(self._clean_json(raw))
            self._apply_capital_fields(extracted)
        except Exception:
            pass

        missing = self._missing_capital_fields()
        if not missing:
            self._compute_promoter_contribution()
            self.store.capital_means.status = SectionStatus.COMPLETE
            self.current_section = "revenue"
            summary  = self._capital_summary()
            next_q   = self._first_question_for("revenue")
            return OrchestratorResponse(
                message=f"{summary}\n\n✅ Capital & means complete.\n\n---\n\n{next_q}",
                section_completed="capital",
                next_section="revenue",
            )
        else:
            return OrchestratorResponse(
                message=self._ask_for_missing_capital(missing)
            )

    async def _handle_revenue(self, user_message: str) -> "OrchestratorResponse":
        """Extract revenue model fields."""
        extraction_prompt = f"""Extract revenue model information.

Business type: {self.store.business_type.value}
Revenue model: {self.store.revenue_model_type.value}
Current products: {json.dumps([{"name":p.name,"unit":p.unit,"price":p.price_per_unit,"capacity":p.capacity_per_day,"yield":p.output_ratio,"split":p.split_percent} for p in self.store.revenue_model.products], indent=2)}
Current util: {{"year1": {self.store.revenue_model.year1_utilization}, "increment": {self.store.revenue_model.annual_utilization_increment}, "max": {self.store.revenue_model.max_utilization}, "working_days": {self.store.revenue_model.working_days_per_month}}}

User message: "{user_message}"

Extract:
- products: list of {{"name":"...", "unit":"...", "price_per_unit":<n>, "capacity_per_day":<n>, "output_ratio":<0-1>, "split_percent":<0-1>}}
- year1_utilization: <fraction 0-1>
- annual_utilization_increment: <fraction>
- max_utilization: <fraction>
- working_days_per_month: <integer>

Respond ONLY with JSON. Include ONLY fields explicitly mentioned.
{{"products":[], "year1_utilization":null, "annual_utilization_increment":null, "max_utilization":null, "working_days_per_month":null}}"""

        raw = await self._llm_call(extraction_prompt, system=self._extraction_system())
        try:
            extracted = json.loads(self._clean_json(raw))
            self._apply_revenue_fields(extracted)
        except Exception:
            pass

        missing = self._missing_revenue_fields()
        if not missing:
            # Now apply Tier 2 benchmarks for revenue section
            tier2_msg = await self._present_tier2_for_section("revenue", user_message)
            if tier2_msg:
                return OrchestratorResponse(message=tier2_msg,
                                             awaiting_tier2="revenue")

            self.store.revenue_model.status = SectionStatus.COMPLETE
            self.current_section = "costs"
            return OrchestratorResponse(
                message=f"✅ Revenue model complete.\n\n---\n\n{self._first_question_for('costs')}",
                section_completed="revenue",
                next_section="costs",
            )
        return OrchestratorResponse(
            message=self._ask_for_missing_revenue(missing)
        )

    async def _handle_costs(self, user_message: str) -> "OrchestratorResponse":
        """Extract cost structure fields."""
        extraction_prompt = (
            "Extract cost structure information from the user message.\n\n"
            f"Already captured raw materials: {__import__('json').dumps([{'name':m.name,'unit':m.unit,'price':m.price_per_unit,'input_per_output':m.input_per_output_unit} for m in self.store.cost_structure.raw_materials])}\n"
            f"Current: transport={self.store.cost_structure.transport_base_lakhs}, misc={self.store.cost_structure.misc_base_lakhs}\n\n"
            f'User message: "{user_message}"\n\n'
            "Extract ALL raw materials mentioned. For each:\n"
            "  - name: string\n"
            "  - unit: string (kg, pieces, litres, rolls, etc.)\n"
            "  - price_per_unit: number (INR, no symbols)\n"
            "  - input_per_output_unit: number (quantity of this input per 1 finished unit)\n"
            "  - price_escalation_pa: number (default 0.05)\n"
            "Also extract transport_base_lakhs and misc_base_lakhs as numbers (null if not mentioned).\n\n"
            "IMPORTANT: price_per_unit must be a plain number like 180 not rupee symbol.\n"
            "IMPORTANT: input_per_output_unit must be a plain number like 1.2.\n"
            "IMPORTANT: If user lists 5 materials return all 5.\n\n"
            'Respond ONLY with JSON: {"raw_materials":[{"name":"...","unit":"...","price_per_unit":0,"input_per_output_unit":0,"price_escalation_pa":0.05}],"transport_base_lakhs":null,"misc_base_lakhs":null}'
        )

        raw = await self._llm_call(extraction_prompt, system=self._extraction_system())
        try:
            import json as _json
            extracted = _json.loads(self._clean_json(raw))
            self._apply_costs_fields(extracted)
        except Exception as e:
            print(f"[costs extraction] failed: {e} | raw response: {raw[:300]}")

    async def _handle_manpower(self, user_message: str) -> "OrchestratorResponse":
        """Extract manpower fields."""
        extraction_prompt = f"""Extract manpower/staffing information.

Current categories: {json.dumps([{"designation":c.designation,"count":c.count,"salary_lakhs":c.monthly_salary_lakhs} for c in self.store.manpower.categories], indent=2)}

User message: "{user_message}"

Extract a list of employee categories. Each must have designation, count, monthly_salary_lakhs.
Respond ONLY with JSON:
{{"categories": [{{"designation":"...", "count":<int>, "monthly_salary_lakhs":<float>}}]}}"""

        raw = await self._llm_call(extraction_prompt, system=self._extraction_system())
        try:
            extracted = json.loads(self._clean_json(raw))
            if extracted.get("categories"):
                self.store.manpower.categories = [
                    EmployeeCategory(
                        designation=c["designation"],
                        count=int(c["count"]),
                        monthly_salary_lakhs=float(c["monthly_salary_lakhs"]),
                        annual_increment_pa=0.05,  # will be set by Tier 2
                    )
                    for c in extracted["categories"]
                    if c.get("designation") and c.get("count") and c.get("monthly_salary_lakhs")
                ]
        except Exception:
            pass

        if self.store.manpower.categories:
            tier2_msg = await self._present_tier2_for_section("manpower", user_message)
            if tier2_msg:
                return OrchestratorResponse(message=tier2_msg,
                                             awaiting_tier2="manpower")
            self.store.manpower.status = SectionStatus.COMPLETE
            self.current_section = "finance"
            return OrchestratorResponse(
                message=f"✅ Manpower complete.\n\n---\n\n{self._first_question_for('finance')}",
                section_completed="manpower",
                next_section="finance",
            )
        return OrchestratorResponse(
            message="Could you tell me about your staffing plan? "
                    "For each role, I need: designation, number of people, and monthly salary. "
                    "For example: '1 manager at ₹40,000, 5 operators at ₹15,000, 1 guard at ₹12,000'"
        )

    async def _handle_finance(self, user_message: str) -> "OrchestratorResponse":
        """Extract working capital days."""
        extraction_prompt = f"""Extract working capital cycle information.

Current: debtor_days={self.store.finance_wc.debtor_days}, creditor_days_rm={self.store.finance_wc.creditor_days_rm}, implementation_months={self.store.finance_wc.implementation_months}

User message: "{user_message}"

Extract:
- debtor_days: how many days customers take to pay
- creditor_days_rm: how many days before paying raw material suppliers
- implementation_months: months to build/install before operations start

Respond ONLY with JSON (include only fields mentioned):
{{"debtor_days":null, "creditor_days_rm":null, "implementation_months":null}}"""

        raw = await self._llm_call(extraction_prompt, system=self._extraction_system())
        try:
            extracted = json.loads(self._clean_json(raw))
            fw = self.store.finance_wc
            if extracted.get("debtor_days") is not None:
                fw.debtor_days = int(extracted["debtor_days"])
            if extracted.get("creditor_days_rm") is not None:
                fw.creditor_days_rm = int(extracted["creditor_days_rm"])
            if extracted.get("implementation_months") is not None:
                fw.implementation_months = int(extracted["implementation_months"])
        except Exception:
            pass

        missing = self._missing_finance_fields()
        if not missing:
            tier2_msg = await self._present_tier2_for_section("finance", user_message)
            if tier2_msg:
                return OrchestratorResponse(message=tier2_msg,
                                             awaiting_tier2="finance")
            self.store.finance_wc.status = SectionStatus.COMPLETE
            self.current_section = "confirm"
            summary = self._full_summary_for_confirmation()
            return OrchestratorResponse(
                message=summary,
                section_completed="finance",
                next_section="confirm",
            )
        return OrchestratorResponse(
            message=self._ask_for_missing_finance(missing)
        )

    async def _handle_confirm(self, user_message: str) -> "OrchestratorResponse":
        """Handle confirmation / correction from user."""
        user_lower = user_message.lower().strip()

        # Check if user is confirming
        confirm_words = ["yes", "confirm", "ok", "looks good", "correct",
                         "proceed", "generate", "go ahead", "perfect", "right"]
        is_confirming = any(w in user_lower for w in confirm_words)

        if is_confirming:
            errors = self.store.validate_completeness()
            if errors:
                return OrchestratorResponse(
                    message=f"There are a few things still missing before I can generate:\n\n"
                            + "\n".join(f"• {e}" for e in errors),
                )
            return OrchestratorResponse(
                message="✅ All details confirmed. Generating your DPR now — this will take about 30 seconds...",
                ready_to_generate=True,
            )

        # User wants to correct something — send back through extraction
        correction_prompt = f"""The user wants to correct something in their DPR data.
Current full summary: {self._compact_summary()}
User correction: "{user_message}"

What field(s) is the user trying to change? Extract the corrected values.
Respond with a brief acknowledgement and the corrected summary."""

        response = await self._llm_call(correction_prompt,
                                         system=self._extraction_system())
        return OrchestratorResponse(
            message=f"{response}\n\nShall I regenerate the summary, or are you ready to generate the DPR?"
        )

    # ── Tier 2 presentation ───────────────────────────────────────────────────

    async def _present_tier2_for_section(
        self, section: str, last_message: str
    ) -> Optional[str]:
        """
        If this section has Tier 2 fields with benchmarks available,
        build the batch presentation message.
        Returns None if no Tier 2 fields for this section.
        """
        t2_fields = tier2_for_section(section)
        if not t2_fields:
            return None

        bm_engine = BenchmarkEngine(api_key=self.api_key)
        msg = BenchmarkEngine.format_batch_for_display(
            benchmarks=self.benchmarks,
            fields=t2_fields,
            context={"business": self._intake_description,
                     "industry": self.store.project_profile.industry},
        )
        return msg

    async def _apply_tier2_responses(
        self, section: str, user_message: str
    ) -> str:
        """
        Parse user's responses to Tier 2 benchmark fields.
        User can press enter (accept) or type a value (override).
        """
        t2_fields = tier2_for_section(section)
        if not t2_fields:
            return ""

        extraction_prompt = f"""The user was shown benchmark defaults for these financial parameters
and asked to accept (enter) or override with their own value.

Fields shown:
{json.dumps([{"key": f.key, "label": f.label, "benchmark": self.benchmarks.get(f.key, {}).value if f.key in self.benchmarks else "N/A"} for f in t2_fields], indent=2)}

User response: "{user_message}"

Extract any overrides the user provided. If user said things like "ok", "yes", "fine",
"accept all", "looks good" — treat all fields as accepted (use benchmark values).

Respond ONLY with JSON dict of key -> value for fields the user explicitly changed:
{{"key1": new_value, "key2": new_value}}
Or {{}} if user accepted all defaults."""

        raw = await self._llm_call(extraction_prompt, system=self._extraction_system())
        try:
            overrides = json.loads(self._clean_json(raw))
        except Exception:
            overrides = {}

        # Apply benchmark defaults, then override where user changed
        cs  = self.store.cost_structure
        fw  = self.store.finance_wc
        rm  = self.store.revenue_model

        for f in t2_fields:
            bm = self.benchmarks.get(f.key)
            if not bm:
                continue

            user_overrode = f.key in overrides
            value = float(overrides[f.key]) if user_overrode else bm.value

            # Write to store
            self._write_tier2_to_store(f.key, value)

            # Log it
            self.assumption_log.add_tier2(
                f.key, f.label, value, f.unit, bm, user_overrode, f.section
            )

        return ""

    def _write_tier2_to_store(self, key: str, value: float):
        """Write a Tier 2 value to the correct store field."""
        cs = self.store.cost_structure
        fw = self.store.finance_wc
        rm = self.store.revenue_model

        mapping = {
            "cost_structure.rm_pct_of_fa":            lambda v: setattr(cs, "rm_pct_of_fa",         v),
            "cost_structure.rm_escalation_pa":         lambda v: setattr(cs, "rm_escalation_pa",      v),
            "cost_structure.insurance_pct_of_fa":      lambda v: setattr(cs, "insurance_pct_of_fa",   v),
            "cost_structure.power_pct_revenue":         lambda v: setattr(cs, "power_pct_revenue",     v),
            "cost_structure.power_escalation_pa":       lambda v: setattr(cs, "power_escalation_pa",   v),
            "cost_structure.marketing_pct_revenue":     lambda v: setattr(cs, "marketing_pct_revenue", v),
            "cost_structure.sga_base_lakhs":            lambda v: setattr(cs, "sga_base_lakhs",        v),
            "cost_structure.transport_escalation_pa":   lambda v: setattr(cs, "transport_escalation_pa",v),
            "cost_structure.misc_escalation_pa":        lambda v: setattr(cs, "misc_escalation_pa",    v),
            "finance_wc.creditor_days_admin":           lambda v: setattr(fw, "creditor_days_admin",   int(v)),
            "finance_wc.stock_days_rm":                 lambda v: setattr(fw, "stock_days_rm",         int(v)),
            "finance_wc.od_rate":                       lambda v: None,  # handled in capital section
        }
        fn = mapping.get(key)
        if fn:
            fn(value)
        # Per-product/per-employee fields handled separately

    # ── LLM call ──────────────────────────────────────────────────────────────

    async def _llm_call(self, user_prompt: str,
                         system: str = "",
                         max_tokens: int = 2000) -> str:
        """Make a Claude API call and return the text response."""
        import aiohttp

        messages = [{"role": "user", "content": user_prompt}]
        payload = {
            "model": self.model,
            "max_tokens": max_tokens,
            "messages": messages,
        }
        if system:
            payload["system"] = system

        headers = {"Content-Type": "application/json",
                   "anthropic-version": "2023-06-01"}
        if self.api_key:
            headers["x-api-key"] = self.api_key

        try:
            async with aiohttp.ClientSession() as session:
                async with session.post(
                    "https://api.anthropic.com/v1/messages",
                    json=payload,
                    headers=headers,
                    timeout=aiohttp.ClientTimeout(total=25)
                ) as resp:
                    data = await resp.json()
                    return data["content"][0]["text"]
        except Exception as e:
            return f'{{"error": "{str(e)}"}}'

    def _extraction_system(self) -> str:
        return (
            "You are a precise data extraction assistant for an Indian MSME "
            "financial modelling system. Extract structured data from user messages "
            "and respond ONLY with valid JSON. Never add prose or markdown. "
            "Convert percentages to decimals (6% → 0.06). "
            "Convert crores to lakhs (1 crore = 100 lakhs). "
            "If a value is not mentioned, use null."
        )

    # ── Store field appliers ──────────────────────────────────────────────────

    def _apply_profile_fields(self, d: dict):
        pp = self.store.project_profile
        if d.get("company_name"):   pp.company_name  = d["company_name"]
        if d.get("promoter_name"):  pp.promoter_name = d["promoter_name"]
        if d.get("entity_type"):
            try: pp.entity_type = EntityType(d["entity_type"])
            except: pass
        if d.get("city"):  pp.city  = d["city"]
        if d.get("state"): pp.state = d["state"]
        if d.get("operation_start_date"):
            pp.operation_start_date = d["operation_start_date"]
        if d.get("projection_years"):
            pp.projection_years = int(d["projection_years"])

    def _apply_capital_fields(self, d: dict):
        cm = self.store.capital_means
        # Assets
        if d.get("assets"):
            existing_names = {a.name.lower() for a in cm.assets}
            for a in d["assets"]:
                if a.get("name") and a.get("cost_lakhs") and \
                   a["name"].lower() not in existing_names:
                    try:
                        cat = AssetCategory(a.get("category", "Other"))
                    except:
                        cat = AssetCategory.OTHER
                    cm.assets.append(Asset(
                        name=a["name"],
                        category=cat,
                        cost_lakhs=float(a["cost_lakhs"]),
                    ))
                    existing_names.add(a["name"].lower())
        # Term loans
        if d.get("term_loans"):
            existing = {f.label for f in cm.finance_sources
                        if f.source_type == FinanceSourceType.TERM_LOAN}
            for tl in d["term_loans"]:
                label = tl.get("label", f"Term Loan {len(cm.term_loans)+1}")
                if label not in existing:
                    cm.finance_sources.append(FinanceSource(
                        source_type=FinanceSourceType.TERM_LOAN,
                        amount_lakhs=float(tl.get("amount_lakhs", 0)),
                        rate_pa=float(tl.get("rate_pa", 0.09)),
                        tenor_months=int(tl.get("tenor_months", 84)),
                        moratorium_months=int(tl.get("moratorium_months", 0)),
                        label=label,
                    ))
        # OD
        if d.get("od_limit_lakhs") and float(d["od_limit_lakhs"]) > 0:
            # Remove existing OD and replace
            cm.finance_sources = [f for f in cm.finance_sources
                                   if f.source_type != FinanceSourceType.OD_LIMIT]
            cm.finance_sources.append(FinanceSource(
                source_type=FinanceSourceType.OD_LIMIT,
                amount_lakhs=float(d["od_limit_lakhs"]),
                rate_pa=float(d.get("od_rate_pa", 0.09)),
            ))

    def _compute_promoter_contribution(self):
        """Add promoter equity = total cost - all other sources."""
        cm = self.store.capital_means
        other = sum(f.amount_lakhs for f in cm.finance_sources
                    if f.source_type != FinanceSourceType.PROMOTER_EQUITY)
        promoter = cm.total_project_cost - other
        if promoter > 0:
            cm.finance_sources = [f for f in cm.finance_sources
                                   if f.source_type != FinanceSourceType.PROMOTER_EQUITY]
            cm.finance_sources.append(FinanceSource(
                source_type=FinanceSourceType.PROMOTER_EQUITY,
                amount_lakhs=round(promoter, 2),
                label="Promoter Contribution",
            ))

    def _apply_revenue_fields(self, d: dict):
        rm = self.store.revenue_model
        if d.get("products"):
            existing = {p.name.lower() for p in rm.products}
            for p in d["products"]:
                if p.get("name") and p["name"].lower() not in existing:
                    rm.products.append(Product(
                        name=p["name"],
                        unit=p.get("unit", "units"),
                        capacity_per_day=float(p.get("capacity_per_day", 0)),
                        output_ratio=float(p.get("output_ratio", 1.0)),
                        split_percent=float(p.get("split_percent", 1.0)),
                        price_per_unit=float(p.get("price_per_unit", 0)),
                        price_escalation_pa=float(p.get("price_escalation_pa", 0.04)),
                    ))
                    existing.add(p["name"].lower())
        if d.get("year1_utilization") is not None:
            rm.year1_utilization = float(d["year1_utilization"])
        if d.get("annual_utilization_increment") is not None:
            rm.annual_utilization_increment = float(d["annual_utilization_increment"])
        if d.get("max_utilization") is not None:
            rm.max_utilization = float(d["max_utilization"])
        if d.get("working_days_per_month") is not None:
            rm.working_days_per_month = int(d["working_days_per_month"])

    def _apply_costs_fields(self, d: dict):
        cs = self.store.cost_structure
        if d.get("raw_materials"):
            existing = {m.name.lower() for m in cs.raw_materials}
            for m in d["raw_materials"]:
                if m.get("name") and m["name"].lower() not in existing:
                    cs.raw_materials.append(RawMaterial(
                        name=m["name"],
                        unit=m.get("unit", "kg"),
                        price_per_unit=float(m.get("price_per_unit", 0)),
                        input_per_output_unit=float(m.get("input_per_output_unit", 1)),
                        price_escalation_pa=float(m.get("price_escalation_pa", 0.05)),
                    ))
                    existing.add(m["name"].lower())
        if d.get("transport_base_lakhs") is not None:
            cs.transport_base_lakhs = float(d["transport_base_lakhs"])
        if d.get("misc_base_lakhs") is not None:
            cs.misc_base_lakhs = float(d["misc_base_lakhs"])

    # ── Missing field checkers ────────────────────────────────────────────────

    def _missing_profile_fields(self) -> list[str]:
        pp = self.store.project_profile
        missing = []
        if not pp.company_name:           missing.append("company_name")
        if not pp.promoter_name:          missing.append("promoter_name")
        if not pp.city:                   missing.append("city")
        if not pp.state:                  missing.append("state")
        if not pp.operation_start_date:   missing.append("operation_start_date")
        if pp.projection_years < 1:       missing.append("projection_years")
        return missing

    def _missing_capital_fields(self) -> list[str]:
        cm = self.store.capital_means
        missing = []
        if not cm.assets:                 missing.append("assets")
        if not cm.term_loans:             missing.append("term_loan")
        return missing

    def _missing_revenue_fields(self) -> list[str]:
        rm = self.store.revenue_model
        missing = []
        if not rm.products:               missing.append("products")
        if rm.year1_utilization <= 0:     missing.append("year1_utilization")
        if rm.working_days_per_month <= 0: missing.append("working_days_per_month")
        return missing

    def _missing_costs_fields(self) -> list[str]:
        cs = self.store.cost_structure
        missing = []
        if not cs.raw_materials:          missing.append("raw_materials")
        if cs.transport_base_lakhs < 0:   missing.append("transport_base_lakhs")
        return missing

    def _missing_finance_fields(self) -> list[str]:
        fw = self.store.finance_wc
        missing = []
        if fw.debtor_days < 0:            missing.append("debtor_days")
        if fw.implementation_months < 0:  missing.append("implementation_months")
        return missing

    def _missing_tier1_fields(self, section: str) -> list[str]:
        return []  # Simplified — full implementation uses field_registry

    # ── Question generators ───────────────────────────────────────────────────

    def _first_question_for(self, section: str) -> str:
        questions = {
            "profile":  "Let's start with your project details.\n\n"
                        "**What is the name of your company, and who is the promoter?**\n"
                        "Also tell me your location (city & state) and when you expect to start operations.",
            "capital":  "Now let's cover your project costs.\n\n"
                        "**What assets will you be creating or buying?** "
                        "List each item and its estimated cost (e.g. 'Civil Works: ₹200 lakhs, Plant & Machinery: ₹110 lakhs').\n"
                        "Also tell me your term loan amount and OD limit.",
            "revenue":  "Let's model your revenue.\n\n"
                        "**What do you sell, and at what price?** "
                        "Tell me the product name, unit of sale, selling price per unit, and your production capacity per day.\n"
                        "Also: what % of capacity do you expect to use in year 1?",
            "costs":    "Now let's cover your input costs.\n\n"
                        "**What raw materials do you purchase?** "
                        "For each, tell me: name, unit, price per unit, and how much you need per unit of output.\n"
                        "Also: your estimated transport and miscellaneous costs per year.",
            "manpower": "Tell me about your team.\n\n"
                        "**List each type of employee** with their count and monthly salary.\n"
                        "Example: '1 Manager at ₹40,000, 5 Operators at ₹15,000, 1 Guard at ₹12,000'",
            "finance":  "Almost done! A few working capital questions.\n\n"
                        "**Debtor days**: How long do your customers take to pay after a sale?\n"
                        "**Creditor days**: How long before you pay your suppliers?\n"
                        "**Implementation**: How many months to set up the project before operations start?",
        }
        return questions.get(section, f"Let's move to {SECTION_DISPLAY_NAMES.get(section, section)}.")

    def _ask_for_missing_profile(self, missing: list[str]) -> str:
        pp = self.store.project_profile
        parts = []
        if "company_name" in missing:
            parts.append("the **company name**")
        if "promoter_name" in missing:
            parts.append("the **promoter's name**")
        if "city" in missing or "state" in missing:
            parts.append("the **project location** (city and state)")
        if "operation_start_date" in missing:
            parts.append("the **expected start date** for commercial operations")
        if "projection_years" in missing:
            parts.append("the **number of projection years** (typically 7)")
        return "I still need " + ", ".join(parts) + "."

    def _ask_for_missing_capital(self, missing: list[str]) -> str:
        parts = []
        if "assets" in missing:
            parts.append("Please list the **assets** you'll create or buy and their costs (e.g. Civil Works, Plant & Machinery)")
        if "term_loan" in missing:
            parts.append("What **term loan amount** are you seeking from the bank?")
        return "\n\n".join(parts)

    def _ask_for_missing_revenue(self, missing: list[str]) -> str:
        parts = []
        if "products" in missing:
            parts.append("What **products or services** do you sell? Include name, unit, price per unit, and capacity per day.")
        if "year1_utilization" in missing:
            parts.append("What **% of capacity** will you use in Year 1? (e.g. 50%)")
        if "working_days_per_month" in missing:
            parts.append("How many **days per month** will you operate?")
        return "\n\n".join(parts)

    def _ask_for_missing_costs(self, missing: list[str]) -> str:
        parts = []
        if "raw_materials" in missing:
            parts.append("What **raw materials** do you use? For each: name, unit, price per unit, and input needed per output unit.")
        return "\n\n".join(parts)

    def _ask_for_missing_finance(self, missing: list[str]) -> str:
        parts = []
        if "debtor_days" in missing:
            parts.append("How many **days** do your customers take to pay after a sale?")
        if "implementation_months" in missing:
            parts.append("How many **months** will the construction/installation take before commercial operations start?")
        return "\n\n".join(parts)

    # ── Summary builders ──────────────────────────────────────────────────────

    def _build_intake_confirmation(self, extracted: dict) -> str:
        biz_type = extracted.get("business_type", "MANUFACTURING").replace("_", " ").title()
        industry = extracted.get("industry", "")

        # Build a summary of what was already captured
        completed = []
        pp = self.store.project_profile
        cm = self.store.capital_means
        rv = self.store.revenue_model
        cs = self.store.cost_structure
        mp = self.store.manpower
        fw = self.store.finance_wc

        if pp.status == SectionStatus.COMPLETE:
            completed.append(f"✅ Project profile ({pp.company_name}, {pp.city})")
        if cm.status == SectionStatus.COMPLETE:
            completed.append(f"✅ Capital & means ({len(cm.assets)} assets, ₹{cm.total_project_cost:.0f}L project cost)")
        if rv.status == SectionStatus.COMPLETE:
            completed.append(f"✅ Revenue model ({len(rv.products)} products)")
        if cs.status == SectionStatus.COMPLETE:
            completed.append(f"✅ Cost structure ({len(cs.raw_materials)} raw materials)")
        if mp.status == SectionStatus.COMPLETE:
            completed.append(f"✅ Manpower ({len(mp.categories)} categories)")
        if fw.status == SectionStatus.COMPLETE:
            completed.append(f"✅ Working capital (debtor days: {fw.debtor_days})")

        intro = (
            f"Got it. I understand you're setting up a **{industry}** unit "
            f"({biz_type} type business).\n\n"
            f"I've generated industry-specific benchmarks for your financial assumptions.\n\n"
        )

        if completed:
            intro += f"**Already extracted from your message:**\n" + "\n".join(completed) + "\n\n"

        if self.current_section == "confirm":
            intro += "🎉 All sections complete! Here's your full summary for confirmation:\n\n"
            intro += self._full_summary_for_confirmation()
        else:
            first_q = self._first_question_for(self.current_section)
            missing_label = {
                "profile": "project details",
                "capital": "capital & finance",
                "revenue": "revenue model",
                "costs": "cost structure",
                "manpower": "manpower",
                "finance": "working capital",
            }.get(self.current_section, self.current_section)
            intro += f"Still need a few details about your **{missing_label}**:\n\n---\n\n{first_q}"

        return intro

    def _profile_summary(self) -> str:
        pp = self.store.project_profile
        return (
            f"**Project Profile:**\n"
            f"• Company: {pp.company_name}\n"
            f"• Promoter: {pp.promoter_name}\n"
            f"• Entity: {pp.entity_type.value}\n"
            f"• Location: {pp.city}, {pp.state}\n"
            f"• Operations start: {pp.operation_start_date}\n"
            f"• Projection years: {pp.projection_years}"
        )

    def _capital_summary(self) -> str:
        cm = self.store.capital_means
        lines = ["**Capital & Means:**"]
        for a in cm.assets:
            lines.append(f"• {a.name}: ₹{a.cost_lakhs:.2f}L")
        lines.append(f"  **Total cost: ₹{cm.total_project_cost:.2f}L**")
        lines.append("")
        for f in cm.finance_sources:
            lines.append(f"• {f.source_type.value}: ₹{f.amount_lakhs:.2f}L")
        lines.append(f"  **Balanced: {'✓' if cm.is_balanced else '✗'}**")
        return "\n".join(lines)

    def _full_summary_for_confirmation(self) -> str:
        return (
            f"## Review Your DPR Inputs\n\n"
            f"{self._profile_summary()}\n\n"
            f"{self._capital_summary()}\n\n"
            f"**Please review the above. Type 'confirm' to generate the DPR, "
            f"or tell me what needs to be changed.**"
        )

    def _compact_summary(self) -> str:
        return json.dumps(self.store.to_dict(), indent=2)[:2000]

    # ── Utilities ─────────────────────────────────────────────────────────────

    @staticmethod
    def _clean_json(text: str) -> str:
        text = text.strip()
        text = re.sub(r'^```json\s*', '', text)
        text = re.sub(r'^```\s*',     '', text)
        text = re.sub(r'\s*```$',     '', text)
        return text


# ─── Response object ──────────────────────────────────────────────────────────

@dataclass
class OrchestratorResponse:
    message: str
    section_completed: str = ""
    next_section: str = ""
    ready_to_generate: bool = False
    awaiting_tier2: str = ""
