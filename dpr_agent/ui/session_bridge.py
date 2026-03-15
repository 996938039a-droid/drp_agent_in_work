"""
ui/session_bridge.py
─────────────────────
Connects Streamlit's st.session_state to the DPR domain objects.

Streamlit rerenders the entire script on every user interaction.
All stateful objects (SessionStore, Orchestrator, conversation history)
must live in st.session_state to survive rerenders.

This module provides a single get_or_create() call that initialises
everything on first load and returns the live objects thereafter.

Usage (in app.py):
    bridge = SessionBridge.get()
    response = bridge.send_message(user_input)
"""

from __future__ import annotations
import streamlit as st
from dataclasses import dataclass, field
from typing import Optional

# Lazy imports to avoid circular dependency at module load time
def _import_core():
    from core.session_store import SessionStore
    from agents.orchestrator import Orchestrator
    from api_client.claude_client import ClaudeClient
    return SessionStore, Orchestrator, ClaudeClient


@dataclass
class ChatMessage:
    role: str    # "user" | "assistant"
    content: str


class SessionBridge:
    """
    Manages all session state for the DPR application.
    Call SessionBridge.get() to retrieve or initialise the bridge.
    """

    # Keys used in st.session_state
    _KEY_BRIDGE      = "_dpr_bridge"
    _KEY_API_KEY     = "_dpr_api_key"
    _KEY_HISTORY     = "_dpr_history"
    _KEY_STORE       = "_dpr_store"
    _KEY_ORCH        = "_dpr_orchestrator"
    _KEY_READY       = "_dpr_ready_to_generate"
    _KEY_OUTPUT_PATH = "_dpr_output_path"
    _KEY_SECTION     = "_dpr_current_section"

    # ── Initialisation ────────────────────────────────────────────────────────

    @classmethod
    def get(cls) -> "SessionBridge":
        """Return the existing bridge or create a new one."""
        if cls._KEY_BRIDGE not in st.session_state:
            bridge = cls()
            bridge._init_state()
            st.session_state[cls._KEY_BRIDGE] = bridge
        return st.session_state[cls._KEY_BRIDGE]

    def _init_state(self):
        """Set up all session_state keys with defaults."""
        SessionStore, Orchestrator, ClaudeClient = _import_core()
        if self._KEY_HISTORY not in st.session_state:
            st.session_state[self._KEY_HISTORY] = []
        if self._KEY_STORE not in st.session_state:
            st.session_state[self._KEY_STORE] = SessionStore()
        if self._KEY_READY not in st.session_state:
            st.session_state[self._KEY_READY] = False
        if self._KEY_OUTPUT_PATH not in st.session_state:
            st.session_state[self._KEY_OUTPUT_PATH] = None
        if self._KEY_SECTION not in st.session_state:
            st.session_state[self._KEY_SECTION] = "intake"
        # Orchestrator is initialised lazily (needs API key)
        if self._KEY_ORCH not in st.session_state:
            st.session_state[self._KEY_ORCH] = None

    # ── API key management ────────────────────────────────────────────────────

    def set_api_key(self, key: str):
        """Called when the user enters/changes their API key in the sidebar."""
        from api_client.claude_client import ClaudeClient
        from agents.orchestrator import Orchestrator

        valid, err = ClaudeClient.validate_key(key)
        if not valid:
            raise ValueError(err)

        st.session_state[self._KEY_API_KEY] = key.strip()

        # Re-create the orchestrator with the new key
        store = st.session_state[self._KEY_STORE]
        orch  = Orchestrator(store=store, api_key=key.strip())
        st.session_state[self._KEY_ORCH] = orch

    @property
    def api_key(self) -> Optional[str]:
        return st.session_state.get(self._KEY_API_KEY)

    @property
    def has_api_key(self) -> bool:
        return bool(self.api_key)

    @property
    def orchestrator(self):
        return st.session_state.get(self._KEY_ORCH)

    @property
    def store(self):
        return st.session_state.get(self._KEY_STORE)

    # ── Conversation ──────────────────────────────────────────────────────────

    @property
    def history(self) -> list[ChatMessage]:
        return st.session_state.get(self._KEY_HISTORY, [])

    def add_message(self, role: str, content: str):
        st.session_state[self._KEY_HISTORY].append(
            ChatMessage(role=role, content=content)
        )

    def send_message(self, user_text: str) -> str:
        """
        Process one user message through the Orchestrator.
        Returns the assistant reply as a string.
        Raises ValueError if no API key is set.
        """
        if not self.has_api_key:
            raise ValueError(
                "No API key set. Please enter your Anthropic API key in the sidebar."
            )

        orch = self.orchestrator
        if orch is None:
            raise ValueError("Orchestrator not initialised. Please re-enter your API key.")

        # Add user message to history
        self.add_message("user", user_text)

        # Run the orchestrator (sync wrapper)
        import asyncio
        import concurrent.futures

        async def _run():
            return await orch.process_message(user_text)

        try:
            loop = asyncio.get_event_loop()
            if loop.is_running():
                with concurrent.futures.ThreadPoolExecutor() as pool:
                    future = pool.submit(asyncio.run, _run())
                    response = future.result(timeout=60)
            else:
                response = loop.run_until_complete(_run())
        except Exception as e:
            reply = f"⚠️ Error processing message: {str(e)}"
            self.add_message("assistant", reply)
            return reply

        # Record state changes
        st.session_state[self._KEY_SECTION] = orch.current_section
        if response.ready_to_generate:
            st.session_state[self._KEY_READY] = True

        reply = response.message
        self.add_message("assistant", reply)
        return reply

    # ── Generation ────────────────────────────────────────────────────────────

    @property
    def ready_to_generate(self) -> bool:
        return st.session_state.get(self._KEY_READY, False)

    @property
    def output_path(self) -> Optional[str]:
        return st.session_state.get(self._KEY_OUTPUT_PATH)

    def generate_dpr(self) -> str:
        """
        Build the Excel workbook and return the file path.
        Called when the user confirms all inputs.
        """
        from excel.workbook_builder import WorkbookBuilder
        import os, datetime

        store = self.store
        if not store:
            raise RuntimeError("No session store found.")

        # Create outputs directory if needed
        outputs_dir = os.path.join(
            os.path.dirname(os.path.dirname(__file__)), "outputs"
        )
        os.makedirs(outputs_dir, exist_ok=True)

        # Filename: DPR_CompanyName_YYYYMMDD.xlsx
        safe_name = "".join(
            c for c in store.project_profile.company_name
            if c.isalnum() or c in (" ", "_", "-")
        ).strip().replace(" ", "_")[:30]
        date_str  = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        filename  = f"DPR_{safe_name}_{date_str}.xlsx"
        out_path  = os.path.join(outputs_dir, filename)

        WorkbookBuilder(store).build(out_path)
        st.session_state[self._KEY_OUTPUT_PATH] = out_path
        return out_path

    # ── Progress ──────────────────────────────────────────────────────────────

    @property
    def current_section(self) -> str:
        return st.session_state.get(self._KEY_SECTION, "intake")

    def get_progress(self) -> dict:
        """Return section completion status for the sidebar progress indicator."""
        orch = self.orchestrator
        if orch is None:
            return {}
        store = self.store
        if store is None:
            return {}

        from core.session_store import SectionStatus
        sections = {
            "intake":   "Business Description",
            "profile":  "Project Profile",
            "capital":  "Capital & Means",
            "revenue":  "Revenue Model",
            "costs":    "Cost Structure",
            "manpower": "Manpower",
            "finance":  "Working Capital",
            "confirm":  "Confirmation",
        }
        result = {}
        status_map = {
            SectionStatus.COMPLETE:     "complete",
            SectionStatus.IN_PROGRESS:  "in_progress",
            SectionStatus.PENDING:      "pending",
        }
        for key, label in sections.items():
            sec_status = SectionStatus.PENDING
            if key == "profile":
                sec_status = store.project_profile.status
            elif key == "capital":
                sec_status = store.capital_means.status
            elif key == "revenue":
                sec_status = store.revenue_model.status
            elif key == "costs":
                sec_status = store.cost_structure.status
            elif key == "manpower":
                sec_status = store.manpower.status
            elif key == "finance":
                sec_status = store.finance_wc.status

            result[key] = {
                "label":   label,
                "status":  status_map.get(sec_status, "pending"),
                "current": key == self.current_section,
            }
        return result

    # ── Reset ─────────────────────────────────────────────────────────────────

    def reset(self):
        """Start a fresh DPR session (keeps the API key)."""
        from core.session_store import SessionStore
        from agents.orchestrator import Orchestrator

        api_key = self.api_key
        st.session_state[self._KEY_HISTORY]     = []
        st.session_state[self._KEY_STORE]       = SessionStore()
        st.session_state[self._KEY_READY]       = False
        st.session_state[self._KEY_OUTPUT_PATH] = None
        st.session_state[self._KEY_SECTION]     = "intake"

        if api_key:
            new_store = st.session_state[self._KEY_STORE]
            st.session_state[self._KEY_ORCH] = Orchestrator(
                store=new_store, api_key=api_key
            )
