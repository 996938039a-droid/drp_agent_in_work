"""
ui/chat.py
───────────
Renders the chat interface in the main area.

Handles:
  - Message history rendering (user + assistant bubbles)
  - Chat input form
  - Generation trigger when ready_to_generate is True
  - Generation spinner and completion state
  - Tier 2 benchmark tables (formatted markdown)
"""

import streamlit as st
from ui.session_bridge import SessionBridge


def render_chat():
    """Main entry point — call from app.py."""
    bridge = SessionBridge.get()

    _render_header(bridge)
    _render_messages(bridge)
    _render_generation_block(bridge)
    _render_input(bridge)


# ── Header ────────────────────────────────────────────────────────────────────

def _render_header(bridge: SessionBridge):
    if not bridge.history:
        st.markdown(
            """
            <div style="text-align:center; padding:2rem 0 1rem 0;">
                <h2 style="color:#1F3864; font-weight:700; margin-bottom:0.25rem;">
                    Welcome to DPR Agent
                </h2>
                <p style="color:#555; font-size:1rem; max-width:500px; margin:0 auto;">
                    I'll guide you through building a complete Detailed Project Report
                    for your MSME business — ready for bank submission.
                </p>
                <p style="color:#888; font-size:0.85rem; margin-top:0.75rem;">
                    Just describe your business to get started.
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        # Compact section indicator
        section_labels = {
            "intake":   "📝 Business Description",
            "profile":  "🏢 Project Profile",
            "capital":  "💰 Capital & Means",
            "revenue":  "📈 Revenue Model",
            "costs":    "🏭 Cost Structure",
            "manpower": "👥 Manpower",
            "finance":  "🏦 Working Capital",
            "confirm":  "✅ Confirmation",
        }
        label = section_labels.get(bridge.current_section, bridge.current_section)
        st.markdown(
            f"<div style='font-size:0.8rem; color:#888; margin-bottom:0.5rem;'>"
            f"Current section: <strong style='color:#2E75B6'>{label}</strong></div>",
            unsafe_allow_html=True,
        )


# ── Message history ───────────────────────────────────────────────────────────

def _render_messages(bridge: SessionBridge):
    for msg in bridge.history:
        if msg.role == "user":
            with st.chat_message("user"):
                st.markdown(msg.content)
        else:
            with st.chat_message("assistant", avatar="🤖"):
                # Render markdown — benchmark tables come through as markdown
                st.markdown(msg.content)


# ── Generation block ──────────────────────────────────────────────────────────

def _render_generation_block(bridge: SessionBridge):
    """
    Show the Generate DPR button when ready_to_generate is True
    and no output has been produced yet.
    """
    if bridge.ready_to_generate and not bridge.output_path:
        st.divider()
        st.markdown(
            "### 🎯 All inputs confirmed — ready to generate your DPR"
        )
        col1, col2 = st.columns([2, 1])
        with col1:
            st.info(
                "Your DPR Excel workbook (14 sheets, all formulas) will be "
                "generated now. This takes about 5–10 seconds.",
                icon="ℹ️",
            )
        with col2:
            if st.button("⚡ Generate DPR", type="primary",
                          use_container_width=True):
                _do_generate(bridge)

    elif bridge.output_path:
        st.success(
            "✅ DPR generated successfully! Download it from the sidebar.",
            icon="✅",
        )


def _do_generate(bridge: SessionBridge):
    with st.spinner("Building your DPR Excel workbook…"):
        try:
            out_path = bridge.generate_dpr()
            bridge.add_message(
                "assistant",
                "✅ **Your DPR is ready!**\n\n"
                "The Excel workbook with all 14 sheets has been generated. "
                "Click **⬇ Download Excel** in the sidebar to save it.\n\n"
                "The workbook contains:\n"
                "- Assumption sheet (single source of truth)\n"
                "- Revenue, ManPower, Depreciation\n"
                "- Expenses, Term Loan schedule, Working Capital, Tax\n"
                "- **P&L Account, Balance Sheet, Cash Flow Statement**\n"
                "- **Key Financial Ratios** (DSCR, ROCE, BEP, Margins)\n\n"
                "All figures are formula-linked — change any assumption and "
                "the entire model recalculates automatically."
            )
            st.rerun()
        except Exception as e:
            st.error(f"Generation failed: {e}")


# ── Chat input ────────────────────────────────────────────────────────────────

def _render_input(bridge: SessionBridge):
    """Render the chat input box and process submissions."""
    if not bridge.has_api_key:
        st.warning(
            "👈 Enter your Anthropic API key in the sidebar to start.",
            icon="⚠️",
        )
        return

    # Don't show input if DPR is already generated
    if bridge.output_path:
        st.divider()
        col1, col2 = st.columns([3, 1])
        with col1:
            st.info("DPR generated. Start a new session to build another.")
        return

    placeholder = _get_placeholder(bridge)
    user_input  = st.chat_input(placeholder)

    if user_input and user_input.strip():
        # Show user message immediately (before API round-trip)
        with st.chat_message("user"):
            st.markdown(user_input.strip())

        # Send to Orchestrator and get reply
        with st.chat_message("assistant", avatar="🤖"):
            with st.spinner("Thinking…"):
                try:
                    reply = bridge.send_message(user_input.strip())
                    st.markdown(reply)
                except Exception as e:
                    st.error(f"Error: {e}")

        # Rerun to refresh progress bar in sidebar
        st.rerun()


def _get_placeholder(bridge: SessionBridge) -> str:
    placeholders = {
        "intake":   "Describe your business (e.g. 'mustard oil processing unit in Indore...')",
        "profile":  "Enter company name, promoter, location, start date...",
        "capital":  "List assets and costs (e.g. 'Civil ₹200L, Machinery ₹110L, Loan ₹200L at 9%...')",
        "revenue":  "Describe products, prices, and capacity...",
        "costs":    "List raw materials with prices and usage ratios...",
        "manpower": "List staff roles with headcount and salaries...",
        "finance":  "Debtor days, creditor days, implementation period...",
        "confirm":  "Type 'confirm' to generate, or tell me what to change...",
    }
    return placeholders.get(bridge.current_section, "Type your message...")
