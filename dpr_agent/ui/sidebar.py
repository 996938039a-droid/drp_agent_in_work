"""
ui/sidebar.py
──────────────
Renders the Streamlit sidebar.

Contains:
  1. API Key input (password field, validated client-side)
  2. Session progress tracker (8 sections with tick/pending/current)
  3. New DPR button (reset)
  4. Download button (appears only after generation is complete)
"""

import streamlit as st
from ui.session_bridge import SessionBridge


def render_sidebar():
    """Call this once from app.py to render the entire sidebar."""
    bridge = SessionBridge.get()

    with st.sidebar:
        _render_logo()
        _render_api_key(bridge)
        st.divider()
        _render_progress(bridge)
        st.divider()
        _render_actions(bridge)


# ── Logo / title ──────────────────────────────────────────────────────────────

def _render_logo():
    st.markdown(
        """
        <div style="text-align:center; padding: 0.5rem 0 1rem 0;">
            <div style="font-size:1.6rem; font-weight:700; letter-spacing:-0.5px; color:#1F3864;">
                DPR Agent
            </div>
            <div style="font-size:0.75rem; color:#888; margin-top:2px;">
                Detailed Project Report Generator
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ── API key ───────────────────────────────────────────────────────────────────

def _render_api_key(bridge: SessionBridge):
    st.markdown("#### 🔑 Anthropic API Key")

    # Auto-load from environment variable if not already set
    if not bridge.has_api_key:
        import os
        env_key = os.environ.get("ANTHROPIC_API_KEY", "")
        if env_key:
            try:
                bridge.set_api_key(env_key)
                st.rerun()
            except ValueError:
                pass  # Bad env key — fall through to manual input

    if bridge.has_api_key:
        masked = bridge.api_key[:12] + "..." + bridge.api_key[-4:]
        st.success(f"Key active: `{masked}`")
        if st.button("Change key", use_container_width=True):
            st.session_state["_show_key_input"] = True
            st.rerun()
    else:
        st.session_state["_show_key_input"] = True

    if st.session_state.get("_show_key_input", False) or not bridge.has_api_key:
        with st.form("api_key_form", clear_on_submit=False):
            raw_key = st.text_input(
                "Enter your API key",
                type="password",
                placeholder="sk-ant-api03-...",
                help="Your key is stored only in this browser session and never saved to disk.",
            )
            submitted = st.form_submit_button("✓ Confirm", use_container_width=True)

            if submitted and raw_key:
                try:
                    bridge.set_api_key(raw_key)
                    st.session_state["_show_key_input"] = False
                    st.success("API key set successfully!")
                    st.rerun()
                except ValueError as e:
                    st.error(str(e))

    if not bridge.has_api_key:
        st.info(
            "💡 Get your key at [console.anthropic.com](https://console.anthropic.com)",
            icon="ℹ️",
        )


# ── Progress tracker ──────────────────────────────────────────────────────────

def _render_progress(bridge: SessionBridge):
    st.markdown("#### 📋 Progress")

    if not bridge.has_api_key:
        st.caption("Enter your API key to begin.")
        return

    progress = bridge.get_progress()
    if not progress:
        st.caption("Start chatting to begin your DPR.")
        return

    STATUS_ICONS = {
        "complete":    "✅",
        "in_progress": "🔵",
        "pending":     "⚪",
    }

    for key, info in progress.items():
        icon   = STATUS_ICONS.get(info["status"], "⚪")
        label  = info["label"]
        is_cur = info["current"]

        if is_cur:
            st.markdown(
                f"<div style='background:#EBF5FB; border-left:3px solid #2E75B6; "
                f"padding:4px 8px; border-radius:0 4px 4px 0; margin:2px 0; "
                f"font-size:0.85rem; font-weight:600;'>{icon} {label}</div>",
                unsafe_allow_html=True,
            )
        else:
            color = "#1A7A4A" if info["status"] == "complete" else "#999"
            st.markdown(
                f"<div style='padding:3px 8px; margin:1px 0; "
                f"font-size:0.82rem; color:{color};'>{icon} {label}</div>",
                unsafe_allow_html=True,
            )

    # Overall progress bar
    total    = len(progress)
    complete = sum(1 for v in progress.values() if v["status"] == "complete")
    if total > 0:
        st.progress(complete / total, text=f"{complete}/{total} sections done")


# ── Actions ───────────────────────────────────────────────────────────────────

def _render_actions(bridge: SessionBridge):
    # Download button (only after generation)
    if bridge.output_path:
        import os
        if os.path.exists(bridge.output_path):
            with open(bridge.output_path, "rb") as f:
                xlsx_bytes = f.read()

            filename = os.path.basename(bridge.output_path)
            st.markdown("#### 📥 Download DPR")
            st.download_button(
                label="⬇ Download Excel (.xlsx)",
                data=xlsx_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )
            st.caption(f"File: `{filename}`")
            st.divider()

    # New DPR / Reset button
    if bridge.history:
        if st.button("🔄 Start New DPR", use_container_width=True):
            bridge.reset()
            st.rerun()

    # Session stats at the bottom
    if bridge.history:
        user_count = sum(1 for m in bridge.history if m.role == "user")
        st.caption(f"Messages sent: {user_count}")
