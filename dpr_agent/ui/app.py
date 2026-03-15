"""
ui/app.py
──────────
Main Streamlit application entry point.

Run with:
    cd dpr_agent
    streamlit run ui/app.py

Architecture:
  - Sidebar: API key + progress + download (ui/sidebar.py)
  - Main area: Chat interface (ui/chat.py)
  - State bridge: Connects Streamlit ↔ SessionStore ↔ Orchestrator
    (ui/session_bridge.py)
  - API client: Injects API key into every Claude call
    (api_client/claude_client.py)
"""

import sys
import os

# Make sure dpr_agent root is on the path when running from ui/
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
from ui.sidebar import render_sidebar
from ui.chat    import render_chat


# ── Page config ───────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="DPR Agent",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        "Get help":    None,
        "Report a bug": None,
        "About": "DPR Agent — AI-powered Detailed Project Report generator for MSME businesses.",
    },
)

# ── Global styles ─────────────────────────────────────────────────────────────

st.markdown(
    """
    <style>
    /* Chat message alignment */
    .stChatMessage {
        max-width: 820px;
    }

    /* Sidebar header */
    .css-1d391kg {
        padding-top: 1rem;
    }

    /* Progress section in sidebar */
    section[data-testid="stSidebar"] .stMarkdown p {
        font-size: 0.83rem;
        margin: 0;
    }

    /* Download button styling */
    .stDownloadButton button {
        background-color: #1F3864 !important;
        color: white !important;
        border: none !important;
    }

    /* Remove top padding on main container */
    .block-container {
        padding-top: 1rem;
    }

    /* Code blocks in chat */
    .stChatMessage code {
        background: #f0f4f8;
        padding: 1px 4px;
        border-radius: 3px;
        font-size: 0.85em;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ── Render ────────────────────────────────────────────────────────────────────

render_sidebar()
render_chat()
