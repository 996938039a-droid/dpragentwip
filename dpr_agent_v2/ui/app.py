"""
ui/app.py
══════════
Main Streamlit entry point for the DPR Agent v2.
Run with: streamlit run ui/app.py
"""

import streamlit as st
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ui.session_bridge import get_orchestrator, run_async
from ui.sidebar import render_sidebar
from ui.chat import render_chat, render_input

st.set_page_config(
    page_title="DPR Agent",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Session state init ────────────────────────────────────────────────────────
if "messages" not in st.session_state:
    st.session_state.messages = []
if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = None
if "excel_filename" not in st.session_state:
    st.session_state.excel_filename = "DPR.xlsx"
if "api_key" not in st.session_state:
    st.session_state.api_key = os.getenv("ANTHROPIC_API_KEY", "")

# ── Sidebar ───────────────────────────────────────────────────────────────────
render_sidebar()

# ── Main area ─────────────────────────────────────────────────────────────────
st.title("📊 DPR Agent")
st.caption("Project Finance Report Generator for MSMEs")

# Show chat history
render_chat(st.session_state.messages)

# Input
user_input = render_input()

if user_input:
    api_key = st.session_state.get("api_key", "")
    if not api_key:
        st.error("Please enter your Anthropic API key in the sidebar.")
        st.stop()

    # Add user message to display
    st.session_state.messages.append({"role": "user", "content": user_input})

    # Get orchestrator and process
    orchestrator = get_orchestrator(api_key)

    with st.spinner("Thinking..."):
        response = run_async(orchestrator.process(user_input))

    # Add assistant message
    st.session_state.messages.append({
        "role": "assistant",
        "content": response.message,
    })

    # Handle Excel generation
    if response.ready_to_generate:
        with st.spinner("Generating Excel workbook..."):
            try:
                from excel.workbook_builder import build_workbook_bytes, suggested_filename
                store = orchestrator.get_store()
                excel_bytes = build_workbook_bytes(store)
                st.session_state.excel_bytes    = excel_bytes
                st.session_state.excel_filename = suggested_filename(store)
                st.session_state.messages.append({
                    "role": "assistant",
                    "content": "✅ Your DPR Excel is ready! Click **Download DPR** in the sidebar.",
                })
            except Exception as e:
                st.session_state.messages.append({
                    "role": "assistant",
                    "content": f"⚠️ Error generating Excel: {str(e)}",
                })

    st.rerun()
