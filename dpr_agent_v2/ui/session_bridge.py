"""
ui/session_bridge.py
══════════════════════
Bridges Streamlit's session state with the async Orchestrator.
Uses a dedicated thread for each async call — compatible with Python 3.14.
"""

import asyncio
import concurrent.futures
import streamlit as st
from agents.orchestrator import Orchestrator


def get_orchestrator(api_key: str) -> Orchestrator:
    """Get or create the orchestrator, persisted in session state."""
    if "orchestrator" not in st.session_state or \
       st.session_state.get("orchestrator_key") != api_key:
        st.session_state.orchestrator     = Orchestrator(api_key)
        st.session_state.orchestrator_key = api_key
    return st.session_state.orchestrator


def run_async(coro):
    """
    Run an async coroutine from Streamlit's sync thread.
    Always uses a fresh thread with its own event loop — works on Python 3.14
    compatible with Python 3.14 threading model.
    """
    def _run(c):
        return asyncio.run(c)

    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as pool:
        future = pool.submit(_run, coro)
        return future.result()
