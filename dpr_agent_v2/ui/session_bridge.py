"""
ui/session_bridge.py
══════════════════════
Bridges Streamlit's session state with the async Orchestrator.
Handles asyncio correctly for Streamlit's threading model.
"""

import asyncio
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
    Run an async coroutine from sync Streamlit context.
    Handles both cases: new event loop needed, or existing loop.
    """
    try:
        loop = asyncio.get_event_loop()
        if loop.is_running():
            # In Streamlit's threading context, use a new thread
            import concurrent.futures
            with concurrent.futures.ThreadPoolExecutor() as pool:
                future = pool.submit(asyncio.run, coro)
                return future.result()
        else:
            return loop.run_until_complete(coro)
    except RuntimeError:
        return asyncio.run(coro)
