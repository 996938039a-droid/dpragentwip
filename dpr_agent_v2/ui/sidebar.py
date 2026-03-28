"""ui/sidebar.py — sidebar with API key, progress, download"""
import streamlit as st

SECTIONS = ["profile","capital","revenue","costs","manpower","wc","review","done"]

def render_sidebar():
    with st.sidebar:
        st.header("⚙️ Settings")

        # API key
        api_key = st.text_input(
            "Anthropic API Key",
            value=st.session_state.get("api_key",""),
            type="password",
            help="Your key is used only for this session and never stored."
        )
        if api_key:
            st.session_state.api_key = api_key

        st.divider()

        # Progress
        st.subheader("📝 Progress")
        orch = st.session_state.get("orchestrator")
        if orch:
            store = orch.get_store()
            current = store.current_section
            for sec in ["profile","capital","revenue","costs","manpower","wc"]:
                done = store.section_complete(sec)
                icon = "✅" if done else ("🔄" if sec == current else "⬜")
                st.write(f"{icon} {sec.title()}")
        else:
            st.caption("Start a conversation to see progress.")

        st.divider()

        # Download
        st.subheader("📥 Download")
        if st.session_state.get("excel_bytes"):
            st.download_button(
                label="⬇️ Download DPR Excel",
                data=st.session_state.excel_bytes,
                file_name=st.session_state.get("excel_filename","DPR.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.caption("Excel will appear here after generation.")

        st.divider()

        # Reset
        if st.button("🔄 New DPR", use_container_width=True):
            for key in ["orchestrator","orchestrator_key","messages",
                        "excel_bytes","excel_filename"]:
                st.session_state.pop(key, None)
            st.rerun()
