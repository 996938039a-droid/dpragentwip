"""ui/chat.py — chat message rendering"""
import streamlit as st

def render_chat(messages: list):
    for msg in messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

def render_input() -> str | None:
    return st.chat_input("Type your message...")
