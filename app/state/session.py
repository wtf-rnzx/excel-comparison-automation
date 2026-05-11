"""Session-state bootstrap helpers for the Streamlit app."""

from __future__ import annotations

import streamlit as st

from app.storage.local_json import LocalJSONStorage
from app.ui.mapping import initialize_mapping_state


def initialize_session_state(mapping_storage: LocalJSONStorage) -> None:
    """Initialize mapping and comparison-related session state values."""
    initialize_mapping_state(mapping_storage)

    if "comparison_result" not in st.session_state:
        st.session_state.comparison_result = None
    if "report_bytes" not in st.session_state:
        st.session_state.report_bytes = None
    if "source_filename" not in st.session_state:
        st.session_state.source_filename = None
    if "target_filename" not in st.session_state:
        st.session_state.target_filename = None
    if "active_mappings" not in st.session_state:
        st.session_state.active_mappings = None
