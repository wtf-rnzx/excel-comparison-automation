"""UI components for configurable value mappings."""

from __future__ import annotations

from copy import deepcopy

import streamlit as st

from app.config.settings import DEFAULT_MAPPING_TEMPLATES, MAPPING_SECTION_META
from app.storage.local_json import LocalJSONStorage


def initialize_mapping_state(storage: LocalJSONStorage) -> None:
    """Populate mapping session keys from local storage if missing."""
    stored = storage.load_data()
    for _, _, _, session_key in MAPPING_SECTION_META:
        if session_key not in st.session_state:
            st.session_state[session_key] = stored.get(session_key, [])


def render_column_mapping_ui(storage: LocalJSONStorage) -> None:
    """Render complete mapping UI with tabs and CRUD actions."""
    with st.expander("🔧 Advanced: Column Value Mappings", expanded=False):
        st.markdown(
            """
            Define custom value equivalencies for validation. When a source value maps to a target value,
            they will be treated as matching during comparison.

            **Example:** Map "Main" → "Pablo Borbon" so rows with "Main" in source match "Pablo Borbon" in target.
            """
        )

        tab_titles = [meta[0] for meta in MAPPING_SECTION_META]
        tabs = st.tabs(tab_titles)

        for tab, (_, field_id, field_label, session_key) in zip(tabs, MAPPING_SECTION_META):
            with tab:
                render_mapping_section(storage, field_id, field_label, session_key)


def _update_mapping_input(storage: LocalJSONStorage, field_id: str, idx: int, session_key: str, field: str) -> None:
    widget_key = f"{field_id}_{field}_{idx}"
    value = st.session_state.get(widget_key, "")
    st.session_state[session_key] = storage.update_record(session_key, idx, field, value)


def render_mapping_section(storage: LocalJSONStorage, field_id: str, field_label: str, session_key: str) -> None:
    """Render a single mapping section."""
    st.subheader(f"{field_label} Mappings")

    mappings = st.session_state.get(session_key, [])

    if mappings:
        st.markdown("**Current Mappings:**")
        for idx, mapping in enumerate(mappings):
            col1, col2, col3, col4 = st.columns([3, 1, 3, 1])
            with col1:
                st.text_input(
                    "Source value",
                    value=mapping.get("source", ""),
                    key=f"{field_id}_source_{idx}",
                    label_visibility="collapsed",
                    on_change=lambda i=idx, k=session_key, fid=field_id: _update_mapping_input(
                        storage, fid, i, k, "source"
                    ),
                )
            with col2:
                st.markdown("<div style='text-align: center; padding-top: 8px;'>→</div>", unsafe_allow_html=True)
            with col3:
                st.text_input(
                    "Target value",
                    value=mapping.get("target", ""),
                    key=f"{field_id}_target_{idx}",
                    label_visibility="collapsed",
                    on_change=lambda i=idx, k=session_key, fid=field_id: _update_mapping_input(
                        storage, fid, i, k, "target"
                    ),
                )
            with col4:
                if st.button("🗑️", key=f"{field_id}_delete_{idx}", help="Delete mapping"):
                    st.session_state[session_key] = storage.delete_record(session_key, idx)
                    st.rerun()

        st.divider()

    st.markdown("**Add New Mapping:**")
    col1, col2, col3, col4 = st.columns([3, 1, 3, 1])

    with col1:
        new_source_key = f"{field_id}_new_source"
        new_source = st.text_input(
            "Source value",
            key=new_source_key,
            placeholder="e.g., Main Campus",
            label_visibility="collapsed",
        )

    with col2:
        st.markdown("<div style='text-align: center; padding-top: 8px;'>→</div>", unsafe_allow_html=True)

    with col3:
        new_target_key = f"{field_id}_new_target"
        new_target = st.text_input(
            "Target value",
            key=new_target_key,
            placeholder="e.g., Pablo Borbon",
            label_visibility="collapsed",
        )

    with col4:
        if st.button("➕ Add", key=f"{field_id}_add", type="primary"):
            if new_source and new_target:
                st.session_state[session_key] = storage.add_record(session_key, new_source, new_target)
                st.session_state[new_source_key] = ""
                st.session_state[new_target_key] = ""
                st.rerun()
            else:
                st.warning("Both source and target values are required")

    template_records = DEFAULT_MAPPING_TEMPLATES.get(field_id)
    if template_records and st.button(f"📋 Load Default {field_label} Mappings", key=f"{field_id}_load_defaults"):
        st.session_state[session_key] = storage.replace_records(session_key, deepcopy(template_records))
        st.rerun()
