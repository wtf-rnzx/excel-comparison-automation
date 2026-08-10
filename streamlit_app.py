"""Streamlit UI for Excel/CSV Comparison Automation

Installation:
    pip install streamlit

Run:
    streamlit run streamlit_app.py
"""

from __future__ import annotations

from pathlib import Path

import streamlit as st

from app.config.settings import DEFAULT_MAPPING_STATE, DEFAULT_SHEET_INPUT, SUPPORTED_UPLOAD_TYPES
from app.state.session import initialize_session_state
from app.services.comparison import run_comparison
from app.storage.local_json import LocalJSONStorage
from app.ui.mapping import render_column_mapping_ui
from app.ui.results import display_results


MAPPING_STORAGE = LocalJSONStorage(
    file_path=Path(__file__).parent / "mappings.json",
    default_data=DEFAULT_MAPPING_STATE,
)


st.set_page_config(
    page_title="Excel Comparison Automation",
    page_icon="📊",
    layout="wide",
)


def main():
    initialize_session_state(MAPPING_STORAGE)
    
    st.title("📊 Automation Tool for Excel/CSV Comparison")
    st.subheader("Compare and Validate BatstateU Student Data Across Files")
    st.markdown("""
    Compare two Excel or CSV files by matching names and validate data consistency.
    Upload your **source-of-truth** file and **target** file to get started.
    """)
    
    st.divider()
    
    # File uploads
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Source File (Source of Truth)")
        source_file = st.file_uploader(
            "Upload source file",
            type=SUPPORTED_UPLOAD_TYPES,
            key="source",
            help="The source-of-truth file to compare against"
        )
        
        if source_file:
            st.success(f"✓ {source_file.name}")
            source_sheet = st.text_input(
                "Source sheet name or index",
                value=DEFAULT_SHEET_INPUT,
                key="source_sheet",
                help="Sheet name (e.g., 'Sheet1') or index (0 for first sheet)"
            )
            source_column = st.text_input(
                "Source column name (optional)",
                value="",
                key="source_column",
                help="Leave blank to auto-detect fullname/FULLNAME column"
            )
    
    with col2:
        st.subheader("Target File (To Compare)")
        target_file = st.file_uploader(
            "Upload target file",
            type=SUPPORTED_UPLOAD_TYPES,
            key="target",
            help="The target file to compare against the source"
        )
        
        if target_file:
            st.success(f"✓ {target_file.name}")
            target_sheet = st.text_input(
                "Target sheet name or index",
                value=DEFAULT_SHEET_INPUT,
                key="target_sheet",
                help="Sheet name (e.g., 'Sheet1') or index (0 for first sheet)"
            )
            target_column = st.text_input(
                "Target column name (optional)",
                value="",
                key="target_column",
                help="Leave blank to auto-detect fullname/FULLNAME column"
            )
    
    st.divider()
    
    # Column Mapping Configuration
    render_column_mapping_ui(MAPPING_STORAGE)
    
    st.divider()
    
    # Run comparison button
    if st.button("🚀 Run Comparison", type="primary", disabled=not (source_file and target_file)):
        # Convert mappings to dictionaries
        campus_map = {m["source"]: m["target"] for m in st.session_state.campus_mappings if m["source"] and m["target"]}
        program_map = {m["source"]: m["target"] for m in st.session_state.program_mappings if m["source"] and m["target"]}
        college_map = {m["source"]: m["target"] for m in st.session_state.college_mappings if m["source"] and m["target"]}
        
        run_comparison(
            source_file,
            target_file,
            source_sheet,
            target_sheet,
            source_column if source_column else None,
            target_column if target_column else None,
            campus_map or None,
            program_map or None,
            college_map or None,
        )
    
    # Display results if they exist in session state (persists across reruns)
    if st.session_state.comparison_result is not None:
        # Display active mappings if any were used
        if st.session_state.active_mappings:
            mappings = st.session_state.active_mappings
            if mappings.get("campus") or mappings.get("program") or mappings.get("college"):
                st.success("✅ Comparison completed with custom mappings!")
                with st.expander("ℹ️ Active Column Mappings", expanded=True):
                    if mappings.get("campus"):
                        st.markdown("**Campus Mappings:**")
                        for src, tgt in mappings["campus"].items():
                            st.text(f"  {src} → {tgt}")
                    if mappings.get("program"):
                        st.markdown("**Program Mappings:**")
                        for src, tgt in mappings["program"].items():
                            st.text(f"  {src} → {tgt}")
                    if mappings.get("college"):
                        st.markdown("**College Mappings:**")
                        for src, tgt in mappings["college"].items():
                            st.text(f"  {src} → {tgt}")
        
        # Display results from session state
        display_results(
            st.session_state.comparison_result,
            st.session_state.report_bytes,
            st.session_state.source_filename,
            st.session_state.target_filename
        )
    
    if not (source_file and target_file):
        st.info("👆 Upload both source and target files to begin comparison")


if __name__ == "__main__":
    main()
