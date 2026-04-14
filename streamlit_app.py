"""Streamlit UI for Excel/CSV Comparison Automation

Installation:
    pip install streamlit

Run:
    streamlit run streamlit_app.py
"""

from __future__ import annotations

import traceback
from io import BytesIO
from pathlib import Path
import tempfile
import json

import streamlit as st
import pandas as pd

from automate import (
    compare_files,
    parse_sheet_arg,
    write_report_to_bytes,
    CompareResult,
)


st.set_page_config(
    page_title="Excel Comparison Tool for IPDO",
    page_icon="📊",
    layout="wide",
)


def initialize_session_state():
    """Initialize session state variables for column mappings and results."""
    if "campus_mappings" not in st.session_state:
        st.session_state.campus_mappings = []
    if "program_mappings" not in st.session_state:
        st.session_state.program_mappings = []
    if "college_mappings" not in st.session_state:
        st.session_state.college_mappings = []
    
    # Initialize result storage
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


def main():
    initialize_session_state()
    
    st.title("📊 Automation Tool for Graduate Students")
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
            type=["csv", "txt", "xlsx", "xls", "xlsm", "xlsb", "ods"],
            key="source",
            help="The source-of-truth file to compare against"
        )
        
        if source_file:
            st.success(f"✓ {source_file.name}")
            source_sheet = st.text_input(
                "Source sheet name or index",
                value="0",
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
            type=["csv", "txt", "xlsx", "xls", "xlsm", "xlsb", "ods"],
            key="target",
            help="The target file to compare against the source"
        )
        
        if target_file:
            st.success(f"✓ {target_file.name}")
            target_sheet = st.text_input(
                "Target sheet name or index",
                value="0",
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
    render_column_mapping_ui()
    
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


def render_column_mapping_ui():
    """Render the UI for configuring column value mappings."""
    with st.expander("🔧 Advanced: Column Value Mappings", expanded=False):
        st.markdown("""
        Define custom value equivalencies for validation. When a source value maps to a target value, 
        they will be treated as matching during comparison.
        
        **Example:** Map "Science" → "Physics" so rows with "Science" in source match "Physics" in target.
        """)
        
        # Create tabs for different column types
        tab1, tab2, tab3 = st.tabs(["🏫 Campus", "📚 Program", "🎓 College"])
        
        with tab1:
            render_mapping_section("campus", "Campus", "campus_mappings")
        
        with tab2:
            render_mapping_section("program", "Program", "program_mappings")
        
        with tab3:
            render_mapping_section("college", "College", "college_mappings")


def render_mapping_section(field_id: str, field_label: str, session_key: str):
    """Render a mapping section for a specific column type."""
    st.subheader(f"{field_label} Mappings")
    
    # Display existing mappings
    mappings = st.session_state[session_key]
    
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
                    on_change=lambda i=idx, k=session_key: update_mapping(i, k, "source", st.session_state[f"{field_id}_source_{i}"])
                )
            with col2:
                st.markdown("<div style='text-align: center; padding-top: 8px;'>→</div>", unsafe_allow_html=True)
            with col3:
                st.text_input(
                    "Target value",
                    value=mapping.get("target", ""),
                    key=f"{field_id}_target_{idx}",
                    label_visibility="collapsed",
                    on_change=lambda i=idx, k=session_key: update_mapping(i, k, "target", st.session_state[f"{field_id}_target_{i}"])
                )
            with col4:
                if st.button("🗑️", key=f"{field_id}_delete_{idx}", help="Delete mapping"):
                    st.session_state[session_key].pop(idx)
                    st.rerun()
        
        st.divider()
    
    # Add new mapping
    st.markdown("**Add New Mapping:**")
    col1, col2, col3, col4 = st.columns([3, 1, 3, 1])
    
    with col1:
        new_source = st.text_input(
            "Source value",
            key=f"{field_id}_new_source",
            placeholder="e.g., Main Campus",
            label_visibility="collapsed"
        )
    
    with col2:
        st.markdown("<div style='text-align: center; padding-top: 8px;'>→</div>", unsafe_allow_html=True)
    
    with col3:
        new_target = st.text_input(
            "Target value",
            key=f"{field_id}_new_target",
            placeholder="e.g., Pablo Borbon",
            label_visibility="collapsed"
        )
    
    with col4:
        if st.button("➕ Add", key=f"{field_id}_add", type="primary"):
            if new_source and new_target:
                st.session_state[session_key].append({"source": new_source, "target": new_target})
                st.rerun()
            else:
                st.warning("Both source and target values are required")
    
    # Preset templates
    if field_id == "campus":
        if st.button("📋 Load Default Campus Mappings", key=f"{field_id}_load_defaults"):
            st.session_state[session_key] = [
                {"source": "PABLO BORBON", "target": "MAIN"},
                {"source": "JPLPC-MALVAR", "target": "MALVAR"},
                {"source": "ARASOF-NASUGBU", "target": "NASUGBU"},
            ]
            st.rerun()
    elif field_id == "college":
        if st.button("📋 Load Default College Mappings", key=f"{field_id}_load_defaults"):
            st.session_state[session_key] = [
                {"source": "COE", "target": "CEAFA"},
                {"source": "CAFAD", "target": "CEAFA"},
                {"source": "CET", "target": "CIT"},
                {"source": "CHS", "target": "CONAHS"},
                {"source": "CABE", "target": "CABEIHM"},
                {"source": "CCJE", "target": "CAS"},
            ]
            st.rerun()


def update_mapping(idx: int, session_key: str, field: str, value: str):
    """Update a mapping value in session state."""
    if idx < len(st.session_state[session_key]):
        st.session_state[session_key][idx][field] = value


def run_comparison(
    source_file,
    target_file,
    source_sheet_input: str,
    target_sheet_input: str,
    source_column: str | None,
    target_column: str | None,
    custom_campus_map: dict[str, str] | None = None,
    custom_program_map: dict[str, str] | None = None,
    custom_college_map: dict[str, str] | None = None,
):
    """Execute the comparison and display results."""
    
    with st.spinner("🔄 Processing files..."):
        try:
            # Create temporary directory for uploaded files
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_path = Path(temp_dir)
                
                # Save uploaded files
                source_path = temp_path / source_file.name
                target_path = temp_path / target_file.name
                
                source_path.write_bytes(source_file.getvalue())
                target_path.write_bytes(target_file.getvalue())
                
                # Parse sheet arguments
                source_sheet = parse_sheet_arg(source_sheet_input)
                target_sheet = parse_sheet_arg(target_sheet_input)
                
                # Run comparison
                result = compare_files(
                    source_path=source_path,
                    target_path=target_path,
                    source_sheet=source_sheet,
                    target_sheet=target_sheet,
                    source_column=source_column,
                    target_column=target_column,
                    custom_campus_map=custom_campus_map,
                    custom_program_map=custom_program_map,
                    custom_college_map=custom_college_map,
                )
                
                # Generate report
                report_bytes = write_report_to_bytes(
                    result,
                    source_label=source_file.name,
                    target_label=target_file.name,
                )
                
                # Store results in session state for persistence
                st.session_state.comparison_result = result
                st.session_state.report_bytes = report_bytes
                st.session_state.source_filename = source_file.name
                st.session_state.target_filename = target_file.name
                
                # Store active mappings for display
                st.session_state.active_mappings = {
                    "campus": custom_campus_map,
                    "program": custom_program_map,
                    "college": custom_college_map,
                }
                
        except Exception as e:
            st.error(f"❌ **Error during comparison:**\n\n{str(e)}")
            with st.expander("🔍 Show full traceback"):
                st.code(traceback.format_exc())


def display_results(
    result: CompareResult,
    report_bytes: bytes,
    source_name: str,
    target_name: str,
):
    """Display comparison results in the UI."""
    
    st.success("✅ Comparison completed successfully!")
    
    # Summary metrics
    st.header("📈 Summary")
    
    # Create metric columns
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Source Total Rows", f"{result.source_total_rows:,}")
        st.metric("Source Unique Names", f"{result.source_unique_names:,}")
    
    with col2:
        st.metric("Target Total Rows", f"{result.target_total_rows:,}")
        st.metric("Target Unique Names", f"{result.target_unique_names:,}")
    
    with col3:
        found_pct = (result.found_count / result.target_unique_names * 100) if result.target_unique_names > 0 else 0
        st.metric(
            "✓ Found in Source",
            f"{result.found_count:,}",
            delta=f"{found_pct:.1f}%",
            delta_color="normal"
        )
    
    with col4:
        not_found_pct = (result.not_found_count / result.target_unique_names * 100) if result.target_unique_names > 0 else 0
        st.metric(
            "✗ Not Found",
            f"{result.not_found_count:,}",
            delta=f"{not_found_pct:.1f}%",
            delta_color="inverse"
        )
    
    # Duplicates row
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Source Duplicate Names", f"{result.source_duplicate_names:,}")
    
    with col2:
        st.metric("Target Duplicate Names", f"{result.target_duplicate_names:,}")
    
    st.divider()
    
    # Data preview expanders
    st.header("📋 Data Preview")
    
    # Data Validation preview
    if not result.data_validation_rows.empty:
        with st.expander(f"🔍 Data Validation ({len(result.data_validation_rows)} rows)", expanded=False):
            st.dataframe(
                result.data_validation_rows,
                use_container_width=True,
                height=400,
            )
            
            # Show validation summary
            if "overall_validation" in result.data_validation_rows.columns:
                validation_counts = result.data_validation_rows["overall_validation"].value_counts()
                st.markdown("**Validation Status:**")
                for status, count in validation_counts.items():
                    emoji = "✅" if status == "MATCH" else "⚠️" if status == "MISMATCH" else "❓"
                    st.write(f"{emoji} {status}: {count}")
    
    # Source duplicates preview
    if not result.source_duplicate_rows.empty:
        with st.expander(f"📑 Source Duplicates ({len(result.source_duplicate_rows)} rows)", expanded=False):
            st.dataframe(
                result.source_duplicate_rows,
                use_container_width=True,
                height=300,
            )
    else:
        st.info("✓ No duplicates found in source file")
    
    # Target duplicates preview
    if not result.target_duplicate_rows.empty:
        with st.expander(f"📑 Target Duplicates ({len(result.target_duplicate_rows)} rows)", expanded=False):
            st.dataframe(
                result.target_duplicate_rows,
                use_container_width=True,
                height=300,
            )
    else:
        st.info("✓ No duplicates found in target file")
    
    # Found/Not Found names with filtering
    st.subheader("📊 Name Comparison Details")
    
    # Create tabs for different views - matching Excel sheets
    tab1, tab2 = st.tabs(["🎯 Target → Source View", "🔄 Source → Target View"])
    
    with tab1:
        st.markdown("**Target-centric view**: Shows which target names were found/not found in source")
        st.caption("*Corresponds to 'Target_to_Source' sheet in Excel report*")
        
        # Use the target_centered_rows dataframe (like Excel)
        if not result.target_centered_rows.empty:
            with st.expander(f"📋 Target to Source Comparison ({len(result.target_centered_rows)} names)", expanded=False):
                st.markdown("*Shows all target names with their match status in source*")
                
                # Add filter options
                col_filter1, col_filter2 = st.columns(2)
                
                with col_filter1:
                    status_filter_target = st.selectbox(
                        "Filter by status",
                        ["All", "Found in Source", "Not Found in Source"],
                        key="status_filter_target_source"
                    )
                
                with col_filter2:
                    search_target = st.text_input(
                        "🔍 Search names",
                        key="search_target_source",
                        placeholder="Type to filter..."
                    )
                
                # Apply filters
                filtered_df = result.target_centered_rows.copy()
                
                if status_filter_target == "Found in Source":
                    filtered_df = filtered_df[filtered_df['target_to_source_status'] == 'FOUND']
                elif status_filter_target == "Not Found in Source":
                    filtered_df = filtered_df[filtered_df['target_to_source_status'] == 'NOT_FOUND']
                
                if search_target:
                    filtered_df = filtered_df[
                        filtered_df['target_fullname'].str.contains(search_target, case=False, na=False)
                    ]
                
                if not filtered_df.empty:
                    st.dataframe(
                        filtered_df,
                        use_container_width=True,
                        height=400,
                        hide_index=True
                    )
                    st.caption(f"Showing {len(filtered_df)} of {len(result.target_centered_rows)} names")
                else:
                    st.info("No names match the selected filters")
        else:
            st.info("No target data available")
    
    with tab2:
        st.markdown("**Source-centric view**: Shows source names and their presence in target")
        st.caption("*Corresponds to 'Source_to_Target' sheet in Excel report*")
        
        # Use the source_centered_rows dataframe (like Excel)
        if not result.source_centered_rows.empty:
            with st.expander(f"📋 Source to Target Comparison ({len(result.source_centered_rows)} names)", expanded=False):
                st.markdown("*Shows all source names with their match status in target*")
                
                # Add filter options
                col_filter1, col_filter2 = st.columns(2)
                
                with col_filter1:
                    status_filter_source = st.selectbox(
                        "Filter by status",
                        ["All", "Found in Target", "Not Found in Target"],
                        key="status_filter_source_target"
                    )
                
                with col_filter2:
                    search_source = st.text_input(
                        "🔍 Search names",
                        key="search_source_target",
                        placeholder="Type to filter..."
                    )
                
                # Apply filters
                filtered_df = result.source_centered_rows.copy()
                
                if status_filter_source == "Found in Target":
                    filtered_df = filtered_df[filtered_df['source_to_target_status'] == 'FOUND']
                elif status_filter_source == "Not Found in Target":
                    filtered_df = filtered_df[filtered_df['source_to_target_status'] == 'NOT_FOUND']
                
                if search_source:
                    filtered_df = filtered_df[
                        filtered_df['source_fullname'].str.contains(search_source, case=False, na=False)
                    ]
                
                if not filtered_df.empty:
                    st.dataframe(
                        filtered_df,
                        use_container_width=True,
                        height=400,
                        hide_index=True
                    )
                    st.caption(f"Showing {len(filtered_df)} of {len(result.source_centered_rows)} names")
                else:
                    st.info("No names match the selected filters")
        else:
            st.info("No source data available")
    
    st.divider()
    
    # Download button
    st.header("💾 Download Report")
    
    from datetime import datetime
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"comparison_report_{timestamp}.xlsx"
    
    st.download_button(
        label="📥 Download Excel Report",
        data=report_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    
    st.markdown("""
    **Report includes:**
    - 📊 Summary sheet with key metrics
    - 🎯 Target-to-Source comparison
    - 🔄 Source-to-Target comparison
    - 🔍 Data Validation (campus, program, college)
    - 📑 Duplicates analysis
    """)


if __name__ == "__main__":
    main()
