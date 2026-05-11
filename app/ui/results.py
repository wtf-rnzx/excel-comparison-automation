"""Results rendering for the Streamlit comparison app."""

from __future__ import annotations

import streamlit as st

from automate import CompareResult


def display_results(
    result: CompareResult,
    report_bytes: bytes,
    source_name: str,
    target_name: str,
) -> None:
    """Display comparison results in the UI."""

    st.success("✅ Comparison completed successfully!")

    st.header("📈 Summary")

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

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("Source Duplicate Names", f"{result.source_duplicate_names:,}")

    with col2:
        st.metric("Target Duplicate Names", f"{result.target_duplicate_names:,}")

    st.divider()

    st.header("📋 Data Preview")

    if not result.data_validation_rows.empty:
        with st.expander(f"🔍 Data Validation ({len(result.data_validation_rows)} rows)", expanded=False):
            st.dataframe(
                result.data_validation_rows,
                use_container_width=True,
                height=400,
            )

            if "overall_validation" in result.data_validation_rows.columns:
                validation_counts = result.data_validation_rows["overall_validation"].value_counts()
                st.markdown("**Validation Status:**")
                for status, count in validation_counts.items():
                    emoji = "✅" if status == "MATCH" else "⚠️" if status == "MISMATCH" else "❓"
                    st.write(f"{emoji} {status}: {count}")

    if not result.source_duplicate_rows.empty:
        with st.expander(f"📑 Source Duplicates ({len(result.source_duplicate_rows)} rows)", expanded=False):
            st.dataframe(
                result.source_duplicate_rows,
                use_container_width=True,
                height=300,
            )
    else:
        st.info("✓ No duplicates found in source file")

    if not result.target_duplicate_rows.empty:
        with st.expander(f"📑 Target Duplicates ({len(result.target_duplicate_rows)} rows)", expanded=False):
            st.dataframe(
                result.target_duplicate_rows,
                use_container_width=True,
                height=300,
            )
    else:
        st.info("✓ No duplicates found in target file")

    st.subheader("📊 Name Comparison Details")

    tab1, tab2 = st.tabs(["🎯 Target → Source View", "🔄 Source → Target View"])

    with tab1:
        st.markdown("**Target-centric view**: Shows which target names were found/not found in source")
        st.caption("*Corresponds to 'Target_to_Source' sheet in Excel report*")

        if not result.target_centered_rows.empty:
            with st.expander(f"📋 Target to Source Comparison ({len(result.target_centered_rows)} names)", expanded=False):
                st.markdown("*Shows all target names with their match status in source*")

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

        if not result.source_centered_rows.empty:
            with st.expander(f"📋 Source to Target Comparison ({len(result.source_centered_rows)} names)", expanded=False):
                st.markdown("*Shows all source names with their match status in target*")

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

    st.markdown(
        """
    **Report includes:**
    - 📊 Summary sheet with key metrics
    - 🎯 Target-to-Source comparison
    - 🔄 Source-to-Target comparison
    - 🔍 Data Validation (campus, program, college)
    - 📑 Duplicates analysis
    """
    )
