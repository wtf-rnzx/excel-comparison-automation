"""Comparison orchestration service for Streamlit UI."""

from __future__ import annotations

import traceback
from pathlib import Path
import tempfile

import streamlit as st

from automate import compare_files, parse_sheet_arg, write_report_to_bytes


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
) -> None:
    """Execute comparison, generate report, and store outputs in session state."""

    with st.spinner("🔄 Processing files..."):
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_path = Path(temp_dir)

                source_path = temp_path / source_file.name
                target_path = temp_path / target_file.name

                source_path.write_bytes(source_file.getvalue())
                target_path.write_bytes(target_file.getvalue())

                source_sheet = parse_sheet_arg(source_sheet_input)
                target_sheet = parse_sheet_arg(target_sheet_input)

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

                report_bytes = write_report_to_bytes(
                    result,
                    source_label=source_file.name,
                    target_label=target_file.name,
                )

                st.session_state.comparison_result = result
                st.session_state.report_bytes = report_bytes
                st.session_state.source_filename = source_file.name
                st.session_state.target_filename = target_file.name
                st.session_state.active_mappings = {
                    "campus": custom_campus_map,
                    "program": custom_program_map,
                    "college": custom_college_map,
                }

        except Exception as exc:
            st.error(f"❌ **Error during comparison:**\n\n{str(exc)}")
            with st.expander("🔍 Show full traceback"):
                st.code(traceback.format_exc())
