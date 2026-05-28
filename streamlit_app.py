import streamlit as st
import tempfile
import shutil
from pathlib import Path
import sys
import os

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="BrightView Audit Report Generator",
    page_icon="📊",
    layout="centered"
)

# ── Header ────────────────────────────────────────────────────────────────────
st.title("📊 Branch Audit Report Generator")
st.markdown("Upload a completed audit Excel workbook to generate a PowerPoint report.")
st.divider()

# ── File upload ───────────────────────────────────────────────────────────────
excel_file = st.file_uploader(
    "Upload Audit Excel (.xlsx)",
    type=["xlsx"],
    help="The completed branch audit workbook with FINISH tab"
)

st.divider()

# ── Inputs ────────────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    period = st.text_input(
        "Period under Review",
        placeholder="e.g. March 2026",
        help="The month and year the audit covers"
    )
    branch = st.text_input(
        "Branch Name Override (optional)",
        placeholder="e.g. 35210 BVLS Homestead",
        help="Leave blank to auto-detect from filename"
    )

with col2:
    month = st.text_input(
        "Month of Audit Report",
        placeholder="e.g. May 2026",
        help="The month the audit was conducted"
    )
    approach = st.selectbox(
        "Audit Approach",
        options=["On-Site", "Remote"],
        index=0
    )

auditors = st.text_input(
    "Auditor Name(s)",
    placeholder="e.g. Jane Smith",
    value="[Auditor Name]"
)

st.divider()

# ── Generate button ───────────────────────────────────────────────────────────
generate = st.button("Generate Report", type="primary", use_container_width=True)

if generate:
    # Validate inputs
    errors = []
    if not excel_file:
        errors.append("Please upload an audit Excel file.")
    if not period:
        errors.append("Please enter the Period under Review.")
    if not month:
        errors.append("Please enter the Month of Audit Report.")

    if errors:
        for e in errors:
            st.error(e)
    else:
        with st.spinner("Reading Excel and generating report..."):
            try:
                # Write uploaded Excel to temp file
                tmp_dir = Path(tempfile.mkdtemp())
                excel_path = tmp_dir / excel_file.name
                excel_path.write_bytes(excel_file.read())

                # Template path - must be in same folder as this script
                script_dir = Path(__file__).parent
                template_path = script_dir / "template.pptx"

                if not template_path.exists():
                    st.error(
                        "template.pptx not found. Please upload it to the workspace "
                        "alongside this script."
                    )
                    st.stop()

                # Import the generator
                sys.path.insert(0, str(script_dir))
                from generate_audit_report_pure import (
                    get_branch_info, read_finish_tab, scan_for_exceptions,
                    TEMPLATES, build_pptx
                )
                import openpyxl

                # Load workbook
                wb = openpyxl.load_workbook(excel_path, data_only=True)

                # Branch info
                branch_num, branch_display = get_branch_info(wb, excel_path)
                if branch.strip():
                    branch_display = branch.strip()

                # Scores
                scores, overall = read_finish_tab(wb)

                # Exceptions
                exceptions = scan_for_exceptions(wb)

                # Draft observations
                for test_name, exc_data in exceptions.items():
                    tmpl  = TEMPLATES.get(test_name, {})
                    obs_fn = tmpl.get("obs_fn")
                    if obs_fn:
                        exc_data["observation"] = obs_fn(exc_data["data"], period)

                # Build PPTX
                output_path = tmp_dir / f"{branch_display.replace(' ', '_')}_Audit_Report.pptx"
                build_pptx(
                    template_path   = template_path,
                    output_path     = output_path,
                    branch_display  = branch_display,
                    scores          = scores,
                    overall_rating  = overall,
                    findings        = exceptions,
                    period_under_review = period,
                    audit_month     = month,
                    auditor_names   = auditors,
                    audit_approach  = approach,
                )

                # Read output and offer download
                pptx_bytes = output_path.read_bytes()
                filename   = output_path.name

                st.success(f"✅ Report generated — {len(exceptions)} exception(s) found across {len(scores)} risk areas.")

                # Summary table
                if scores:
                    st.markdown("**Scores by Risk Area:**")
                    cols = st.columns(4)
                    for i, (area, data) in enumerate(scores.items()):
                        with cols[i % 4]:
                            rating = data['rating']
                            color  = {"Exceeds Expectations": "🔵",
                                      "Satisfactory": "🟢",
                                      "Needs Improvement": "🟡",
                                      "Unsatisfactory": "🔴"}.get(rating, "⚪")
                            st.metric(area, data['score_pct'], delta=f"{color} {rating}",
                                      delta_color="off")

                st.divider()
                st.download_button(
                    label     = "⬇️ Download PowerPoint",
                    data      = pptx_bytes,
                    file_name = filename,
                    mime      = "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width = True,
                    type      = "primary"
                )

            except Exception as e:
                st.error(f"Something went wrong: {e}")
                st.exception(e)
            finally:
                shutil.rmtree(tmp_dir, ignore_errors=True)
