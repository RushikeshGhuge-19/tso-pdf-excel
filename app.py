"""
TSO Converter — Streamlit Web App
==================================
Upload a TSO PDF or Excel + the M&M TSO Download template.
Downloads the populated output Excel instantly.

Run: streamlit run app.py
"""

import streamlit as st
import tempfile, shutil, io
from pathlib import Path

from tso_converter_v3 import parse_input, write_excel

# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="TSO Converter",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── Minimal custom styling ────────────────────────────────────────────────────
st.markdown("""
<style>
    .block-container { padding-top: 2rem; max-width: 760px; }
    .stAlert { border-radius: 8px; }
    div[data-testid="stFileUploader"] { border-radius: 8px; }
    .flag-box {
        background: #FAEEDA; border-left: 3px solid #BA7517;
        padding: 10px 14px; border-radius: 4px;
        font-size: 13px; font-family: monospace; color: #412402;
        margin-bottom: 6px; white-space: pre-wrap; word-break: break-word;
    }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.title("TSO Converter")
st.caption("Upload a TSO source file (PDF or Excel) + the M&M TSO Download template → get a populated Excel ready for upload.")

st.divider()

# ── File uploaders ─────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. TSO source file")
    input_file = st.file_uploader(
        "PDF or Excel",
        type=["pdf", "xlsx"],
        help="The TSO document — either the PDF received from supplier or a previously filled Excel.",
        label_visibility="collapsed",
    )
    if input_file:
        ext = Path(input_file.name).suffix.upper()
        st.success(f"{ext} uploaded — **{input_file.name}**")

with col2:
    st.subheader("2. TSO template Excel")
    template_file = st.file_uploader(
        "M&M TSO Download template (.xlsx)",
        type=["xlsx"],
        help="The blank M&M TSO Download template Excel — must contain the Library sheet with all dropdowns.",
        label_visibility="collapsed",
    )
    if template_file:
        st.success(f"XLSX uploaded — **{template_file.name}**")

st.divider()

# ── Convert button ─────────────────────────────────────────────────────────────
if not input_file or not template_file:
    st.info("Upload both files above to enable conversion.", icon="ℹ️")
    st.stop()

if st.button("Convert to Excel", type="primary", use_container_width=True):
    with st.spinner("Reading input and writing Excel…"):
        try:
            with tempfile.TemporaryDirectory() as tmp:
                tmp = Path(tmp)

                # Write uploaded bytes to temp files
                input_path    = tmp / input_file.name
                template_path = tmp / template_file.name
                out_name      = Path(input_file.name).stem + "_TSO_output.xlsx"
                out_path      = tmp / out_name

                input_path.write_bytes(input_file.getvalue())
                template_path.write_bytes(template_file.getvalue())

                # Parse + write
                data  = parse_input(input_path)
                flags = write_excel(data, template_path, out_path)

                # Read output into memory before temp dir is deleted
                output_bytes = out_path.read_bytes()

            # ── Results ───────────────────────────────────────────────────────
            st.success("Conversion complete!", icon="✅")

            # Summary metrics
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Source",    data.get('source','').upper())
            m2.metric("Project",   data['meta'].get('project', '—'))
            m3.metric("BOM parts", len(data['bom']))
            m4.metric("Tool ops",  len(data['tool_ops']))

            # Download button
            st.download_button(
                label="⬇ Download output Excel",
                data=output_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )

            # Flags
            if flags:
                st.warning(f"{len(flags)} field(s) need manual review after download:", icon="⚠️")
                for flag in flags:
                    st.markdown(f'<div class="flag-box">{flag.strip()}</div>', unsafe_allow_html=True)
            else:
                st.info("All fields matched from Library dropdowns — no manual review needed.", icon="✅")

        except Exception as e:
            import traceback
            st.error(f"Conversion failed: {e}", icon="❌")
            with st.expander("Error details"):
                st.code(traceback.format_exc())

# ── Footer ────────────────────────────────────────────────────────────────────
st.divider()
st.caption("TSO Converter v3 · Supports PDF and Excel input · All dropdowns sourced from Library sheet at runtime")
