"""
HR Attendance Processor — Streamlit Web App
Upload an Excel attendance file → process WO allocation → download results.
"""

import io
import random
import sys
from pathlib import Path

import streamlit as st
import pandas as pd

# ---------------------------------------------------------------------------
# Import the core processing logic from the parent directory
# ---------------------------------------------------------------------------
sys.path.insert(0, str(Path(__file__).resolve().parent))
from process_wo_columns import (
    _read_excel,
    build_wo_columns,
    add_new_timing_columns,
    validate_wo_output,
)

# ---------------------------------------------------------------------------
# Page configuration
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="HR Attendance Processor",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ---------------------------------------------------------------------------
# Custom CSS for a premium, modern look
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
    /* ---------- Google Fonts ---------- */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    html, body, p, h1, h2, h3, h4, h5, h6, span, div, label, input, button, select, textarea {
        font-family: 'Inter', sans-serif;
    }
    /* Preserve Streamlit's icon font */
    [data-testid="stIconMaterial"] {
        font-family: 'Material Symbols Rounded' !important;
    }

    /* ---------- Header / Hero ---------- */
    .hero {
        background: linear-gradient(135deg, #1e3a5f 0%, #2d6a9f 50%, #4a90d9 100%);
        padding: 2.5rem 2rem;
        border-radius: 16px;
        margin-bottom: 2rem;
        color: white;
        text-align: center;
    }
    .hero h1 {
        font-size: 2.2rem;
        font-weight: 700;
        margin-bottom: 0.3rem;
    }
    .hero p {
        font-size: 1.05rem;
        opacity: 0.85;
        margin: 0;
    }

    /* ---------- Step cards ---------- */
    .step-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 1rem;
        transition: box-shadow 0.2s ease;
    }
    .step-card:hover {
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.06);
    }
    .step-label {
        display: inline-block;
        background: linear-gradient(135deg, #1e3a5f, #2d6a9f);
        color: white;
        padding: 3px 12px;
        border-radius: 20px;
        font-size: 0.75rem;
        font-weight: 600;
        letter-spacing: 0.03em;
        margin-bottom: 0.5rem;
    }

    /* ---------- Stats row ---------- */
    .stat-box {
        background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%);
        border: 1px solid #bae6fd;
        border-radius: 12px;
        padding: 1rem 1.2rem;
        text-align: center;
    }
    .stat-box .value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1e3a5f;
    }
    .stat-box .label {
        font-size: 0.82rem;
        color: #64748b;
        margin-top: 2px;
    }

    /* ---------- Success banner ---------- */
    .success-banner {
        background: linear-gradient(135deg, #ecfdf5 0%, #d1fae5 100%);
        border: 1px solid #6ee7b7;
        border-radius: 12px;
        padding: 1.2rem 1.5rem;
        text-align: center;
        margin: 1rem 0;
    }
    .success-banner h3 {
        color: #065f46;
        margin: 0 0 0.3rem 0;
    }
    .success-banner p {
        color: #047857;
        margin: 0;
        font-size: 0.92rem;
    }

    /* ---------- Footer ---------- */
    .footer {
        text-align: center;
        color: #94a3b8;
        font-size: 0.78rem;
        margin-top: 3rem;
        padding: 1rem 0;
        border-top: 1px solid #e2e8f0;
    }

    /* ---------- Hide Streamlit branding ---------- */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* ---------- Button overrides ---------- */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #1e3a5f 0%, #2d6a9f 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.6rem 2rem !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        transition: transform 0.15s ease, box-shadow 0.15s ease !important;
    }
    .stDownloadButton > button:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 6px 20px rgba(30, 58, 95, 0.3) !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Hero banner
# ---------------------------------------------------------------------------
st.markdown(
    """
    <div class="hero">
        <h1>📋 HR Attendance Processor</h1>
        <p>Upload attendance Excel &nbsp;→&nbsp; Process WO allocation &nbsp;→&nbsp; Download results</p>
    </div>
    """,
    unsafe_allow_html=True,
)


# ═══════════════════════════════════════════════════════════════════════════
#  Helper: read uploaded file into a dict of DataFrames (one per sheet)
# ═══════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes: bytes, file_name: str) -> list[str]:
    """Return sheet names from an uploaded Excel file."""
    buf = io.BytesIO(file_bytes)
    try:
        xls = pd.ExcelFile(buf, engine="calamine")
    except Exception:
        xls = pd.ExcelFile(buf, engine="openpyxl")
    return xls.sheet_names


@st.cache_data(show_spinner="Reading sheet…")
def read_sheet(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    buf = io.BytesIO(file_bytes)
    try:
        return pd.read_excel(buf, sheet_name=sheet_name, engine="calamine")
    except Exception:
        buf.seek(0)
        return pd.read_excel(buf, sheet_name=sheet_name, engine="openpyxl")


def process_dataframe(df: pd.DataFrame, seed: int | None) -> pd.DataFrame:
    """Run the full processing pipeline on a DataFrame."""
    if "Hrs" not in df.columns and "Total Hours Worked" in df.columns:
        df = df.rename(columns={"Total Hours Worked": "Hrs"})
    if "Remarks" not in df.columns:
        df["Remarks"] = ""

    if seed is not None:
        random.seed(seed)

    df = build_wo_columns(df)
    df = add_new_timing_columns(df)
    return df


def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    """Convert a DataFrame to .xlsx bytes for download."""
    buf = io.BytesIO()
    try:
        df.to_excel(buf, index=False, sheet_name="Processed_Data", engine="xlsxwriter")
    except ImportError:
        df.to_excel(buf, index=False, sheet_name="Processed_Data", engine="openpyxl")
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════
#  Step 1 — Upload
# ═══════════════════════════════════════════════════════════════════════════
st.markdown(
    '<div class="step-card"><span class="step-label">STEP 1</span>',
    unsafe_allow_html=True,
)
st.markdown("### 📁 Upload Attendance File")
uploaded = st.file_uploader(
    "Drag and drop your Excel file here",
    type=["xlsx", "xls"],
    label_visibility="collapsed",
)
st.markdown("</div>", unsafe_allow_html=True)

if uploaded is None:
    st.info("👆 Upload an attendance Excel file to get started.")
    st.markdown(
        '<div class="footer">HR Attendance Processor • Private & Secure • No data is stored</div>',
        unsafe_allow_html=True,
    )
    st.stop()

# ═══════════════════════════════════════════════════════════════════════════
#  Step 2 — Select sheet
# ═══════════════════════════════════════════════════════════════════════════
file_bytes = uploaded.getvalue()
sheets = get_sheet_names(file_bytes, uploaded.name)

st.markdown(
    '<div class="step-card"><span class="step-label">STEP 2</span>',
    unsafe_allow_html=True,
)
st.markdown("### 📑 Select Worksheet")

col_sheet, col_seed = st.columns([3, 1])
with col_sheet:
    selected_sheet = st.selectbox(
        "Which sheet contains the attendance data?",
        options=sheets,
        index=0,
    )
with col_seed:
    seed_val = st.number_input(
        "Random seed (optional)",
        value=None,
        min_value=0,
        step=1,
        help="Set a seed for reproducible synthetic timings. Leave blank for random.",
    )

st.markdown("</div>", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
#  Step 3 — Process
# ═══════════════════════════════════════════════════════════════════════════
st.markdown(
    '<div class="step-card"><span class="step-label">STEP 3</span>',
    unsafe_allow_html=True,
)
st.markdown("### ⚙️ Process Attendance")

process_btn = st.button("🚀 Process File", use_container_width=True, type="primary")
st.markdown("</div>", unsafe_allow_html=True)

if not process_btn and "processed_df" not in st.session_state:
    st.markdown(
        '<div class="footer">HR Attendance Processor • Private & Secure • No data is stored</div>',
        unsafe_allow_html=True,
    )
    st.stop()

# ---------------------------------------------------------------------------
#  Run processing
# ---------------------------------------------------------------------------
if process_btn:
    with st.spinner("Reading worksheet…"):
        df_raw = read_sheet(file_bytes, selected_sheet)

    # Validate required columns
    required = {"Empl Id", "Attendance Date", "Duty Status", "Shift", "First In", "Last Out", "Hrs"}
    present_cols = set(df_raw.columns)
    # Also accept alternate name
    if "Total Hours Worked" in present_cols:
        present_cols.add("Hrs")
    missing = required - present_cols
    if missing:
        st.error(f"❌ Missing required columns: **{', '.join(sorted(missing))}**")
        st.info("The sheet must contain: " + ", ".join(sorted(required)))
        st.stop()

    with st.spinner("Processing WO allocation & timings…"):
        processed = process_dataframe(df_raw.copy(), seed=seed_val)

    st.session_state["processed_df"] = processed
    st.session_state["raw_df"] = df_raw
    st.session_state["source_name"] = uploaded.name

# ═══════════════════════════════════════════════════════════════════════════
#  Step 4 — Results
# ═══════════════════════════════════════════════════════════════════════════
if "processed_df" in st.session_state:
    df_result = st.session_state["processed_df"]
    df_raw = st.session_state["raw_df"]
    source = st.session_state["source_name"]

    # Success banner
    st.markdown(
        """
        <div class="success-banner">
            <h3>✅ Processing Complete</h3>
            <p>Your attendance data has been processed successfully.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Stats row
    wo_count = (df_result["New_WO"].astype(str) == "WO").sum()
    emp_count = df_result["Empl Id"].nunique()
    total_rows = len(df_result)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(
            f'<div class="stat-box"><div class="value">{total_rows:,}</div><div class="label">Total Rows</div></div>',
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f'<div class="stat-box"><div class="value">{emp_count}</div><div class="label">Employees</div></div>',
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            f'<div class="stat-box"><div class="value">{wo_count}</div><div class="label">WO Days Assigned</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("")

    # Validation
    errors = validate_wo_output(df_result)
    if errors:
        with st.expander(f"⚠️ Validation warnings ({len(errors)})", expanded=False):
            for e in errors[:50]:
                st.write(f"- {e}")
    else:
        st.success("✔ Validation passed — WO spacing, week_day, and week_number are correct.")

    # Preview
    with st.expander("🔍 Preview processed data", expanded=True):
        st.dataframe(df_result, use_container_width=True, height=400)

    # Download
    st.markdown("---")
    out_name = Path(source).stem + "_processed.xlsx"

    xlsx_data = df_to_xlsx_bytes(df_result)

    st.download_button(
        label=f"⬇️  Download {out_name}",
        data=xlsx_data,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ---------------------------------------------------------------------------
#  Footer
# ---------------------------------------------------------------------------
st.markdown(
    '<div class="footer">HR Attendance Processor • Private & Secure • No data is stored</div>',
    unsafe_allow_html=True,
)
