"""
WO Approval Audit Tool — Streamlit Web App
Wraps the v12.5 audit engine for browser-based use.
Upload your WO report, optionally add a Haas RMA file, click Run Audit, download the result.
"""
import streamlit as st
import os, time

# ── Page config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="WO Approval Audit",
    page_icon="🔍",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── Custom CSS ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Main container */
    .block-container { max-width: 720px; padding-top: 2rem; }

    /* Header bar */
    .header-bar {
        background: linear-gradient(135deg, #0D1F3C 0%, #1F3864 100%);
        color: white;
        padding: 1.5rem 2rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
    }
    .header-bar h1 { margin: 0; font-size: 1.6rem; font-weight: 700; color: white; }
    .header-bar p  { margin: 0.3rem 0 0 0; font-size: 0.9rem; opacity: 0.85; color: #ccc; }

    /* Metric cards */
    .metric-row {
        display: flex; gap: 0.75rem; margin: 1rem 0;
    }
    .metric-card {
        flex: 1;
        background: #f8f9fa;
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
        border: 1px solid #e9ecef;
    }
    .metric-card.pass  { border-left: 4px solid #28a745; }
    .metric-card.fail  { border-left: 4px solid #dc3545; }
    .metric-card.warn  { border-left: 4px solid #ffc107; }
    .metric-card.info  { border-left: 4px solid #2E75B6; }
    .metric-val  { font-size: 2rem; font-weight: 700; line-height: 1; }
    .metric-val.pass { color: #28a745; }
    .metric-val.fail { color: #dc3545; }
    .metric-val.warn { color: #e6a100; }
    .metric-val.info { color: #2E75B6; }
    .metric-label { font-size: 0.78rem; color: #6c757d; margin-top: 0.3rem; }

    /* Upload area tweaks */
    [data-testid="stFileUploader"] { margin-bottom: 0.5rem; }

    /* Footer */
    .footer { text-align: center; color: #aaa; font-size: 0.75rem; margin-top: 3rem; padding: 1rem 0; border-top: 1px solid #eee; }
</style>
""", unsafe_allow_html=True)

# ── Header ─────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-bar">
    <h1>WO Approval Audit Tool</h1>
    <p>Upload your Salesforce WO report &mdash; get a colour-coded audit workbook in seconds.</p>
</div>
""", unsafe_allow_html=True)

# ── File uploads ───────────────────────────────────────────────────────────
wo_file = st.file_uploader(
    "**WO Approval Report** (.xlsx)",
    type=["xlsx"],
    help="The Salesforce WO export containing WO_Output, Parts_Output, RO, and related sheets.",
)
haas_file = st.file_uploader(
    "**Haas RMA Status** (.xlsx) — optional",
    type=["xlsx"],
    help="Haas RMA status export. If provided, an extra 'Parts + Haas RMA' sheet merges the data.",
)

# ── Run button ─────────────────────────────────────────────────────────────
run_disabled = wo_file is None
run_clicked  = st.button(
    "Run Audit",
    type="primary",
    use_container_width=True,
    disabled=run_disabled,
)

if run_disabled and not run_clicked:
    st.info("Upload a WO Approval Report to get started.")

# ── Execute audit ──────────────────────────────────────────────────────────
if run_clicked and wo_file is not None:
    with st.spinner("Running audit engine..."):
        start = time.time()
        from audit_engine import run_audit   # import here to keep startup fast
        result = run_audit(wo_file, haas_file)
        elapsed = time.time() - start

    # ── Results ────────────────────────────────────────────────────────
    gs = result["gate_summary"]
    wo_n = result["wo_count"]
    parts_n = result["parts_count"]
    built = result["built"]
    failed = result["failed"]

    # Success / error banner
    if built:
        st.success(f"Audit complete — {len(built)} sheet{'s' if len(built)!=1 else ''} built in {elapsed:.1f}s")
    else:
        st.error("Audit failed — no sheets could be built. Check the log below.")

    # Metric cards
    total_gates = gs["pass"] + gs["fail"] + gs["warn"]
    st.markdown(f"""
    <div class="metric-row">
        <div class="metric-card info">
            <div class="metric-val info">{wo_n}</div>
            <div class="metric-label">Work Orders</div>
        </div>
        <div class="metric-card info">
            <div class="metric-val info">{parts_n}</div>
            <div class="metric-label">Part Lines</div>
        </div>
        <div class="metric-card pass">
            <div class="metric-val pass">{gs['pass']}</div>
            <div class="metric-label">WOs Ready</div>
        </div>
        <div class="metric-card fail">
            <div class="metric-val fail">{gs['fail']}</div>
            <div class="metric-label">WOs Blocked</div>
        </div>
        <div class="metric-card warn">
            <div class="metric-val warn">{gs['warn']}</div>
            <div class="metric-label">WOs Warned</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Sheets built
    if built:
        st.markdown("**Sheets:** " + " · ".join(f"`{s}`" for s in built))

    # Failed sheets
    if failed:
        with st.expander(f"⚠ {len(failed)} sheet(s) failed to build", expanded=True):
            for name, err in failed:
                st.error(f"**{name}:** {err}")

    # Download button
    if result["xlsx_bytes"]:
        base_name = os.path.splitext(wo_file.name)[0]
        out_name  = base_name + "_AUDIT.xlsx"
        st.download_button(
            label=f"Download {out_name}",
            data=result["xlsx_bytes"],
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

    # Processing log
    with st.expander("Processing log"):
        st.code("\n".join(result["log"]), language="text")

# ── Footer ─────────────────────────────────────────────────────────────────
st.markdown("""
<div class="footer">
    WO Approval Audit Tool v12.5 &mdash; Audit engine by K &middot; Streamlit wrapper by Claude
</div>
""", unsafe_allow_html=True)
