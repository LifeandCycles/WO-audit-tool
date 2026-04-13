"""
WO Audit Console — Streamlit Web App V1.27
Cyberpunk command console interface.
Tab 1: WO Approval Audit (engine v12.7)
Tab 2: Quick WO Gates — pass/fail gates without CA grading
Tab 3: Orphan Work Order Analyzer (engine v1.0)
"""
import streamlit as st
import streamlit.components.v1 as components
import os, time, urllib.parse
from audit_engine import run_audit
from orphan_engine import run_orphan_analysis

# ── Page config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="WO Audit Console",
    page_icon="⚡",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── Matrix G-code background ──────────────────────────────────────────────
MATRIX_HTML = """
<html>
<head><style>html,body{margin:0;padding:0;overflow:hidden;background:transparent;}</style></head>
<body>
<canvas id="c"></canvas>
<script>
const canvas = document.getElementById('c');
const ctx = canvas.getContext('2d');
canvas.width = window.innerWidth;
canvas.height = window.innerHeight;
window.addEventListener('resize', () => {
    canvas.width = window.innerWidth;
    canvas.height = window.innerHeight;
});

const gcodes = [
    'G00','G01','G02','G03','G17','G18','G19','G20','G21','G28',
    'G40','G41','G42','G43','G49','G50','G52','G53','G54','G55',
    'G56','G57','G58','G59','G73','G76','G80','G81','G82','G83',
    'G84','G85','G86','G90','G91','G92','G93','G94','G95','G96',
    'G97','G98','G99','M00','M01','M03','M04','M05','M06','M08',
    'M09','M19','M30','M97','M98','M99','T01','T02','T03','T04',
    'SPINDLE','COOLANT','TOOL CHG','FEEDRATE','RAPID','HOME',
    'ATC','PROBE','OFFSET','DATUM','CLAMP','UNCLAMP'
];

const fs = 14;
ctx.font = fs + 'px monospace';
const colW = fs * 5;
const cols = Math.floor(canvas.width / colW);
const TRAIL = 16;
const rowH = fs + 4;

let columns = [];
for (let i = 0; i < cols; i++) {
    columns.push({
        y: Math.random() * canvas.height * -1,
        speed: 2 + Math.random() * 4,
        glyph: gcodes[Math.floor(Math.random() * gcodes.length)],
        frame: 0,
        trail: []
    });
}

function draw() {
    ctx.fillStyle = 'rgba(0, 0, 0, 0.12)';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    for (let i = 0; i < columns.length; i++) {
        const col = columns[i];
        col.frame++;

        // Cycle glyph as it falls
        if (col.frame % 3 === 0) {
            col.trail.unshift(col.glyph);
            if (col.trail.length > TRAIL) col.trail.pop();
            col.glyph = gcodes[Math.floor(Math.random() * gcodes.length)];
        }

        const x = i * colW;

        // Draw white-hot leader
        ctx.fillStyle = '#ffffff';
        ctx.font = 'bold ' + fs + 'px monospace';
        ctx.fillText(col.glyph, x, col.y);

        // Draw bright green second char
        if (col.trail.length > 0) {
            ctx.fillStyle = '#00ff9c';
            ctx.font = fs + 'px monospace';
            ctx.fillText(col.trail[0], x, col.y - rowH);
        }

        // Draw fading green trail
        ctx.font = fs + 'px monospace';
        for (let t = 1; t < col.trail.length; t++) {
            const alpha = (1 - t / TRAIL) * 0.7;
            ctx.fillStyle = 'rgba(0, 255, 156, ' + alpha + ')';
            ctx.fillText(col.trail[t], x, col.y - (t + 1) * rowH);
        }

        col.y += col.speed;
        if (col.y > canvas.height + TRAIL * rowH) {
            col.y = -(Math.random() * canvas.height * 0.5);
            col.speed = 2 + Math.random() * 4;
            col.trail = [];
        }
    }
    requestAnimationFrame(draw);
}
draw();
</script>
</body>
</html>
"""
components.html(MATRIX_HTML, height=300)

# ── Cyberpunk Console CSS ──────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@300;400;500;700&family=Rajdhani:wght@400;500;600;700&display=swap');

    /* Global */
    .stApp { background-color: #0a0e17; }
    .block-container { max-width: 780px; padding-top: 1.5rem; }

    /* Matrix G-code background — pin the iframe behind everything */
    div:has(> iframe) {
        position: fixed !important;
        top: 0; left: 0; right: 0; bottom: 0;
        width: 100vw !important;
        height: 100vh !important;
        z-index: 0;
        opacity: 0.12;
        pointer-events: none;
    }
    div:has(> iframe) iframe {
        width: 100vw !important;
        height: 100vh !important;
    }
    /* Ensure all real content sits above the background */
    .block-container { position: relative; z-index: 1; }

    /* Kill default Streamlit header/footer chrome */
    header[data-testid="stHeader"] { background: transparent; }
    .stDeployButton { display: none; }

    /* ── Header bar ── */
    .console-header {
        background: linear-gradient(135deg, #0d1117 0%, #161b22 50%, #0d1117 100%);
        border: 1px solid #00ff9c33;
        border-left: 3px solid #00ff9c;
        padding: 1.5rem 2rem;
        border-radius: 4px;
        margin-bottom: 1.5rem;
        position: relative;
        overflow: hidden;
    }
    .console-header::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 1px;
        background: linear-gradient(90deg, transparent, #00ff9c66, transparent);
    }
    .console-header h1 {
        margin: 0;
        font-family: 'Rajdhani', sans-serif;
        font-size: 1.7rem;
        font-weight: 700;
        color: #00ff9c;
        letter-spacing: 2px;
        text-transform: uppercase;
    }
    .console-header .subtitle {
        margin: 0.3rem 0 0 0;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.75rem;
        color: #8b949e;
        letter-spacing: 0.5px;
    }
    .console-header .version {
        position: absolute;
        top: 1rem; right: 1.5rem;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.7rem;
        color: #00ff9c88;
        border: 1px solid #00ff9c33;
        padding: 2px 8px;
        border-radius: 3px;
    }

    /* ── Status indicator ── */
    .status-line {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.8rem;
        color: #8b949e;
        padding: 0.5rem 0;
        border-bottom: 1px solid #21262d;
        margin-bottom: 1rem;
    }
    .status-line .dot {
        display: inline-block;
        width: 8px; height: 8px;
        border-radius: 50%;
        background: #00ff9c;
        margin-right: 8px;
        box-shadow: 0 0 6px #00ff9c66;
        animation: pulse 2s ease-in-out infinite;
    }
    @keyframes pulse {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.4; }
    }

    /* ── Metric cards ── */
    .metrics-grid {
        display: grid;
        grid-template-columns: repeat(6, 1fr);
        gap: 0.5rem;
        margin: 1rem 0;
    }
    .metrics-grid-3 {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 0.5rem;
        margin: 1rem 0;
    }
    .m-card {
        background: #0d1117;
        border: 1px solid #21262d;
        border-radius: 4px;
        padding: 0.8rem 0.5rem;
        text-align: center;
        position: relative;
        overflow: hidden;
    }
    .m-card::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 2px;
    }
    .m-card.ready::before  { background: #00ff9c; }
    .m-card.blocked::before { background: #ff3e3e; }
    .m-card.warned::before  { background: #ffb800; }
    .m-card.held::before    { background: #a371f7; }
    .m-card.info::before    { background: #58a6ff; }
    .m-val {
        font-family: 'Rajdhani', sans-serif;
        font-size: 2rem;
        font-weight: 700;
        line-height: 1;
    }
    .m-val.ready  { color: #00ff9c; }
    .m-val.blocked { color: #ff3e3e; }
    .m-val.warned  { color: #ffb800; }
    .m-val.held    { color: #a371f7; }
    .m-val.info    { color: #58a6ff; }
    .m-label {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.65rem;
        color: #8b949e;
        margin-top: 0.3rem;
        text-transform: uppercase;
        letter-spacing: 1px;
    }

    /* ── Gate breakdown table ── */
    .gate-table {
        width: 100%;
        border-collapse: collapse;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.78rem;
        margin: 1rem 0;
    }
    .gate-table th {
        text-align: left;
        color: #8b949e;
        font-weight: 500;
        padding: 0.4rem 0.6rem;
        border-bottom: 1px solid #21262d;
        font-size: 0.7rem;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .gate-table td {
        padding: 0.5rem 0.6rem;
        border-bottom: 1px solid #161b2211;
        color: #c9d1d9;
    }
    .gate-table tr:hover td { background: #161b22; }
    .gate-table .gate-name { color: #c9d1d9; }
    .g-pass { color: #00ff9c; font-weight: 600; }
    .g-fail { color: #ff3e3e; font-weight: 600; }
    .g-warn { color: #ffb800; font-weight: 600; }
    .g-zero { color: #30363d; }

    /* ── Section headers ── */
    .section-hdr {
        font-family: 'Rajdhani', sans-serif;
        font-size: 1rem;
        font-weight: 600;
        color: #58a6ff;
        text-transform: uppercase;
        letter-spacing: 2px;
        padding: 0.8rem 0 0.3rem 0;
        border-bottom: 1px solid #21262d;
        margin: 1.5rem 0 0.8rem 0;
    }

    /* ── Upload area ── */
    [data-testid="stFileUploader"] label p {
        font-family: 'JetBrains Mono', monospace !important;
        color: #c9d1d9 !important;
        font-size: 0.85rem !important;
    }
    [data-testid="stFileUploader"] {
        border: 1px dashed #21262d;
        border-radius: 4px;
        padding: 0.3rem;
        margin-bottom: 0.5rem;
    }

    /* ── Buttons ── */
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #00ff9c22, #00ff9c11) !important;
        border: 1px solid #00ff9c55 !important;
        color: #00ff9c !important;
        font-family: 'Rajdhani', sans-serif !important;
        font-weight: 600 !important;
        letter-spacing: 2px !important;
        text-transform: uppercase !important;
        transition: all 0.3s ease !important;
    }
    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(135deg, #00ff9c33, #00ff9c22) !important;
        border-color: #00ff9c !important;
        box-shadow: 0 0 15px #00ff9c22 !important;
    }
    .stDownloadButton > button {
        background: linear-gradient(135deg, #58a6ff22, #58a6ff11) !important;
        border: 1px solid #58a6ff55 !important;
        color: #58a6ff !important;
        font-family: 'Rajdhani', sans-serif !important;
        font-weight: 600 !important;
        letter-spacing: 2px !important;
        text-transform: uppercase !important;
    }
    .stDownloadButton > button:hover {
        background: linear-gradient(135deg, #58a6ff33, #58a6ff22) !important;
        border-color: #58a6ff !important;
        box-shadow: 0 0 15px #58a6ff22 !important;
    }

    /* ── Expander ── */
    .streamlit-expanderHeader {
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 0.8rem !important;
        color: #8b949e !important;
    }

    /* ── Log output ── */
    .log-output {
        background: #0d1117;
        border: 1px solid #21262d;
        border-radius: 4px;
        padding: 1rem;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.75rem;
        color: #8b949e;
        line-height: 1.6;
        max-height: 300px;
        overflow-y: auto;
    }
    .log-output .log-ok { color: #00ff9c; }
    .log-output .log-err { color: #ff3e3e; }
    .log-output .log-info { color: #58a6ff; }

    /* ── Feedback section ── */
    .feedback-box {
        background: #0d1117;
        border: 1px solid #ffb80033;
        border-left: 3px solid #ffb800;
        border-radius: 4px;
        padding: 1.2rem 1.5rem;
        margin: 1rem 0;
    }
    .feedback-box h3 {
        font-family: 'Rajdhani', sans-serif;
        color: #ffb800;
        font-size: 1rem;
        margin: 0 0 0.5rem 0;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .feedback-box p {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.75rem;
        color: #8b949e;
        margin: 0 0 0.8rem 0;
        line-height: 1.5;
    }
    .feedback-btn {
        display: inline-block;
        background: linear-gradient(135deg, #ffb80022, #ffb80011);
        border: 1px solid #ffb80055;
        color: #ffb800;
        font-family: 'Rajdhani', sans-serif;
        font-weight: 600;
        font-size: 0.85rem;
        letter-spacing: 1.5px;
        text-transform: uppercase;
        padding: 0.5rem 1.2rem;
        border-radius: 4px;
        text-decoration: none;
        transition: all 0.3s ease;
    }
    .feedback-btn:hover {
        background: linear-gradient(135deg, #ffb80033, #ffb80022);
        border-color: #ffb800;
        box-shadow: 0 0 15px #ffb80022;
        color: #ffb800;
    }

    /* ── Success/error banners ── */
    [data-testid="stAlert"] {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.82rem;
        border-radius: 4px;
    }

    /* ── Tab styling ── */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0;
        border-bottom: 1px solid #21262d;
    }
    .stTabs [data-baseweb="tab"] {
        font-family: 'Rajdhani', sans-serif !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
        letter-spacing: 1.5px !important;
        text-transform: uppercase !important;
        color: #8b949e !important;
        padding: 0.6rem 1.5rem !important;
        border: none !important;
        background: transparent !important;
    }
    .stTabs [aria-selected="true"] {
        color: #00ff9c !important;
        border-bottom: 2px solid #00ff9c !important;
    }

    /* ── Orphan reason table ── */
    .reason-table {
        width: 100%;
        border-collapse: collapse;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.78rem;
        margin: 1rem 0;
    }
    .reason-table th {
        text-align: left;
        color: #8b949e;
        font-weight: 500;
        padding: 0.4rem 0.6rem;
        border-bottom: 1px solid #21262d;
        font-size: 0.7rem;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .reason-table td {
        padding: 0.5rem 0.6rem;
        border-bottom: 1px solid #161b2211;
        color: #c9d1d9;
    }
    .reason-table tr:hover td { background: #161b22; }

    /* ── CA Reminder banner ── */
    .ca-reminder {
        background: #0d1117;
        border: 1px solid #58a6ff33;
        border-left: 3px solid #58a6ff;
        border-radius: 4px;
        padding: 1.2rem 1.5rem;
        margin: 1.2rem 0;
    }
    .ca-reminder h3 {
        font-family: 'Rajdhani', sans-serif;
        color: #58a6ff;
        font-size: 1rem;
        margin: 0 0 0.5rem 0;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .ca-reminder p {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.78rem;
        color: #8b949e;
        margin: 0;
        line-height: 1.6;
    }

    /* ── Footer ── */
    .console-footer {
        text-align: center;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.7rem;
        color: #30363d;
        margin-top: 3rem;
        padding: 1rem 0;
        border-top: 1px solid #161b22;
    }
    .console-footer span { color: #00ff9c44; }
</style>
""", unsafe_allow_html=True)

# ── Header ─────────────────────────────────────────────────────────────────
st.markdown("""
<div class="console-header">
    <div class="version">V1.27</div>
    <h1>WO Audit Console</h1>
    <p class="subtitle">// work order approval audit &amp; quick gates &amp; orphan analyzer</p>
</div>
""", unsafe_allow_html=True)

# ── Status line ────────────────────────────────────────────────────────────
st.markdown("""
<div class="status-line">
    <span class="dot"></span>SYSTEM ONLINE &mdash; audit engine v12.7 &middot; quick gates &middot; orphan engine v1.0
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════
tab_audit, tab_quick, tab_orphan = st.tabs(["WO Approval Audit", "Quick WO Gates", "Orphan Analyzer"])

# ══════════════════════════════════════════════════════════════════════════
# TAB 1 — WO APPROVAL AUDIT
# ══════════════════════════════════════════════════════════════════════════
with tab_audit:
    wo_file = st.file_uploader(
        "**WO Approval Report** (.xlsx)",
        type=["xlsx"],
        help="Salesforce WO export with WO_Output, Parts_Output, RO, and related sheets.",
        key="audit_wo_file",
    )
    haas_file = st.file_uploader(
        "**Haas RMA Status** (.xlsx) — optional",
        type=["xlsx"],
        help="Haas RMA status export for Parts + Haas RMA merge sheet.",
        key="audit_haas_file",
    )

    run_disabled = wo_file is None
    run_clicked  = st.button(
        "Execute Audit",
        type="primary",
        use_container_width=True,
        disabled=run_disabled,
        key="audit_run",
    )

    if run_disabled and not run_clicked:
        st.markdown("""
        <div style="font-family: 'JetBrains Mono', monospace; font-size: 0.8rem; color: #30363d;
                    text-align: center; padding: 2rem 0;">
            &gt; awaiting file upload_
        </div>
        """, unsafe_allow_html=True)

    if run_clicked and wo_file is not None:
        with st.spinner("Executing audit sequence..."):
            start = time.time()
            result = run_audit(wo_file, haas_file)
            elapsed = time.time() - start

        # ── Validation failure ─────────────────────────────────────────
        if not result.get("validation_ok", True):
            st.error("VALIDATION FAILED — file is not a valid WO Approval Report")
            for e in result.get("validation_errors", []):
                st.markdown(f"<div style='font-family: JetBrains Mono, monospace; font-size: 0.8rem; color: #ff3e3e; padding: 0.2rem 0;'>  &gt; {e}</div>", unsafe_allow_html=True)
        else:
            # ── Results ────────────────────────────────────────────────
            gs = result["gate_summary"]
            gd = result.get("gate_details", {})
            wo_n = result["wo_count"]
            parts_n = result["parts_count"]
            built = result["built"]
            failed = result["failed"]

            if built:
                st.success(f"AUDIT COMPLETE — {len(built)} sheets built in {elapsed:.1f}s")
            else:
                st.error("AUDIT FAILED — no sheets could be built")

            # ── Metric cards ───────────────────────────────────────────
            held_n = gs.get('hold', 0)
            st.markdown(f"""
            <div class="metrics-grid">
                <div class="m-card info">
                    <div class="m-val info">{wo_n}</div>
                    <div class="m-label">Work Orders</div>
                </div>
                <div class="m-card info">
                    <div class="m-val info">{parts_n}</div>
                    <div class="m-label">Part Lines</div>
                </div>
                <div class="m-card ready">
                    <div class="m-val ready">{gs['pass']}</div>
                    <div class="m-label">Ready</div>
                </div>
                <div class="m-card blocked">
                    <div class="m-val blocked">{gs['fail']}</div>
                    <div class="m-label">Blocked</div>
                </div>
                <div class="m-card warned">
                    <div class="m-val warned">{gs['warn']}</div>
                    <div class="m-label">Warned</div>
                </div>
                <div class="m-card held">
                    <div class="m-val held">{held_n}</div>
                    <div class="m-label">Held</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # ── Gate-level breakdown ───────────────────────────────────
            if gd:
                st.markdown('<div class="section-hdr">Gate Breakdown</div>', unsafe_allow_html=True)
                rows_html = ""
                for gname, counts in gd.items():
                    p = counts.get("Pass", 0)
                    f_count = counts.get("Fail", 0)
                    w = counts.get("Warn", 0)
                    p_cls = "g-pass" if p > 0 else "g-zero"
                    f_cls = "g-fail" if f_count > 0 else "g-zero"
                    w_cls = "g-warn" if w > 0 else "g-zero"
                    rows_html += f"""
                    <tr>
                        <td class="gate-name">{gname}</td>
                        <td class="{p_cls}">{p}</td>
                        <td class="{f_cls}">{f_count}</td>
                        <td class="{w_cls}">{w}</td>
                    </tr>"""

                st.markdown(f"""
                <table class="gate-table">
                    <thead>
                        <tr><th>Gate</th><th>Pass</th><th>Fail</th><th>Warn</th></tr>
                    </thead>
                    <tbody>{rows_html}</tbody>
                </table>
                """, unsafe_allow_html=True)

            # ── Sheets built ───────────────────────────────────────────
            if built:
                st.markdown(
                    '<div class="section-hdr">Output Sheets</div>'
                    + '<div style="font-family: JetBrains Mono, monospace; font-size: 0.78rem; color: #8b949e; padding: 0.3rem 0;">'
                    + " &middot; ".join(f'<span style="color:#c9d1d9">{s}</span>' for s in built)
                    + '</div>',
                    unsafe_allow_html=True
                )

            if failed:
                with st.expander(f"  {len(failed)} sheet(s) failed"):
                    for name, err in failed:
                        st.error(f"**{name}:** {err}")

            # ── Download button ────────────────────────────────────────
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
                    key="audit_download",
                )

            # ── Feedback / Dispute section ─────────────────────────────
            st.markdown('<div class="section-hdr">Scoring Feedback</div>', unsafe_allow_html=True)

            feedback_subject = "WO Audit Scoring Feedback"
            feedback_body = """WO Audit Scoring Dispute / Feedback
=====================================

1. WO Number: [e.g. WO-00142587]

2. Gate Flagged: [e.g. Documentation Quality - Warn]

3. Current Score & Grade: [e.g. 72/100, Grade C]

4. What Was Flagged Missing: [e.g. closure confirmation]

5. Full Corrective Action Text:
[PASTE THE COMPLETE CORRECTIVE ACTION FROM THE CORRECTIVE ACTION REPORT HERE — not the truncated WO field. This is the most important piece.]

6. Why You Disagree:
[Explain what the tech documented that should satisfy the missing element]

7. Subtype: [Paid / Courtesy / Warranty]

=====================================
This feedback helps calibrate the scoring model.
"""
            mailto_link = (
                "mailto:knutfinnjr@gmail.com"
                "?subject=" + urllib.parse.quote(feedback_subject)
                + "&body=" + urllib.parse.quote(feedback_body)
            )

            st.markdown(f"""
            <div class="feedback-box">
                <h3>Dispute a Score?</h3>
                <p>If a Documentation Quality score seems wrong, click below to open a
                pre-filled email with the details we need to review and improve the model.</p>
                <a href="{mailto_link}" class="feedback-btn" target="_blank">
                    Open Feedback Email
                </a>
            </div>
            """, unsafe_allow_html=True)

            # ── Processing log ─────────────────────────────────────────
            with st.expander("Processing log"):
                log_text = "\n".join(result["log"])
                log_html = log_text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                log_html = log_html.replace("\u2713", '<span class="log-ok">\u2713</span>')
                log_html = log_html.replace("\u2717", '<span class="log-err">\u2717</span>')
                log_html = log_html.replace("ERROR", '<span class="log-err">ERROR</span>')
                st.markdown(f'<div class="log-output"><pre>{log_html}</pre></div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
# TAB 2 — QUICK WO GATES (no Corrective Action grading)
# ══════════════════════════════════════════════════════════════════════════
with tab_quick:
    st.markdown("""
    <div style="font-family: 'JetBrains Mono', monospace; font-size: 0.8rem; color: #8b949e;
                padding: 0.5rem 0 1rem 0;">
        Pass/fail gate check &mdash; all gates <em>except</em> Documentation Quality scoring.
    </div>
    """, unsafe_allow_html=True)

    qg_wo_file = st.file_uploader(
        "**WO Approval Report** (.xlsx)",
        type=["xlsx"],
        help="Same Salesforce WO export. Gates run without CA grading.",
        key="qg_wo_file",
    )
    qg_haas_file = st.file_uploader(
        "**Haas RMA Status** (.xlsx) — optional",
        type=["xlsx"],
        help="Haas RMA status export for Parts + Haas RMA merge sheet.",
        key="qg_haas_file",
    )

    qg_disabled = qg_wo_file is None
    qg_clicked  = st.button(
        "Run Quick Gates",
        type="primary",
        use_container_width=True,
        disabled=qg_disabled,
        key="qg_run",
    )

    if qg_disabled and not qg_clicked:
        st.markdown("""
        <div style="font-family: 'JetBrains Mono', monospace; font-size: 0.8rem; color: #30363d;
                    text-align: center; padding: 2rem 0;">
            &gt; awaiting file upload_
        </div>
        """, unsafe_allow_html=True)

    if qg_clicked and qg_wo_file is not None:
        with st.spinner("Running quick gate check..."):
            start = time.time()
            qg_result = run_audit(qg_wo_file, qg_haas_file, skip_dq=True)
            elapsed = time.time() - start

        # ── Validation failure ─────────────────────────────────────────
        if not qg_result.get("validation_ok", True):
            st.error("VALIDATION FAILED — file is not a valid WO Approval Report")
            for e in qg_result.get("validation_errors", []):
                st.markdown(f"<div style='font-family: JetBrains Mono, monospace; font-size: 0.8rem; color: #ff3e3e; padding: 0.2rem 0;'>  &gt; {e}</div>", unsafe_allow_html=True)
        else:
            gs = qg_result["gate_summary"]
            gd = qg_result.get("gate_details", {})
            wo_n = qg_result["wo_count"]
            parts_n = qg_result["parts_count"]
            built = qg_result["built"]
            failed = qg_result["failed"]

            if built:
                st.success(f"QUICK GATES COMPLETE — {len(built)} sheets built in {elapsed:.1f}s")
            else:
                st.error("AUDIT FAILED — no sheets could be built")

            # ── Metric cards ───────────────────────────────────────────
            held_n = gs.get('hold', 0)
            st.markdown(f"""
            <div class="metrics-grid">
                <div class="m-card info">
                    <div class="m-val info">{wo_n}</div>
                    <div class="m-label">Work Orders</div>
                </div>
                <div class="m-card info">
                    <div class="m-val info">{parts_n}</div>
                    <div class="m-label">Part Lines</div>
                </div>
                <div class="m-card ready">
                    <div class="m-val ready">{gs['pass']}</div>
                    <div class="m-label">Ready</div>
                </div>
                <div class="m-card blocked">
                    <div class="m-val blocked">{gs['fail']}</div>
                    <div class="m-label">Blocked</div>
                </div>
                <div class="m-card warned">
                    <div class="m-val warned">{gs['warn']}</div>
                    <div class="m-label">Warned</div>
                </div>
                <div class="m-card held">
                    <div class="m-val held">{held_n}</div>
                    <div class="m-label">Held</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # ── Gate-level breakdown (no Doc Quality) ─────────────────
            if gd:
                st.markdown('<div class="section-hdr">Gate Breakdown</div>', unsafe_allow_html=True)
                rows_html = ""
                for gname, counts in gd.items():
                    p = counts.get("Pass", 0)
                    f_count = counts.get("Fail", 0)
                    w = counts.get("Warn", 0)
                    p_cls = "g-pass" if p > 0 else "g-zero"
                    f_cls = "g-fail" if f_count > 0 else "g-zero"
                    w_cls = "g-warn" if w > 0 else "g-zero"
                    rows_html += f"""
                    <tr>
                        <td class="gate-name">{gname}</td>
                        <td class="{p_cls}">{p}</td>
                        <td class="{f_cls}">{f_count}</td>
                        <td class="{w_cls}">{w}</td>
                    </tr>"""

                st.markdown(f"""
                <table class="gate-table">
                    <thead>
                        <tr><th>Gate</th><th>Pass</th><th>Fail</th><th>Warn</th></tr>
                    </thead>
                    <tbody>{rows_html}</tbody>
                </table>
                """, unsafe_allow_html=True)

            # ── CA Reminder Banner ────────────────────────────────────
            st.markdown("""
            <div class="ca-reminder">
                <h3>&#9888; Corrective Action Reminder</h3>
                <p>Documentation Quality scoring is <strong style="color:#58a6ff;">not included</strong>
                in this quick gate check. Please review the Corrective Action report manually
                at <strong style="color:#58a6ff;">final approval</strong> before sending to D365.</p>
            </div>
            """, unsafe_allow_html=True)

            # ── Sheets built ──────────────────────────────────────────
            if built:
                st.markdown(
                    '<div class="section-hdr">Output Sheets</div>'
                    + '<div style="font-family: JetBrains Mono, monospace; font-size: 0.78rem; color: #8b949e; padding: 0.3rem 0;">'
                    + " &middot; ".join(f'<span style="color:#c9d1d9">{s}</span>' for s in built)
                    + '</div>',
                    unsafe_allow_html=True
                )

            if failed:
                with st.expander(f"  {len(failed)} sheet(s) failed"):
                    for name, err in failed:
                        st.error(f"**{name}:** {err}")

            # ── Download button ───────────────────────────────────────
            if qg_result["xlsx_bytes"]:
                base_name = os.path.splitext(qg_wo_file.name)[0]
                out_name  = base_name + "_AUDIT.xlsx"
                st.download_button(
                    label=f"Download {out_name}",
                    data=qg_result["xlsx_bytes"],
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                    key="qg_download",
                )

            # ── Processing log ────────────────────────────────────────
            with st.expander("Processing log"):
                log_text = "\n".join(qg_result["log"])
                log_html = log_text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                log_html = log_html.replace("\u2713", '<span class="log-ok">\u2713</span>')
                log_html = log_html.replace("\u2717", '<span class="log-err">\u2717</span>')
                log_html = log_html.replace("ERROR", '<span class="log-err">ERROR</span>')
                st.markdown(f'<div class="log-output"><pre>{log_html}</pre></div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
# TAB 3 — ORPHAN WORK ORDER ANALYZER
# ══════════════════════════════════════════════════════════════════════════
with tab_orphan:
    st.markdown("""
    <div style="font-family: 'JetBrains Mono', monospace; font-size: 0.8rem; color: #8b949e;
                padding: 0.5rem 0 1rem 0;">
        Upload a WO/SA export to identify orphaned work orders &mdash;
        open WOs with no valid future service appointment.
    </div>
    """, unsafe_allow_html=True)

    orphan_file = st.file_uploader(
        "**WO / SA Export** (.xlsx)",
        type=["xlsx"],
        help="Salesforce export with Work Order Number, Account Name, Status, Appointment Number, SA Status, ESP, Scheduled Start, Due Date columns.",
        key="orphan_file",
    )

    orphan_disabled = orphan_file is None
    orphan_clicked  = st.button(
        "Analyze Orphans",
        type="primary",
        use_container_width=True,
        disabled=orphan_disabled,
        key="orphan_run",
    )

    if orphan_disabled and not orphan_clicked:
        st.markdown("""
        <div style="font-family: 'JetBrains Mono', monospace; font-size: 0.8rem; color: #30363d;
                    text-align: center; padding: 2rem 0;">
            &gt; awaiting WO/SA export_
        </div>
        """, unsafe_allow_html=True)

    if orphan_clicked and orphan_file is not None:
        with st.spinner("Scanning for orphaned work orders..."):
            start = time.time()
            result = run_orphan_analysis(orphan_file, source_filename=orphan_file.name)
            elapsed = time.time() - start

        # ── Validation failure ─────────────────────────────────────────
        if not result.get("validation_ok", True):
            st.error("VALIDATION FAILED — file is not a valid WO/SA export")
            for e in result.get("validation_errors", []):
                st.markdown(f"<div style='font-family: JetBrains Mono, monospace; font-size: 0.8rem; color: #ff3e3e; padding: 0.2rem 0;'>  &gt; {e}</div>", unsafe_allow_html=True)
        else:
            # ── Results ────────────────────────────────────────────────
            total    = result["total_wos"]
            orphans  = result["orphan_count"]
            healthy  = result["healthy_count"]
            rate     = result["orphan_rate"]
            reasons  = result.get("reason_counts", {})

            rate_pct = f"{rate * 100:.1f}%"

            if orphans == 0:
                st.success(f"ALL CLEAR — {total} active WOs, zero orphans ({elapsed:.1f}s)")
            else:
                st.warning(f"SCAN COMPLETE — {orphans} orphan(s) found across {total} active WOs ({elapsed:.1f}s)")

            # ── Metric cards ───────────────────────────────────────────
            st.markdown(f"""
            <div class="metrics-grid-3">
                <div class="m-card info">
                    <div class="m-val info">{total}</div>
                    <div class="m-label">Active WOs</div>
                </div>
                <div class="m-card blocked">
                    <div class="m-val blocked">{orphans}</div>
                    <div class="m-label">Orphans</div>
                </div>
                <div class="m-card ready">
                    <div class="m-val ready">{healthy}</div>
                    <div class="m-label">Healthy</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # ── Orphan rate ────────────────────────────────────────────
            st.markdown(f"""
            <div style="text-align: center; padding: 0.5rem 0;">
                <span style="font-family: 'JetBrains Mono', monospace; font-size: 0.8rem; color: #8b949e;">
                    Orphan Rate: </span>
                <span style="font-family: 'Rajdhani', sans-serif; font-size: 1.3rem; font-weight: 700;
                             color: {'#ff3e3e' if rate > 0.1 else '#ffb800' if rate > 0.05 else '#00ff9c'};">
                    {rate_pct}</span>
            </div>
            """, unsafe_allow_html=True)

            # ── Reason breakdown ───────────────────────────────────────
            if reasons:
                st.markdown('<div class="section-hdr">Orphan Reasons</div>', unsafe_allow_html=True)
                rows_html = ""
                for reason, cnt in sorted(reasons.items(), key=lambda x: -x[1]):
                    rows_html += f"""
                    <tr>
                        <td style="color: #ff3e3e; font-weight: 600; width: 60px; text-align: center;">{cnt}</td>
                        <td>{reason}</td>
                    </tr>"""

                st.markdown(f"""
                <table class="reason-table">
                    <thead>
                        <tr><th>Count</th><th>Reason</th></tr>
                    </thead>
                    <tbody>{rows_html}</tbody>
                </table>
                """, unsafe_allow_html=True)

            # ── Download button ────────────────────────────────────────
            if result.get("xlsx_bytes"):
                base_name = os.path.splitext(orphan_file.name)[0]
                out_name  = base_name + "_ORPHAN_REPORT.xlsx"
                st.download_button(
                    label=f"Download {out_name}",
                    data=result["xlsx_bytes"],
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                    key="orphan_download",
                )

            # ── Processing log ─────────────────────────────────────────
            with st.expander("Processing log"):
                log_text = "\n".join(result["log"])
                log_html = log_text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                log_html = log_html.replace("\u2713", '<span class="log-ok">\u2713</span>')
                log_html = log_html.replace("\u2717", '<span class="log-err">\u2717</span>')
                log_html = log_html.replace("ERROR", '<span class="log-err">ERROR</span>')
                st.markdown(f'<div class="log-output"><pre>{log_html}</pre></div>', unsafe_allow_html=True)


# ── Footer ─────────────────────────────────────────────────────────────────
st.markdown("""
<div class="console-footer">
    <span>&#9889;</span> WO Audit Console V1.27 &mdash; Engine v12.7 &middot; Quick Gates &middot; Orphan v1.0
    <span>&middot;</span> Built by K
</div>
""", unsafe_allow_html=True)
