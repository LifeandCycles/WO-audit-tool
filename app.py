"""
WO Audit Console — Streamlit Web App V1.30
Cyberpunk command console interface.
Modules: WO Approval Audit (v12.6) + Orphan Work Order Analyzer (v1.0)
"""
import streamlit as st
import os, time, urllib.parse

# ── Page config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="WO Audit Console v1.30",
    page_icon="⚡",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── Matrix CNC Rain Background ────────────────────────────────────────────
# Uses st.components.v1.html() because Streamlit strips <script> from st.markdown
import streamlit.components.v1 as components

MATRIX_HTML = """
<style>
    body, html { margin:0; padding:0; overflow:hidden; background:transparent; }
</style>
<canvas id="matrixCNC"></canvas>
<script>
(function() {
    const canvas = document.getElementById('matrixCNC');
    const ctx = canvas.getContext('2d');

    const glyphs = [
        'G00','G01','G02','G03','G28','G40','G41','G42','G43','G54',
        'G80','G81','G83','G90','G91','M00','M01','M03','M05','M06',
        'M08','M09','M30','S12000','F250','T01','T02','T03','T04',
        'X0.0','Y0.0','Z-1.5','Z0.1','A90','B45','H01','D01',
        'RPM','IPM','SFM','DOC','WCS','ATC','DRO','CNC',
        'HAAS','VF2','VF3','VF4','UMC','ST10','ST20','ST30',
        'FANUC','MAZAK','DMG','OKUMA','MORI',
        'SPINDLE','TURRET','COOLANT','PROBE','TOOL',
        'RAPID','FEED','DWELL','CYCLE','HOME',
        'WO','SA','PM','RMA','SN'
    ];

    function resize() {
        canvas.width = window.innerWidth;
        canvas.height = window.innerHeight;
    }
    resize();
    window.addEventListener('resize', resize);

    const fontSize = 13;
    let columns = Math.floor(canvas.width / (fontSize * 2.5));
    let drops = new Array(columns).fill(0).map(() => Math.random() * -50);
    let speeds = new Array(columns).fill(0).map(() => 0.3 + Math.random() * 0.7);

    function draw() {
        ctx.fillStyle = 'rgba(10, 14, 23, 0.15)';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.font = fontSize + 'px monospace';

        for (let i = 0; i < columns; i++) {
            const glyph = glyphs[Math.floor(Math.random() * glyphs.length)];
            const x = i * (fontSize * 2.5);
            const y = drops[i] * fontSize;

            ctx.fillStyle = '#00ff9c';
            ctx.fillText(glyph, x, y);

            if (drops[i] > 1) {
                const trailGlyph = glyphs[Math.floor(Math.random() * glyphs.length)];
                ctx.fillStyle = 'rgba(0, 255, 156, 0.3)';
                ctx.fillText(trailGlyph, x, y - fontSize * 2);
            }

            drops[i] += speeds[i];

            if (y > canvas.height && Math.random() > 0.975) {
                drops[i] = Math.random() * -20;
                speeds[i] = 0.3 + Math.random() * 0.7;
            }
        }

        const newCols = Math.floor(canvas.width / (fontSize * 2.5));
        if (newCols !== columns) {
            columns = newCols;
            drops = new Array(columns).fill(0).map(() => Math.random() * -50);
            speeds = new Array(columns).fill(0).map(() => 0.3 + Math.random() * 0.7);
        }
        requestAnimationFrame(draw);
    }
    draw();
})();
</script>
"""

# Inject the matrix iframe as a fixed background layer via CSS
st.markdown("""
<style>
    iframe[title="streamlit_app.static.matrix"],
    iframe[title="components.v1.html"] {
        position: fixed !important;
        top: 0 !important;
        left: 0 !important;
        width: 100vw !important;
        height: 100vh !important;
        z-index: 0 !important;
        pointer-events: none !important;
        opacity: 0.12 !important;
        border: none !important;
    }
</style>
""", unsafe_allow_html=True)

# Render the matrix animation in an iframe Streamlit will actually execute
components.html(MATRIX_HTML, height=0, scrolling=False)

# ── Cyberpunk Console CSS ──────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@300;400;500;700&family=Rajdhani:wght@400;500;600;700&display=swap');

    /* Global */
    .stApp { background-color: #0a0e17; }
    .block-container { max-width: 780px; padding-top: 1.5rem; position: relative; z-index: 1; }

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
        grid-template-columns: repeat(5, 1fr);
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

    /* ── Tabs ── */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0;
        border-bottom: 1px solid #21262d;
    }
    .stTabs [data-baseweb="tab"] {
        font-family: 'Rajdhani', sans-serif !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        letter-spacing: 1.5px !important;
        text-transform: uppercase !important;
        color: #8b949e !important;
        background: transparent !important;
        border: none !important;
        border-bottom: 2px solid transparent !important;
        padding: 0.6rem 1.2rem !important;
    }
    .stTabs [aria-selected="true"] {
        color: #00ff9c !important;
        border-bottom: 2px solid #00ff9c !important;
    }
    .stTabs [data-baseweb="tab"]:hover {
        color: #c9d1d9 !important;
    }
    .stTabs [data-baseweb="tab-panel"] {
        padding-top: 1rem;
    }

    /* ── Orphan metrics (3-col) ── */
    .metrics-grid-3 {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 0.5rem;
        margin: 1rem 0;
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
    .reason-table .r-count { color: #ff3e3e; font-weight: 600; font-family: 'Rajdhani', sans-serif; font-size: 1.1rem; }

    /* ── Orphan data table ── */
    .orphan-table {
        width: 100%;
        border-collapse: collapse;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.72rem;
        margin: 1rem 0;
    }
    .orphan-table th {
        text-align: left;
        color: #8b949e;
        font-weight: 500;
        padding: 0.4rem 0.5rem;
        border-bottom: 1px solid #21262d;
        font-size: 0.65rem;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .orphan-table td {
        padding: 0.4rem 0.5rem;
        border-bottom: 1px solid #161b2211;
        color: #c9d1d9;
    }
    .orphan-table tr:hover td { background: #161b22; }
    .orphan-table .wo-num { color: #58a6ff; font-weight: 600; }
    .orphan-table .reason-cell { color: #ff3e3e; font-size: 0.68rem; }

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
    <div class="version">V1.30</div>
    <h1>WO Audit Console</h1>
    <p class="subtitle">// approval audit &middot; orphan analyzer &middot; upload &gt; execute &gt; download</p>
</div>
""", unsafe_allow_html=True)

# ── Status line ────────────────────────────────────────────────────────────
st.markdown("""
<div class="status-line">
    <span class="dot"></span>SYSTEM ONLINE &mdash; audit engine v12.6 + orphan analyzer v1.0
</div>
""", unsafe_allow_html=True)

# ── Tabs ──────────────────────────────────────────────────────────────────
tab_audit, tab_orphan = st.tabs(["WO Approval Audit", "Orphan Analyzer"])

# ══════════════════════════════════════════════════════════════════════════
# TAB 1: WO APPROVAL AUDIT
# ══════════════════════════════════════════════════════════════════════════
with tab_audit:
    # ── File uploads ───────────────────────────────────────────────────
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

    # ── Run button ─────────────────────────────────────────────────────
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

    # ── Execute audit ──────────────────────────────────────────────────
    if run_clicked and wo_file is not None:
        with st.spinner("Executing audit sequence..."):
            start = time.time()
            from audit_engine import run_audit
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

            # Build mailto template
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
# TAB 2: ORPHAN WORK ORDER ANALYZER
# ══════════════════════════════════════════════════════════════════════════
with tab_orphan:
    orphan_file = st.file_uploader(
        "**WO / SA Export** (.xlsx)",
        type=["xlsx"],
        help="Salesforce export with Work Order Number, Account, Status, Appointment Number, SA Status, Earliest Start Permitted, Scheduled Start, Due Date.",
        key="orphan_file",
    )

    orphan_run_disabled = orphan_file is None
    orphan_run_clicked  = st.button(
        "Analyze Orphans",
        type="primary",
        use_container_width=True,
        disabled=orphan_run_disabled,
        key="orphan_run",
    )

    if orphan_run_disabled and not orphan_run_clicked:
        st.markdown("""
        <div style="font-family: 'JetBrains Mono', monospace; font-size: 0.8rem; color: #30363d;
                    text-align: center; padding: 2rem 0;">
            &gt; awaiting WO/SA export upload_
        </div>
        """, unsafe_allow_html=True)

    # ── Execute orphan analysis ────────────────────────────────────────
    if orphan_run_clicked and orphan_file is not None:
        with st.spinner("Running orphan analysis..."):
            start = time.time()
            from orphan_engine import run_orphan_analysis
            o_result = run_orphan_analysis(orphan_file, source_filename=orphan_file.name)
            elapsed = time.time() - start

        # ── Validation failure ─────────────────────────────────────────
        if not o_result.get("validation_ok", True):
            st.error("VALIDATION FAILED — file is not a valid WO/SA export")
            for e in o_result.get("validation_errors", []):
                st.markdown(f"<div style='font-family: JetBrains Mono, monospace; font-size: 0.8rem; color: #ff3e3e; padding: 0.2rem 0;'>  &gt; {e}</div>", unsafe_allow_html=True)
        else:
            total   = o_result["total_wos"]
            orphans = o_result["orphan_count"]
            healthy = o_result["healthy_count"]
            rate    = o_result["orphan_rate"]
            reasons = o_result["reason_counts"]
            orph_list = o_result["orphans"]

            if orphans == 0:
                st.success(f"ANALYSIS COMPLETE — 0 orphans found across {total} WOs in {elapsed:.1f}s")
            else:
                st.warning(f"ANALYSIS COMPLETE — {orphans} orphan(s) found across {total} WOs in {elapsed:.1f}s")

            # ── Metric cards (3-col) ───────────────────────────────────
            rate_pct = f"{rate * 100:.1f}%"
            st.markdown(f"""
            <div class="metrics-grid-3">
                <div class="m-card info">
                    <div class="m-val info">{total}</div>
                    <div class="m-label">Active WOs</div>
                </div>
                <div class="m-card blocked">
                    <div class="m-val blocked">{orphans}</div>
                    <div class="m-label">Orphaned</div>
                </div>
                <div class="m-card ready">
                    <div class="m-val ready">{healthy}</div>
                    <div class="m-label">Healthy</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # ── Orphan rate bar ────────────────────────────────────────
            st.markdown(f"""
            <div style="font-family: 'JetBrains Mono', monospace; font-size: 0.75rem; color: #8b949e; margin: 0.5rem 0 0.3rem 0;">
                ORPHAN RATE: <span style="color: #ff3e3e; font-weight: 600;">{rate_pct}</span>
            </div>
            <div style="background: #161b22; border-radius: 3px; height: 6px; overflow: hidden; margin-bottom: 1rem;">
                <div style="background: linear-gradient(90deg, #ff3e3e, #ff6b6b); width: {min(rate * 100, 100):.1f}%; height: 100%; border-radius: 3px;"></div>
            </div>
            """, unsafe_allow_html=True)

            # ── Reason breakdown ───────────────────────────────────────
            if reasons:
                st.markdown('<div class="section-hdr">Orphan Reasons</div>', unsafe_allow_html=True)
                reason_rows = ""
                for reason, count in sorted(reasons.items(), key=lambda x: -x[1]):
                    reason_rows += f"""
                    <tr>
                        <td class="r-count">{count}</td>
                        <td>{reason}</td>
                    </tr>"""
                st.markdown(f"""
                <table class="reason-table">
                    <thead><tr><th>Count</th><th>Reason</th></tr></thead>
                    <tbody>{reason_rows}</tbody>
                </table>
                """, unsafe_allow_html=True)

            # ── Orphan detail table ────────────────────────────────────
            if orph_list:
                st.markdown('<div class="section-hdr">Orphaned Work Orders</div>', unsafe_allow_html=True)
                detail_rows = ""
                for o in orph_list:
                    sched_str = o["sched_start"].strftime("%Y-%m-%d %H:%M") if o["sched_start"] else "&mdash;"
                    due_str   = o["due_date"].strftime("%Y-%m-%d %H:%M") if o["due_date"] else "&mdash;"
                    detail_rows += f"""
                    <tr>
                        <td class="wo-num">{o['wo_num']}</td>
                        <td>{o['wo_status']}</td>
                        <td>{o['sa_status']}</td>
                        <td>{sched_str}</td>
                        <td>{due_str}</td>
                        <td class="reason-cell">{o['reason']}</td>
                    </tr>"""
                st.markdown(f"""
                <table class="orphan-table">
                    <thead>
                        <tr>
                            <th>WO Number</th>
                            <th>WO Status</th>
                            <th>SA Status</th>
                            <th>Sched Start</th>
                            <th>Due Date</th>
                            <th>Orphan Reason</th>
                        </tr>
                    </thead>
                    <tbody>{detail_rows}</tbody>
                </table>
                """, unsafe_allow_html=True)

            # ── Download button ────────────────────────────────────────
            if o_result["xlsx_bytes"]:
                out_name = "Orphaned_Work_Order_Report.xlsx"
                st.download_button(
                    label=f"Download {out_name}",
                    data=o_result["xlsx_bytes"],
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                    key="orphan_download",
                )

            # ── Processing log ─────────────────────────────────────────
            with st.expander("Processing log"):
                log_text = "\n".join(o_result["log"])
                log_html = log_text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                log_html = log_html.replace("[OK]", '<span class="log-ok">[OK]</span>')
                log_html = log_html.replace("[INFO]", '<span class="log-info">[INFO]</span>')
                log_html = log_html.replace("ERROR", '<span class="log-err">ERROR</span>')
                st.markdown(f'<div class="log-output"><pre>{log_html}</pre></div>', unsafe_allow_html=True)

# ── Footer ─────────────────────────────────────────────────────────────────
st.markdown("""
<div class="console-footer">
    <span>&#9889;</span> WO Audit Console V1.30 &mdash; Audit v12.6 + Orphan v1.0
    <span>&middot;</span> Built by K
</div>
""", unsafe_allow_html=True)
