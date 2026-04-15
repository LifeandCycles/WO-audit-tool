"""
ledger_engine.py — Per-HFO append-only WO ledger.

Purpose:
    Track every Field-Work-Complete WO over time, per HFO (3010/3020/3030/3040),
    so team and per-technician monthly trends can be computed without losing
    row-level fidelity.

Design:
    - One SQLite file per HFO at data/ledger_<hfo>.db (honors "each HFO
      individually" per user spec).
    - Single table `wo_ledger` keyed by WO# (stable identity per user spec).
    - Re-ingesting the same report safely upserts: first_seen_run_date is
      preserved, last_seen_run_date + final_result update in place.
    - "Passed all gates up front" = engine's Final row != "WO blocked"
      (Destin-01-only failures are approver's responsibility and count as
      passed, per v12.10 design).
    - Per-HFO rules_<hfo>.json can override tech/territory name normalization,
      declare extra hard blockers, or treat Destin-01 as a hard fail.

Additive module — does not modify any existing v12.7 file.
"""
from __future__ import annotations

import json
import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd

# Reuse the v12.7 engine's loaders and gate functions — never duplicate logic.
from audit_engine import (
    AUDIT_STATUSES,
    load_data,
    build_destination_lookups,
    build_ca_lookup,
    get_ca,
    check_hold,
    gate_wo_status,
    gate_destination,
    gate_ro_eligible_parts,
    gate_pr_prli,
    gate_consumed_vs_nnu,
    gate_ro_coverage,
    gate_ro_status,
    gate_ca_duplicate,
    dq_gate,
    _build_tech_ca_index,
)

# ── Paths ────────────────────────────────────────────────────────────────
_APP_DIR   = Path(__file__).resolve().parent
_ROOT      = _APP_DIR.parent                       # WO approval (1)/
DATA_DIR   = _ROOT / "data"
CONFIG_DIR = _ROOT / "config"
VALID_HFOS = ("3010", "3020", "3030", "3040")


# ── Rules ────────────────────────────────────────────────────────────────
def load_rules(hfo: str) -> dict:
    """Load rules_<hfo>.json. Returns empty-defaults dict if file missing."""
    if hfo not in VALID_HFOS:
        raise ValueError(f"Invalid HFO {hfo!r}; expected one of {VALID_HFOS}")
    path = CONFIG_DIR / f"rules_{hfo}.json"
    defaults = {
        "hfo": hfo,
        "technician_aliases": {},
        "territory_aliases": {},
        "technician_home_territory": {},
        "team_leads": {},
        "former_technicians": [],
        "extra_blockers": [],
        "exclude_from_rate": {"subtypes": [], "work_types": []},
        "destin_is_hard_fail": False,
    }
    if not path.exists():
        return defaults
    with path.open() as f:
        loaded = json.load(f)
    # Shallow-merge: any key the user hasn't set falls back to default.
    for k, v in defaults.items():
        loaded.setdefault(k, v)
    return loaded


def _norm(value: str, alias_map: dict) -> str:
    """Apply alias map with case-insensitive match; return original if no hit."""
    if not value:
        return value
    s = str(value).strip()
    if s in alias_map:
        return alias_map[s]
    lower = {k.lower(): v for k, v in alias_map.items()}
    return lower.get(s.lower(), s)


# ── Schema ───────────────────────────────────────────────────────────────
_SCHEMA = """
CREATE TABLE IF NOT EXISTS wo_ledger (
    wo_id               TEXT PRIMARY KEY,
    hfo                 TEXT NOT NULL,
    wo_status           TEXT,
    territory           TEXT,       -- where the WO was worked (from sheet col D)
    home_territory      TEXT,       -- tech's home team (falls back to territory)
    technician          TEXT,
    customer            TEXT,
    subtype             TEXT,
    work_type           TEXT,
    final_result        TEXT,
    passed_up_front     INTEGER,
    blocking_gates      TEXT,       -- JSON list
    destin_only_fail    INTEGER,
    dq_result           TEXT,
    dq_detail           TEXT,
    field_complete_date TEXT,       -- PRIMARY month key: WO End Date, fallback to Created Date
    date_source         TEXT,       -- 'end_date' | 'created_date' | 'unknown'
    first_seen_run_date TEXT,       -- audit trail: first ingest
    last_seen_run_date  TEXT,       -- audit trail: most recent ingest
    source_file         TEXT
);

CREATE INDEX IF NOT EXISTS idx_wo_ledger_field_date    ON wo_ledger(field_complete_date);
CREATE INDEX IF NOT EXISTS idx_wo_ledger_technician    ON wo_ledger(technician);
CREATE INDEX IF NOT EXISTS idx_wo_ledger_territory     ON wo_ledger(territory);
CREATE INDEX IF NOT EXISTS idx_wo_ledger_home_territory ON wo_ledger(home_territory);

CREATE TABLE IF NOT EXISTS tech_activity (
    wo_id            TEXT NOT NULL,
    hfo              TEXT NOT NULL,
    technician       TEXT NOT NULL,
    home_territory   TEXT,
    line_type        TEXT NOT NULL,   -- 'Travel', 'Labor', 'Expense', ...
    month            TEXT NOT NULL,   -- YYYY-MM (from line-item EndDate)
    billable_hours   REAL DEFAULT 0,
    nonbill_hours    REAL DEFAULT 0,
    courtesy_hours   REAL DEFAULT 0,
    total_hours      REAL DEFAULT 0,
    source           TEXT,            -- 'woli2' | 'woli_legacy'
    first_seen_run_date TEXT,
    last_seen_run_date  TEXT,
    source_file      TEXT,
    PRIMARY KEY (wo_id, technician, line_type, month)
);

CREATE INDEX IF NOT EXISTS idx_ta_tech   ON tech_activity(technician);
CREATE INDEX IF NOT EXISTS idx_ta_month  ON tech_activity(month);
CREATE INDEX IF NOT EXISTS idx_ta_type   ON tech_activity(line_type);
CREATE INDEX IF NOT EXISTS idx_ta_wo     ON tech_activity(wo_id);
"""


def _db_path(hfo: str) -> Path:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    return DATA_DIR / f"ledger_{hfo}.db"


def _connect(path: Path) -> sqlite3.Connection:
    """
    Connect with journal_mode=MEMORY — required because the ledger lives on a
    FUSE-mounted workspace where SQLite's default WAL journal hits I/O errors.
    """
    conn = sqlite3.connect(path)
    conn.execute("PRAGMA journal_mode=MEMORY")
    conn.execute("PRAGMA synchronous=NORMAL")
    return conn


def init_ledger(hfo: str) -> Path:
    """Create (if needed) and return the path to the per-HFO ledger DB."""
    path = _db_path(hfo)
    with _connect(path) as conn:
        conn.executescript(_SCHEMA)
    return path


# ── Date lookup (field complete date) ────────────────────────────────────
def _build_date_lookup(data: dict) -> dict:
    """
    Build {wo_num: (iso_date, source)} where source is 'end_date' |
    'created_date' | 'unknown'. Primary: WO sheet "End Date". Fallback:
    "Created Date (Reporting)". Handles mixed ISO + m/d/yyyy formats.
    """
    out: dict[str, tuple[str, str]] = {}
    wo = data.get("WO")
    if wo is None or "WO#" not in wo.columns:
        return out
    for _, r in wo.iterrows():
        wn = str(r.get("WO#", "")).strip()
        if not wn:
            continue
        end_raw = r.get("End Date")
        crt_raw = r.get("Created Date (Reporting)")
        parsed_end = pd.to_datetime(end_raw, errors="coerce")
        parsed_crt = pd.to_datetime(crt_raw, errors="coerce")
        if pd.notna(parsed_end):
            out[wn] = (parsed_end.date().isoformat(), "end_date")
        elif pd.notna(parsed_crt):
            out[wn] = (parsed_crt.date().isoformat(), "created_date")
        else:
            out[wn] = ("", "unknown")
    return out


# ── Tech fallback lookup ─────────────────────────────────────────────────
def _build_tech_fallback(data: dict) -> dict:
    """
    Build {wo_num: tech_name} from WOLI and WO sheets for use when WO_Output's
    WO_Technician is blank or "-". Priority order (highest wins on overwrite):
      1. WO sheet (Service Technician: Name)           — lowest priority
      2. WOLI legacy sheet (Service Resource: Name)    — preferred (line-item level)
    WOs where every source is blank stay unresolved.
    """
    lookup: dict[str, str] = {}

    def _clean(v) -> str:
        s = str(v or "").strip()
        return "" if s in {"", "-"} else s

    wo_df = data.get("WO")
    if wo_df is not None and "WO#" in wo_df.columns and "Service Technician: Name" in wo_df.columns:
        for _, r in wo_df.iterrows():
            wn = str(r["WO#"]).strip()
            t = _clean(r["Service Technician: Name"])
            if wn and t:
                lookup[wn] = t

    woli = data.get("WOLI")
    if woli is not None and "WO#" in woli.columns and "Service Resource: Name" in woli.columns:
        for _, r in woli.iterrows():
            wn = str(r["WO#"]).strip()
            t = _clean(r["Service Resource: Name"])
            if wn and t:
                lookup[wn] = t

    return lookup


# ── Gate execution (pure — no xlsx output) ───────────────────────────────
def _compute_wo_records(data: dict, rules: dict) -> list[dict]:
    """
    Run the same gate pipeline the audit engine runs, but return per-WO dicts
    instead of writing a workbook. Final-result logic mirrors build_gate_summary
    exactly so the ledger stays consistent with the engine's xlsx output.
    """
    wo_all = data["WO_Output"]
    wo_df  = (
        wo_all[wo_all["WO_Status"].isin(AUDIT_STATUSES)]
        if "WO_Status" in wo_all.columns else wo_all
    )
    pa_df = data["Parts_Output"]
    ro_df = data["RO"]
    ca_lk = data.get("ca_lookup", {}) or build_ca_lookup(data)

    woli_lk, sa_lk, wo_cust, destin_lk = build_destination_lookups(data)
    tech_ca_idx     = _build_tech_ca_index(wo_df, ca_lk)
    tech_fallback   = _build_tech_fallback(data)
    date_lookup     = _build_date_lookup(data)

    tech_aliases  = rules.get("technician_aliases", {}) or {}
    terr_aliases  = rules.get("territory_aliases",  {}) or {}
    home_map      = rules.get("technician_home_territory", {}) or {}
    # Normalize home_map keys through the alias map so a roster entry under the
    # display name (e.g. "Joe Blodgett Jr") matches a sheet value that came in
    # under the pre-alias name ("Joseph Blodgett Jr").
    home_map_ci   = {k.strip().lower(): v for k, v in home_map.items()}
    extra_blockers = set(rules.get("extra_blockers", []) or [])
    destin_hard   = bool(rules.get("destin_is_hard_fail", False))

    out: list[dict] = []

    for _, wrow in wo_df.iterrows():
        wn   = str(wrow.get("WO#", "")).strip()
        if not wn:
            continue
        cust = str(wrow.get("Customer", "") or "")
        raw_tech = str(wrow.get("WO_Technician", "") or "").strip()
        if raw_tech in {"", "-"}:
            raw_tech = tech_fallback.get(wn, raw_tech)  # WOLI → WO sheet fallback
        tech = _norm(raw_tech, tech_aliases)
        terr = _norm(str(wrow.get("Territory",     "") or ""), terr_aliases)
        sub  = str(wrow.get("WO_Subtype", "") or "")
        wt   = str(wrow.get("WorkType",   "") or "")
        st   = str(wrow.get("WO_Status",  "") or "")

        # Home territory = configured mapping, or fall back to WO's territory.
        # This is what lets Bryson's training WOs in East roll up to Southwest
        # for per-tech views, while per-territory views still see them as East.
        home_terr = home_map_ci.get(tech.strip().lower(), terr)

        # Field-complete date: End Date primary, Created Date fallback.
        fcd, date_src = date_lookup.get(wn, ("", "unknown"))

        wp    = pa_df[pa_df["WO#"] == wn] if len(pa_df) else pd.DataFrame()
        parts = wp.to_dict("records")

        cause, ca = get_ca(ca_lk, wn)

        is_held, hold_detail = check_hold(wn, sa_lk)
        if is_held:
            out.append({
                "wo_id": wn, "wo_status": st, "territory": terr,
                "home_territory": home_terr, "technician": tech,
                "customer": cust, "subtype": sub, "work_type": wt,
                "final_result": "HOLD", "passed_up_front": False,
                "blocking_gates": ["HOLD"], "destin_only_fail": False,
                "dq_result": None, "dq_detail": hold_detail,
                "field_complete_date": fcd, "date_source": date_src,
            })
            continue

        gates = [
            ("WO Status Ready",              gate_wo_status(wrow.to_dict())),
            ("Destination Review",           gate_destination(wrow.to_dict(), woli_lk, sa_lk, wo_cust, destin_lk)),
            ("RO-Eligible Parts Identified", gate_ro_eligible_parts(parts)),
            ("PR / PRLI Integrity",          gate_pr_prli(parts)),
            ("Consumed vs Not Used",         gate_consumed_vs_nnu(parts)),
            ("Required RO Coverage",         gate_ro_coverage(parts, sub, ro_df, wn)),
            ("RO Status Valid",              gate_ro_status(parts, ro_df, wn)),
        ]
        failed = [g for g, (r, _) in gates if r == "Fail"]

        # Doc Quality gate (Warn counts as pass, only Fail blocks — matches engine)
        has_parts = len(parts) > 0
        dq_result, dq_detail, dq_scores = dq_gate(cause, ca, sub, wn, has_parts=has_parts)
        # CA duplicate merge — same logic as build_gate_summary
        dup_r, dup_d = gate_ca_duplicate(wn, tech, tech_ca_idx)
        if dup_r == "Warn":
            dq_detail = f"{dq_detail} | COPY/PASTE WARNING: {dup_d}"
            if dq_result == "Pass":
                dq_result = "Warn"
        if dq_result == "Fail":
            failed.append("Documentation Quality")

        # Final-row replication
        hard_fails  = [g for g in failed if g != "Destination Review"] + [
            g for g in failed if g in extra_blockers and g not in {"Destination Review"}
        ]
        # Per-HFO option: promote Destination Review to hard fail
        if destin_hard and "Destination Review" in failed:
            hard_fails.append("Destination Review")

        destin_only = ("Destination Review" in failed) and (not hard_fails)

        if hard_fails:
            final_result = "WO blocked"
        elif destin_only:
            final_result = "WO approved — add destination/fees before sending to D365"
        else:
            final_result = "WO ready to send to D365"

        passed_up_front = (final_result != "WO blocked")

        out.append({
            "wo_id": wn,
            "wo_status": st,
            "territory": terr,
            "home_territory": home_terr,
            "technician": tech,
            "customer": cust,
            "subtype": sub,
            "work_type": wt,
            "final_result": final_result,
            "passed_up_front": passed_up_front,
            "blocking_gates": list(dict.fromkeys(failed)),  # preserve order, dedup
            "destin_only_fail": destin_only,
            "dq_result": dq_result,
            "dq_detail": dq_detail,
            "field_complete_date": fcd,
            "date_source": date_src,
        })

    return out


# ── Tech activity computation (WOLI hours per tech per type per month) ───
def _compute_tech_activity(data: dict, rules: dict) -> list[dict]:
    """
    Build per-(wo, tech, type, month) hour totals from WOLI sources.

    Sources (priority order):
      1. WOLI2 (Work Order Line Item, SF ID join) — complete data, hours split
         across billable/nonbill/courtesy quantities.
      2. Legacy WOLI — fills in WOs not represented in WOLI2. Only total
         minutes available; billable/nonbill/courtesy stored as 0.

    Tech names pass through the same technician_aliases used by wo_ledger.
    """
    tech_aliases = rules.get("technician_aliases", {}) or {}
    home_map_ci  = {
        k.strip().lower(): v
        for k, v in (rules.get("technician_home_territory", {}) or {}).items()
    }

    # Build 15-char SF ID → WO# lookup via WO sheet.
    wo_sheet = data.get("WO")
    sf15_to_wo: dict[str, str] = {}
    if wo_sheet is not None and "Work Order ID" in wo_sheet.columns and "WO#" in wo_sheet.columns:
        for _, r in wo_sheet.iterrows():
            sf = str(r.get("Work Order ID") or "").strip()[:15]
            wn = str(r.get("WO#") or "").strip()
            if sf and wn:
                sf15_to_wo[sf] = wn

    # Month key = line-item EndDate (fallback StartDate).
    def _month_from(row, end_col: str, start_col: str) -> str:
        for col in (end_col, start_col):
            if col in row and row[col] is not None and str(row[col]).strip():
                ts = pd.to_datetime(row[col], errors="coerce")
                if pd.notna(ts):
                    return ts.strftime("%Y-%m")
        return ""

    def _num(v) -> float:
        try:
            return float(v)
        except (TypeError, ValueError):
            return 0.0

    # {(wo_id, tech, type, month): dict_of_sums}
    agg: dict[tuple[str, str, str, str], dict] = {}

    def _bump(wo_id, tech, line_type, month, *, billable=0.0, nonbill=0.0,
              courtesy=0.0, total=None, source="woli2"):
        if not wo_id or not tech or not line_type or not month:
            return
        total_hrs = total if total is not None else (billable + nonbill + courtesy)
        key = (wo_id, tech, line_type, month)
        entry = agg.setdefault(key, {
            "billable": 0.0, "nonbill": 0.0, "courtesy": 0.0,
            "total": 0.0, "source": source,
        })
        entry["billable"] += billable
        entry["nonbill"]  += nonbill
        entry["courtesy"] += courtesy
        entry["total"]    += total_hrs
        # If any row in this bucket came from WOLI2, mark the bucket woli2.
        if source == "woli2":
            entry["source"] = "woli2"

    # ── 1. WOLI2 (primary) ───────────────────────────────────────────────
    w2 = data.get("WOLI2")
    seen_in_w2: set[str] = set()
    if w2 is not None and "WorkOrderId" in w2.columns:
        for _, r in w2.iterrows():
            sf15 = str(r.get("WorkOrderId") or "").strip()[:15]
            wn   = sf15_to_wo.get(sf15)
            if not wn:
                continue
            seen_in_w2.add(wn)
            raw_tech = str(r.get("Service_Resource_Name__c") or "").strip()
            if not raw_tech or raw_tech == "-":
                continue
            tech = _norm(raw_tech, tech_aliases)
            ltype = str(r.get("Type__c") or "").strip()
            month = _month_from(r, "EndDate", "StartDate")
            _bump(wn, tech, ltype, month,
                  billable=_num(r.get("Billable_Quantity__c")),
                  nonbill =_num(r.get("Non_Billable_Qty__c")),
                  courtesy=_num(r.get("Courtesy_Qty__c")),
                  source="woli2")

    # ── 2. Legacy WOLI (fallback for WOs absent from WOLI2) ─────────────
    wl = data.get("WOLI")
    if wl is not None and "WO#" in wl.columns:
        for _, r in wl.iterrows():
            wn = str(r.get("WO#") or "").strip()
            if not wn or wn in seen_in_w2:
                continue
            raw_tech = str(r.get("Service Resource: Name") or "").strip()
            if not raw_tech or raw_tech == "-":
                continue
            tech  = _norm(raw_tech, tech_aliases)
            ltype = str(r.get("Type") or "").strip()
            month = _month_from(r, "End Date", "Start Time")
            minutes = _num(r.get("Time Duration (Minutes)"))
            _bump(wn, tech, ltype, month, total=minutes / 60.0, source="woli_legacy")

    # Flatten with home_territory attached.
    out: list[dict] = []
    for (wo_id, tech, ltype, month), vals in agg.items():
        home = home_map_ci.get(tech.strip().lower(), "")
        out.append({
            "wo_id": wo_id,
            "technician": tech,
            "home_territory": home,
            "line_type": ltype,
            "month": month,
            "billable_hours": round(vals["billable"], 4),
            "nonbill_hours":  round(vals["nonbill"], 4),
            "courtesy_hours": round(vals["courtesy"], 4),
            "total_hours":    round(vals["total"], 4),
            "source": vals["source"],
        })
    return out


# ── Ingest ───────────────────────────────────────────────────────────────
def ingest(report_path: str | Path, hfo: str, run_date: str | None = None) -> dict:
    """
    Load a WO Approval Report workbook and upsert its WOs into the HFO ledger.

    Args:
        report_path: path to the .xlsx export
        hfo: one of "3010", "3020", "3030", "3040"
        run_date: ISO date string (YYYY-MM-DD); defaults to today.

    Returns a summary dict: {inserted, updated, total, hfo, db_path}.
    """
    if hfo not in VALID_HFOS:
        raise ValueError(f"Invalid HFO {hfo!r}; expected one of {VALID_HFOS}")

    run_date = run_date or datetime.now().strftime("%Y-%m-%d")
    report_path = Path(report_path)
    if not report_path.exists():
        raise FileNotFoundError(report_path)

    rules = load_rules(hfo)
    data  = load_data(str(report_path))
    if "ca_lookup" not in data:
        data["ca_lookup"] = build_ca_lookup(data)

    records = _compute_wo_records(data, rules)
    activity = _compute_tech_activity(data, rules)

    db = init_ledger(hfo)
    inserted = 0
    updated  = 0
    activity_rows = 0

    with _connect(db) as conn:
        conn.execute("PRAGMA foreign_keys = ON")
        cur = conn.cursor()
        for rec in records:
            existing = cur.execute(
                "SELECT first_seen_run_date FROM wo_ledger WHERE wo_id = ?",
                (rec["wo_id"],),
            ).fetchone()

            payload = (
                rec["wo_id"],
                hfo,
                rec["wo_status"],
                rec["territory"],
                rec["home_territory"],
                rec["technician"],
                rec["customer"],
                rec["subtype"],
                rec["work_type"],
                rec["final_result"],
                1 if rec["passed_up_front"] else 0,
                json.dumps(rec["blocking_gates"]),
                1 if rec["destin_only_fail"] else 0,
                rec["dq_result"],
                rec["dq_detail"],
                rec["field_complete_date"],
                rec["date_source"],
            )

            if existing is None:
                cur.execute(
                    """
                    INSERT INTO wo_ledger (
                        wo_id, hfo, wo_status, territory, home_territory, technician, customer,
                        subtype, work_type, final_result, passed_up_front,
                        blocking_gates, destin_only_fail, dq_result, dq_detail,
                        field_complete_date, date_source,
                        first_seen_run_date, last_seen_run_date, source_file
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """,
                    payload + (run_date, run_date, report_path.name),
                )
                inserted += 1
            else:
                cur.execute(
                    """
                    UPDATE wo_ledger SET
                        hfo=?, wo_status=?, territory=?, home_territory=?, technician=?, customer=?,
                        subtype=?, work_type=?, final_result=?, passed_up_front=?,
                        blocking_gates=?, destin_only_fail=?, dq_result=?, dq_detail=?,
                        field_complete_date=?, date_source=?,
                        last_seen_run_date=?, source_file=?
                    WHERE wo_id=?
                    """,
                    payload[1:] + (run_date, report_path.name, rec["wo_id"]),
                )
                updated += 1

        # ── tech_activity: delete-and-replace for WOs touched by this report ─
        # Include every WO that appears anywhere in the report's WOLI data,
        # not just WOs in records — most tech_activity rows come from WOs that
        # aren't yet Field Work Complete (so they never enter `records`).
        wos_in_report = sorted(
            {r["wo_id"] for r in records} | {a["wo_id"] for a in activity}
        )
        first_seen_keep: dict[tuple, str] = {}
        if wos_in_report:
            # SQLite has a parameter limit (~999). Chunk if needed.
            CHUNK = 500
            for i in range(0, len(wos_in_report), CHUNK):
                chunk = wos_in_report[i:i + CHUNK]
                ph = ",".join("?" * len(chunk))
                cur.execute(
                    f"SELECT wo_id, technician, line_type, month, first_seen_run_date "
                    f"FROM tech_activity WHERE wo_id IN ({ph})",
                    chunk,
                )
                for row in cur.fetchall():
                    first_seen_keep[(row[0], row[1], row[2], row[3])] = row[4]
                cur.execute(
                    f"DELETE FROM tech_activity WHERE wo_id IN ({ph})",
                    chunk,
                )

        for a in activity:
            key = (a["wo_id"], a["technician"], a["line_type"], a["month"])
            first_seen = first_seen_keep.get(key, run_date)
            cur.execute(
                """
                INSERT INTO tech_activity (
                    wo_id, hfo, technician, home_territory, line_type, month,
                    billable_hours, nonbill_hours, courtesy_hours, total_hours,
                    source, first_seen_run_date, last_seen_run_date, source_file
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """,
                (
                    a["wo_id"], hfo, a["technician"], a["home_territory"],
                    a["line_type"], a["month"],
                    a["billable_hours"], a["nonbill_hours"], a["courtesy_hours"], a["total_hours"],
                    a["source"], first_seen, run_date, report_path.name,
                ),
            )
            activity_rows += 1

        conn.commit()

    return {
        "hfo": hfo,
        "total": len(records),
        "inserted": inserted,
        "updated": updated,
        "activity_rows": activity_rows,
        "db_path": str(db),
        "run_date": run_date,
        "source_file": report_path.name,
    }


# ── Read helpers (used by Trends page in next step) ──────────────────────
def read_ledger(hfo: str) -> pd.DataFrame:
    """Return the full ledger for an HFO as a DataFrame (empty if none yet)."""
    db = _db_path(hfo)
    if not db.exists():
        return pd.DataFrame()
    with _connect(db) as conn:
        df = pd.read_sql_query("SELECT * FROM wo_ledger", conn)
    if not df.empty:
        df["first_seen_run_date"] = pd.to_datetime(df["first_seen_run_date"])
        df["last_seen_run_date"]  = pd.to_datetime(df["last_seen_run_date"])
        df["field_complete_date"] = pd.to_datetime(df["field_complete_date"], errors="coerce")
        df["passed_up_front"]     = df["passed_up_front"].astype(bool)
        df["destin_only_fail"]    = df["destin_only_fail"].astype(bool)
        df["blocking_gates"]      = df["blocking_gates"].apply(
            lambda s: json.loads(s) if s else []
        )
    return df


def read_tech_activity(hfo: str) -> pd.DataFrame:
    """Return tech_activity rows for an HFO (empty frame if none yet)."""
    db = _db_path(hfo)
    if not db.exists():
        return pd.DataFrame()
    with _connect(db) as conn:
        df = pd.read_sql_query("SELECT * FROM tech_activity", conn)
    if not df.empty:
        df["first_seen_run_date"] = pd.to_datetime(df["first_seen_run_date"])
        df["last_seen_run_date"]  = pd.to_datetime(df["last_seen_run_date"])
    return df


def read_all_ledgers() -> pd.DataFrame:
    """Concatenate every per-HFO ledger into one frame (for cross-HFO views)."""
    frames = []
    for hfo in VALID_HFOS:
        df = read_ledger(hfo)
        if not df.empty:
            frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


# ── CLI for quick manual ingest/testing ─────────────────────────────────
if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Ingest a WO Approval Report into an HFO ledger.")
    ap.add_argument("report", help="Path to .xlsx report")
    ap.add_argument("hfo", choices=VALID_HFOS, help="HFO entity")
    ap.add_argument("--run-date", help="ISO date (default: today)")
    args = ap.parse_args()
    summary = ingest(args.report, args.hfo, args.run_date)
    print(json.dumps(summary, indent=2))
