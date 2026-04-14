"""
Regression test for Documentation Quality scoring.
Locks in the v12.8 fixes so future changes can't silently regress these cases.

Run with:
    python3 test_regression.py /path/to/Python\ WO\ Approval\ Report.xlsx

Each test case records:
  - the WO number
  - the expected (score_100, grade, verdict)
  - why this case matters (the defect it was originally meant to catch)

Tolerances: score within +/- 5; grade and verdict must match exactly.
"""
import sys, os, pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from audit_engine import dq_gate, dq_clean

# ── Locked-in cases (v12.8) ──────────────────────────────────────────────────
# Format: wo_number: (expected_score, expected_grade, expected_verdict, rationale)
CASES = {
    '402898': (85, 'B', 'Pass',
        "Cross-machine drift: CA ends with 'the other umc will need...'. "
        "Must cap at 85 and flag drift, not score 100/A."),
    '402787': (100, 'A', 'Pass',
        "Customer-directed hold (Ginn, Bristol Precision): customer requested "
        "machine left disassembled pending replace-vs-retire decision. Must "
        "count as valid closure, not missing closure."),
    '402421': (95, 'A', 'Pass',
        "Customer will install parts themselves ('Customer will order and "
        "replace coupling'). Must count as valid closure."),
    '396315': (95, 'A', 'Pass',
        "Has standalone closure ('returned the machine to service'). "
        "Should NOT rely on the tightened 'per customer request' pattern "
        "(that match was removed for being too broad)."),
    '400651': (95, 'A', 'Pass',
        "'Customer will monitor for signs of pump cavitation' — legitimate "
        "customer-takeover hold. Plus 'machine was returned to service'."),
    '401623': (95, 'A', 'Pass',
        "'Customer will continue to monitor and report if issue persists' "
        "after software revert resolved the issue."),
}

SCORE_TOLERANCE = 5


def run(xlsx_path):
    df = pd.read_excel(xlsx_path, sheet_name='Work Order Corrective action')
    results = []
    for wo, (exp_score, exp_grade, exp_verdict, rationale) in CASES.items():
        match = df[df['WorkOrderNumber'].astype(str).str.contains(wo)]
        if match.empty:
            results.append((wo, 'MISSING', f"WO {wo} not found in workbook"))
            continue
        r = match.iloc[0]
        v, detail, sd = dq_gate(r['Cause__c'] or '', r['Corrective_Action__c'] or '')
        actual_score = sd.get('score_100', 0)
        actual_grade = sd.get('grade', 'F')
        score_ok = abs(actual_score - exp_score) <= SCORE_TOLERANCE
        grade_ok = actual_grade == exp_grade
        verdict_ok = v == exp_verdict
        passed = score_ok and grade_ok and verdict_ok
        status = 'PASS' if passed else 'FAIL'
        msg = (f"exp {exp_score}/{exp_grade}/{exp_verdict} | "
               f"got {actual_score}/{actual_grade}/{v}")
        if not passed:
            msg += f"\n    rationale: {rationale}"
            msg += f"\n    detail: {detail}"
        results.append((wo, status, msg))

    print(f"{'='*72}\nDoc Quality regression — {len(CASES)} cases\n{'='*72}")
    n_pass = sum(1 for _, s, _ in results if s == 'PASS')
    for wo, status, msg in results:
        mark = '[OK]  ' if status == 'PASS' else '[FAIL]'
        print(f"{mark} 00{wo}: {msg}")
    print(f"{'='*72}\n{n_pass}/{len(CASES)} cases passed")
    return n_pass == len(CASES)


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Usage: python3 test_regression.py <path-to-WO-workbook.xlsx>")
        sys.exit(2)
    sys.exit(0 if run(sys.argv[1]) else 1)
