"""
WSR capture for the Sterling on-time WO dashboard (lab flow).

Every run:
  1. Download the latest WSR export (WSR_URL) or read a local file (WSR_FILE).
  2. Skip if the file is identical to the last one processed.
  3. Compare each test against stored state and record, the first time we see them:
       - Analyzed time  (StatusDate when the test is seen at Analyzed)
       - Reviewed time  (StatusDate when the test is seen at Reviewed)
       - every WODueDate value a WO has had (each one becomes its own record)
  4. Tests that disappear from the report:
       - last seen at Analyzed or later  -> lab-complete at the time of the
         first pull where they are missing (conservative)
       - last seen earlier               -> flagged "dropped unexpectedly"
  5. Write data/state.json (memory between runs) and data/wo_records.json
     (what the dashboard reads), plus data/meta.json and data/pulls.csv.

Rules (agreed):
  - Clock = WODueDate. Lab analysis due 3:30 PM, data validation due 4:00 PM
    (cutoffs are applied by the dashboard, not here).
  - A WO is Analyzed / Reviewed when its LAST active test is.
  - Cancelled, Hold and Subcontracted tests (or tests with a SubLab) are skipped.
  - A test seen at Reviewed without an Analyzed time gets Analyzed = Reviewed
    time (an upper bound), flagged approximate.
  - Drop-off times from a gap of more than 30 minutes between pulls are
    flagged approximate.
"""

from __future__ import annotations

import csv
import hashlib
import io
import json
import os
import re
import sys
from datetime import datetime, timedelta
from email.utils import parsedate_to_datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd

LOCAL_TZ = ZoneInfo("America/Chicago")
DATA_DIR = Path(os.environ.get("DATA_DIR", "data"))
GAP_LIMIT = timedelta(minutes=30)      # drop-off gaps longer than this are approximate
PRUNE_AFTER = timedelta(days=21)       # tests gone this long are frozen and removed
PULL_LOG_MAX = 3000                    # rows kept in pulls.csv

REQUIRED = ["Wrk", "LabNumber", "AnalysisName", "RECode", "Status",
            "StatusDate", "WODueDate", "WrkPriority", "Client"]
OPTIONAL = ["DepartmentName", "LabManager", "SubLab"]

# Lab progression. Anything at or past ANALYZED counts as analyzed;
# at or past REVIEWED counts as data-validated.
RANK = {"Pending": 0, "Received": 1, "Available": 2, "Leached": 3,
        "Batched": 4, "Prepared": 5, "Analyzed": 6, "Reviewed": 7,
        "Peer Reviewed": 8, "Reported": 9}
ANALYZED, REVIEWED = RANK["Analyzed"], RANK["Reviewed"]
EXCLUDED_STATUS = {"Cancelled", "Hold", "Subcontracted"}


# ---------------------------------------------------------------- helpers
def iso(ts) -> str | None:
    if ts is None or pd.isna(ts):
        return None
    return pd.Timestamp(ts).to_pydatetime().replace(microsecond=0).isoformat()


def parse(s: str | None) -> datetime | None:
    return datetime.fromisoformat(s) if s else None


def log(msg: str) -> None:
    print(msg, flush=True)


def download_url(url: str) -> str:
    """Turn any Google Drive / Google Sheets share link into a direct download URL.
    Other links (e.g. SharePoint) are returned unchanged."""
    m = re.search(r"/d/([\w-]{20,})", url) or re.search(r"[?&]id=([\w-]{20,})", url)
    if "docs.google.com/spreadsheets" in url and m:
        # The file was converted to a Google Sheet: export it as .xlsx
        return f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=xlsx"
    if ("drive.google.com" in url or "drive.usercontent.google.com" in url) and m:
        return (f"https://drive.usercontent.google.com/download?id={m.group(1)}"
                "&export=download&confirm=t")
    return url


# ---------------------------------------------------------------- input
def fetch_pull() -> tuple[bytes, datetime, str]:
    """Return (file bytes, pull time in local time, how the time was found)."""
    override = os.environ.get("PULL_TIME")          # used for testing
    local_file = os.environ.get("WSR_FILE")
    if local_file:
        content = Path(local_file).read_bytes()
        if override:
            return content, datetime.fromisoformat(override), "override"
        return content, datetime.now(LOCAL_TZ).replace(tzinfo=None, microsecond=0), "clock"

    import requests
    url = os.environ.get("WSR_URL", "").strip()
    if not url:
        sys.exit("WSR_URL is not set (add it under Settings > Secrets and variables > Actions).")
    url = download_url(url)
    resp = requests.get(url, timeout=120, allow_redirects=True)
    resp.raise_for_status()
    content = resp.content
    if not content.startswith(b"PK"):
        text = content[:4000].decode("utf-8", "ignore")
        title = re.search(r"<title>(.*?)</title>", text, re.I | re.S)
        hint = title.group(1).strip() if title else text[:200].strip()
        sys.exit("Download did not return an .xlsx file.\n"
                 f"  Got back: {hint!r}\n"
                 "  Check that the file is shared as 'Anyone with the link' (Viewer) and that "
                 "the link opens without sign-in in a private browser window.")
    if override:
        return content, datetime.fromisoformat(override), "override"
    lm = resp.headers.get("Last-Modified")
    if lm:
        t = parsedate_to_datetime(lm).astimezone(LOCAL_TZ).replace(tzinfo=None, microsecond=0)
        return content, t, "last-modified"
    return content, datetime.now(LOCAL_TZ).replace(tzinfo=None, microsecond=0), "clock"


def read_wsr(content: bytes) -> pd.DataFrame:
    header = pd.read_excel(io.BytesIO(content), nrows=0).columns
    missing = [c for c in REQUIRED if c not in header]
    if missing:
        sys.exit(f"WSR export is missing required columns: {missing}")
    cols = REQUIRED + [c for c in OPTIONAL if c in header]
    df = pd.read_excel(io.BytesIO(content), usecols=cols)
    df = df[df["Wrk"].notna() & df["LabNumber"].notna()].copy()
    for c in ("StatusDate", "WODueDate"):
        df[c] = pd.to_datetime(df[c], errors="coerce")
    for c in OPTIONAL:
        if c not in df.columns:
            df[c] = None
    df["RECode"] = df["RECode"].fillna("I")
    df["key"] = (df["LabNumber"].astype(str) + "|" + df["AnalysisName"].astype(str)
                 + "|" + df["RECode"].astype(str))
    dups = df["key"].duplicated(keep="last")
    if dups.any():
        log(f"Note: {int(dups.sum())} duplicate test rows; keeping the last of each.")
        df = df[~dups]
    return df


# ---------------------------------------------------------------- state
def load_state() -> dict:
    p = DATA_DIR / "state.json"
    if p.exists():
        return json.loads(p.read_text())
    return {"version": 1, "capture_start": None, "last_hash": None,
            "last_pull": None, "tests": {}, "wos": {}}


def update_test(t: dict, row, pull_s: str) -> None:
    status = row.Status
    sd = iso(row.StatusDate)
    rank = RANK.get(status)
    t.update(w=row.Wrk, s=status, sd=sd, last=pull_s,
             dept=None if pd.isna(row.DepartmentName) else row.DepartmentName,
             sub=bool(status == "Subcontracted" or (row.SubLab is not None and not pd.isna(row.SubLab))))
    t.setdefault("seen", pull_s)
    if t.get("gone"):                      # came back onto the report
        t["gone"] = None
        t["drop"] = None
        for k in ("an", "rv"):
            if t.get(k + "_src") == "dropoff":
                t[k] = t[k + "_src"] = None
                t[k + "_approx"] = False

    if rank is None:                       # Cancelled / Hold / Subcontracted / unknown
        return

    # Status moved backwards (e.g. review rejected) -> completion reopens.
    if t.get("rv") and rank < REVIEWED:
        t["rv"] = t["rv_src"] = None
        t["rv_approx"] = False
        t["reopened"] = pull_s
        if status == "Analyzed" and sd:    # re-analyzed: the later analysis counts
            t.update(an=sd, an_src="status", an_approx=False)
    if t.get("an") and rank < ANALYZED:
        t["an"] = t["an_src"] = None
        t["an_approx"] = False
        t["reopened"] = pull_s

    if status == "Analyzed" and not t.get("an") and sd:
        t.update(an=sd, an_src="status", an_approx=False)
    if rank >= REVIEWED and not t.get("rv") and sd:
        exact = status == "Reviewed"
        t.update(rv=sd, rv_src="status" if exact else "upper", rv_approx=not exact)
    if rank >= ANALYZED and not t.get("an"):
        if t.get("rv"):                    # never saw Analyzed: bounded by Reviewed time
            t.update(an=t["rv"], an_src="upper", an_approx=True)


def mark_gone(t: dict, pull: datetime, pull_s: str) -> None:
    t["gone"] = pull_s
    status = t.get("s")
    if status in EXCLUDED_STATUS or t.get("sub"):
        t["drop"] = "excluded"
        return
    rank = RANK.get(status, -1)
    if rank < ANALYZED:
        t["drop"] = "unexpected"
        return
    approx = (pull - parse(t["last"])) > GAP_LIMIT
    if not t.get("an"):
        t.update(an=pull_s, an_src="dropoff", an_approx=approx)
    if not t.get("rv"):
        t.update(rv=pull_s, rv_src="dropoff", rv_approx=approx)
    t["drop"] = "done"


# ---------------------------------------------------------------- records
def is_active(t: dict) -> bool:
    return not (t.get("s") in EXCLUDED_STATUS or t.get("sub")
                or t.get("drop") in ("excluded", "unexpected"))


def build_records(wrk: str, wo: dict, tests: list[dict], capture_start: str) -> list[dict]:
    active = [t for t in tests if is_active(t)]
    unexpected = sum(1 for t in tests if t.get("drop") == "unexpected")

    def finish(field: str):
        if not active or any(not t.get(field) for t in active):
            return None, False, sum(1 for t in active if not t.get(field))
        last = max(active, key=lambda t: t[field])
        approx = any(t.get(field + "_approx") for t in active if t[field] == last[field])
        return last[field], approx, 0

    an, an_approx, an_open = finish("an")
    rv, rv_approx, rv_open = finish("rv")
    depts = sorted({t["dept"] for t in active if t.get("dept")})
    start_date = capture_start[:10]
    dues = wo["dues"]
    out = []
    for i, d in enumerate(dues):
        out.append({
            "wo": wrk,
            "due": d["due"],
            "due_seq": i + 1,
            "due_count": len(dues),
            "superseded": i < len(dues) - 1,
            "due_first_seen": d["first_seen"],
            "priority": wo.get("prio"),
            "client": wo.get("client"),
            "csp": wo.get("lm"),
            "depts": depts,
            "tests": len(active),
            "analyzed_at": an, "analyzed_approx": an_approx, "open_analyzed": an_open,
            "reviewed_at": rv, "reviewed_approx": rv_approx, "open_reviewed": rv_open,
            "unexpected_drops": unexpected,
            "pre_capture": d["due"] <= start_date,
        })
    return out


# ---------------------------------------------------------------- main
def main() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    state = load_state()
    content, pull, pull_src = fetch_pull()
    digest = hashlib.sha256(content).hexdigest()
    if digest == state.get("last_hash"):
        log("WSR file unchanged since last run; nothing to do.")
        return
    if state.get("last_pull") and pull <= parse(state["last_pull"]):
        log(f"Pull time {pull} is not newer than last pull {state['last_pull']}; skipping.")
        return

    df = read_wsr(content)
    pull_s = pull.isoformat()
    if not state["capture_start"]:
        state["capture_start"] = pull_s
    tests, wos = state["tests"], state["wos"]

    # WO-level facts and due-date history
    first = df.drop_duplicates("Wrk")
    for row in first.itertuples(index=False):
        wo = wos.setdefault(row.Wrk, {"dues": []})
        wo.update(prio=None if pd.isna(row.WrkPriority) else str(row.WrkPriority),
                  client=row.Client,
                  lm=None if pd.isna(row.LabManager) else row.LabManager,
                  last=pull_s)
        if pd.notna(row.WODueDate):
            due = row.WODueDate.date().isoformat()
            if due not in [d["due"] for d in wo["dues"]]:
                wo["dues"].append({"due": due, "first_seen": pull_s})

    # Tests on this pull
    present = set()
    for row in df.itertuples(index=False):
        present.add(row.key)
        update_test(tests.setdefault(row.key, {}), row, pull_s)

    # Tests that dropped off
    dropped = 0
    for key, t in tests.items():
        if key not in present and not t.get("gone"):
            mark_gone(t, pull, pull_s)
            dropped += 1

    # Freeze and prune WOs whose tests have all been gone a long time
    keys_by_wo: dict[str, list[str]] = {}
    for key, t in tests.items():
        keys_by_wo.setdefault(t["w"], []).append(key)
    cutoff = pull - PRUNE_AFTER
    for wrk, keys in keys_by_wo.items():
        ts = [tests[k] for k in keys]
        if wrk in wos and all(t.get("gone") and parse(t["gone"]) < cutoff for t in ts):
            wos[wrk]["frozen"] = build_records(wrk, wos[wrk], ts, state["capture_start"])
            for k in keys:
                del tests[k]
    by_wo = {}
    for t in tests.values():
        by_wo.setdefault(t["w"], []).append(t)

    records = []
    for wrk, wo in wos.items():
        if "frozen" in wo:
            records.extend(wo["frozen"])
        elif wo["dues"]:
            records.extend(build_records(wrk, wo, by_wo.get(wrk, []), state["capture_start"]))
    records.sort(key=lambda r: (r["due"], r["wo"], r["due_seq"]))

    state.update(last_hash=digest, last_pull=pull_s)
    (DATA_DIR / "state.json").write_text(json.dumps(state, separators=(",", ":")))
    (DATA_DIR / "wo_records.json").write_text(json.dumps(records, separators=(",", ":")))
    (DATA_DIR / "meta.json").write_text(json.dumps({
        "last_pull": pull_s, "pull_time_source": pull_src,
        "capture_start": state["capture_start"], "records": len(records),
        "rules": {"analyzed_cutoff": "15:30", "reviewed_cutoff": "16:00",
                  "approx_gap_minutes": int(GAP_LIMIT.total_seconds() // 60)},
    }, indent=2))

    log_path = DATA_DIR / "pulls.csv"
    rows = list(csv.reader(log_path.open())) if log_path.exists() else []
    if not rows:
        rows = [["pull_time", "source", "test_rows", "wos", "dropped", "hash"]]
    rows.append([pull_s, pull_src, len(df), df["Wrk"].nunique(), dropped, digest[:12]])
    rows = [rows[0]] + rows[1:][-PULL_LOG_MAX:]
    with log_path.open("w", newline="") as f:
        csv.writer(f).writerows(rows)

    (DATA_DIR / ".changed").write_text(pull_s)
    log(f"Pull {pull_s} ({pull_src}): {len(df)} tests, {df['Wrk'].nunique()} WOs, "
        f"{dropped} dropped, {len(records)} WO records written.")


if __name__ == "__main__":
    main()
