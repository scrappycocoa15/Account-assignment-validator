#!/usr/bin/env python3
"""
Account Assignment Validator
Streamlit Cloud compatible — only dependency is streamlit.
Run locally:  streamlit run app_v2.py

LinkedIn enrichment uses DuckDuckGo HTML search — no API key required.
"""

import streamlit as st
import urllib.request
import urllib.parse
import urllib.error
import json
import re
import time
import csv
import io
from html import escape as he

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Account Assignment Validator", layout="wide")

# ── SAP light-mode styling ────────────────────────────────────────────────────
st.markdown("""
<style>
* { font-family: '72','72full',Arial,Helvetica,sans-serif !important; }
section.main > div { padding-top:1rem; }
div[data-testid="stMetric"] {
    background:#F5F6F7; border:1px solid #E1E2E6;
    border-radius:8px; padding:.8rem 1rem;
}
div[data-testid="stMetric"] label { font-size:.78rem !important; color:#6A7275 !important; }
div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
    font-size:1.6rem !important; font-weight:700 !important; color:#0070F2 !important;
}
.stButton button {
    background:#0070F2 !important; color:white !important;
    border:none !important; border-radius:4px !important; font-weight:600 !important;
}
.stButton button:hover { background:#0057C2 !important; }
.stButton button:disabled { background:#BCC0C5 !important; }
.stDownloadButton button {
    background:white !important; color:#0070F2 !important;
    border:2px solid #0070F2 !important; border-radius:4px !important; font-weight:600 !important;
}
.stProgress > div > div { background-color:#0070F2 !important; }
.badge { display:inline-block; padding:.18rem .5rem; border-radius:4px;
         font-size:.78rem; font-weight:600; white-space:nowrap; }
.bc  { background:#E8F5E9; color:#188918; }
.bi  { background:#FFEBEE; color:#BB0000; }
.bli { background:#E1F4FF; color:#0057C2; }
.br  { background:#FFF3E0; color:#E76500; }
.bg  { background:#EAECEE; color:#6A7275; }
.rtable { width:100%; border-collapse:collapse; font-size:.84rem; }
.rtable th { background:#0070F2; color:white; padding:.6rem .9rem;
             text-align:left; font-weight:600; white-space:nowrap; }
.rtable td { padding:.5rem .9rem; border-bottom:1px solid #E1E2E6; vertical-align:middle; }
.rtable tr:nth-child(even) td { background:#FAFAFA; }
.rtable tr:hover td { background:#E1F4FF; }
.rtable a { color:#0070F2; text-decoration:none; font-weight:500; }
.rtable a:hover { text-decoration:underline; }
.preview-table { width:100%; border-collapse:collapse; font-size:.8rem; margin-top:.4rem; }
.preview-table th { background:#EAECEE; color:#32363A; padding:.35rem .7rem;
                    text-align:left; font-weight:600; }
.preview-table td { padding:.3rem .7rem; border-bottom:1px solid #E1E2E6; color:#555; }
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
SFDC_API_VERSION = "v59.0"
DNB_THRESHOLD    = 300
DNB_MIN_RELIABLE = 10          # D&B counts below this are treated as unreliable
DDG_SEARCH_URL   = "https://html.duckduckgo.com/html/"
DDG_DELAY        = 1.2         # seconds between DDG requests (be respectful)
DEFAULT_INSTANCE = "https://sapconcur.my.salesforce.com"
DEFAULT_REPORT   = "00OPg00000QbzTl"

BAND_TO_SEGMENT = {
    "1-10":         "General Business",
    "11-50":        "General Business",
    "51-200":       "General Business",
    "201-500":      "General Business",
    "501-1,000":    "US National",
    "1,001-5,000":  "US National",
    "5,001-10,000": "US National",
    "10,001+":      "US National",
}

BAND_PATTERNS = [
    (r"10[,.]?001\+?\s+employees?",                    "10,001+"),
    (r"5[,.]?001\s*[-\u2013]\s*10[,.]?000\s+employees?",   "5,001-10,000"),
    (r"1[,.]?001\s*[-\u2013]\s*5[,.]?000\s+employees?",    "1,001-5,000"),
    (r"501\s*[-\u2013]\s*1[,.]?000\s+employees?",           "501-1,000"),
    (r"201\s*[-\u2013]\s*500\s+employees?",                 "201-500"),
    (r"51\s*[-\u2013]\s*200\s+employees?",                  "51-200"),
    (r"11\s*[-\u2013]\s*50\s+employees?",                   "11-50"),
    (r"\b1\s*[-\u2013]\s*10\s+employees?",                  "1-10"),
]

# Hints ordered most-specific → least-specific
HINTS = {
    "account_name": ["account name", "company name", "name"],
    "segment":      ["us market segment", "market segment", "sales segment",
                     "account segment", "segment", "territory", "assignment"],
    "dnb":          ["d&b employees worldwide", "d&b employee worldwide",
                     "employees worldwide", "d&b employees", "dnb employees",
                     "employee worldwide", "d&b", "dnb", "dun",
                     "numberofemployees", "employees (d&b)"],
    "city":         ["billing city", "city"],
    "state":        ["billing state/province", "billing state", "state", "province"],
    "website":      ["account website", "website url", "website",
                     "web address", "web", "url", "domain"],
}

DATE_PAT = re.compile(
    r"^\d{1,2}/\d{1,2}/\d{4}$"      # MM/DD/YYYY
    r"|^\d{4}-\d{2}-\d{2}"           # YYYY-MM-DD
    r"|^\d{1,2}-\d{1,2}-\d{4}$"      # M-D-YYYY
)
ID_PAT = re.compile(r"\bid\b", re.IGNORECASE)

# ── Helper functions ──────────────────────────────────────────────────────────

def normalize_segment(s):
    if not s:
        return ""
    sl = str(s).lower().strip()
    if re.search(r"\bnational\b|\bnats?\b", sl):
        return "US National"
    if re.search(r"\bgeneral\s+business\b|\bgb\b|\bsmb\b|\bgen\s+bus\b", sl):
        return "General Business"
    return str(s).strip()

def parse_number(val):
    if val is None:
        return None
    sv = str(val).strip()
    if not sv or sv.lower() in ["none", "null", "n/a", "-", "\u2014", ""]:
        return None
    try:
        return int(float(re.sub(r"[,\s]", "", sv)))
    except Exception:
        return None

def extract_band(text):
    if not text:
        return None
    for pat, label in BAND_PATTERNS:
        if re.search(pat, text, re.IGNORECASE):
            return label
    return None

def detect_date_columns(cols, rows):
    """Return set of column labels whose first non-empty value looks like a date."""
    date_cols = set()
    for col in cols:
        for row in rows[:10]:
            val = str(row.get(col, "")).strip()
            if val and DATE_PAT.match(val):
                date_cols.add(col)
                break
    return date_cols

def auto_detect(cols, field, rows=None, required=True):
    """
    Match column labels against hints (most-specific first).
    Skips ID columns for account_name/segment.
    Skips date columns for account_name/segment (using actual cell values).
    """
    hints     = HINTS.get(field, [])
    date_cols = detect_date_columns(cols, rows) if rows else set()

    def should_skip(col):
        if field in ("account_name", "segment"):
            if ID_PAT.search(col):
                return True
            if col in date_cols:
                return True
        return False

    for h in hints:
        for col in cols:
            if should_skip(col):
                continue
            if h in col.lower():
                return col

    # Fallback: first non-skipped column
    if required and cols:
        for col in cols:
            if not should_skip(col):
                return col
        return cols[0]
    return "(not available)"

# ── Salesforce API ────────────────────────────────────────────────────────────

def fetch_sfdc_report(instance_url, session_id, report_id):
    url = (f"{instance_url.rstrip('/')}/services/data/{SFDC_API_VERSION}"
           f"/analytics/reports/{report_id}?includeDetails=true")
    req = urllib.request.Request(
        url,
        headers={"Authorization": f"Bearer {session_id}",
                 "Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read())
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="ignore")
        if e.code == 401:
            raise Exception("Invalid or expired Session ID. Refresh your Salesforce session.")
        if e.code == 403:
            raise Exception("Access denied. Check report permissions.")
        if e.code == 404:
            raise Exception(f"Report {report_id} not found.")
        raise Exception(f"Salesforce error {e.code}: {body[:200]}")
    except urllib.error.URLError as e:
        raise Exception(f"Cannot reach Salesforce: {e.reason}")

def parse_report(data):
    meta      = data.get("reportMetadata", {})
    api_names = meta.get("detailColumns", [])
    col_info  = (data.get("reportExtendedMetadata", {})
                     .get("detailColumnInfo", {}))
    labels    = [col_info.get(n, {}).get("label", n) for n in api_names]
    rows_raw  = data.get("factMap", {}).get("T!T", {}).get("rows", [])
    if not rows_raw:
        raise Exception(
            "No rows found. Ensure the report is Tabular format "
            "and contains data for the current year."
        )
    rows = []
    for row in rows_raw:
        cells = row.get("dataCells", [])
        rows.append({
            lbl: (cells[i].get("label", "") if i < len(cells) else "")
            for i, lbl in enumerate(labels)
        })
    return rows, labels

# ── DuckDuckGo LinkedIn search (no API key required) ─────────────────────────

def search_linkedin(name, city, state, website):
    """
    Search DuckDuckGo for the company's LinkedIn page.
    No API key required.
    Returns (linkedin_url, band, note).
    """
    # Build query
    parts = [f'site:linkedin.com/company "{name}"']
    if city and str(city).strip():
        parts.append(str(city).strip())
    if state and str(state).strip():
        parts.append(str(state).strip())
    if website:
        domain = re.sub(r"https?://(www\.)?", "", str(website)).split("/")[0].strip()
        if domain:
            parts.append(domain)

    query   = " ".join(parts)
    req_url = DDG_SEARCH_URL + "?" + urllib.parse.urlencode({"q": query})

    req = urllib.request.Request(
        req_url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            "Accept":          "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
        },
    )

    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            html = resp.read().decode("utf-8", errors="ignore")
    except Exception as e:
        return None, None, f"Search unavailable: {e}"

    # ── 1. Find LinkedIn company URL ──────────────────────────────────────────

    li_url = None

    # Strategy A: decode uddg= redirect parameters (DDG's redirect wrapper)
    for enc in re.findall(r'uddg=([^&"\'>\s]+)', html):
        try:
            dec = urllib.parse.unquote(enc)
            if "linkedin.com/company/" in dec:
                # Must not be a /jobs or /in/ sub-path
                m = re.search(
                    r"(https?://(?:www\.)?linkedin\.com/company/[^\s\"'&/?#]+)",
                    dec, re.IGNORECASE
                )
                if m:
                    candidate = m.group(1)
                    if "/jobs" not in candidate and "/in/" not in candidate:
                        li_url = candidate
                        break
        except Exception:
            continue

    # Strategy B: direct href attributes
    if not li_url:
        for m in re.finditer(
            r'href=["\']([^"\']*linkedin\.com/company/[^"\'?#\s]+)["\']',
            html, re.IGNORECASE
        ):
            candidate = m.group(1)
            if "/jobs" not in candidate and "/in/" not in candidate and "/school" not in candidate:
                li_url = candidate if candidate.startswith("http") else "https://" + candidate
                break

    # ── 2. Extract LinkedIn band from result snippets ─────────────────────────

    li_band = None

    # DDG result snippets appear in <a class="result__snippet">…</a>
    # or <div class="result__snippet">…</div>
    for raw_snippet in re.findall(
        r'class=["\']result__snippet["\'][^>]*>(.*?)</(?:a|div)>',
        html, re.DOTALL | re.IGNORECASE
    ):
        clean = re.sub(r"<[^>]+>", " ", raw_snippet).strip()
        band  = extract_band(clean)
        if band:
            li_band = band
            break

    # ── 3. Return ─────────────────────────────────────────────────────────────

    if li_url:
        return li_url, li_band, None

    # Fallback: return a LinkedIn company search URL the user can open manually
    fallback_url = (
        "https://www.linkedin.com/search/results/companies/?"
        + urllib.parse.urlencode({"keywords": name})
    )
    return fallback_url, None, "No exact LinkedIn match found — search link provided"

# ── Row processor ─────────────────────────────────────────────────────────────

def process_row(row, mapping, search_count):
    name        = str(row.get(mapping.get("account_name") or "", "")).strip()
    current_raw = str(row.get(mapping.get("segment")      or "", "")).strip()
    current_seg = normalize_segment(current_raw)
    dnb_raw     = str(row.get(mapping.get("dnb")          or "", "")).strip()
    dnb_val     = parse_number(dnb_raw)
    city        = str(row.get(mapping.get("city")    or "", "")).strip()
    state       = str(row.get(mapping.get("state")   or "", "")).strip()
    website     = str(row.get(mapping.get("website") or "", "")).strip()

    rec = {
        "account_name":     name,
        "current_segment":  current_raw,
        "dnb_employees":    dnb_raw or "\u2014",
        "expected_segment": "\u2014",
        "status":           "\u2014",
        "basis":            "\u2014",
        "linkedin_band":    "\u2014",
        "linkedin_url":     "",
    }
    did_search = False

    # Decide if D&B is usable
    dnb_usable = (dnb_val is not None) and (dnb_val >= DNB_MIN_RELIABLE)
    dnb_low    = (dnb_val is not None) and (dnb_val < DNB_MIN_RELIABLE)

    if dnb_usable:
        # ── Standard ROE path ─────────────────────────────────────────────────
        if dnb_val >= DNB_THRESHOLD:
            rec["expected_segment"] = "US National"
            rec["basis"]            = f"D\u0026B: {dnb_val:,} \u2265 {DNB_THRESHOLD}"
        else:
            rec["expected_segment"] = "General Business"
            rec["basis"]            = f"D\u0026B: {dnb_val:,} < {DNB_THRESHOLD}"
        rec["status"] = (
            "Correct"
            if current_seg == normalize_segment(rec["expected_segment"])
            else "Incorrect"
        )

    else:
        # ── No D&B or unreliable D&B → LinkedIn enrichment via DuckDuckGo ────
        if dnb_low:
            rec["basis"] = (
                f"D\u0026B count unreliable ({dnb_val} < {DNB_MIN_RELIABLE})"
                f" \u2014 using LinkedIn"
            )
        else:
            rec["basis"] = "No D\u0026B data \u2014 using LinkedIn"

        # Rate-limit DuckDuckGo requests
        if search_count > 0:
            time.sleep(DDG_DELAY)

        li_url, li_band, err = search_linkedin(name, city, state, website)
        did_search = True

        if li_url:
            rec["linkedin_url"] = li_url
            if li_band:
                expected               = BAND_TO_SEGMENT.get(li_band, "\u2014")
                rec["linkedin_band"]   = li_band
                rec["expected_segment"]= expected
                rec["basis"]          += f" | LinkedIn band: {li_band}"
                rec["status"] = (
                    "Correct (LinkedIn)"
                    if current_seg == normalize_segment(expected)
                    else "Incorrect (LinkedIn)"
                )
            else:
                rec["status"] = "Needs Review"
                note = err or "band not in snippet"
                rec["basis"] += f" | LinkedIn found \u2014 {note}"
        else:
            rec["status"] = "Data Gap"
            rec["basis"] += f" | {err or 'No LinkedIn match'}"

    return rec, did_search

# ── Rendering ─────────────────────────────────────────────────────────────────

def badge_html(status):
    cls = {
        "Correct":              "bc",
        "Incorrect":            "bi",
        "Correct (LinkedIn)":   "bli",
        "Incorrect (LinkedIn)": "bi",
        "Needs Review":         "br",
        "Data Gap":             "bg",
    }.get(status, "bg")
    return f'<span class="badge {cls}">{he(status)}</span>'

def render_table(rows):
    tbody = ""
    for r in rows:
        li_cell = (
            f'<a href="{he(r["linkedin_url"])}" target="_blank">Open &#8599;</a>'
            if r.get("linkedin_url") else "\u2014"
        )
        tbody += f"""<tr>
          <td>{he(r['account_name'])}</td>
          <td>{he(r['current_segment'])}</td>
          <td>{he(str(r['dnb_employees']))}</td>
          <td>{he(r['expected_segment'])}</td>
          <td>{badge_html(r['status'])}</td>
          <td>{he(r['basis'])}</td>
          <td>{he(r['linkedin_band'])}</td>
          <td>{li_cell}</td>
        </tr>"""
    st.markdown(f"""
<div style="overflow-x:auto">
<table class="rtable">
  <thead><tr>
    <th>Account Name</th><th>Current Segment</th><th>D&amp;B Employees</th>
    <th>Expected Segment</th><th>Status</th><th>Basis</th>
    <th>LinkedIn Band</th><th>LinkedIn Page</th>
  </tr></thead>
  <tbody>{tbody}</tbody>
</table></div>""", unsafe_allow_html=True)

def col_preview_html(cols, rows, mapping):
    """Small preview table showing the first 3 values of each mapped column."""
    fields = [
        ("Account Name",    mapping.get("account_name")),
        ("Current Segment", mapping.get("segment")),
        ("D&B Employees",   mapping.get("dnb")),
        ("City",            mapping.get("city")),
        ("State",           mapping.get("state")),
        ("Website",         mapping.get("website")),
    ]
    sample_rows = rows[:3]
    thead   = "".join(f"<th>{he(f)}</th>" for f, c in fields if c)
    tbodies = []
    for sr in sample_rows:
        cells = "".join(
            f"<td>{he(str(sr.get(c, '')))}</td>"
            for f, c in fields if c
        )
        tbodies.append(f"<tr>{cells}</tr>")
    return f"""
<details open>
  <summary style="cursor:pointer;font-size:.82rem;font-weight:600;
                  color:#0070F2;margin-bottom:.4rem">
    Column preview (first 3 rows) \u2014 verify the mapping is correct before validating
  </summary>
  <div style="overflow-x:auto">
  <table class="preview-table">
    <thead><tr>{thead}</tr></thead>
    <tbody>{''.join(tbodies)}</tbody>
  </table></div>
</details>"""

# ── App layout ────────────────────────────────────────────────────────────────

st.markdown("""
<div style="background:#0070F2;color:white;padding:1.2rem 1.6rem;
            border-radius:8px;margin-bottom:1rem">
  <div style="font-size:1.4rem;font-weight:700;color:white">
    Account Assignment Validator
  </div>
  <div style="font-size:.88rem;opacity:.85;margin-top:.25rem">
    US SMB &middot; ROE validation via D&amp;B employee count
    and LinkedIn band enrichment
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div style="background:#E1F4FF;border:1px solid #4CB1FF;border-radius:6px;
            padding:.7rem 1rem;font-size:.84rem;margin-bottom:1.2rem">
  <b>ROE:</b>
  D&amp;B &ge; 300 &rarr; <b>US National</b> &nbsp;|&nbsp;
  D&amp;B &lt; 300 &rarr; <b>General Business</b> &nbsp;|&nbsp;
  D&amp;B &lt; 10 (unreliable) or missing &rarr; LinkedIn band above 201&ndash;500
  &rarr; <b>US National</b>, rest &rarr; <b>General Business</b>
  &nbsp;&nbsp;<span style="color:#6A7275">
  &bull; LinkedIn enrichment uses DuckDuckGo search &mdash; no API key required
  </span>
</div>
""", unsafe_allow_html=True)

# ── Step 1: Connect ───────────────────────────────────────────────────────────
st.markdown("**1. Connect to Salesforce**")

c1, c2 = st.columns(2)
with c1:
    instance_url = st.text_input(
        "Salesforce Instance URL",
        value=DEFAULT_INSTANCE,
        key="k_instance",
    )
    session_id = st.text_input(
        "Session ID",
        type="password",
        placeholder="Paste your Salesforce Session ID here",
        key="k_sid",
    )
with c2:
    report_id = st.text_input(
        "Report ID",
        value=DEFAULT_REPORT,
        key="k_rid",
    )

if st.button("Load Report", key="btn_load"):
    if not session_id:
        st.error("Paste your Salesforce Session ID to continue.")
    else:
        with st.spinner("Connecting to Salesforce..."):
            try:
                raw_data         = fetch_sfdc_report(instance_url, session_id, report_id)
                rows, col_labels = parse_report(raw_data)
                report_name      = raw_data.get("reportMetadata", {}).get("name", "Report")
                st.session_state["rows"]        = rows
                st.session_state["col_labels"]  = col_labels
                st.session_state["report_name"] = report_name
                st.session_state["results"]     = None
                st.success(f"Loaded **{report_name}** \u2014 {len(rows):,} accounts.")
            except Exception as e:
                st.error(str(e))

# ── Step 2: Column mapping ────────────────────────────────────────────────────
if st.session_state.get("rows"):
    rows       = st.session_state["rows"]
    col_labels = st.session_state["col_labels"]
    opt        = ["(not available)"] + col_labels

    st.markdown("**2. Map Report Columns**")
    st.caption(
        "Auto-detected from column names and cell values. "
        "Check the preview below and adjust dropdowns if needed."
    )

    def _idx(cols, val):
        return cols.index(val) if val in cols else 0

    c1, c2, c3 = st.columns(3)
    c4, c5, c6 = st.columns(3)

    with c1:
        acct_col = st.selectbox(
            "Account Name \u2733", col_labels,
            index=_idx(col_labels, auto_detect(col_labels, "account_name", rows)),
        )
    with c2:
        seg_col = st.selectbox(
            "Current Segment \u2733", col_labels,
            index=_idx(col_labels, auto_detect(col_labels, "segment", rows)),
        )
    with c3:
        dnb_col = st.selectbox(
            "D\u0026B Employee Worldwide \u2733", col_labels,
            index=_idx(col_labels, auto_detect(col_labels, "dnb", rows)),
        )
    with c4:
        city_col = st.selectbox(
            "Billing City", opt,
            index=_idx(opt, auto_detect(col_labels, "city", rows, required=False)),
        )
    with c5:
        state_col = st.selectbox(
            "Billing State", opt,
            index=_idx(opt, auto_detect(col_labels, "state", rows, required=False)),
        )
    with c6:
        web_col = st.selectbox(
            "Website", opt,
            index=_idx(opt, auto_detect(col_labels, "website", rows, required=False)),
        )

    mapping = {
        "account_name": acct_col,
        "segment":      seg_col,
        "dnb":          dnb_col,
        "city":         None if city_col  == "(not available)" else city_col,
        "state":        None if state_col == "(not available)" else state_col,
        "website":      None if web_col   == "(not available)" else web_col,
    }

    # Column preview
    st.markdown(col_preview_html(col_labels, rows, mapping), unsafe_allow_html=True)

    # Duplicate-column guard
    if acct_col == seg_col:
        st.error(
            f"**Account Name** and **Current Segment** are both mapped to "
            f"**{acct_col}**. Adjust the dropdowns \u2014 they must point to "
            f"different columns."
        )

    if st.button(
        "Validate Assignments",
        key="btn_validate",
        disabled=(acct_col == seg_col),
    ):
        results      = []
        search_count = 0
        total        = len(rows)
        prog         = st.progress(0.0, text="Starting\u2026")

        for i, row in enumerate(rows):
            name = str(row.get(acct_col, "")).strip()
            prog.progress(
                (i + 1) / total,
                text=f"Processing {i + 1} of {total}\u2003\u2014\u2003{name[:70]}",
            )
            result, did_search = process_row(row, mapping, search_count)
            if did_search:
                search_count += 1
            results.append(result)

        prog.empty()
        st.session_state["results"] = results

# ── Step 3: Results ───────────────────────────────────────────────────────────
if st.session_state.get("results"):
    results = st.session_state["results"]
    st.markdown("**3. Validation Results**")

    total   = len(results)
    correct = sum(1 for r in results if "Correct"   in r["status"])
    wrong   = sum(1 for r in results if "Incorrect" in r["status"])
    gap     = sum(1 for r in results if r["status"] in ("Data Gap", "Needs Review"))
    pct     = lambda n: f"{round(n / total * 100)}%" if total else "0%"

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Accounts",       total)
    m2.metric("Correctly Assigned",   f"{correct} ({pct(correct)})")
    m3.metric("Incorrectly Assigned", f"{wrong} ({pct(wrong)})")
    m4.metric("Data Gap / Review",    f"{gap} ({pct(gap)})")

    st.write("")
    all_statuses = ["All"] + sorted({r["status"] for r in results})
    sel          = st.selectbox("Filter by Status", all_statuses, key="k_filter")
    filtered     = results if sel == "All" else [r for r in results if r["status"] == sel]
    st.caption(f"Showing {len(filtered):,} of {total:,} accounts")

    render_table(filtered)

    st.write("")
    buf    = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=[
        "account_name", "current_segment", "dnb_employees",
        "expected_segment", "status", "basis", "linkedin_band", "linkedin_url",
    ])
    writer.writeheader()
    writer.writerows(results)
    st.download_button(
        "Download Results as CSV",
        buf.getvalue(),
        "account_assignment_validation.csv",
        "text/csv",
        key="btn_csv",
    )
