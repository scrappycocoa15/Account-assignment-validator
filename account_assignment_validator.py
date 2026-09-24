#!/usr/bin/env python3
"""
Account Assignment Validator — Open Opportunities
Validates ROE segment assignment for accounts with open opportunities.
Streamlit Cloud compatible. Dependencies: streamlit, openpyxl.
"""

import streamlit as st
import streamlit.components.v1 as stc
import urllib.request
import urllib.parse
import urllib.error
import json
import re
import csv
import io
import base64
from pathlib import Path
from html import escape as he

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Opportunity Assignment Validator", layout="wide")

# ── SAP light-mode styling ────────────────────────────────────────────────────
st.markdown("""
<style>
* { font-family: '72','72full',Arial,Helvetica,sans-serif !important; }
section.main > div { padding-top:1rem; }
div[data-testid="stMetric"] {
    background:#F5F6F7; border:1px solid #E1E2E6;
    border-radius:8px; padding:.8rem 1rem;
}
div[data-testid="stMetric"] label {
    font-size:.78rem !important; color:#6A7275 !important;
}
div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
    font-size:1.6rem !important; font-weight:700 !important; color:#0070F2 !important;
}
.stButton button {
    background:#0070F2 !important; color:white !important;
    border:none !important; border-radius:4px !important; font-weight:600 !important;
}
.stButton button:hover  { background:#0057C2 !important; }
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
DEFAULT_INSTANCE = "https://sapconcur.my.salesforce.com"
DEFAULT_REPORT   = "00OPg00000Qgik1"

_HERE              = Path(__file__).parent
GB_TERRITORY_FILE  = _HERE / "2026-01-01 US General Business Territories.xlsx"
NAT_TERRITORY_FILE = _HERE / "2026-01-01 US National Territories.xlsx"
GB_SHEET_NAME      = "2026 GB Zip Assignments"
NAT_SHEET_NAME     = "2026 Nat Zip Assignments"

# ── Column hints (most-specific → least-specific) ─────────────────────────────
HINTS = {
    "account_name": ["account name", "company name", "account"],
    "account_id":   ["account id", "account: id", "18 digit", "sfdc id", "acct id"],
    "segment":      ["us market segment", "market segment", "sales segment",
                     "account segment", "owner division", "acct owner division",
                     "acct division", "segment", "territory", "division", "assignment"],
    "dnb":          ["d&b employees worldwide", "d&b employee worldwide",
                     "employees worldwide", "d&b employees", "dnb employees",
                     "employee worldwide", "d&b", "dnb", "dun",
                     "numberofemployees", "employees (d&b)"],
    "city":         ["billing city", "city"],
    "state":        ["billing state/province", "billing state", "state", "province"],
    "zip":          ["billing zip/postal", "billing zip", "zip/postal",
                     "postal code", "zip code", "zip", "postal"],
    "website":      ["account website", "website url", "website",
                     "web address", "web", "url", "domain"],
    "opp_name":     ["opportunity name", "opp name", "opportunity"],
    "stage":        ["stage name", "opportunity stage", "stage"],
    "close_date":   ["close date", "close"],
    "amount":       ["amount", "arr", "acv", "value"],
}

DATE_PAT = re.compile(
    r"^\d{1,2}/\d{1,2}/\d{4}$"
    r"|^\d{4}-\d{2}-\d{2}"
    r"|^\d{1,2}-\d{1,2}-\d{4}$"
)
ID_PAT = re.compile(r"\bid\b", re.IGNORECASE)

# ── Core helpers ──────────────────────────────────────────────────────────────

def normalize_segment(s):
    if not s: return ""
    sl = str(s).lower().strip()
    if re.search(r"\bnational\b|\bnats?\b", sl):
        return "US National"
    if re.search(r"\bgeneral\s+business\b|\bsmall\s+business\b|\bgb\b|\bsmb\b|\bgen\s+bus\b", sl):
        return "General Business"
    return str(s).strip()

def parse_number(val):
    if val is None: return None
    sv = str(val).strip()
    if not sv or sv.lower() in ["none", "null", "n/a", "-", "\u2014", ""]: return None
    try:
        return int(float(re.sub(r"[,\s]", "", sv)))
    except Exception:
        return None

def detect_date_columns(cols, rows):
    date_cols = set()
    for col in cols:
        for row in rows[:10]:
            val = str(row.get(col, "")).strip()
            if val and DATE_PAT.match(val):
                date_cols.add(col)
                break
    return date_cols

def auto_detect(cols, field, rows=None, required=True):
    hints     = HINTS.get(field, [])
    date_cols = detect_date_columns(cols, rows) if rows else set()

    def should_skip(col):
        if field in ("account_name", "segment"):
            if ID_PAT.search(col): return True
            if col in date_cols:   return True
        return False

    for h in hints:
        for col in cols:
            if should_skip(col): continue
            if h in col.lower():
                return col

    if required and cols:
        for col in cols:
            if not should_skip(col): return col
        return cols[0]
    return "(not available)"

# ── Salesforce ────────────────────────────────────────────────────────────────

def fetch_sfdc_report(instance_url, session_id, report_id):
    url = (f"{instance_url.rstrip('/')}/services/data/{SFDC_API_VERSION}"
           f"/analytics/reports/{report_id}?includeDetails=true")
    req = urllib.request.Request(
        url,
        headers={"Authorization": f"Bearer {session_id}",
                 "Content-Type":  "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read())
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="ignore")
        if e.code == 401: raise Exception("Invalid or expired Session ID.")
        if e.code == 403: raise Exception("Access denied. Check report permissions.")
        if e.code == 404: raise Exception(f"Report {report_id} not found.")
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
            "No rows found. Ensure the report is Tabular format and contains data."
        )
    rows = []
    for row in rows_raw:
        cells = row.get("dataCells", [])
        rows.append({
            lbl: (cells[i].get("label", "") if i < len(cells) else "")
            for i, lbl in enumerate(labels)
        })
    return rows, labels

# ── Territory helpers ─────────────────────────────────────────────────────────

def normalize_zip(val):
    if val is None: return ""
    z = re.sub(r"[^0-9]", "", str(val).strip().split("-")[0])[:5]
    return z.zfill(5) if z else ""

def _parse_territory_wb(wb, sheet_name):
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else (
        next((wb[s] for s in wb.sheetnames if "zip" in s.lower()), wb.active)
    )
    rows = list(ws.iter_rows(values_only=True))
    if not rows: return {}, "empty sheet"
    hdr_raw = [str(h).strip() if h is not None else "" for h in rows[0]]
    hdr     = [h.lower() for h in hdr_raw]

    def _col(*needles):
        for needle in needles:
            for i, h in enumerate(hdr):
                if needle in h: return i
        return None

    zip_idx  = _col("zip code", "zip", "postal")

    # Owner name: prefer columns whose header ends with "owner" or "rep name",
    # explicitly skip any column whose header ends with "id"
    def _col_name_only(*needles):
        for needle in needles:
            for i, h in enumerate(hdr):
                if needle in h and not h.rstrip().endswith("id"):
                    return i
        return None

    own_idx  = _col_name_only("fy26 account owner", "account owner name",
                               "account owner", "rep name", "rep")
    id_idx   = _col("fy26 account owner id", "account owner id",
                    "owner id", "user id", "rep id", "sfdc id", "salesforce id")
    terr_idx = _col("territory name", "territory")

    if zip_idx is None or own_idx is None:
        return {}, f"could not find zip/owner columns in {hdr_raw}"

    result = {}
    for row in rows[1:]:
        raw_zip = row[zip_idx] if zip_idx < len(row) else None
        owner   = str(row[own_idx]).strip() if own_idx < len(row) and row[own_idx] else ""
        z       = normalize_zip(raw_zip)
        if not z or not owner or owner.lower() in ("none", "nan", ""): continue
        entry = {"owner": owner}
        if id_idx  is not None and id_idx  < len(row) and row[id_idx]:
            entry["owner_id"]  = str(row[id_idx]).strip()
        if terr_idx is not None and terr_idx < len(row) and row[terr_idx]:
            entry["territory"] = str(row[terr_idx]).strip()
        result[z] = entry
    return result, None

def load_territory_path(path, sheet_name):
    try:
        import openpyxl
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
        d, e = _parse_territory_wb(wb, sheet_name)
        wb.close()
        return d, e
    except Exception as ex:
        return {}, str(ex)

def load_territory_upload(uploaded_file, sheet_name):
    try:
        import openpyxl
        if uploaded_file.name.lower().endswith(".csv"):
            content = uploaded_file.read().decode("utf-8-sig", errors="ignore")
            reader  = csv.DictReader(io.StringIO(content))
            hdr_raw = reader.fieldnames or []
            data    = list(reader)
            hdr     = [h.strip().lower() for h in hdr_raw]
            def _col(*needles):
                for needle in needles:
                    for i, h in enumerate(hdr):
                        if needle in h: return hdr_raw[i]
                return None
            zip_col  = _col("zip code", "zip", "postal")
            own_col  = _col("fy26 account owner", "account owner", "rep name", "rep")
            terr_col = _col("territory name", "territory")
            if not zip_col or not own_col: return {}, "Could not detect zip/owner columns"
            result = {}
            for row in data:
                z     = normalize_zip(row.get(zip_col, ""))
                owner = str(row.get(own_col, "")).strip()
                if z and owner and owner.lower() not in ("none","nan",""):
                    entry = {"owner": owner}
                    if terr_col and row.get(terr_col):
                        entry["territory"] = str(row[terr_col]).strip()
                    result[z] = entry
            return result, None
        else:
            wb = openpyxl.load_workbook(
                io.BytesIO(uploaded_file.read()), read_only=True, data_only=True
            )
            d, e = _parse_territory_wb(wb, sheet_name)
            wb.close()
            return d, e
    except Exception as ex:
        return {}, str(ex)

def lookup_owner(zip_raw, territory_dict):
    if not zip_raw or not territory_dict: return None
    entry = territory_dict.get(normalize_zip(zip_raw))
    if not entry: return None
    if isinstance(entry, dict):
        owner     = entry.get("owner", "")
        territory = entry.get("territory", "")
        return f"{owner} \u2014 {territory}" if territory else owner
    return str(entry)

def lookup_owner_entry(zip_raw, territory_dict):
    """Return raw dict entry {owner, owner_id?, territory?} for email building."""
    if not zip_raw or not territory_dict: return None
    return territory_dict.get(normalize_zip(zip_raw))

# ── LinkedIn ──────────────────────────────────────────────────────────────────

def linkedin_search_url(name):
    return (
        "https://www.linkedin.com/search/results/companies/?"
        + urllib.parse.urlencode({"keywords": name})
    )

# ── Row processor ─────────────────────────────────────────────────────────────

def process_row(row, mapping, gb_dict=None, national_dict=None):
    gb_dict       = gb_dict       or {}
    national_dict = national_dict or {}

    name        = str(row.get(mapping.get("account_name") or "", "")).strip()
    acct_id     = str(row.get(mapping.get("account_id")   or "", "")).strip()
    current_raw = str(row.get(mapping.get("segment")      or "", "")).strip()
    current_seg = normalize_segment(current_raw)
    dnb_raw     = str(row.get(mapping.get("dnb")          or "", "")).strip()
    dnb_val     = parse_number(dnb_raw)
    city        = str(row.get(mapping.get("city")         or "", "")).strip()
    state       = str(row.get(mapping.get("state")        or "", "")).strip()
    zip_raw     = str(row.get(mapping.get("zip")          or "", "")).strip()
    opp_name    = str(row.get(mapping.get("opp_name")     or "", "")).strip()
    stage       = str(row.get(mapping.get("stage")        or "", "")).strip()
    close_date  = str(row.get(mapping.get("close_date")   or "", "")).strip()
    amount      = str(row.get(mapping.get("amount")       or "", "")).strip()

    rec = {
        "account_name":     name,
        "account_id":       acct_id,
        "zip":              zip_raw,
        "city":             city       or "\u2014",
        "state":            state      or "\u2014",
        "opp_name":         opp_name   or "\u2014",
        "stage":            stage      or "\u2014",
        "close_date":       close_date or "\u2014",
        "amount":           amount     or "\u2014",
        "current_segment":  current_raw,
        "dnb_employees":    dnb_raw    or "\u2014",
        "expected_segment": "\u2014",
        "status":           "\u2014",
        "basis":            "\u2014",
        "linkedin_url":     "",
        "suggested_owner":  "\u2014",
    }

    if dnb_val is not None and dnb_val >= 5:
        # ── ROE path ──────────────────────────────────────────────────────────
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
        if rec["status"] == "Incorrect":
            correct_dict = national_dict if rec["expected_segment"] == "US National" else gb_dict
            owner = lookup_owner(zip_raw, correct_dict)
            rec["suggested_owner"] = owner or ("Zip not in territory file" if correct_dict else "\u2014")

    else:
        # ── No D&B or D&B < 5 (suspect value) ────────────────────────────────
        if dnb_val is not None:
            rec["basis"] = (
                f"D\u0026B: {dnb_val} \u2014 value < 5, likely incorrect. "
                "Verify headcount via LinkedIn."
            )
        else:
            rec["basis"] = "No D\u0026B data \u2014 manual LinkedIn check needed"
        rec["linkedin_url"] = linkedin_search_url(name)
        rec["status"]       = "Needs Review"
        if gb_dict or national_dict:
            gb_owner  = lookup_owner(zip_raw, gb_dict)       or ("Zip not found" if gb_dict       else None)
            nat_owner = lookup_owner(zip_raw, national_dict) or ("Zip not found" if national_dict else None)
            parts = []
            if gb_owner  is not None: parts.append(f"If GB: {gb_owner}")
            if nat_owner is not None: parts.append(f"If National: {nat_owner}")
            if parts: rec["suggested_owner"] = " \u2502 ".join(parts)

    return rec

# ── Rendering ─────────────────────────────────────────────────────────────────

def badge_html(status):
    cls = {
        "Correct":      "bc",
        "Incorrect":    "bi",
        "Needs Review": "br",
    }.get(status, "bg")
    return f'<span class="badge {cls}">{he(status)}</span>'

def render_table(rows):
    tbody = ""
    for r in rows:
        li_cell = (
            f'<a href="{he(r["linkedin_url"])}" target="_blank">Search &#8599;</a>'
            if r.get("linkedin_url") else "\u2014"
        )
        tbody += f"""<tr>
          <td>{he(r['account_name'])}</td>
          <td>{he(str(r['city']))}, {he(str(r['state']))}</td>
          <td>{he(str(r['opp_name']))}</td>
          <td>{he(str(r['stage']))}</td>
          <td>{he(str(r['close_date']))}</td>
          <td>{he(str(r['amount']))}</td>
          <td>{he(r['current_segment'])}</td>
          <td>{he(str(r['dnb_employees']))}</td>
          <td>{he(r['expected_segment'])}</td>
          <td>{badge_html(r['status'])}</td>
          <td>{he(r['basis'])}</td>
          <td>{he(str(r.get('suggested_owner', '\u2014')))}</td>
          <td>{li_cell}</td>
        </tr>"""
    st.markdown(f"""
<div style="overflow-x:auto">
<table class="rtable">
  <thead><tr>
    <th>Account Name</th><th>City, State</th>
    <th>Opportunity</th><th>Stage</th><th>Close Date</th><th>Amount</th>
    <th>Current Segment</th><th>D&amp;B Employees</th>
    <th>Expected Segment</th><th>Status</th><th>Basis</th>
    <th>Suggested Owner</th><th>LinkedIn</th>
  </tr></thead>
  <tbody>{tbody}</tbody>
</table></div>""", unsafe_allow_html=True)

def col_preview_html(col_labels, rows, mapping):
    fields = [
        ("Account Name",    mapping.get("account_name")),
        ("Account ID",      mapping.get("account_id")),
        ("Current Segment", mapping.get("segment")),
        ("D&B Employees",   mapping.get("dnb")),
        ("City",            mapping.get("city")),
        ("State",           mapping.get("state")),
        ("Zip",             mapping.get("zip")),
        ("Opp Name",        mapping.get("opp_name")),
        ("Stage",           mapping.get("stage")),
        ("Close Date",      mapping.get("close_date")),
        ("Amount",          mapping.get("amount")),
    ]
    sample = rows[:3]
    thead  = "".join(f"<th>{he(f)}</th>" for f, c in fields if c)
    tbody  = ""
    for sr in sample:
        cells = "".join(
            f"<td>{he(str(sr.get(c, '')))}</td>"
            for f, c in fields if c
        )
        tbody += f"<tr>{cells}</tr>"
    return f"""
<details open>
  <summary style="cursor:pointer;font-size:.82rem;font-weight:600;
                  color:#0070F2;margin-bottom:.4rem">
    Column preview (first 3 rows) \u2014 verify mappings before validating
  </summary>
  <div style="overflow-x:auto">
  <table class="preview-table">
    <thead><tr>{thead}</tr></thead>
    <tbody>{tbody}</tbody>
  </table></div>
</details>"""

# ── Auto-load territory files ─────────────────────────────────────────────────
if "gb_dict" not in st.session_state:
    if GB_TERRITORY_FILE.exists():
        _d, _e = load_territory_path(GB_TERRITORY_FILE, GB_SHEET_NAME)
        st.session_state["gb_dict"]       = _d
        st.session_state["gb_dict_label"] = (
            f"Loaded \u2014 {len(_d):,} zip codes" if not _e else f"Error: {_e}"
        )
    else:
        st.session_state["gb_dict"]       = {}
        st.session_state["gb_dict_label"] = "File not found"

if "nat_dict" not in st.session_state:
    if NAT_TERRITORY_FILE.exists():
        _d, _e = load_territory_path(NAT_TERRITORY_FILE, NAT_SHEET_NAME)
        st.session_state["nat_dict"]       = _d
        st.session_state["nat_dict_label"] = (
            f"Loaded \u2014 {len(_d):,} zip codes" if not _e else f"Error: {_e}"
        )
    else:
        st.session_state["nat_dict"]       = {}
        st.session_state["nat_dict_label"] = "File not found"

# ── App header ────────────────────────────────────────────────────────────────
st.markdown("""
<div style="background:#0070F2;color:white;padding:1.2rem 1.6rem;
            border-radius:8px;margin-bottom:1rem">
  <div style="font-size:1.4rem;font-weight:700;color:white">
    Open Opportunity Assignment Validator
  </div>
  <div style="font-size:.88rem;opacity:.85;margin-top:.25rem">
    Validates ROE segment assignment for accounts with open opportunities
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div style="background:#E1F4FF;border:1px solid #4CB1FF;border-radius:6px;
            padding:.7rem 1rem;font-size:.84rem;margin-bottom:1.2rem">
  <b>ROE:</b>
  D&amp;B &ge; 300 &rarr; <b>US National</b> &nbsp;|&nbsp;
  D&amp;B &lt; 300 &rarr; <b>General Business</b> &nbsp;|&nbsp;
  D&amp;B missing &rarr; <b>Needs Review</b> with LinkedIn search link
</div>
""", unsafe_allow_html=True)

# ── Step 1: Connect ───────────────────────────────────────────────────────────
st.markdown("**1. Connect to Salesforce**")

c1, c2 = st.columns([3, 1])
with c1:
    instance_url = st.text_input("Salesforce Instance URL", value=DEFAULT_INSTANCE, key="k_instance")
with c2:
    report_id = st.text_input("Report ID", value=DEFAULT_REPORT, key="k_rid")
session_id = st.text_input("Session ID", type="password",
                            placeholder="Paste your Salesforce Session ID here", key="k_sid")

# Territory status banner
_gb_label  = st.session_state.get("gb_dict_label",  "Not loaded")
_nat_label = st.session_state.get("nat_dict_label", "Not loaded")
st.markdown(
    f"<div style='background:#E8F5E9;border:1px solid #A5D6A7;border-radius:6px;"
    f"padding:.5rem .9rem;font-size:.82rem;margin-top:.5rem;margin-bottom:.4rem'>"
    f"<b>Territory files</b> &nbsp;&nbsp;"
    f"<span style='color:#188918'>&#10003; GB: {he(_gb_label)}</span>"
    f"&nbsp;&nbsp;&bull;&nbsp;&nbsp;"
    f"<span style='color:#188918'>&#10003; National: {he(_nat_label)}</span>"
    f"</div>",
    unsafe_allow_html=True,
)

if st.toggle("Update territory files", key="k_terr_toggle"):
    oc1, oc2 = st.columns(2)
    with oc1:
        gb_upload = st.file_uploader("General Business Territory", type=["csv","xlsx"], key="k_gb_up")
    with oc2:
        nat_upload = st.file_uploader("US National Territory", type=["csv","xlsx"], key="k_nat_up")
    for upload, skey, sheet in [
        (gb_upload,  "gb_dict",  GB_SHEET_NAME),
        (nat_upload, "nat_dict", NAT_SHEET_NAME),
    ]:
        if upload is not None:
            fid = f"{upload.name}_{upload.size}"
            if st.session_state.get(f"{skey}_upload_id") != fid:
                _d, _e = load_territory_upload(upload, sheet)
                label  = f"Loaded — {len(_d):,} zip codes" if not _e else f"Error: {_e}"
                st.session_state[skey]                      = _d
                st.session_state[f"{skey}_label"]           = label
                st.session_state[f"{skey}_upload_id"]       = fid
                st.success(f"Loaded override: {label}")

if st.button("Load Report", key="btn_load"):
    if not session_id:
        st.error("Paste your Salesforce Session ID to continue.")
    else:
        with st.spinner("Connecting to Salesforce\u2026"):
            try:
                raw_data         = fetch_sfdc_report(instance_url, session_id, report_id)
                rows, col_labels = parse_report(raw_data)
                report_name      = raw_data.get("reportMetadata", {}).get("name", "Report")
                st.session_state["rows"]        = rows
                st.session_state["col_labels"]  = col_labels
                st.session_state["report_name"] = report_name
                st.session_state["results"]     = None
                st.success(f"Loaded **{report_name}** \u2014 {len(rows):,} rows.")
            except Exception as e:
                st.error(str(e))

# ── Step 2: Column mapping ────────────────────────────────────────────────────
if st.session_state.get("rows"):
    rows       = st.session_state["rows"]
    col_labels = st.session_state["col_labels"]
    opt        = ["(not available)"] + col_labels

    st.markdown("**2. Map Report Columns**")
    st.caption("Auto-detected from column names and cell values. Adjust if needed.")

    def _idx(cols, val):
        return cols.index(val) if val in cols else 0

    # Row 1 — required account fields + account ID
    r1c1, r1c2, r1c3, r1c4 = st.columns(4)
    with r1c1:
        acct_col = st.selectbox("Account Name \u2733", col_labels,
            index=_idx(col_labels, auto_detect(col_labels, "account_name", rows)))
    with r1c2:
        acct_id_col = st.selectbox("Account ID (18-digit)", opt,
            index=_idx(opt, auto_detect(col_labels, "account_id", rows, required=False)))
    with r1c3:
        seg_col  = st.selectbox("Current Segment \u2733", col_labels,
            index=_idx(col_labels, auto_detect(col_labels, "segment", rows)))
    with r1c4:
        dnb_col  = st.selectbox("D\u0026B Employees Worldwide \u2733", col_labels,
            index=_idx(col_labels, auto_detect(col_labels, "dnb", rows)))

    # Row 2 — location + zip
    r2c1, r2c2, r2c3 = st.columns(3)
    with r2c1:
        city_col  = st.selectbox("Billing City", opt,
            index=_idx(opt, auto_detect(col_labels, "city", rows, required=False)))
    with r2c2:
        state_col = st.selectbox("Billing State", opt,
            index=_idx(opt, auto_detect(col_labels, "state", rows, required=False)))
    with r2c3:
        zip_col   = st.selectbox("Billing Zip", opt,
            index=_idx(opt, auto_detect(col_labels, "zip", rows, required=False)))

    # Row 3 — opportunity fields
    r3c1, r3c2, r3c3, r3c4 = st.columns(4)
    with r3c1:
        opp_col   = st.selectbox("Opportunity Name", opt,
            index=_idx(opt, auto_detect(col_labels, "opp_name", rows, required=False)))
    with r3c2:
        stage_col = st.selectbox("Stage", opt,
            index=_idx(opt, auto_detect(col_labels, "stage", rows, required=False)))
    with r3c3:
        close_col = st.selectbox("Close Date", opt,
            index=_idx(opt, auto_detect(col_labels, "close_date", rows, required=False)))
    with r3c4:
        amt_col   = st.selectbox("Amount", opt,
            index=_idx(opt, auto_detect(col_labels, "amount", rows, required=False)))

    mapping = {
        "account_name": acct_col,
        "account_id":   None if acct_id_col == "(not available)" else acct_id_col,
        "segment":      seg_col,
        "dnb":          dnb_col,
        "city":         None if city_col  == "(not available)" else city_col,
        "state":        None if state_col == "(not available)" else state_col,
        "zip":          None if zip_col   == "(not available)" else zip_col,
        "opp_name":     None if opp_col   == "(not available)" else opp_col,
        "stage":        None if stage_col == "(not available)" else stage_col,
        "close_date":   None if close_col == "(not available)" else close_col,
        "amount":       None if amt_col   == "(not available)" else amt_col,
    }

    st.markdown(col_preview_html(col_labels, rows, mapping), unsafe_allow_html=True)

    if acct_col == seg_col:
        st.error(
            f"**Account Name** and **Current Segment** are both mapped to **{acct_col}**. "
            "Adjust the dropdowns — they must point to different columns."
        )

    if st.button("Validate Assignments", key="btn_validate", disabled=(acct_col == seg_col)):
        results = []
        total   = len(rows)
        prog    = st.progress(0.0, text="Starting\u2026")
        gb_dict      = st.session_state.get("gb_dict",  {})
        national_dict= st.session_state.get("nat_dict", {})

        for i, row in enumerate(rows):
            name = str(row.get(acct_col, "")).strip()
            prog.progress(
                (i + 1) / total,
                text=f"Processing {i + 1} of {total}\u2003\u2014\u2003{name[:70]}",
            )
            results.append(process_row(row, mapping, gb_dict, national_dict))

        prog.empty()
        st.session_state["results"] = results


# ── Step 3: Results ───────────────────────────────────────────────────────────
if st.session_state.get("results"):
    results = st.session_state["results"]
    st.markdown("**3. Validation Results**")

    total   = len(results)
    correct = sum(1 for r in results if r["status"] == "Correct")
    wrong   = sum(1 for r in results if r["status"] == "Incorrect")
    review  = sum(1 for r in results if r["status"] == "Needs Review")
    pct     = lambda n: f"{round(n / total * 100)}%" if total else "0%"

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Opportunities",  total)
    m2.metric("Correctly Assigned",   f"{correct} ({pct(correct)})")
    m3.metric("Incorrectly Assigned", f"{wrong} ({pct(wrong)})")
    m4.metric("Needs Review",         f"{review} ({pct(review)})")
    st.write("")

    correct_rows = [r for r in results if r["status"] == "Correct"]
    wrong_rows   = [r for r in results if r["status"] == "Incorrect"]
    review_rows  = [r for r in results if r["status"] == "Needs Review"]

    gb_dict_r  = st.session_state.get("gb_dict",  {})
    nat_dict_r = st.session_state.get("nat_dict", {})

    # ── Small helpers ─────────────────────────────────────────────────────────

    def _cs(r):
        c, s = str(r["city"]), str(r["state"])
        return "\u2014" if c == "\u2014" and s == "\u2014" else f"{c}, {s}"

    def _owner_str(zip_raw, seg):
        """Return (display_str, owner, owner_id, territory) for a zip + segment."""
        tgt   = nat_dict_r if seg == "US National" else gb_dict_r
        entry = lookup_owner_entry(zip_raw, tgt) or {}
        owner    = entry.get("owner",    "")
        owner_id = entry.get("owner_id", "")
        terr     = entry.get("territory","")
        display  = f"{owner} \u2014 {terr}" if terr else (owner or "Zip not in territory file")
        return display, owner, owner_id, terr

    # ── Render helpers ────────────────────────────────────────────────────────

    _HDR = "font-size:.78rem;font-weight:700;color:#0070F2;padding:.3rem 0;white-space:nowrap;"
    _CEL = "font-size:.82rem;padding:.1rem 0;line-height:1.4;"
    _MID = "font-size:.78rem;color:#6A7275;font-family:monospace;padding:.1rem 0;"
    _LNK = "font-size:.82rem;color:#0070F2;text-decoration:none;font-weight:500;"

    def _hdr(txt):
        st.markdown(f"<div style='{_HDR}'>{txt}</div>", unsafe_allow_html=True)

    def _cell(txt, sty=None):
        s = sty or _CEL
        st.markdown(f"<div style='{s}'>{he(str(txt))}</div>", unsafe_allow_html=True)

    def _sep():
        st.markdown(
            "<hr style='margin:.25rem 0 .4rem 0;border:none;"
            "border-top:1px solid #E1E2E6'>",
            unsafe_allow_html=True,
        )

    # ── Column ratios ─────────────────────────────────────────────────────────
    #  [chk | name | id | city,st | dnb | roe/basis | seg-dropdown | owner ]
    _INC = [0.32, 2.4, 1.0, 1.15, 0.65, 1.2, 1.6, 2.3]
    _NR  = [0.32, 2.4, 1.0, 1.15, 0.65, 2.5, 1.6, 2.3, 0.72]

    _INC_HDRS = ["", "Account Name", "Account ID", "City, State",
                 "D&B", "ROE Expects", "Segment to Assign", "Account Owner"]
    _NR_HDRS  = ["", "Account Name", "Account ID", "City, State",
                 "D&B", "Basis", "Segment to Assign", "Account Owner", "LinkedIn"]

    # Selections built during tab rendering
    inc_selections = []
    nr_selections  = []

    # ── Tabs ──────────────────────────────────────────────────────────────────
    tab1, tab2, tab3 = st.tabs([
        f"Correct ({len(correct_rows)})",
        f"Incorrect ({len(wrong_rows)})",
        f"Needs Review ({len(review_rows)})",
    ])

    with tab1:
        if correct_rows:
            render_table(correct_rows)
        else:
            st.success("All accounts in this report are correctly assigned.")

    # ── Tab 2: Incorrect ──────────────────────────────────────────────────────
    with tab2:
        if wrong_rows:
            st.caption(
                "Segment to Assign is pre-filled from the ROE rule. "
                "Change it if needed \u2014 Account Owner updates immediately."
            )
            hcols = st.columns(_INC)
            for hc, lbl in zip(hcols, _INC_HDRS):
                with hc: _hdr(lbl)
            _sep()

            for i, r in enumerate(wrong_rows):
                exp = r.get("expected_segment", "General Business")
                if exp not in ("General Business", "US National"):
                    exp = "General Business"
                cols = st.columns(_INC)

                with cols[0]:
                    checked = st.checkbox("", key=f"inc_chk_{i}",
                                          value=True, label_visibility="collapsed")
                with cols[1]:
                    st.markdown(
                        f"<div style='{_CEL}'><b>{he(r['account_name'])}</b></div>",
                        unsafe_allow_html=True,
                    )
                with cols[2]:
                    _cell(r.get("account_id", "\u2014"), _MID)
                with cols[3]:
                    _cell(_cs(r))
                with cols[4]:
                    _cell(r["dnb_employees"])
                with cols[5]:
                    _cell(exp)
                with cols[6]:
                    chosen_seg = st.selectbox(
                        "seg", ["General Business", "US National"],
                        index=1 if exp == "US National" else 0,
                        key=f"inc_seg_{i}",
                        label_visibility="collapsed",
                    )
                with cols[7]:
                    disp, owner, owner_id, terr = _owner_str(r.get("zip", ""), chosen_seg)
                    _cell(disp)

                if checked:
                    inc_selections.append({
                        "account_name":      r["account_name"],
                        "account_id":        r.get("account_id", ""),
                        "current_segment":   r["current_segment"],
                        "segment_to_assign": chosen_seg,
                        "owner_name":        owner,
                        "owner_id":          owner_id,
                        "territory":         terr,
                        "city_state":        _cs(r),
                        "dnb":               str(r["dnb_employees"]),
                        "basis":             r["basis"],
                    })
        else:
            st.success("No incorrectly assigned accounts found.")

    # ── Tab 3: Needs Review ───────────────────────────────────────────────────
    with tab3:
        if review_rows:
            st.caption(
                "D&B is missing or < 5. Verify headcount via LinkedIn, "
                "select the correct segment \u2014 Account Owner updates immediately."
            )
            hcols = st.columns(_NR)
            for hc, lbl in zip(hcols, _NR_HDRS):
                with hc: _hdr(lbl)
            _sep()

            for i, r in enumerate(review_rows):
                cols = st.columns(_NR)

                with cols[0]:
                    checked = st.checkbox("", key=f"nr_chk_{i}",
                                          value=True, label_visibility="collapsed")
                with cols[1]:
                    st.markdown(
                        f"<div style='{_CEL}'><b>{he(r['account_name'])}</b></div>",
                        unsafe_allow_html=True,
                    )
                with cols[2]:
                    _cell(r.get("account_id", "\u2014"), _MID)
                with cols[3]:
                    _cell(_cs(r))
                with cols[4]:
                    _cell(r["dnb_employees"])
                with cols[5]:
                    _cell(r["basis"])
                with cols[6]:
                    chosen_seg = st.selectbox(
                        "seg", ["General Business", "US National"],
                        index=0,
                        key=f"nr_seg_{i}",
                        label_visibility="collapsed",
                    )
                with cols[7]:
                    disp, owner, owner_id, terr = _owner_str(r.get("zip", ""), chosen_seg)
                    _cell(disp)
                with cols[8]:
                    li_url = r.get("linkedin_url", "")
                    if li_url:
                        st.markdown(
                            f"<a href='{he(li_url)}' target='_blank'"
                            f" style='{_LNK}'>Search \u2197</a>",
                            unsafe_allow_html=True,
                        )
                    else:
                        _cell("\u2014")

                if checked:
                    nr_selections.append({
                        "account_name":      r["account_name"],
                        "account_id":        r.get("account_id", ""),
                        "current_segment":   r["current_segment"],
                        "segment_to_assign": chosen_seg,
                        "owner_name":        owner,
                        "owner_id":          owner_id,
                        "territory":         terr,
                        "city_state":        _cs(r),
                        "dnb":               str(r["dnb_employees"]),
                        "basis":             r["basis"],
                    })
        else:
            st.success("No accounts need review.")

    # ── Reassignment actions ──────────────────────────────────────────────────
    all_selections   = inc_selections + nr_selections
    total_actionable = len(wrong_rows) + len(review_rows)

    # ── Excel builder ─────────────────────────────────────────────────────────
    def _build_excel(sels):
        import openpyxl
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Reassignment Request"

        headers = [
            "Account Name", "Account ID", "Current Segment", "New Segment",
            "New Owner", "Owner ID", "Territory",
            "City, State", "D&B Employees", "Basis / Reason",
        ]
        col_widths = [38, 22, 22, 22, 32, 20, 32, 22, 14, 55]

        # ── Header row ────────────────────────────────────────────────────────
        hdr_fill  = PatternFill("solid", fgColor="0070F2")
        hdr_font  = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
        hdr_align = Alignment(horizontal="center", vertical="center",
                               wrap_text=True)
        thin = Side(style="thin", color="FFFFFF")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for ci, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=ci, value=h)
            cell.fill    = hdr_fill
            cell.font    = hdr_font
            cell.alignment = hdr_align
            cell.border  = border

        ws.row_dimensions[1].height = 32

        # ── Data rows ─────────────────────────────────────────────────────────
        alt_fill  = PatternFill("solid", fgColor="E1F4FF")
        dat_font  = Font(name="Calibri", size=10)
        dat_align = Alignment(vertical="top", wrap_text=True)
        dat_border = Border(
            left=Side(style="thin", color="D0D0D0"),
            right=Side(style="thin", color="D0D0D0"),
            bottom=Side(style="thin", color="D0D0D0"),
        )

        for ri, s in enumerate(sels, 2):
            values = [
                s["account_name"],    s["account_id"],
                s["current_segment"], s["segment_to_assign"],
                s["owner_name"],      s["owner_id"],
                s["territory"],       s["city_state"],
                s["dnb"],             s["basis"],
            ]
            fill = alt_fill if ri % 2 == 0 else None
            for ci, val in enumerate(values, 1):
                cell = ws.cell(row=ri, column=ci, value=val or "")
                cell.font      = dat_font
                cell.alignment = dat_align
                cell.border    = dat_border
                if fill:
                    cell.fill = fill
            ws.row_dimensions[ri].height = 18

        # ── Column widths + freeze ─────────────────────────────────────────────
        for ci, w in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.freeze_panes = "A2"

        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    # ── Preview table helper ──────────────────────────────────────────────────
    def _preview_html(sels):
        if not sels:
            return "<p style='color:#6A7275;font-size:.85rem'>No accounts selected.</p>"
        cols = ["Account Name", "Account ID", "New Segment",
                "New Owner", "Owner ID", "Territory", "City, State", "D&B", "Basis"]
        thead = "".join(
            f"<th style='background:#0070F2;color:#fff;padding:.45rem .7rem;"
            f"text-align:left;font-size:.78rem;white-space:nowrap'>{c}</th>"
            for c in cols
        )
        tbody = ""
        for ri, s in enumerate(sels):
            bg = "#E1F4FF" if ri % 2 == 0 else "#fff"
            vals = [
                he(s["account_name"]),
                f"<span style='font-family:monospace;font-size:.72rem'>{he(s['account_id'])}</span>",
                he(s["segment_to_assign"]),
                he(s["owner_name"] or "\u2014"),
                f"<span style='font-family:monospace;font-size:.72rem'>{he(s['owner_id'] or '\u2014')}</span>",
                he(s["territory"]  or "\u2014"),
                he(s["city_state"]),
                he(s["dnb"]),
                f"<span style='font-size:.75rem'>{he(s['basis'])}</span>",
            ]
            tds = "".join(
                f"<td style='padding:.4rem .7rem;border-bottom:1px solid #E1E2E6;"
                f"font-size:.82rem;vertical-align:top'>{v}</td>"
                for v in vals
            )
            tbody += f"<tr style='background:{bg}'>{tds}</tr>"
        return (
            "<div style='overflow-x:auto'>"
            f"<table style='border-collapse:collapse;width:100%'>"
            f"<thead><tr>{thead}</tr></thead>"
            f"<tbody>{tbody}</tbody>"
            "</table></div>"
        )

    if total_actionable > 0:
        st.markdown("---")
        st.caption(
            f"{len(all_selections)} of {total_actionable} actionable accounts selected. "
            "Uncheck rows above to exclude from the file and email."
        )

        # ── Preview toggle ────────────────────────────────────────────────────
        if st.toggle("Preview reassignment file", key="k_xl_preview"):
            st.markdown(_preview_html(all_selections), unsafe_allow_html=True)
            st.write("")

        # ── Action buttons ────────────────────────────────────────────────────
        excel_bytes = _build_excel(all_selections) if all_selections else b""
        excel_b64   = base64.b64encode(excel_bytes).decode() if excel_bytes else ""

        # Build compact email body (details are in the Excel attachment)
        def _email_body(sels):
            lines = [
                "Hello Field Services team,",
                "",
                "Please reassign the attached accounts.",
                "",
                "Thank you",
            ]
            return lines

        email_lines = _email_body(all_selections)
        email_body  = "\r\n".join(email_lines)
        mailto_url  = (
            "mailto:concur_fieldservices@sap.com"
            + "?subject=" + urllib.parse.quote("Account Reassignment Request", safe="")
            + "&body="    + urllib.parse.quote(email_body, safe="")
        )

        # Escape for embedding in JS string
        mailto_js = mailto_url.replace("\\", "\\\\").replace("'", "\\'")

        ac1, ac2, _spacer = st.columns([1.5, 2.0, 4])

        with ac1:
            st.download_button(
                "Download Excel (.xlsx)",
                excel_bytes,
                "reassignment_request.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_dl_excel",
                disabled=not all_selections,
            )

        with ac2:
            # Single button: downloads Excel + opens email draft simultaneously
            n_sel = len(all_selections)
            btn_label = (
                f"Attach &amp; Send to FS for Reassignment "
                f"({n_sel} account{'s' if n_sel != 1 else ''})"
            )
            stc.html(
                f"""
                <style>
                  #fs-btn {{
                    background: {'#0070F2' if all_selections else '#BCC0C5'};
                    color: #fff;
                    border: none;
                    padding: .45rem 1rem;
                    border-radius: 4px;
                    font-weight: 600;
                    font-size: .84rem;
                    cursor: {'pointer' if all_selections else 'not-allowed'};
                    font-family: '72', Arial, Helvetica, sans-serif;
                    white-space: nowrap;
                  }}
                  #fs-btn:hover {{ background: {'#0057C2' if all_selections else '#BCC0C5'}; }}
                </style>
                <button id="fs-btn"
                  {'disabled' if not all_selections else ''}
                  onclick="(function(){{
                    var a = document.createElement('a');
                    a.href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{excel_b64}';
                    a.download = 'reassignment_request.xlsx';
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    setTimeout(function(){{ window.location.href = '{mailto_js}'; }}, 900);
                  }})()">
                  {btn_label}
                </button>
                """,
                height=48,
            )

        if st.toggle("Preview email body", key="k_email_preview"):
            st.code("\n".join(email_lines), language=None)

    # ── Full results CSV ──────────────────────────────────────────────────────
    st.write("")
    all_buf = io.StringIO()
    all_w   = csv.DictWriter(all_buf, fieldnames=[
        "account_name", "account_id", "city", "state", "opp_name", "stage",
        "close_date", "amount", "current_segment", "dnb_employees",
        "expected_segment", "status", "basis", "suggested_owner", "linkedin_url",
    ], extrasaction="ignore")
    all_w.writeheader()
    all_w.writerows(results)
    st.download_button(
        "Download Full Validation Results (CSV)",
        all_buf.getvalue(),
        "open_opp_assignment_validation.csv",
        "text/csv",
        key="btn_all_csv",
    )
