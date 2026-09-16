#!/usr/bin/env python3
"""
Account Assignment Validator
No pip installs required — Python standard library only.
Run: python3 validator.py
"""

import http.server
import socketserver
import json
import re
import time
import urllib.request
import urllib.parse
import urllib.error
import webbrowser
import threading
import uuid
import sys

PORT = 8765
SFDC_API_VERSION = "v59.0"
DNB_THRESHOLD    = 300
BING_ENDPOINT    = "https://api.bing.microsoft.com/v7.0/search"
BING_DELAY       = 0.4  # seconds between Bing calls (rate limit)

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

# Most specific patterns first to avoid partial matches
BAND_PATTERNS = [
    (r"10[,.]?001\+?\s+employees?",                  "10,001+"),
    (r"5[,.]?001\s*[-\u2013]\s*10[,.]?000\s+employees?", "5,001-10,000"),
    (r"1[,.]?001\s*[-\u2013]\s*5[,.]?000\s+employees?",  "1,001-5,000"),
    (r"501\s*[-\u2013]\s*1[,.]?000\s+employees?",        "501-1,000"),
    (r"201\s*[-\u2013]\s*500\s+employees?",              "201-500"),
    (r"51\s*[-\u2013]\s*200\s+employees?",               "51-200"),
    (r"11\s*[-\u2013]\s*50\s+employees?",                "11-50"),
    (r"\b1\s*[-\u2013]\s*10\s+employees?",               "1-10"),
]

# In-memory job store (single-user local app)
_jobs      = {}
_jobs_lock = threading.Lock()


# ── Core logic ────────────────────────────────────────────────────────────────

def normalize_segment(s):
    if not s:
        return ""
    sl = str(s).lower().strip()
    if any(x in sl for x in ["national", " nat", "nats"]):
        return "US National"
    if any(x in sl for x in ["general business", " gb", "smb", "gen bus"]):
        return "General Business"
    return str(s).strip()

def parse_number(val):
    if val is None:
        return None
    sv = str(val).strip()
    if not sv or sv.lower() in ("none", "null", "n/a", "-", "\u2014", ""):
        return None
    try:
        return int(float(re.sub(r"[,\s]", "", sv)))
    except Exception:
        return None

def extract_band(text):
    if not text:
        return None
    for pattern, label in BAND_PATTERNS:
        if re.search(pattern, text, re.IGNORECASE):
            return label
    return None


# ── Salesforce ────────────────────────────────────────────────────────────────

def fetch_sfdc_report(instance_url, session_id, report_id):
    url = (
        f"{instance_url.rstrip('/')}/services/data/{SFDC_API_VERSION}"
        f"/analytics/reports/{report_id}?includeDetails=true"
    )
    req = urllib.request.Request(
        url,
        headers={"Authorization": f"Bearer {session_id}", "Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read())
    except urllib.error.HTTPError as ex:
        if ex.code == 401:
            raise Exception("Authentication failed — check your Session ID.")
        if ex.code == 404:
            raise Exception("Report not found — check your Report ID.")
        raise Exception(f"Salesforce API error {ex.code}: {ex.reason}")
    except urllib.error.URLError as ex:
        raise Exception(f"Connection error: {ex.reason}")

def parse_report(data):
    api_names = data.get("reportMetadata", {}).get("detailColumns", [])
    col_info  = data.get("reportExtendedMetadata", {}).get("detailColumnInfo", {})
    labels    = [col_info.get(n, {}).get("label", n) for n in api_names]

    rows_raw = data.get("factMap", {}).get("T!T", {}).get("rows", [])
    if not rows_raw:
        raise Exception(
            "No rows found in report. "
            "Ensure it is a Tabular report and contains data."
        )

    rows = []
    for row in rows_raw:
        cells = row.get("dataCells", [])
        rows.append({
            lbl: (cells[i].get("label", "") if i < len(cells) else "")
            for i, lbl in enumerate(labels)
        })
    return rows, labels


# ── Bing / LinkedIn enrichment ────────────────────────────────────────────────

def bing_linkedin(name, city, state, website, api_key):
    """Return (linkedin_url, band, error). Does NOT fetch LinkedIn pages."""
    parts = [f'site:linkedin.com/company "{name}"']
    if city:    parts.append(city)
    if state:   parts.append(state)
    if website:
        domain = re.sub(r"https?://(www\.)?", "", website or "").split("/")[0]
        if domain: parts.append(domain)

    url = BING_ENDPOINT + "?" + urllib.parse.urlencode(
        {"q": " ".join(parts), "count": 5, "mkt": "en-US"}
    )
    req = urllib.request.Request(url, headers={"Ocp-Apim-Subscription-Key": api_key})
    try:
        with urllib.request.urlopen(req, timeout=12) as resp:
            results = json.loads(resp.read()).get("webPages", {}).get("value", [])

        for r in results:
            r_url   = r.get("url", "")
            snippet = r.get("snippet", "")
            title   = r.get("name", "")
            if (
                "linkedin.com/company/" in r_url
                and "/jobs"   not in r_url
                and "/in/"    not in r_url
                and "/school" not in r_url
            ):
                band = extract_band(snippet + " " + title)
                return r_url, band, None

        return None, None, "No LinkedIn company page found"

    except urllib.error.HTTPError as ex:
        if ex.code == 401:
            return None, None, "Invalid Bing API key"
        return None, None, f"Bing API error {ex.code}"
    except Exception as ex:
        return None, None, str(ex)


# ── Per-row processing ────────────────────────────────────────────────────────

def process_row(row, mapping, bing_key, bing_call_count):
    """Returns (result_dict, used_bing: bool)."""
    def col(key):
        v = mapping.get(key)
        return str(row.get(v, "")).strip() if v and v != "(not available)" else ""

    name        = col("account_name")
    current_raw = col("segment")
    current_seg = normalize_segment(current_raw)
    dnb_raw     = col("dnb")
    dnb_val     = parse_number(dnb_raw)
    city        = col("city")
    state       = col("state")
    website     = col("website")

    rec = {
        "account_name":    name,
        "current_segment": current_raw,
        "dnb_employees":   dnb_raw if dnb_raw else "\u2014",
        "expected_segment": "\u2014",
        "status":          "\u2014",
        "basis":           "\u2014",
        "linkedin_band":   "\u2014",
        "linkedin_url":    None,
    }

    used_bing = False

    # ── D&B present → apply ROE directly ─────────────────────────────────────
    if dnb_val is not None:
        if dnb_val >= DNB_THRESHOLD:
            rec["expected_segment"] = "US National"
            rec["basis"]            = f"D&B: {dnb_val:,} \u2265 {DNB_THRESHOLD}"
        else:
            rec["expected_segment"] = "General Business"
            rec["basis"]            = f"D&B: {dnb_val:,} < {DNB_THRESHOLD}"

        if current_seg == normalize_segment(rec["expected_segment"]):
            rec["status"] = "Correct"
        else:
            rec["status"] = "Incorrect"

    # ── No D&B → search LinkedIn via Bing ────────────────────────────────────
    else:
        rec["basis"] = "No D&B data"

        if bing_key:
            if bing_call_count > 0:
                time.sleep(BING_DELAY)

            li_url, li_band, err = bing_linkedin(name, city, state, website, bing_key)
            used_bing = True

            if li_url:
                rec["linkedin_url"] = li_url
                if li_band:
                    expected               = BAND_TO_SEGMENT.get(li_band, "\u2014")
                    rec["linkedin_band"]   = li_band
                    rec["expected_segment"]= expected
                    rec["basis"]           = f"LinkedIn band: {li_band}"

                    if current_seg == normalize_segment(expected):
                        rec["status"] = "Correct (LinkedIn)"
                    else:
                        rec["status"] = "Incorrect (LinkedIn)"
                else:
                    rec["status"] = "Needs Review"
                    rec["basis"]  = "LinkedIn found \u2014 band not in snippet"
            else:
                rec["status"] = "Data Gap"
                rec["basis"]  = err or "No match"
        else:
            rec["status"] = "Data Gap"
            rec["basis"]  = "No D&B; no Bing key provided"

    return rec, used_bing


# ── HTTP server ───────────────────────────────────────────────────────────────

class ThreadingHTTPServer(socketserver.ThreadingMixIn, http.server.HTTPServer):
    daemon_threads = True


class Handler(http.server.BaseHTTPRequestHandler):

    def log_message(self, *args):
        pass  # suppress console logs

    # ── Helpers ───────────────────────────────────────────────────────────────

    def send_json(self, data, status=200):
        body = json.dumps(data).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def read_body(self):
        length = int(self.headers.get("Content-Length", 0))
        return json.loads(self.rfile.read(length)) if length else {}

    # ── Routing ───────────────────────────────────────────────────────────────

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self):
        path = self.path.split("?")[0]

        if path in ("/", "/index.html"):
            body = HTML.encode()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        elif path == "/api/stream":
            self._handle_stream()

        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        try:
            body = self.read_body()
        except Exception:
            self.send_json({"error": "Invalid JSON"}, 400)
            return

        if self.path == "/api/load":
            self._handle_load(body)
        elif self.path == "/api/start-validate":
            self._handle_start_validate(body)
        else:
            self.send_json({"error": "Not found"}, 404)

    # ── Handlers ──────────────────────────────────────────────────────────────

    def _handle_load(self, body):
        try:
            raw = fetch_sfdc_report(
                body.get("instance_url", "").strip(),
                body.get("session_id",   "").strip(),
                body.get("report_id",    "").strip(),
            )
            rows, labels = parse_report(raw)
            name         = raw.get("reportMetadata", {}).get("name", "Report")
            self.send_json({"rows": rows, "col_labels": labels, "report_name": name})
        except Exception as ex:
            self.send_json({"error": str(ex)})

    def _handle_start_validate(self, body):
        job_id = str(uuid.uuid4())
        with _jobs_lock:
            _jobs[job_id] = {
                "rows":       body.get("rows", []),
                "mapping":    body.get("mapping", {}),
                "bing_key":   body.get("bing_key", ""),
            }
        self.send_json({"job_id": job_id})

    def _handle_stream(self):
        params  = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
        job_id  = params.get("job_id", [None])[0]

        with _jobs_lock:
            job = _jobs.get(job_id)
        if not job:
            self.send_response(404)
            self.end_headers()
            return

        self.send_response(200)
        self.send_header("Content-Type",  "text/event-stream")
        self.send_header("Cache-Control", "no-cache")
        self.send_header("Connection",    "keep-alive")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()

        def emit(data):
            self.wfile.write(f"data: {json.dumps(data)}\n\n".encode())
            self.wfile.flush()

        rows      = job["rows"]
        mapping   = job["mapping"]
        bing_key  = job["bing_key"]
        total     = len(rows)
        bing_calls = 0

        try:
            for i, row in enumerate(rows):
                name = str(row.get(mapping.get("account_name", ""), "")).strip()
                emit({"type": "progress", "current": i + 1, "total": total, "account": name[:60]})

                result, used = process_row(row, mapping, bing_key, bing_calls)
                if used:
                    bing_calls += 1

                emit({"type": "result", "row": result})

            emit({"type": "done"})

        except BrokenPipeError:
            pass
        except Exception as ex:
            try:
                emit({"type": "error", "message": str(ex)})
            except Exception:
                pass

        with _jobs_lock:
            _jobs.pop(job_id, None)


# ── Embedded HTML/CSS/JS ──────────────────────────────────────────────────────

HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Account Assignment Validator</title>
<style>
*{font-family:'72','72full',Arial,Helvetica,sans-serif;box-sizing:border-box;margin:0;padding:0}
body{background:#F5F6F7;color:#32363A;padding:1.5rem}
.wrap{max-width:1380px;margin:0 auto}

/* Header */
.hdr{background:#0070F2;color:#fff;padding:1.3rem 1.8rem;border-radius:8px;margin-bottom:1.2rem}
.hdr h1{font-size:1.45rem;font-weight:700}
.hdr p{font-size:.88rem;opacity:.85;margin-top:.3rem}

/* ROE bar */
.roe{background:#E1F4FF;border:1px solid #4CB1FF;border-radius:6px;padding:.75rem 1.1rem;font-size:.84rem;margin-bottom:1.2rem}
.roe b{color:#0057C2}

/* Card */
.card{background:#fff;border:1px solid #E1E2E6;border-radius:8px;padding:1.4rem;margin-bottom:1.2rem}
.card h2{font-size:.95rem;font-weight:700;border-left:4px solid #0070F2;padding-left:.55rem;margin-bottom:1.1rem}

/* Form */
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:.9rem;margin-bottom:1rem}
.fg{display:flex;flex-direction:column;gap:.3rem}
.fg label{font-size:.8rem;font-weight:600}
.fg input,.fg select{border:1px solid #D0D4D8;border-radius:4px;padding:.45rem .65rem;font-size:.85rem;color:#32363A;background:#fff}
.fg input:focus,.fg select:focus{border-color:#0070F2;outline:none;box-shadow:0 0 0 2px rgba(0,112,242,.15)}

/* Buttons */
.btn{background:#0070F2;color:#fff;border:none;border-radius:4px;padding:.5rem 1.3rem;font-size:.85rem;font-weight:600;cursor:pointer}
.btn:hover{background:#0057C2}
.btn:disabled{background:#BCC0C5;cursor:not-allowed}
.btn-out{background:#fff;color:#0070F2;border:2px solid #0070F2}
.btn-out:hover{background:#E1F4FF}

/* Stats */
.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:.9rem;margin-bottom:1.2rem}
.stat{background:#F5F6F7;border:1px solid #E1E2E6;border-radius:8px;padding:.9rem 1rem;text-align:center}
.stat .v{font-size:1.9rem;font-weight:700;color:#0070F2}
.stat .l{font-size:.77rem;color:#6A7275;margin-top:.1rem}
.stat.ok .v{color:#188918}
.stat.bad .v{color:#BB0000}
.stat.warn .v{color:#E76500}

/* Progress */
.prog{margin:.8rem 0}
.pbg{background:#E1E2E6;border-radius:4px;height:7px}
.pfill{background:#0070F2;height:7px;border-radius:4px;transition:width .15s;width:0}
.ptxt{font-size:.78rem;color:#6A7275;margin-top:.35rem}

/* Table */
.tscroll{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:.83rem}
th{background:#0070F2;color:#fff;padding:.6rem .85rem;text-align:left;font-weight:600;white-space:nowrap}
td{padding:.5rem .85rem;border-bottom:1px solid #E1E2E6;vertical-align:middle}
tr:nth-child(even) td{background:#FAFAFA}
tr:hover td{background:#E1F4FF}

/* Badges */
.badge{display:inline-block;padding:.17rem .5rem;border-radius:4px;font-size:.75rem;font-weight:600}
.bc{background:#E8F5E9;color:#188918}
.bi{background:#FFEBEE;color:#BB0000}
.bl{background:#E1F4FF;color:#0057C2}
.br{background:#FFF3E0;color:#E76500}
.bg{background:#EAECEE;color:#6A7275}

a.ll{color:#0070F2;text-decoration:none;font-weight:500}
a.ll:hover{text-decoration:underline}

/* Filter bar */
.fbar{display:flex;align-items:center;gap:.8rem;margin-bottom:.9rem}
.fbar label{font-size:.8rem;font-weight:600}
.fbar select{border:1px solid #D0D4D8;border-radius:4px;padding:.3rem .55rem;font-size:.82rem}

.err{background:#FFEBEE;border:1px solid #FFCDD2;border-radius:4px;padding:.65rem 1rem;color:#BB0000;font-size:.83rem;margin:.4rem 0}
.hidden{display:none}

.spin{display:inline-block;width:14px;height:14px;border:2px solid #BCC0C5;border-top-color:#0070F2;border-radius:50%;animation:sp .7s linear infinite;vertical-align:middle;margin-right:.35rem}
@keyframes sp{to{transform:rotate(360deg)}}
</style>
</head>
<body>
<div class="wrap">

<div class="hdr">
  <h1>Account Assignment Validator</h1>
  <p>US SMB &mdash; validates segment assignments per Rules of Engagement using D&amp;B employee count and LinkedIn band enrichment.</p>
</div>

<div class="roe">
  <b>ROE applied:</b>
  &nbsp;D&amp;B &ge; 300 &rarr; <b>US National</b>
  &nbsp;|&nbsp; D&amp;B &lt; 300 &rarr; <b>General Business</b>
  &nbsp;|&nbsp; No D&amp;B: LinkedIn band above 201&ndash;500 (501&ndash;1,000+) &rarr; <b>US National</b>, rest &rarr; <b>General Business</b>
</div>

<!-- Step 1 -->
<div class="card">
  <h2>1 &nbsp; Connect to Salesforce</h2>
  <div class="grid">
    <div class="fg"><label>Salesforce Instance URL *</label><input id="iUrl" value="https://sapconcur.my.salesforce.com"></div>
    <div class="fg"><label>Session ID *</label><input id="iSid" type="password" placeholder="00D..."></div>
    <div class="fg"><label>Report ID *</label><input id="iRid" value="00OPg00000QbzTl"></div>
    <div class="fg"><label>Bing Web Search API Key <span style="font-weight:400;color:#6A7275">(for LinkedIn enrichment)</span></label><input id="iBing" type="password" placeholder="Azure Bing key"></div>
  </div>
  <div id="errLoad" class="err hidden"></div>
  <button class="btn" id="btnLoad" onclick="loadReport()">Load Report</button>
</div>

<!-- Step 2 -->
<div class="card hidden" id="cardMap">
  <h2>2 &nbsp; Map Report Columns</h2>
  <p style="font-size:.82rem;color:#6A7275;margin-bottom:.9rem">Auto-detected where possible &mdash; adjust if needed.</p>
  <div class="grid" id="mapGrid"></div>
  <div id="errMap" class="err hidden"></div>
  <button class="btn" onclick="validate()">Validate Assignments</button>
</div>

<!-- Step 3 -->
<div class="card hidden" id="cardRes">
  <h2>3 &nbsp; Validation Results</h2>

  <div id="progWrap" class="prog">
    <div class="pbg"><div class="pfill" id="pfill"></div></div>
    <div class="ptxt" id="ptxt">Starting&hellip;</div>
  </div>

  <div class="stats hidden" id="stats">
    <div class="stat">    <div class="v" id="sTotal">0</div><div class="l">Total Accounts</div></div>
    <div class="stat ok"> <div class="v" id="sOk">0</div>   <div class="l">Correctly Assigned</div></div>
    <div class="stat bad"><div class="v" id="sBad">0</div>  <div class="l">Incorrectly Assigned</div></div>
    <div class="stat warn"><div class="v" id="sGap">0</div> <div class="l">Data Gap / Review</div></div>
  </div>

  <div class="fbar hidden" id="fbar">
    <label>Filter:</label>
    <select id="fsel" onchange="applyFilter()">
      <option>All</option>
      <option>Correct</option>
      <option>Incorrect</option>
      <option>Correct (LinkedIn)</option>
      <option>Incorrect (LinkedIn)</option>
      <option>Needs Review</option>
      <option>Data Gap</option>
    </select>
    <button class="btn btn-out" onclick="dlCSV()">Download CSV</button>
  </div>

  <div class="tscroll">
    <table>
      <thead><tr>
        <th>Account Name</th><th>Current Segment</th><th>D&amp;B Employees</th>
        <th>Expected Segment</th><th>Status</th><th>Basis</th>
        <th>LinkedIn Band</th><th>LinkedIn Page</th>
      </tr></thead>
      <tbody id="tbody"></tbody>
    </table>
  </div>
</div>

</div><!-- /wrap -->

<script>
var reportData=null, allRows=[], es=null;

var HINTS={
  account_name:['account name','account','name'],
  segment:['segment','assignment','territory','sales segment'],
  dnb:["d&b","dnb","dun","employee worldwide","employees (d&b)","numberofemployees","employees"],
  city:['billing city','city'],
  state:['billing state','state','province'],
  website:['website','web','domain']
};

function esc(s){
  if(!s||s==='—'||s==='\u2014')return s||'—';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function show(id,msg){var e=document.getElementById(id);e.textContent=msg;e.classList.remove('hidden')}
function hide(id){document.getElementById(id).classList.add('hidden')}
function vis(id){document.getElementById(id).classList.remove('hidden')}

function badgeHtml(s){
  var c={
    'Correct':'bc','Incorrect':'bi',
    'Correct (LinkedIn)':'bl','Incorrect (LinkedIn)':'bi',
    'Needs Review':'br','Data Gap':'bg'
  }[s]||'bg';
  return '<span class="badge '+c+'">'+esc(s)+'</span>';
}

function rowHtml(r){
  var li=r.linkedin_url?'<a class="ll" href="'+esc(r.linkedin_url)+'" target="_blank">Open \u2197</a>':'—';
  return '<tr data-st="'+esc(r.status)+'">'
    +'<td>'+esc(r.account_name)+'</td>'
    +'<td>'+esc(r.current_segment)+'</td>'
    +'<td>'+esc(r.dnb_employees)+'</td>'
    +'<td>'+esc(r.expected_segment)+'</td>'
    +'<td>'+badgeHtml(r.status)+'</td>'
    +'<td>'+esc(r.basis)+'</td>'
    +'<td>'+esc(r.linkedin_band)+'</td>'
    +'<td>'+li+'</td>'
    +'</tr>';
}

function updateStats(){
  var t=allRows.length,
      ok=allRows.filter(function(r){return r.status.indexOf('Correct')>-1}).length,
      bad=allRows.filter(function(r){return r.status.indexOf('Incorrect')>-1}).length,
      gap=allRows.filter(function(r){return r.status==='Data Gap'||r.status==='Needs Review'}).length;
  document.getElementById('sTotal').textContent=t;
  document.getElementById('sOk').textContent=ok+(t?' ('+Math.round(ok/t*100)+'%)':'');
  document.getElementById('sBad').textContent=bad+(t?' ('+Math.round(bad/t*100)+'%)':'');
  document.getElementById('sGap').textContent=gap+(t?' ('+Math.round(gap/t*100)+'%)':'');
}

function applyFilter(){
  var v=document.getElementById('fsel').value;
  document.querySelectorAll('#tbody tr').forEach(function(tr){
    tr.style.display=(v==='All'||tr.dataset.st===v)?'':'none';
  });
}

function dlCSV(){
  var hdrs=['Account Name','Current Segment','D&B Employees','Expected Segment','Status','Basis','LinkedIn Band','LinkedIn URL'];
  var lines=allRows.map(function(r){
    return [r.account_name,r.current_segment,r.dnb_employees,r.expected_segment,
            r.status,r.basis,r.linkedin_band,r.linkedin_url||'']
      .map(function(v){return '"'+String(v||'').replace(/"/g,'""')+'"'}).join(',');
  });
  var csv=[hdrs.join(',')].concat(lines).join('\n');
  var a=document.createElement('a');
  a.href='data:text/csv;charset=utf-8,'+encodeURIComponent(csv);
  a.download='account_assignment_validation.csv';
  a.click();
}

function autoDetect(cols,field){
  var hints=HINTS[field]||[];
  for(var i=0;i<cols.length;i++){
    for(var j=0;j<hints.length;j++){
      if(cols[i].toLowerCase().indexOf(hints[j])>-1) return cols[i];
    }
  }
  return null;
}

function buildMappingUI(cols){
  var fields=[
    {id:'m_an', label:'Account Name *',                 key:'account_name', req:true},
    {id:'m_sg', label:'Current Segment / Assignment *', key:'segment',      req:true},
    {id:'m_db', label:'D&amp;B Employee Worldwide *',   key:'dnb',          req:true},
    {id:'m_cy', label:'Billing City',                   key:'city',         req:false},
    {id:'m_st', label:'Billing State / Province',       key:'state',        req:false},
    {id:'m_ws', label:'Website',                        key:'website',      req:false},
  ];
  var g=document.getElementById('mapGrid');
  g.innerHTML='';
  fields.forEach(function(f){
    var det=autoDetect(cols,f.key);
    var opts=(f.req?[]:['(not available)']).concat(cols);
    var idx=det?opts.indexOf(det):0;
    if(idx<0)idx=0;
    var sel=opts.map(function(o,i){
      return '<option value="'+esc(o)+'"'+(i===idx?' selected':'')+'>'+esc(o)+'</option>';
    }).join('');
    g.insertAdjacentHTML('beforeend',
      '<div class="fg"><label>'+f.label+'</label><select id="'+f.id+'">'+sel+'</select></div>');
  });
}

async function loadReport(){
  hide('errLoad');
  var url=document.getElementById('iUrl').value.trim(),
      sid=document.getElementById('iSid').value.trim(),
      rid=document.getElementById('iRid').value.trim();
  if(!url||!sid||!rid){show('errLoad','Instance URL, Session ID, and Report ID are required.');return}
  var btn=document.getElementById('btnLoad');
  btn.disabled=true;
  btn.innerHTML='<span class="spin"></span>Loading\u2026';
  try{
    var r=await fetch('/api/load',{
      method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({instance_url:url,session_id:sid,report_id:rid})
    });
    var d=await r.json();
    if(d.error) throw new Error(d.error);
    reportData=d;
    buildMappingUI(d.col_labels);
    vis('cardMap');
    document.getElementById('cardMap').scrollIntoView({behavior:'smooth'});
  }catch(e){show('errLoad',e.message||'Failed to load report.')}
  finally{btn.disabled=false;btn.textContent='Load Report'}
}

async function validate(){
  hide('errMap');
  if(!reportData) return;
  var mapping={
    account_name: document.getElementById('m_an').value,
    segment:      document.getElementById('m_sg').value,
    dnb:          document.getElementById('m_db').value,
    city:         document.getElementById('m_cy').value,
    state:        document.getElementById('m_st').value,
    website:      document.getElementById('m_ws').value,
  };
  var bingKey=document.getElementById('iBing').value.trim();

  allRows=[];
  document.getElementById('tbody').innerHTML='';
  hide('stats'); hide('fbar');
  vis('progWrap'); vis('cardRes');
  document.getElementById('pfill').style.width='0%';
  document.getElementById('ptxt').textContent='Starting\u2026';
  document.getElementById('cardRes').scrollIntoView({behavior:'smooth'});

  if(es){es.close();es=null}

  var jobId;
  try{
    var r=await fetch('/api/start-validate',{
      method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({rows:reportData.rows,col_labels:reportData.col_labels,mapping:mapping,bing_key:bingKey})
    });
    var d=await r.json();
    if(d.error) throw new Error(d.error);
    jobId=d.job_id;
  }catch(e){show('errMap',e.message||'Failed to start.');return}

  es=new EventSource('/api/stream?job_id='+encodeURIComponent(jobId));
  es.onmessage=function(evt){
    var d=JSON.parse(evt.data);
    if(d.type==='progress'){
      var pct=Math.round(d.current/d.total*100);
      document.getElementById('pfill').style.width=pct+'%';
      document.getElementById('ptxt').textContent=
        'Processing '+d.current+' of '+d.total+(d.account?'  \u2014  '+d.account:'');
    } else if(d.type==='result'){
      allRows.push(d.row);
      document.getElementById('tbody').insertAdjacentHTML('beforeend',rowHtml(d.row));
      updateStats();
    } else if(d.type==='done'){
      es.close(); es=null;
      hide('progWrap');
      vis('stats'); vis('fbar');
      updateStats();
    } else if(d.type==='error'){
      es.close(); es=null;
      show('errMap',d.message);
    }
  };
  es.onerror=function(){if(es){es.close();es=null}};
}
</script>
</body>
</html>"""


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    print(f"\n  Account Assignment Validator")
    print(f"  Open in your browser: http://localhost:{PORT}")
    print(f"  Press Ctrl+C to stop.\n")
    threading.Timer(1.2, lambda: webbrowser.open(f"http://localhost:{PORT}")).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n  Stopped.")
        sys.exit(0)
