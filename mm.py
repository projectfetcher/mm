"""
MIMU Jobs PDF Extractor
========================
Reads PDF URLs from the published Google Sheet CSV, downloads each PDF,
extracts full text, parses all fields, and saves to mimu_jobs.csv.

REQUIREMENTS:
    pip install requests pdfplumber pandas

USAGE:
    python mimu_jobs_extractor.py
"""

import requests
import pdfplumber
import pandas as pd
import re
import io
import time
import sys
import math
from datetime import datetime

# ── Config ─────────────────────────────────────────────────────────────────────

SHEET_CSV_URL = (
    "https://docs.google.com/spreadsheets/d/e/"
    "2PACX-1vS9_Zza6If2T_LT-iVvQjBTIFTeqt_OVBws70v_s3NJavT-ZosZ28qtE7xds7iS5rLmU2UbhzxWnOsY"
    "/pub?gid=964760760&single=true&output=csv"
)

OUTPUT_FILE = "mimu_jobs.csv"

OUTPUT_COLUMNS = [
    "Job Title",
    "Job Type",
    "Job Qualifications",
    "Job Experience",
    "Job Location",
    "Job Field",
    "Date Posted",
    "Deadline",
    "Job Description",
    "Application",
    "Company URL",
    "Company Name",
    "Company Logo",
    "Company Industry",
    "Company Founded",
    "Company Type",
    "Company Website",
    "Company Address",
    "Company Details",
    "Job URL",
    "Estimated Deadline",
    "Salary Range",
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}

# ── Field keyword maps ─────────────────────────────────────────────────────────

FIELD_KEYWORDS = {
    "Health":         ["health", "medical", "doctor", "nurse", "clinical", "pharmacy", "nutrition", "epidemic"],
    "Logistics":      ["logistics", "supply chain", "fleet", "transport", "warehouse", "procurement"],
    "Finance":        ["finance", "accounting", "audit", "budget", "financial", "grants"],
    "WASH":           ["wash", "water", "sanitation", "hygiene"],
    "Protection":     ["protection", "gbv", "gender", "child protection", "safeguarding"],
    "Education":      ["education", "teacher", "school", "training", "learning"],
    "HR":             ["human resource", "recruitment", "personnel", "talent"],
    "Administration": ["admin", "administration", "office management"],
    "Food Security":  ["food security", "livelihoods", "agriculture"],
    "Shelter":        ["shelter", "nfi", "construction", "engineer", "infrastructure"],
    "Nutrition":      ["malnutrition", "stunting", "wasting", "feeding"],
    "IT":             ["information technology", "software", "developer", "database", "network", "ict"],
    "Communications": ["communication", "media", "journalist", "reporting", "public relations"],
    "Legal":          ["legal", "lawyer", "compliance", "policy"],
}

JOB_TYPE_KEYWORDS = {
    "Consultancy": ["consultant", "consultancy", "terms of reference", "tor "],
    "Internship":  ["intern", "internship"],
    "Part-time":   ["part-time", "part time"],
    "Contract":    ["contract", "fixed-term", "fixed term", "temporary"],
}

# ── Safe string helper ─────────────────────────────────────────────────────────

def s(val) -> str:
    """Safely convert any value (including NaN/None/float) to string."""
    if val is None:
        return ""
    if isinstance(val, float) and math.isnan(val):
        return ""
    return str(val).strip()

# ── PDF downloader ─────────────────────────────────────────────────────────────

def fetch_pdf_text(pdf_url: str) -> str:
    if not pdf_url or not str(pdf_url).startswith("http"):
        return ""
    try:
        resp = requests.get(pdf_url, headers=HEADERS, timeout=30)
        if resp.status_code != 200:
            print(f"      ✗ HTTP {resp.status_code}")
            return ""
        with pdfplumber.open(io.BytesIO(resp.content)) as pdf:
            pages = []
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    pages.append(text)
            return "\n".join(pages)
    except Exception as e:
        print(f"      ✗ PDF error: {e}")
        return ""

# ── Text field extractors ──────────────────────────────────────────────────────

def search_pattern(text: str, patterns: list) -> str:
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return m.group(1).strip(" :.,\t-")
    return ""

def search_block(text: str, header_patterns: list, max_lines: int = 6) -> str:
    lines = text.split("\n")
    for i, line in enumerate(lines):
        for pat in header_patterns:
            if re.search(pat, line, re.IGNORECASE):
                block = []
                for j in range(i + 1, min(i + 1 + max_lines, len(lines))):
                    l = lines[j].strip()
                    if not l:
                        continue
                    if re.match(r'^[A-Z][A-Z\s]{4,}:?\s*$', l) or (l.endswith(":") and len(l) < 50):
                        break
                    block.append(l)
                if block:
                    return " ".join(block)
    return ""

def detect_job_type(text: str, title: str) -> str:
    combined = (text + " " + title).lower()
    for jtype, keywords in JOB_TYPE_KEYWORDS.items():
        if any(kw in combined for kw in keywords):
            return jtype
    return "Full-time"

def detect_job_field(text: str, title: str) -> str:
    combined = (text + " " + title).lower()
    for field, keywords in FIELD_KEYWORDS.items():
        if any(kw in combined for kw in keywords):
            return field
    return "Other"

def detect_company_type(text: str) -> str:
    tl = text.lower()
    if re.search(r'ngo|non.governmental', tl):
        return "NGO"
    if re.search(r'united nations|\bun\b|undp|unicef|wfp|unhcr|\bwho\b|ilo', tl):
        return "UN Agency"
    if re.search(r'government|ministry|department of', tl):
        return "Government"
    if re.search(r'private|company|ltd|limited|corporation', tl):
        return "Private Sector"
    return ""

def detect_company_founded(text: str) -> str:
    m = re.search(r'(?:established|founded|since)\s+(?:in\s+)?(\d{4})', text, re.IGNORECASE)
    if m:
        return m.group(1)
    return ""

# ── Verbose printer ────────────────────────────────────────────────────────────

def print_extracted(record: dict, pdf_text: str):
    pad = "      "
    div = pad + "-" * 56
    print(div)
    fields_to_show = [
        ("Job Title",          20),
        ("Job Type",           20),
        ("Job Field",          20),
        ("Job Location",       20),
        ("Date Posted",        20),
        ("Deadline",           20),
        ("Estimated Deadline", 20),
        ("Salary Range",       20),
        ("Job Qualifications", 80),
        ("Job Experience",     80),
        ("Application",        80),
        ("Company Name",       40),
        ("Company Type",       20),
        ("Company Founded",    10),
        ("Company Website",    80),
        ("Company Address",    80),
    ]
    for field, maxlen in fields_to_show:
        val = record.get(field, "")
        if val:
            label = (field + ":").ljust(22)
            display = val[:maxlen] + ("…" if len(val) > maxlen else "")
            print(f"{pad}{label} {display}")

    desc = record.get("Job Description", "")
    if desc:
        print(f"{pad}{'Job Description:'.ljust(22)} {desc[:150]}{'…' if len(desc) > 150 else ''}")

    details = record.get("Company Details", "")
    if details:
        print(f"{pad}{'Company Details:'.ljust(22)} {details[:120]}{'…' if len(details) > 120 else ''}")

    if pdf_text:
        snippet = " ".join(pdf_text[:400].split())
        print(f"\n{pad}--- PDF RAW SNIPPET (first 400 chars) ---")
        print(f"{pad}{snippet[:400]}")
    print(div)

# ── Field parser ───────────────────────────────────────────────────────────────

def parse_pdf_fields(text: str, row: dict) -> dict:
    title    = s(row.get("Job Title", "")) or search_pattern(text, [
        r'position[:\s]+([^\n]{3,80})',
        r'job title[:\s]+([^\n]{3,80})',
    ])
    location = s(row.get("Job Location", "")) or search_pattern(text, [
        r'location[:\s]+([^\n]{3,80})',
        r'duty station[:\s]+([^\n]{3,80})',
    ])
    deadline = s(row.get("Deadline", "")) or search_pattern(text, [
        r'closing date[:\s]+([^\n]{3,40})',
        r'deadline[:\s]+([^\n]{3,40})',
        r'application deadline[:\s]+([^\n]{3,40})',
    ])

    # Qualifications
    quals = search_block(text, [r'qualif', r'education', r'academic', r'degree required'], 5)
    if not quals:
        quals = search_pattern(text, [
            r'(?:qualifications?|education)[:\s]+([^\n]{10,200})',
            r"(bachelor'?s?|master'?s?|phd|diploma|degree)[^\n]{0,100}",
        ])
    if not quals:
        quals = s(row.get("Job Qualifications", ""))

    # Experience
    exp = search_pattern(text, [
        r'experience[:\s]+([^\n]{5,80})',
        r'(\d+\+?\s*(?:to\s*\d+\s*)?years?\s+(?:of\s+)?experience[^\n]{0,50})',
        r'(minimum\s+\d+\s+years?[^\n]{0,50})',
        r'(at least\s+\d+\s+years?[^\n]{0,50})',
    ])
    if not exp:
        exp = s(row.get("Job Experience", ""))

    # Description
    desc = search_block(text, [r'responsibilit', r'duties', r'key tasks', r'scope of work', r'objective'], 8)
    if not desc:
        long_lines = [l.strip() for l in text.split("\n") if len(l.strip()) > 40]
        desc = " ".join(long_lines[:4])
    if not desc:
        desc = s(row.get("Job Description", ""))
    desc = str(desc or "")[:500]

    # Salary
    salary = search_pattern(text, [
        r'salary[:\s]+([^\n]{5,60})',
        r'remuneration[:\s]+([^\n]{5,60})',
        r'([\d,]+\s*(?:MMK|USD|Ks|Kyats)[^\n]{0,30})',
    ])
    if not salary:
        salary = s(row.get("Salary Range", ""))

    # Application URL
    app_url = s(row.get("Application", "")) or search_pattern(text, [r'apply[:\s]+(https?://\S+)'])

    # Company website
    website = search_pattern(text, [
        r'website[:\s]+(https?://\S+)',
        r'(https?://(?!.*apply|.*career|.*job)\S{10,})',
    ])
    if not website:
        website = s(row.get("Company Website", "")) or s(row.get("Company URL", ""))

    # Address
    address = search_pattern(text, [
        r'address[:\s]+([^\n]{10,120})',
        r'office[:\s]+([^\n]{10,120})',
    ])
    if not address:
        address = s(row.get("Company Address", ""))

    # Company details
    org = s(row.get("Company Name", ""))
    org_escaped = re.escape(org[:10]) if org else ""
    detail_pats = ([rf'about\s+(?:us|the\s+organization|{org_escaped})'] if org_escaped else []) + [r'background']
    details = search_block(text, detail_pats, 6)
    if not details:
        details = s(row.get("Company Details", ""))

    job_type  = detect_job_type(text, title)
    job_field = detect_job_field(text, title)
    comp_type = detect_company_type(text) or s(row.get("Company Type", ""))
    founded   = detect_company_founded(text)

    return {
        "Job Title":          str(title or ""),
        "Job Type":           str(job_type or ""),
        "Job Qualifications": str(quals or "")[:300],
        "Job Experience":     str(exp or "")[:200],
        "Job Location":       str(location or ""),
        "Job Field":          str(job_field or ""),
        "Date Posted":        str(s(row.get("Date Posted", ""))),
        "Deadline":           str(deadline or ""),
        "Job Description":    str(desc or ""),
        "Application":        str(app_url or ""),
        "Company URL":        str(website or ""),
        "Company Name":       str(org or ""),
        "Company Logo":       "",
        "Company Industry":   str(job_field or ""),
        "Company Founded":    str(founded or ""),
        "Company Type":       str(comp_type or ""),
        "Company Website":    str(website or ""),
        "Company Address":    str(address or ""),
        "Company Details":    str(details or "")[:400],
        "Job URL":            str(s(row.get("Job URL", ""))),
        "Estimated Deadline": str(deadline or ""),
        "Salary Range":       str(salary or ""),
    }

# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("MIMU Jobs PDF Extractor")
    print(f"Started : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    # 1. Load sheet
    print(f"\n[1/3] Fetching Google Sheet CSV …")
    try:
        df = pd.read_csv(SHEET_CSV_URL)
        df = df.dropna(how="all")
        print(f"      Loaded {len(df)} rows.")
    except Exception as e:
        print(f"ERROR loading sheet: {e}")
        sys.exit(1)

    if "PDF URL" not in df.columns:
        print("ERROR: 'PDF URL' column not found.")
        print("Columns:", list(df.columns))
        sys.exit(1)

    # 2. Process each row
    print(f"\n[2/3] Downloading PDFs and extracting fields …")
    records = []
    total         = len(df)
    success_count = 0
    skip_count    = 0

    for idx, row in df.iterrows():
        pdf_url = s(row.get("PDF URL", ""))
        title   = s(row.get("Job Title", ""))
        num     = idx + 1

        print(f"\n  [{num}/{total}] {title}")

        if pdf_url and pdf_url.startswith("http"):
            print(f"      PDF : {pdf_url}")
            pdf_text = fetch_pdf_text(pdf_url)
            if pdf_text:
                print(f"      ✓ {len(pdf_text):,} characters extracted")
                success_count += 1
            else:
                print(f"      ⚠ No text extracted from PDF")
                skip_count += 1
        else:
            print(f"      ⚠ No PDF URL — using sheet data only")
            pdf_text = ""
            skip_count += 1

        enriched = parse_pdf_fields(pdf_text, row.to_dict())
        print_extracted(enriched, pdf_text)
        records.append(enriched)

        time.sleep(0.3)

    # 3. Save
    print(f"\n[3/3] Saving to {OUTPUT_FILE} …")
    out_df = pd.DataFrame(records, columns=OUTPUT_COLUMNS)
    out_df.to_csv(OUTPUT_FILE, index=False, encoding="utf-8-sig")

    print(f"\n{'=' * 60}")
    print(f"✅ COMPLETE")
    print(f"   Total        : {total}")
    print(f"   PDF success  : {success_count}")
    print(f"   Sheet-only   : {skip_count}")
    print(f"   Output       : {OUTPUT_FILE}")
    print(f"   Finished     : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)


if __name__ == "__main__":
    main()
