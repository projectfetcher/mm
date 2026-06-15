import requests 
import pdfplumber
import pandas as pd
import re
import io
import time
import sys
from datetime import datetime

# ── Config ─────────────────────────────────────────────────────────────────────

SHEET_CSV_URL = (
    "https://docs.google.com/spreadsheets/d/e/"
    "2PACX-1vS9_Zza6If2T_LT-iVvQjBTIFTeqt_OVBws70v_s3NJavT-ZosZ28qtE7xds7iS5rLmU2UbhzxWnOsY"
    "/pub?gid=964760760&single=true&output=csv"
)

OUTPUT_FILE = f"mimu_jobs_enriched_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"

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

# ── PDF downloader ─────────────────────────────────────────────────────────────

def fetch_pdf_text(pdf_url: str) -> str:
    """Download a PDF and extract all text using pdfplumber."""
    if not pdf_url or not pdf_url.startswith("http"):
        return ""
    try:
        resp = requests.get(pdf_url, headers=HEADERS, timeout=30)
        if resp.status_code != 200:
            print(f"    ✗ HTTP {resp.status_code}")
            return ""
        with pdfplumber.open(io.BytesIO(resp.content)) as pdf:
            pages = []
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    pages.append(text)
            return "\n".join(pages)
    except Exception as e:
        print(f"    ✗ PDF error: {e}")
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

def parse_pdf_fields(text: str, row: dict) -> dict:
    """Merge PDF-extracted fields with existing sheet metadata."""
    tl = text.lower()

    title    = row.get("Job Title", "") or search_pattern(text, [
        r'position[:\s]+([^\n]{3,80})',
        r'job title[:\s]+([^\n]{3,80})',
    ])
    location = row.get("Job Location", "") or search_pattern(text, [
        r'location[:\s]+([^\n]{3,80})',
        r'duty station[:\s]+([^\n]{3,80})',
    ])
    deadline = row.get("Deadline", "") or search_pattern(text, [
        r'closing date[:\s]+([^\n]{3,40})',
        r'deadline[:\s]+([^\n]{3,40})',
        r'application deadline[:\s]+([^\n]{3,40})',
    ])

    # Qualifications — prefer PDF text over sheet
    quals = search_block(text, [r'qualif', r'education', r'academic', r'degree required'], 5)
    if not quals:
        quals = search_pattern(text, [
            r'(?:qualifications?|education)[:\s]+([^\n]{10,200})',
            r"(bachelor'?s?|master'?s?|phd|diploma|degree)[^\n]{0,100}",
        ])
    if not quals:
        quals = str(row.get("Job Qualifications", "") or "")

    # Experience
    exp = search_pattern(text, [
        r'experience[:\s]+([^\n]{5,80})',
        r'(\d+\+?\s*(?:to\s*\d+\s*)?years?\s+(?:of\s+)?experience[^\n]{0,50})',
        r'(minimum\s+\d+\s+years?[^\n]{0,50})',
        r'(at least\s+\d+\s+years?[^\n]{0,50})',
    ])
    if not exp:
        exp = str(row.get("Job Experience", "") or "")

    # Description — prefer PDF
    desc = search_block(text, [r'responsibilit', r'duties', r'key tasks', r'scope of work', r'objective'], 8)
    if not desc:
        long_lines = [l.strip() for l in text.split("\n") if len(l.strip()) > 40]
        desc = " ".join(long_lines[:4])
    if not desc:
        desc = str(row.get("Job Description", "") or "")
    desc = str(desc or "")[:500]

    # Salary
    salary = search_pattern(text, [
        r'salary[:\s]+([^\n]{5,60})',
        r'remuneration[:\s]+([^\n]{5,60})',
        r'([\d,]+\s*(?:MMK|USD|Ks|Kyats)[^\n]{0,30})',
    ])
    if not salary:
        salary = str(row.get("Salary Range", "") or "")

    # Application URL
    app_url = row.get("Application", "") or search_pattern(text, [r'apply[:\s]+(https?://\S+)'])

    # Company website
    website = search_pattern(text, [
        r'website[:\s]+(https?://\S+)',
        r'(https?://(?!.*apply|.*career|.*job)\S{10,})',
    ])
    if not website:
        website = row.get("Company Website", "") or row.get("Company URL", "")

    # Address
    address = search_pattern(text, [
        r'address[:\s]+([^\n]{10,120})',
        r'office[:\s]+([^\n]{10,120})',
    ])
    if not address:
        address = row.get("Company Address", "")

    # Company details
    org = row.get("Company Name", "")
    org_escaped = re.escape(org[:10]) if org else ""
    detail_pats = ([rf'about\s+(?:us|the\s+organization|{org_escaped})'] if org_escaped else []) + [r'background']
    details = search_block(text, detail_pats, 6)
    if not details:
        details = row.get("Company Details", "")

    job_type    = detect_job_type(text, title)
    job_field   = detect_job_field(text, title)
    comp_type   = detect_company_type(text) or row.get("Company Type", "")

    return {
        "Job Title":          str(title or ""),
        "Job Type":           str(job_type or ""),
        "Job Qualifications": str(quals or "")[:300],
        "Job Experience":     str(exp or "")[:200],
        "Job Location":       str(location or ""),
        "Job Field":          str(job_field or ""),
        "Date Posted":        str(row.get("Date Posted", "") or ""),
        "Deadline":           str(deadline or ""),
        "Job Description":    str(desc or ""),
        "Application":        str(app_url or ""),
        "Company URL":        str(website or ""),
        "Company Name":       str(org or ""),
        "Company Industry":   str(job_field or ""),
        "Company Type":       str(comp_type or ""),
        "Company Website":    str(website or ""),
        "Company Address":    str(address or ""),
        "Company Details":    str(details or "")[:400],
        "Job URL":            str(row.get("Job URL", "") or ""),
        "Estimated Deadline": str(deadline or ""),
        "Salary Range":       str(salary or ""),
        "PDF URL":            str(row.get("PDF URL", "") or ""),
        "PDF Text Length":    len(text),
    }
# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("MIMU Jobs PDF Extractor")
    print("=" * 60)

    # 1. Load sheet
    print(f"\n[1/3] Fetching Google Sheet CSV …")
    try:
        df = pd.read_csv(SHEET_CSV_URL)
        # Drop completely empty rows
        df = df.dropna(how="all")
        print(f"      Loaded {len(df)} rows.")
    except Exception as e:
        print(f"ERROR loading sheet: {e}")
        sys.exit(1)

    if "PDF URL" not in df.columns:
        print("ERROR: 'PDF URL' column not found in sheet.")
        print("Columns found:", list(df.columns))
        sys.exit(1)

    # 2. Process each row
    print(f"\n[2/3] Downloading PDFs and extracting fields …\n")
    records = []
    total = len(df)

    for idx, row in df.iterrows():
        pdf_url = str(row.get("PDF URL", "")).strip()
        title   = str(row.get("Job Title", "")).strip()
        num     = idx + 1

        print(f"  [{num}/{total}] {title}")

        if pdf_url and pdf_url.startswith("http"):
            print(f"    → {pdf_url}")
            pdf_text = fetch_pdf_text(pdf_url)
            print(f"    ✓ {len(pdf_text):,} chars extracted")
        else:
            print(f"    ⚠ No PDF URL — using sheet data only")
            pdf_text = ""

        enriched = parse_pdf_fields(pdf_text, row.to_dict())
        records.append(enriched)

        time.sleep(0.3)  # be polite

    # 3. Save output
    print(f"\n[3/3] Saving to {OUTPUT_FILE} …")
    out_df = pd.DataFrame(records)
    out_df.to_csv(OUTPUT_FILE, index=False, encoding="utf-8-sig")
    print(f"      ✅ Done! {len(records)} jobs saved to '{OUTPUT_FILE}'")
    print("=" * 60)


if __name__ == "__main__":
    main()
