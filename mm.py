"""
MIMU Jobs PDF Extractor
========================
Reads PDF URLs from the published Google Sheet CSV, downloads each PDF,
extracts full text, parses all fields, and saves to mimu_jobs.csv + mimu_jobs.xlsx

REQUIREMENTS:
    pip install requests pdfplumber pandas openpyxl

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

OUTPUT_CSV  = "mimu_jobs.csv"
OUTPUT_XLSX = "mimu_jobs.xlsx"

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

def fetch_logo_from_website(website_url: str) -> str:
    """Try to find a logo URL from the org's website."""
    if not website_url or not website_url.startswith("http"):
        return ""
    try:
        resp = requests.get(website_url, headers=HEADERS, timeout=10)
        if resp.status_code != 200:
            return ""
        html = resp.text

        # Common logo patterns in HTML
        patterns = [
            r'<link[^>]+rel=["\'](?:shortcut )?icon["\'][^>]+href=["\']([^"\']+)["\']',
            r'<img[^>]+(?:class|id)=["\'][^"\']*logo[^"\']*["\'][^>]+src=["\']([^"\']+)["\']',
            r'<img[^>]+src=["\']([^"\']*logo[^"\']*)["\']',
            r'<img[^>]+src=["\']([^"\']*brand[^"\']*)["\']',
        ]
        base = re.match(r'(https?://[^/]+)', website_url)
        base_url = base.group(1) if base else ""

        for pat in patterns:
            m = re.search(pat, html, re.IGNORECASE)
            if m:
                logo = m.group(1)
                if logo.startswith("//"):
                    logo = "https:" + logo
                elif logo.startswith("/"):
                    logo = base_url + logo
                elif not logo.startswith("http"):
                    logo = base_url + "/" + logo
                return logo
    except Exception:
        pass
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

def extract_emails(text: str) -> str:
    """Extract all email addresses from text, return first valid one."""
    # Decode common encoded formats like %40 -> @
    decoded = text.replace("%40", "@").replace("%2E", ".").replace("%2F", "/")
    emails = re.findall(r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}', decoded)
    # Filter out generic/noisy emails
    skip = ['example.com', 'domain.com', 'email.com', 'yourmail']
    for email in emails:
        if not any(sk in email for sk in skip):
            return email
    return ""

def extract_urls(text: str) -> str:
    """Extract the most relevant apply/application URL from text."""
    urls = re.findall(r'https?://[^\s\'"<>]+', text)
    # Prefer application/career/apply links
    for url in urls:
        if re.search(r'apply|career|job|recruit|workday|bamboo|greenhouse|lever|smartrecruiters|smrtr', url, re.IGNORECASE):
            return url.rstrip('.,)')
    # Return first URL if none match
    if urls:
        return urls[0].rstrip('.,)')
    return ""

def extract_application(text: str, existing: str) -> str:
    """
    Aggressively find application method:
    1. Existing URL from sheet
    2. Email addresses in PDF
    3. Apply URLs in PDF
    4. Any URL near 'apply', 'send', 'submit', 'contact'
    5. Physical address near 'submit', 'send'
    """
    # 1. Use existing if it's a real URL (not a partial like /hro%40...)
    if existing and existing.startswith("http"):
        return existing

    # Decode encoded characters in the full text
    decoded_text = text.replace("%40", "@").replace("%2E", ".").replace("%2F", "/")

    # 2. Look for email near application instructions
    app_section = ""
    for pat in [
        r'(?:to apply|how to apply|application|submit|send)[^\n]{0,200}',
        r'(?:interested candidates?)[^\n]{0,200}',
        r'(?:please send|please submit|kindly send)[^\n]{0,200}',
        r'(?:contact us|for more info)[^\n]{0,200}',
    ]:
        m = re.search(pat, decoded_text, re.IGNORECASE)
        if m:
            app_section += m.group(0) + " "

    # Try email in application section first
    email = extract_emails(app_section)
    if not email:
        email = extract_emails(decoded_text)
    if email:
        return email

    # 3. Try URL in application section
    url = extract_urls(app_section)
    if not url:
        url = extract_urls(decoded_text)
    if url:
        return url

    # 4. Existing partial value (email encoded, address, etc.)
    if existing:
        # Try to decode it
        decoded_existing = existing.replace("%40", "@").replace("%2E", ".").replace("%2F", "/")
        # If it looks like an email after decoding
        if "@" in decoded_existing:
            return decoded_existing.lstrip("/")
        return existing

    return ""

def extract_address(text: str, existing: str) -> str:
    """Extract physical address from PDF text."""
    if existing and len(existing) > 10:
        return existing

    patterns = [
        r'(?:address|office)[:\s]+([^\n]{15,150})',
        r'(?:located at|based at|head office)[:\s]+([^\n]{15,150})',
        r'(\d+[,\s]+[A-Z][^\n]{15,100}(?:Street|Road|Avenue|Lane|Township|Yangon|Mandalay|Myanmar)[^\n]{0,50})',
        r'(No\.?\s*\d+[^\n]{10,100}(?:Street|Road|Township|Yangon|Myanmar)[^\n]{0,30})',
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return m.group(1).strip()
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
    if re.search(r'\bngo\b|non.governmental', tl):
        return "NGO"
    if re.search(r'united nations|\bundp\b|\bunicef\b|\bwfp\b|\bunhcr\b|\bwho\b|\bilo\b|\bun\b agency', tl):
        return "UN Agency"
    if re.search(r'\bgovernment\b|ministry of|department of', tl):
        return "Government"
    if re.search(r'\bprivate\b|\bltd\b|\blimited\b|\bcorporation\b|\binc\b', tl):
        return "Private Sector"
    return ""

def detect_company_founded(text: str) -> str:
    m = re.search(r'(?:established|founded|since|incorporated)\s+(?:in\s+)?(\d{4})', text, re.IGNORECASE)
    if m:
        year = int(m.group(1))
        if 1900 <= year <= datetime.now().year:
            return m.group(1)
    return ""

def extract_website(text: str, existing: str) -> str:
    """Extract org website — prefer official org site over application/form links."""
    # Known application platform domains to skip
    skip_domains = [
        'smrtr.io', 'workday', 'bamboohr', 'greenhouse', 'lever.co',
        'forms.office', 'google.com/forms', 'hr-manager', 'myworkday',
        'jobs.', 'career', 'themimu.info'
    ]

    urls = re.findall(r'https?://[^\s\'"<>)]+', text)
    candidates = []
    for url in urls:
        url = url.rstrip('.,)')
        if not any(sk in url for sk in skip_domains):
            candidates.append(url)

    # Prefer URLs that look like org homepages (short path)
    for url in candidates:
        path = re.sub(r'https?://[^/]+', '', url)
        if len(path) < 20:  # short path = likely homepage
            return url

    if candidates:
        return candidates[0]

    return existing if existing else ""

def extract_company_details(text: str, org: str, existing: str) -> str:
    """Extract org background/about section from PDF."""
    org_escaped = re.escape(org[:15]) if org else ""
    patterns = []
    if org_escaped:
        patterns += [
            rf'about\s+{org_escaped}[^\n]{{0,200}}',
            rf'{org_escaped}[^\n]{{0,50}}\nis\s+(?:a|an|the)[^\n]{{0,200}}',
        ]
    patterns += [
        r'(?:about us|background|organization overview|who we are)[:\s]*\n([^\n]{{20,}}(?:\n[^\n]{{20,}}){{0,5}})',
        r'(?:about the organization|about \w+)[:\s]+([^\n]{{30,300}})',
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE | re.DOTALL)
        if m:
            result = m.group(0) if m.lastindex is None else m.group(1) if m.lastindex >= 1 else m.group(0)
            result = re.sub(r'\s+', ' ', result).strip()
            if len(result) > 30:
                return result[:400]

    # Fallback to block search
    block = search_block(text, [r'background', r'about us', r'who we are', r'organization overview'], 8)
    if block:
        return block[:400]

    return existing[:400] if existing else ""

# ── Verbose printer ────────────────────────────────────────────────────────────

def print_extracted(record: dict, pdf_text: str):
    pad = "      "
    div = pad + "-" * 56
    print(div)
    fields_to_show = [
        ("Job Title",          60),
        ("Job Type",           20),
        ("Job Field",          20),
        ("Job Location",       60),
        ("Date Posted",        20),
        ("Deadline",           20),
        ("Salary Range",       60),
        ("Job Qualifications", 100),
        ("Job Experience",     100),
        ("Application",        100),
        ("Company Name",       60),
        ("Company Logo",       100),
        ("Company Type",       20),
        ("Company Founded",    10),
        ("Company Website",    100),
        ("Company Address",    100),
    ]
    for field, maxlen in fields_to_show:
        val = record.get(field, "")
        if val:
            label = (field + ":").ljust(22)
            display = val[:maxlen] + ("…" if len(val) > maxlen else "")
            print(f"{pad}{label} {display}")

    for field in ("Job Description", "Company Details"):
        val = record.get(field, "")
        if val:
            label = (field + ":").ljust(22)
            print(f"{pad}{label} {val[:150]}{'…' if len(val) > 150 else ''}")

    if pdf_text:
        snippet = " ".join(pdf_text[:500].split())
        print(f"\n{pad}--- PDF SNIPPET (500 chars) ---")
        print(f"{pad}{snippet[:500]}")
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
        r'place of work[:\s]+([^\n]{3,80})',
        r'based in[:\s]+([^\n]{3,60})',
    ])
    deadline = s(row.get("Deadline", "")) or search_pattern(text, [
        r'closing date[:\s]+([^\n]{3,40})',
        r'application deadline[:\s]+([^\n]{3,40})',
        r'deadline[:\s]+([^\n]{3,40})',
        r'submit (?:by|before)[:\s]+([^\n]{3,40})',
    ])

    # Qualifications — deep search
    quals = search_block(text, [r'qualif', r'academic requirement', r'minimum requirement', r'degree required'], 6)
    if not quals:
        quals = search_pattern(text, [
            r'(?:qualifications?|education(?:al)? requirements?)[:\s]+([^\n]{10,250})',
            r"(bachelor'?s?|master'?s?|phd|diploma|degree)[^\n]{0,120}",
            r'(minimum\s+(?:diploma|degree|bachelor)[^\n]{0,100})',
        ])
    if not quals:
        quals = s(row.get("Job Qualifications", ""))

    # Experience — deep search
    exp = search_pattern(text, [
        r'(\d+\+?\s*(?:to\s*\d+\s*)?years?\s+(?:of\s+)?(?:relevant\s+)?(?:work\s+)?experience[^\n]{0,80})',
        r'(minimum\s+(?:of\s+)?\d+\s+years?[^\n]{0,80})',
        r'(at least\s+\d+\s+years?[^\n]{0,80})',
        r'experience[:\s]+([^\n]{5,100})',
    ])
    if not exp:
        exp = s(row.get("Job Experience", ""))

    # Description — pull responsibilities section
    desc = search_block(text, [
        r'key responsibilit', r'main responsibilit', r'responsibilit',
        r'duties and responsibilit', r'key tasks', r'scope of work',
        r'main duties', r'job purpose', r'objective', r'role summary'
    ], 10)
    if not desc:
        long_lines = [l.strip() for l in text.split("\n") if len(l.strip()) > 50]
        desc = " ".join(long_lines[:5])
    if not desc:
        desc = s(row.get("Job Description", ""))
    desc = str(desc or "")[:600]

    # Salary
    salary = search_pattern(text, [
        r'salary[:\s]+([^\n]{5,80})',
        r'remuneration[:\s]+([^\n]{5,80})',
        r'compensation[:\s]+([^\n]{5,80})',
        r'pay[:\s]+([^\n]{5,60})',
        r'([\d,]+\s*(?:MMK|USD|Ks|Kyats?)[^\n]{0,40})',
        r'(\$[\d,]+[^\n]{0,30})',
        r'(competitive salary[^\n]{0,60})',
    ])
    if not salary:
        salary = s(row.get("Salary Range", ""))

    # Application — aggressive extraction
    existing_app = s(row.get("Application", ""))
    app_url = extract_application(text, existing_app)

    # Website
    existing_website = s(row.get("Company Website", "")) or s(row.get("Company URL", ""))
    website = extract_website(text, existing_website)

    # Address
    existing_address = s(row.get("Company Address", ""))
    address = extract_address(text, existing_address)

    # Company details / background
    org = s(row.get("Company Name", ""))
    existing_details = s(row.get("Company Details", ""))
    details = extract_company_details(text, org, existing_details)

    # Company logo — try to fetch from website
    logo = ""
    if website:
        logo = fetch_logo_from_website(website)

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
        "Company Logo":       str(logo or ""),
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

        time.sleep(0.2)

    # 3. Save CSV + XLSX
    print(f"\n[3/3] Saving output files …")
    out_df = pd.DataFrame(records, columns=OUTPUT_COLUMNS)

    out_df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
    print(f"      ✓ Saved {OUTPUT_CSV}")

    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False, sheet_name="MIMU Jobs")
        ws = writer.sheets["MIMU Jobs"]
        # Auto-width columns
        for col in ws.columns:
            max_len = max((len(str(cell.value)) if cell.value else 0) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)
        # Freeze header
        ws.freeze_panes = "A2"
    print(f"      ✓ Saved {OUTPUT_XLSX}")

    print(f"\n{'=' * 60}")
    print(f"✅ COMPLETE")
    print(f"   Total        : {total}")
    print(f"   PDF success  : {success_count}")
    print(f"   Sheet-only   : {skip_count}")
    print(f"   Outputs      : {OUTPUT_CSV}, {OUTPUT_XLSX}")
    print(f"   Finished     : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)


if __name__ == "__main__":
    main()
