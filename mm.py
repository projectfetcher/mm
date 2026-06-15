#!/usr/bin/env python3
"""
MIMU Jobs Scraper — Python version
====================================
Scrapes https://themimu.info/jobs-for-myanmar-nationals
Downloads each PDF, extracts text, parses fields with regex,
and writes everything to a CSV (and optionally Excel).

Requirements:
    pip install requests beautifulsoup4 pdfplumber openpyxl
"""

import re
import time
import logging
import requests
import pdfplumber
import openpyxl
import csv
from io import BytesIO
from bs4 import BeautifulSoup
from openpyxl.styles import PatternFill, Font, Alignment

# ── Config ─────────────────────────────────────────────────────────────────────
JOBS_URL   = "https://themimu.info/jobs-for-myanmar-nationals"
BASE_URL   = "https://themimu.info"
OUTPUT_CSV  = "mimu_jobs.csv"
OUTPUT_XLSX = "mimu_jobs.xlsx"

OUTPUT_COLUMNS = [
    "Job Title", "Job Type", "Job Qualifications", "Job Experience",
    "Job Location", "Job Field", "Date Posted", "Deadline",
    "Job Description", "Application", "Company URL", "Company Name",
    "Company Industry", "Company Type", "Company Website", "Company Address",
    "Company Details", "Job URL", "Estimated Deadline", "Salary Range",
    "PDF URL"
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5"
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(message)s",
    datefmt="%H:%M:%S"
)
log = logging.getLogger(__name__)

# ── Keyword maps ───────────────────────────────────────────────────────────────
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
    "Legal":          ["legal", "lawyer", "compliance", "policy"]
}

JOB_TYPE_KEYWORDS = {
    "Consultancy": ["consultant", "consultancy", "terms of reference", "tor "],
    "Internship":  ["intern", "internship"],
    "Part-time":   ["part-time", "part time"],
    "Contract":    ["contract", "fixed-term", "fixed term", "temporary"]
}


# ── Fetch helpers ──────────────────────────────────────────────────────────────

def fetch_html(url: str) -> str:
    try:
        r = requests.get(url, headers=HEADERS, timeout=30, allow_redirects=True)
        r.raise_for_status()
        return r.text
    except Exception as e:
        log.error("fetch_html error: %s", e)
        return ""


def fetch_pdf_text(pdf_url: str) -> str:
    try:
        r = requests.get(pdf_url, headers={"User-Agent": HEADERS["User-Agent"]},
                         timeout=30, allow_redirects=True)
        if r.status_code != 200:
            return ""
        with pdfplumber.open(BytesIO(r.content)) as pdf:
            pages = [page.extract_text() or "" for page in pdf.pages]
        return "\n".join(pages)
    except Exception as e:
        log.error("PDF fetch/extract error: %s", e)
        return ""


# ── HTML table parser ──────────────────────────────────────────────────────────

def resolve_url(href: str) -> str:
    if not href:
        return ""
    if href.startswith("http"):
        return href
    if href.startswith("//"):
        return "https:" + href
    return BASE_URL.rstrip("/") + "/" + href.lstrip("/")


def parse_jobs_table(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    jobs = []

    # Find the right table
    job_table = None
    for table in soup.find_all("table"):
        text = table.get_text(" ", strip=True).lower()
        if "closing date" in text or "job title" in text:
            job_table = table
            break
    if not job_table:
        # Fallback: largest table
        tables = soup.find_all("table")
        if tables:
            job_table = max(tables, key=lambda t: len(t.find_all("tr")))

    if not job_table:
        log.warning("Could not locate jobs table")
        return jobs

    rows = job_table.find_all("tr")
    log.info("Table rows (inc. header): %d", len(rows))
    if len(rows) < 2:
        return jobs

    # Detect columns from header
    col_map = {"title": 0, "pdf": 1, "app": 2, "location": 3,
               "org": 4, "date_posted": 5, "deadline": 6, "remarks": 7}
    header_cells = rows[0].find_all(["th", "td"])
    for i, cell in enumerate(header_cells):
        txt = cell.get_text(" ", strip=True).lower()
        if "job title" in txt:                          col_map["title"]       = i
        elif re.search(r"description|download", txt):   col_map["pdf"]         = i
        elif re.search(r"application|online", txt):     col_map["app"]         = i
        elif re.search(r"location|post", txt):          col_map["location"]    = i
        elif re.search(r"organisation|organization", txt): col_map["org"]      = i
        elif re.search(r"upload|posted|date of", txt):  col_map["date_posted"] = i
        elif re.search(r"closing|deadline", txt):       col_map["deadline"]    = i
        elif "remark" in txt:                           col_map["remarks"]     = i

    def safe_cell_text(cells, key):
        idx = col_map.get(key, -1)
        if 0 <= idx < len(cells):
            return cells[idx].get_text(" ", strip=True)
        return ""

    def safe_cell_html(cells, key):
        idx = col_map.get(key, -1)
        if 0 <= idx < len(cells):
            return str(cells[idx])
        return ""

    for row in rows[1:]:
        cells = row.find_all(["td", "th"])
        if not cells:
            continue

        title = safe_cell_text(cells, "title")
        if not title or re.match(r"^job title$", title, re.I):
            continue

        # PDF URL
        pdf_url = ""
        pdf_cell_html = safe_cell_html(cells, "pdf")
        pdf_match = re.search(r'href=["\']([^"\']+\.pdf[^"\']*)', pdf_cell_html, re.I)
        if pdf_match:
            pdf_url = resolve_url(pdf_match.group(1))

        # Application URL
        app_url = ""
        app_cell = cells[col_map.get("app", -1)] if 0 <= col_map.get("app", -1) < len(cells) else None
        if app_cell:
            a_tag = app_cell.find("a", href=True)
            if a_tag:
                app_url = a_tag["href"]
            else:
                raw = app_cell.get_text(" ", strip=True)
                if raw.startswith("http") or "@" in raw:
                    app_url = raw

        jobs.append({
            "title":      title,
            "pdf_url":    pdf_url,
            "app_url":    app_url,
            "location":   safe_cell_text(cells, "location"),
            "org":        safe_cell_text(cells, "org"),
            "date_posted": safe_cell_text(cells, "date_posted"),
            "deadline":   safe_cell_text(cells, "deadline"),
            "remarks":    safe_cell_text(cells, "remarks"),
            "job_url":    JOBS_URL
        })

    return jobs


# ── Regex field extractor ──────────────────────────────────────────────────────

def search_pattern(text: str, patterns: list) -> str:
    for pat in patterns:
        m = re.search(pat, text, re.I)
        if m and m.group(1):
            return re.sub(r'^[\s:.,\-\t]+|[\s:.,\-\t]+$', '', m.group(1))
    return ""


def search_block(text: str, header_patterns: list, max_lines: int = 6) -> str:
    lines = text.split("\n")
    for i, line in enumerate(lines):
        for pat in header_patterns:
            if re.search(pat, line, re.I):
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


def parse_fields(text: str, meta: dict) -> dict:
    t  = text or ""
    tl = t.lower()

    title    = meta.get("title")    or search_pattern(t, [r'position[:\s]+([^\n]{3,80})', r'job title[:\s]+([^\n]{3,80})'])
    location = meta.get("location") or search_pattern(t, [r'location[:\s]+([^\n]{3,80})', r'duty station[:\s]+([^\n]{3,80})'])
    deadline = meta.get("deadline") or search_pattern(t, [r'closing date[:\s]+([^\n]{3,40})', r'deadline[:\s]+([^\n]{3,40})', r'application deadline[:\s]+([^\n]{3,40})'])

    # Job Type
    job_type = "Full-time"
    for jt, kws in JOB_TYPE_KEYWORDS.items():
        if any(kw in tl or kw in (meta.get("title") or "").lower() for kw in kws):
            job_type = jt
            break

    # Job Field
    job_field = "Other"
    for field, kws in FIELD_KEYWORDS.items():
        if any(kw in tl or kw in (meta.get("title") or "").lower() for kw in kws):
            job_field = field
            break

    # Qualifications
    quals = search_block(t, [r'qualif', r'education', r'academic', r'degree required'], 5)
    if not quals:
        quals = search_pattern(t, [
            r'(?:qualifications?|education)[:\s]+([^\n]{10,200})',
            r"(bachelor'?s?|master'?s?|phd|diploma|degree)[^\n]{0,100}"
        ])

    # Experience
    exp = search_pattern(t, [
        r'experience[:\s]+([^\n]{5,80})',
        r'(\d+\+?\s*(?:to\s*\d+\s*)?years?\s+(?:of\s+)?experience[^\n]{0,50})',
        r'(minimum\s+\d+\s+years?[^\n]{0,50})',
        r'(at least\s+\d+\s+years?[^\n]{0,50})'
    ])

    # Description
    desc = search_block(t, [r'responsibilit', r'duties', r'key tasks', r'scope of work', r'objective'], 8)
    if not desc:
        long_lines = [l for l in t.split("\n") if len(l.strip()) > 40]
        desc = " ".join(long_lines[:4])
    desc = (desc or "")[:400]

    # Salary
    salary = search_pattern(t, [
        r'salary[:\s]+([^\n]{5,60})',
        r'remuneration[:\s]+([^\n]{5,60})',
        r'([\d,]+\s*(?:MMK|USD|Ks|Kyats)[^\n]{0,30})'
    ])

    # Application URL
    app_url = meta.get("app_url") or search_pattern(t, [r'apply[:\s]+(https?://[^\s]+)'])

    # Company website
    website = search_pattern(t, [
        r'website[:\s]+(https?://[^\s]+)',
        r'(https?://(?!.*apply|.*career|.*job)[^\s]{10,})'
    ])
    if not website and app_url and app_url.startswith("http"):
        m = re.match(r'(https?://[^/]+)', app_url)
        if m:
            website = m.group(1)

    # Address
    address = search_pattern(t, [
        r'address[:\s]+([^\n]{10,120})',
        r'office[:\s]+([^\n]{10,120})'
    ])

    # Company details
    org_escaped = re.escape((meta.get("org") or "")[:10])
    detail_patterns = [rf'about\s+(?:us|the\s+organization|{org_escaped})', r'background'] if org_escaped else [r'background']
    details = search_block(t, detail_patterns, 6)

    # Company type
    company_type = ""
    if re.search(r'ngo|non.governmental', tl):
        company_type = "NGO"
    elif re.search(r'united nations|\bun\b|undp|unicef|wfp|unhcr|\bwho\b|ilo', tl):
        company_type = "UN Agency"
    elif re.search(r'government|ministry|department of', tl):
        company_type = "Government"
    elif re.search(r'private|company|ltd|limited|corporation', tl):
        company_type = "Private Sector"

    return {
        "Job Title":          title or "",
        "Job Type":           job_type,
        "Job Qualifications": (quals or "")[:300],
        "Job Experience":     (exp or "")[:200],
        "Job Location":       location or "",
        "Job Field":          job_field,
        "Date Posted":        meta.get("date_posted") or "",
        "Deadline":           deadline or "",
        "Job Description":    desc,
        "Application":        app_url or "",
        "Company URL":        website or "",
        "Company Name":       meta.get("org") or "",
        "Company Industry":   job_field,
        "Company Type":       company_type,
        "Company Website":    website or "",
        "Company Address":    address or "",
        "Company Details":    (details or "")[:400],
        "Job URL":            meta.get("job_url") or "",
        "Estimated Deadline": deadline or "",
        "Salary Range":       salary or "",
        "PDF URL":            meta.get("pdf_url") or ""
    }


# ── Write outputs ──────────────────────────────────────────────────────────────

def write_csv(records: list[dict]):
    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=OUTPUT_COLUMNS)
        writer.writeheader()
        writer.writerows(records)
    log.info("CSV saved → %s", OUTPUT_CSV)


def write_xlsx(records: list[dict]):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MIMU Jobs"

    # Header style
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True, size=10)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.append(OUTPUT_COLUMNS)
    for cell in ws[1]:
        cell.fill   = header_fill
        cell.font   = header_font
        cell.alignment = header_align
    ws.row_dimensions[1].height = 40
    ws.freeze_panes = "A2"

    # Data rows
    even_fill = PatternFill("solid", fgColor="EBF3FB")
    data_font  = Font(size=9)
    data_align = Alignment(vertical="top", wrap_text=True)

    for i, rec in enumerate(records, start=2):
        row = [rec.get(col, "") for col in OUTPUT_COLUMNS]
        ws.append(row)
        fill = PatternFill("solid", fgColor="EBF3FB") if i % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
        for cell in ws[i]:
            cell.fill      = fill
            cell.font      = data_font
            cell.alignment = data_align

    # Column widths (in Excel units ≈ chars)
    col_widths = [32,14,36,24,24,16,13,13,54,32,28,24,20,16,28,28,32,36,16,20,28]
    for i, width in enumerate(col_widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width

    wb.save(OUTPUT_XLSX)
    log.info("Excel saved → %s", OUTPUT_XLSX)


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    log.info("Step 1/4 — Fetching MIMU jobs page …")
    html = fetch_html(JOBS_URL)
    if not html:
        log.error("Could not fetch the MIMU jobs page.")
        return

    log.info("Step 2/4 — Parsing jobs table …")
    jobs = parse_jobs_table(html)
    log.info("Found %d job listings.", len(jobs))

    if not jobs:
        log.error("No jobs found. The page structure may have changed.")
        return

    log.info("Step 3/4 — Downloading PDFs and extracting fields …")
    records = []
    for i, job in enumerate(jobs, start=1):
        log.info("[%d/%d] %s", i, len(jobs), job["title"])
        pdf_text = ""
        if job["pdf_url"]:
            pdf_text = fetch_pdf_text(job["pdf_url"])
            log.info("  PDF text chars: %d", len(pdf_text))
        else:
            log.info("  No PDF link — using metadata only")

        records.append(parse_fields(pdf_text, job))
        time.sleep(0.3)

    log.info("Step 4/4 — Writing outputs …")
    write_csv(records)
    write_xlsx(records)
    log.info("Done! %d jobs written.", len(records))


if __name__ == "__main__":
    main()
