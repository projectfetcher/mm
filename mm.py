import os
import re
import json
import time
import datetime
import requests
import pdfplumber
import pandas as pd
from pathlib import Path
from bs4 import BeautifulSoup
import anthropic

# ── Configuration ──────────────────────────────────────────────────────────────
BASE_URL   = "https://themimu.info"
JOBS_URL   = f"{BASE_URL}/jobs-for-myanmar-nationals"
PDF_DIR    = Path("pdfs")
OUTPUT_XLS = f"mimu_jobs_{datetime.date.today():%Y%m%d}.xlsx"

# NOTE: The MIMU site renders ALL jobs on a single page — no pagination needed.
# We do, however, handle the case where the site returns partial HTML by
# checking the total row count after parsing.

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}

# Fields we want in the output
OUTPUT_COLUMNS = [
    "Job Title", "Job Type", "Job Qualifications", "Job Experience",
    "Job Location", "Job Field", "Date Posted", "Deadline",
    "Job Description", "Application", "Company URL", "Company Name",
    "Company Logo", "Company Industry", "Company Founded", "Company Type",
    "Company Website", "Company Address", "Company Details",
    "Job URL", "Estimated Deadline", "Salary Range",
]

# ── Helpers ────────────────────────────────────────────────────────────────────

def fetch_page(url: str, retries: int = 5) -> str:
    """Fetch a URL with retry logic, return HTML text."""
    for attempt in range(1, retries + 1):
        try:
            r = requests.get(url, headers=HEADERS, timeout=60)
            r.raise_for_status()
            return r.text
        except Exception as e:
            wait = 3 * attempt
            print(f"  [!] Attempt {attempt}/{retries} failed for {url}: {e}")
            if attempt < retries:
                print(f"      Retrying in {wait}s …")
                time.sleep(wait)
    return ""


def download_pdf(url: str, dest: Path) -> bool:
    """Download a PDF to dest. Returns True on success."""
    if dest.exists() and dest.stat().st_size > 0:
        return True
    try:
        r = requests.get(url, headers=HEADERS, timeout=120, stream=True)
        r.raise_for_status()
        dest.write_bytes(r.content)
        return True
    except Exception as e:
        print(f"  [!] PDF download failed ({url}): {e}")
        return False


def extract_pdf_text(pdf_path: Path) -> str:
    """Extract all text from a PDF using pdfplumber."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            pages_text = []
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    pages_text.append(t.strip())
            return "\n".join(pages_text)
    except Exception as e:
        print(f"  [!] PDF text extraction failed ({pdf_path}): {e}")
        return ""


def parse_jobs_table(html: str) -> list[dict]:
    """
    Parse every job row from the MIMU jobs page.

    The page contains a single <table> with columns:
        Job Title | Job Description (PDF link) | Online Application Link |
        Post Location | Organisation | Date of upload | Closing Date | Remarks

    All jobs are on one page — no pagination.
    Returns a list of dicts.
    """
    soup = BeautifulSoup(html, "html.parser")
    jobs = []

    # ── Locate the jobs table ──────────────────────────────────────────────────
    # Strategy 1: find a table whose header row contains "Closing Date"
    job_table = None
    for table in soup.find_all("table"):
        header_row = table.find("tr")
        if header_row:
            header_text = header_row.get_text(" ", strip=True).lower()
            if "closing date" in header_text or "job title" in header_text:
                job_table = table
                break

    # Strategy 2: fallback — first table with ≥ 5 rows
    if not job_table:
        for table in soup.find_all("table"):
            if len(table.find_all("tr")) >= 5:
                job_table = table
                break

    if not job_table:
        print("[!] Could not locate the jobs table in the HTML.")
        return jobs

    rows = job_table.find_all("tr")
    print(f"      Table rows found (inc. header): {len(rows)}")

    # Detect header row to find column indices dynamically
    col_map = {}
    header_row = rows[0]
    for idx, cell in enumerate(header_row.find_all(["th", "td"])):
        text = cell.get_text(" ", strip=True).lower()
        if "job title" in text:
            col_map["title"] = idx
        elif "description" in text or "download" in text:
            col_map["pdf"] = idx
        elif "application" in text or "online" in text:
            col_map["app"] = idx
        elif "location" in text or "post" in text:
            col_map["location"] = idx
        elif "organisation" in text or "organization" in text:
            col_map["org"] = idx
        elif "upload" in text or "posted" in text or "date of" in text:
            col_map["date_posted"] = idx
        elif "closing" in text or "deadline" in text:
            col_map["deadline"] = idx
        elif "remark" in text:
            col_map["remarks"] = idx

    # Fallback column positions (matches observed site structure)
    defaults = {
        "title": 0, "pdf": 1, "app": 2,
        "location": 3, "org": 4, "date_posted": 5,
        "deadline": 6, "remarks": 7,
    }
    for k, v in defaults.items():
        col_map.setdefault(k, v)

    # ── Parse data rows ────────────────────────────────────────────────────────
    for row in rows[1:]:
        cols = row.find_all(["td", "th"])
        if not cols:
            continue

        def get_col(key: str) -> "BeautifulSoup | None":
            idx = col_map.get(key, -1)
            return cols[idx] if 0 <= idx < len(cols) else None

        title_cell    = get_col("title")
        pdf_cell      = get_col("pdf")
        app_cell      = get_col("app")
        loc_cell      = get_col("location")
        org_cell      = get_col("org")
        posted_cell   = get_col("date_posted")
        deadline_cell = get_col("deadline")
        remarks_cell  = get_col("remarks")

        if not title_cell:
            continue

        title = title_cell.get_text(strip=True)

        # Skip header-like rows
        if not title or title.lower() in ("job title", "title"):
            continue

        # ── PDF link ─────────────────────────────────────────────────────────
        pdf_url = ""
        if pdf_cell:
            dl_link = pdf_cell.find("a", href=True)
            if dl_link:
                href = dl_link["href"]
                # Handle relative, absolute, and protocol-relative paths
                if href.startswith("http"):
                    pdf_url = href
                elif href.startswith("//"):
                    pdf_url = "https:" + href
                else:
                    pdf_url = BASE_URL.rstrip("/") + "/" + href.lstrip("/")

        # ── Application URL ───────────────────────────────────────────────────
        app_url = ""
        if app_cell:
            app_link = app_cell.find("a", href=True)
            if app_link:
                app_url = app_link["href"]
            else:
                raw = app_cell.get_text(strip=True)
                # Only keep it if it looks like a URL or email
                if raw.startswith("http") or "@" in raw:
                    app_url = raw

        location   = loc_cell.get_text("  ", strip=True) if loc_cell   else ""
        org        = org_cell.get_text(strip=True)        if org_cell   else ""
        date_posted= posted_cell.get_text(strip=True)     if posted_cell   else ""
        deadline   = deadline_cell.get_text(strip=True)   if deadline_cell else ""
        remarks    = remarks_cell.get_text(strip=True)    if remarks_cell  else ""

        # Normalise multi-whitespace in location
        location = re.sub(r"\s{2,}", ", ", location).strip(", ")

        jobs.append({
            "title":       title,
            "pdf_url":     pdf_url,
            "app_url":     app_url,
            "location":    location,
            "org":         org,
            "date_posted": date_posted,
            "deadline":    deadline,
            "remarks":     remarks,
            "job_url":     JOBS_URL,
        })

    return jobs


def ai_extract_fields(raw_text: str, meta: dict, client: anthropic.Anthropic) -> dict:
    """
    Use Claude to extract structured fields from PDF text + table metadata.
    Falls back to table metadata if AI call fails.
    """
    system_prompt = """You are a job-listing data extractor for an NGO/humanitarian sector job board.
Given raw text from a job vacancy PDF and some metadata already known from the website table,
return ONLY a valid JSON object (no markdown fences, no explanation) with these exact keys:

Job Title, Job Type, Job Qualifications, Job Experience, Job Location,
Job Field, Date Posted, Deadline, Job Description, Application,
Company URL, Company Name, Company Logo, Company Industry,
Company Founded, Company Type, Company Website, Company Address,
Company Details, Job URL, Estimated Deadline, Salary Range

Rules:
- Use metadata values when the PDF does not contain better information.
- "Job Description" — concise 2–4 sentence summary of duties/responsibilities.
- "Job Qualifications" — key education / certification requirements only.
- "Job Experience" — years and type of experience required.
- "Job Type" — infer from context: Full-time / Part-time / Contract / Consultancy / Internship.
- "Job Field" — sector keyword: Health / Logistics / Finance / WASH / Protection / Education /
  HR / Administration / Food Security / Shelter / CCCM / Nutrition / IT / Communications / Legal / Other.
- "Salary Range" — extract if explicitly stated, else leave blank.
- "Estimated Deadline" — same as Deadline if present, else blank.
- "Company Name" — the hiring organisation abbreviation expanded where obvious, e.g. IRC → International Rescue Committee.
- For fields not found leave empty string "".
- NEVER invent information not in the text or metadata.
"""

    user_msg = f"""METADATA (from website table):
Title:       {meta.get('title', '')}
Organisation:{meta.get('org', '')}
Location:    {meta.get('location', '')}
Date Posted: {meta.get('date_posted', '')}
Deadline:    {meta.get('deadline', '')}
App URL:     {meta.get('app_url', '')}
Job URL:     {meta.get('job_url', '')}

PDF TEXT (first 4000 chars):
{raw_text[:4000]}
"""

    try:
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1200,
            system=system_prompt,
            messages=[{"role": "user", "content": user_msg}]
        )
        raw_json = response.content[0].text.strip()
        # Strip markdown fences if present
        raw_json = re.sub(r"^```(?:json)?\s*|```$", "", raw_json, flags=re.MULTILINE).strip()
        return json.loads(raw_json)
    except json.JSONDecodeError as e:
        print(f"  [!] JSON parse error: {e}")
    except Exception as e:
        print(f"  [!] AI extraction failed: {e}")

    # ── Fallback: populate from metadata only ──────────────────────────────────
    return {
        "Job Title":          meta.get("title", ""),
        "Job Location":       meta.get("location", ""),
        "Date Posted":        meta.get("date_posted", ""),
        "Deadline":           meta.get("deadline", ""),
        "Estimated Deadline": meta.get("deadline", ""),
        "Application":        meta.get("app_url", ""),
        "Company Name":       meta.get("org", ""),
        "Job URL":            meta.get("job_url", ""),
        "Job Description":    raw_text[:500] if raw_text else "",
        **{k: "" for k in OUTPUT_COLUMNS if k not in [
            "Job Title", "Job Location", "Date Posted", "Deadline",
            "Estimated Deadline", "Application", "Company Name",
            "Job URL", "Job Description",
        ]}
    }


def save_excel(records: list[dict], path: str):
    """Save structured records to a nicely formatted Excel file."""
    df = pd.DataFrame(records, columns=OUTPUT_COLUMNS)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Jobs")
        ws = writer.sheets["Jobs"]

        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        header_fill = PatternFill(fill_type="solid", fgColor="1F4E79")
        header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        thin_border = Border(
            left=Side(style="thin"),  right=Side(style="thin"),
            top=Side(style="thin"),   bottom=Side(style="thin"),
        )

        for col_idx, col_name in enumerate(OUTPUT_COLUMNS, start=1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill   = header_fill
            cell.font   = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        data_font = Font(name="Arial", size=9)
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.font      = data_font
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border    = thin_border

        col_widths = {
            "Job Title": 35, "Job Type": 15, "Job Qualifications": 40,
            "Job Experience": 25, "Job Location": 25, "Job Field": 18,
            "Date Posted": 14, "Deadline": 14, "Job Description": 60,
            "Application": 35, "Company URL": 30, "Company Name": 25,
            "Company Logo": 15, "Company Industry": 20, "Company Founded": 15,
            "Company Type": 15, "Company Website": 30, "Company Address": 30,
            "Company Details": 35, "Job URL": 40, "Estimated Deadline": 18,
            "Salary Range": 20,
        }
        for col_idx, col_name in enumerate(OUTPUT_COLUMNS, start=1):
            ws.column_dimensions[get_column_letter(col_idx)].width = col_widths.get(col_name, 20)

        ws.row_dimensions[1].height = 30
        ws.freeze_panes = "A2"

    print(f"\n✅  Saved {len(records)} jobs → {path}")


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    PDF_DIR.mkdir(exist_ok=True)
    client = anthropic.Anthropic()   # reads ANTHROPIC_API_KEY from environment

    print(f"[1/4] Fetching jobs page: {JOBS_URL}")
    html = fetch_page(JOBS_URL)
    if not html:
        print("ERROR: Could not fetch the jobs page. Check your network and try again.")
        return

    print("[2/4] Parsing jobs table …")
    jobs = parse_jobs_table(html)
    print(f"      Found {len(jobs)} job listings.")

    if not jobs:
        print("No jobs found — the page structure may have changed.")
        # Dump HTML snippet for debugging
        debug_path = Path("debug_page.html")
        debug_path.write_text(html[:5000], encoding="utf-8")
        print(f"      First 5000 chars saved to {debug_path} for inspection.")
        return

    records = []
    print("[3/4] Downloading PDFs and extracting data …")

    for i, job in enumerate(jobs, start=1):
        title   = job["title"]
        pdf_url = job["pdf_url"]
        print(f"\n  [{i}/{len(jobs)}] {title}")

        raw_text = ""
        if pdf_url:
            # Build a safe filename from the URL's last path segment
            url_slug = re.sub(r"[^\w\-.]", "_", pdf_url.split("/")[-1])[:80]
            if not url_slug.lower().endswith(".pdf"):
                url_slug += ".pdf"
            pdf_path = PDF_DIR / url_slug

            ok = download_pdf(pdf_url, pdf_path)
            if ok:
                raw_text = extract_pdf_text(pdf_path)
                print(f"      PDF: {len(raw_text)} chars extracted from {pdf_path.name}")
            else:
                print("      PDF download failed — using metadata only")
        else:
            print("      No PDF link found — using metadata only")

        print("      Calling Claude to extract fields …")
        extracted = ai_extract_fields(raw_text, job, client)

        # Ensure every output column is present
        row = {col: extracted.get(col, "") for col in OUTPUT_COLUMNS}
        records.append(row)

        # Be polite to servers
        time.sleep(0.5)

    print("\n[4/4] Saving to Excel …")
    save_excel(records, OUTPUT_XLS)
    print(f"\nDone! Open '{OUTPUT_XLS}' to see all {len(records)} jobs.\n")


if __name__ == "__main__":
    main()
