"""
MIMU Jobs PDF Extractor — Full Version
=======================================
- Reads PDF URLs from published Google Sheet CSV
- Downloads & extracts full PDF text
- Standardised Job Field / Qualification / Experience bands
- Visits org website to fill missing company info + logo
- Saves mimu_jobs.csv + mimu_jobs.xlsx

REQUIREMENTS:
    pip install requests pdfplumber pandas openpyxl beautifulsoup4

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
from bs4 import BeautifulSoup

# ── Config ─────────────────────────────────────────────────────────────────────

SHEET_CSV_URL = (
    "https://docs.google.com/spreadsheets/d/e/"
    "2PACX-1vS9_Zza6If2T_LT-iVvQjBTIFTeqt_OVBws70v_s3NJavT-ZosZ28qtE7xds7iS5rLmU2UbhzxWnOsY"
    "/pub?gid=964760760&single=true&output=csv"
)

OUTPUT_CSV  = "mimu_jobs.csv"
OUTPUT_XLSX = "mimu_jobs.xlsx"

OUTPUT_COLUMNS = [
    "Job Title", "Job Type", "Job Qualifications", "Job Experience",
    "Job Location", "Job Field", "Date Posted", "Deadline",
    "Job Description", "Application", "Company URL", "Company Name",
    "Company Logo", "Company Industry", "Company Founded", "Company Type",
    "Company Website", "Company Address", "Company Details",
    "Job URL", "Estimated Deadline", "Salary Range",
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
}

# Skip these domains when looking for org websites
APP_PLATFORM_DOMAINS = [
    'smrtr.io', 'workday', 'myworkday', 'bamboohr', 'greenhouse.io',
    'lever.co', 'forms.office', 'google.com/forms', 'hr-manager',
    'smartrecruiters', 'jobs.', 'careers.', 'themimu.info',
    'candidate.', 'apply.', 'recruiting.',
]

# =============================================================================
#  STANDARDISED JOB FIELD
# =============================================================================

FIELD_KEYWORD_MAP = [
    ("Information Technology",
     ["software engineer","developer","devops","frontend","backend","full stack","fullstack",
      "sysadmin","cloud","cybersecurity","data engineer","machine learning","artificial intelligence",
      "ai/ml","it support","network engineer","database administrator","kubernetes","docker",
      "react","node.js","python developer","java developer","ict officer","ict manager"],
     ["programming","coding","api","agile","scrum","git","linux","server","infrastructure","software"]),

    ("Finance & Accounting",
     ["accountant","auditor","finance manager","financial analyst","cfo","treasurer","tax",
      "bookkeeper","payroll","budget analyst","credit analyst","investment","portfolio manager",
      "risk analyst","actuary","acca","cfa","cpa","finance officer","grants officer",
      "budget officer","financial management"],
     ["financial","accounting","balance sheet","p&l","reconciliation","ifrs","gaap",
      "ledger","invoicing","grants","budget","donor funds"]),

    ("Sales & Business Development",
     ["sales executive","sales manager","business development","account manager",
      "sales representative","bd manager","regional sales","key account","sales director",
      "commercial manager","sales officer","demand generator"],
     ["revenue","pipeline","crm","leads","prospects","quota","target","upsell","b2b","b2c"]),

    ("Marketing & Communications",
     ["marketing manager","digital marketing","seo","sem","content marketer","social media manager",
      "brand manager","marketing executive","communications manager","pr manager","copywriter",
      "growth hacker","email marketing","campaign manager","communications officer"],
     ["marketing","branding","advertising","social media","content","campaign","analytics",
      "google ads","facebook ads","influencer","public relations"]),

    ("Human Resources",
     ["hr manager","human resources","recruiter","talent acquisition","hr business partner",
      "hrbp","hr officer","compensation","benefits manager","organisational development",
      "learning and development","l&d","hr generalist","payroll manager","hr coordinator",
      "hr assistant","hr & admin"],
     ["recruitment","onboarding","performance management","employee relations","hr","workforce",
      "personnel","staffing","hiring"]),

    ("Healthcare & Medicine",
     ["doctor","physician","nurse","pharmacist","medical officer","surgeon","anaesthetist",
      "physiotherapist","radiographer","lab technician","clinical officer","healthcare manager",
      "occupational therapist","dentist","midwife","health officer","nurse counsellor",
      "medical assistant","community health","health specialist","nutrition specialist",
      "x-ray technician","cxr technician","health and nutrition","wash officer",
      "wash expert","emergency wash","foot unit worker","domestic health"],
     ["hospital","clinic","patient","medical","health","pharmaceutical","diagnosis","treatment",
      "tb","hiv","malaria","nutrition","wash","sanitation","hygiene","epidemic"]),

    ("Protection & Social Work",
     ["protection officer","protection assistant","gbv officer","social worker","case manager",
      "community mobilizer","community development officer","welfare officer","safeguarding",
      "child protection"],
     ["protection","gbv","gender","child protection","safeguarding","community","welfare",
      "beneficiary","case management","psychosocial"]),

    ("Education & Training",
     ["teacher","lecturer","professor","trainer","educator","tutor","school principal",
      "academic","curriculum","e-learning","instructional designer","teaching assistant",
      "training coordinator"],
     ["school","university","college","classroom","students","pedagogy","curriculum","education",
      "training","capacity building","learning"]),

    ("Logistics & Supply Chain",
     ["supply chain manager","logistics coordinator","warehouse manager","fleet manager",
      "procurement manager","purchasing manager","import export","freight","shipping coordinator",
      "inventory manager","demand planner","logistics officer"],
     ["logistics","supply chain","warehouse","inventory","freight","procurement","sourcing",
      "transport","fleet","distribution"]),

    ("Engineering & Construction",
     ["mechanical engineer","civil engineer","electrical engineer","structural engineer",
      "process engineer","project engineer","maintenance engineer","production engineer",
      "quality engineer","safety engineer","site engineer","design engineer","quantity surveyor",
      "site supervisor","architect","draughtsman","building inspector","construction manager",
      "civil engineering specialist"],
     ["engineering","cad","autocad","solidworks","manufacturing","plant","machinery",
      "construction","building","site","contractor","infrastructure"]),

    ("Administration & Operations",
     ["office manager","executive assistant","administrative officer","operations manager",
      "personal assistant","receptionist","data entry","office administrator","admin officer",
      "company secretary","admin and finance","admin coordinator","housekeeper","office assistant",
      "program coordinator","programme coordinator","programme manager","project coordinator",
      "project officer","field officer","project support","branch manager"],
     ["administration","operations","office","coordination","scheduling","reporting","clerical",
      "planning","implementation","monitoring","reporting"]),

    ("Legal",
     ["lawyer","attorney","legal counsel","paralegal","compliance officer","legal advisor",
      "solicitor","barrister","corporate counsel","legal manager","contract manager"],
     ["legal","law","contracts","litigation","regulatory","compliance","gdpr","policy"]),

    ("Research & Data",
     ["research scientist","data scientist","lab researcher","research analyst",
      "clinical researcher","environmental scientist","chemist","biologist","statistician",
      "data analyst","data assistant","m&e","monitoring and evaluation","data analytics"],
     ["research","analysis","data","laboratory","science","experiment","findings","methodology",
      "monitoring","evaluation","indicators","log frame","data collection"]),

    ("Media & Journalism",
     ["journalist","editor","reporter","broadcast","news anchor","content creator",
      "media manager","radio","television","producer","scriptwriter"],
     ["media","journalism","broadcast","news","editorial","publishing","press"]),

    ("Security",
     ["security officer","security guard","security manager","cctv","loss prevention",
      "risk manager","health and safety","hse officer","osh","fire safety",
      "national security support"],
     ["security","safety","risk","surveillance","patrol","access control","emergency"]),
]

def infer_job_field(title: str, description: str) -> str:
    if not title and not description:
        return "Other"
    combined = ((title or "") + " " + (description or "")).lower()
    best_field, best_score = "Other", 0
    for label, high_keys, supporting in FIELD_KEYWORD_MAP:
        score  = sum(3 for k in high_keys      if k in combined)
        score += sum(1 for k in supporting     if k in combined)
        if score > best_score:
            best_score, best_field = score, label
    return best_field if best_score >= 2 else "Other"

# =============================================================================
#  STANDARDISED QUALIFICATIONS
# =============================================================================

QUALIFICATION_TIERS = [
    ("PhD / Doctorate",
     ["phd","ph.d","doctorate","doctoral","doctor of philosophy"]),
    ("Master's Degree",
     ["master","msc","m.sc","mba","m.b.a","meng","m.eng","mphil","postgraduate",
      "post-graduate","post graduate","master of"]),
    ("Bachelor's Degree",
     ["bachelor","bsc","b.sc","b.a ","beng","b.eng","bcom","b.com","bba","llb",
      "degree in","undergraduate","honours","hons","b.med","mbbs","m.b.,b.s"]),
    ("Higher National Diploma",
     ["hnd","hnc","higher national diploma","higher national certificate",
      "higher diploma","advanced diploma"]),
    ("Diploma",
     ["diploma","dip ","dip.","associate degree","foundation degree","lcci"]),
    ("Professional Certification",
     ["acca","cpa","cfa","cima","pmp","prince2","cissp","aws certified","comptia",
      "cisco","ccna","ccnp","shrm","cipd","chartered","certified public",
      "certified financial","certified project","professional certification",
      "professional certificate","b.med.tech"]),
    ("A-Levels / High School",
     ["a-level","a level","hsc","higher school certificate","ib diploma",
      "international baccalaureate","gce advanced","high school","secondary school"]),
    ("No Formal Qualification Required",
     ["no qualification","no degree","no formal","school leaver","entry level",
      "no experience required","training provided","will train","primary school",
      "minimum high school"]),
]

def extract_qualification(text: str) -> str:
    if not text:
        return ""
    lower = text.lower()
    for label, keywords in QUALIFICATION_TIERS:
        if any(k in lower for k in keywords):
            return label
    return ""

# =============================================================================
#  STANDARDISED EXPERIENCE
# =============================================================================

NO_EXP_KW = [
    "no experience","no prior experience","fresh graduate","freshers","entry level",
    "entry-level","0 years","zero experience","training provided","will train",
    "no experience required","open to fresh",
]
LESS1_KW = [
    "less than 1 year","under 1 year","6 months","less than a year",
    "some experience","minimal experience",
]

def years_to_band(n: int) -> str:
    if n <= 0:  return "No Experience Required"
    if n <= 2:  return "1 - 2 Years"
    if n <= 5:  return "3 - 5 Years"
    if n <= 10: return "6 - 10 Years"
    return "10+ Years"

def extract_experience(text: str) -> str:
    if not text:
        return ""
    lower = text.lower()
    if any(k in lower for k in NO_EXP_KW):
        return "No Experience Required"
    if any(k in lower for k in LESS1_KW):
        return "Less than 1 Year"
    patterns = [
        r"(\d+)\s*[-–to]+\s*(\d+)\s*\+?\s*years?",
        r"(\d+)\s*\+\s*years?\s*(?:of\s+)?(?:experience)?",
        r"(?:minimum|at\s+least|over|more\s+than)\s+(\d+)\s*\+?\s*years?",
        r"(\d+)\s*years?\s*(?:of\s+)?(?:relevant\s+)?(?:work\s+)?experience",
        r"experience\s*(?:of\s+)?(\d+)\s*years?",
    ]
    for p in patterns:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            raw = int(m.group(1))
            if 0 < raw <= 25:
                return years_to_band(raw)
    return ""

# =============================================================================
#  STANDARDISED JOB TYPE
# =============================================================================

def detect_job_type(text: str, title: str) -> str:
    combined = (text + " " + title).lower()
    if any(k in combined for k in ["intern", "internship"]):
        return "Internship"
    if any(k in combined for k in ["part-time", "part time"]):
        return "Part-time"
    if any(k in combined for k in ["volunteer"]):
        return "Volunteer"
    if any(k in combined for k in ["consultant", "consultancy", "terms of reference", "tor "]):
        return "Consultancy / Contract"
    if any(k in combined for k in ["contract", "fixed-term", "fixed term", "temporary", "service agreement"]):
        return "Consultancy / Contract"
    return "Full-time"

# =============================================================================
#  COMPANY TYPE
# =============================================================================

def detect_company_type(text: str) -> str:
    tl = text.lower()
    if re.search(r'\bngo\b|non.governmental|nonprofit|non-profit', tl):
        return "NGO / Non-Profit"
    if re.search(r'\bundp\b|\bunicef\b|\bwfp\b|\bunhcr\b|\bwho\b|\bilo\b|united nations|un agency|\bioм\b|\biom\b|\bunesco\b', tl):
        return "UN Agency"
    if re.search(r'\bgovernment\b|ministry of|department of|\bgovernmental\b', tl):
        return "Government"
    if re.search(r'\bingo\b|international ngо|international non', tl):
        return "INGO"
    if re.search(r'\bprivate\b|\bltd\b|\blimited\b|\bcorporation\b|\binc\b|\bplc\b', tl):
        return "Private Sector"
    return ""

# =============================================================================
#  SAFE STRING HELPER
# =============================================================================

def s(val) -> str:
    if val is None:
        return ""
    if isinstance(val, float) and math.isnan(val):
        return ""
    return str(val).strip()

# =============================================================================
#  HTTP HELPERS
# =============================================================================

def get_html(url: str, timeout: int = 12) -> str:
    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout)
        if resp.status_code == 200:
            return resp.text
    except Exception:
        pass
    return ""

def fetch_pdf_text(pdf_url: str) -> str:
    if not pdf_url or not pdf_url.startswith("http"):
        return ""
    try:
        resp = requests.get(pdf_url, headers=HEADERS, timeout=30)
        if resp.status_code != 200:
            print(f"      ✗ HTTP {resp.status_code}")
            return ""
        with pdfplumber.open(io.BytesIO(resp.content)) as pdf:
            pages = [p.extract_text() for p in pdf.pages if p.extract_text()]
            return "\n".join(pages)
    except Exception as e:
        print(f"      ✗ PDF error: {e}")
        return ""

# =============================================================================
#  WEBSITE SCRAPER — fills missing company fields
# =============================================================================

def scrape_website(url: str) -> dict:
    """
    Visit the org website and extract:
      - description / about text
      - logo URL
      - founded year
      - address
      - company type hints
    Returns a dict with keys: description, logo, founded, address, company_type
    """
    result = {"description": "", "logo": "", "founded": "", "address": "", "company_type": ""}
    if not url or not url.startswith("http"):
        return result

    html = get_html(url)
    if not html:
        return result

    soup = BeautifulSoup(html, "html.parser")
    base = re.match(r'(https?://[^/]+)', url)
    base_url = base.group(1) if base else ""

    # ── Logo ──────────────────────────────────────────────────────────────────
    logo = ""
    # 1. OpenGraph image
    og = soup.find("meta", property="og:image") or soup.find("meta", attrs={"name": "og:image"})
    if og and og.get("content"):
        logo = og["content"]
    # 2. img with "logo" in class/id/src/alt
    if not logo:
        for img in soup.find_all("img"):
            src = img.get("src", "")
            alt = img.get("alt", "")
            cls = " ".join(img.get("class", []))
            iid = img.get("id", "")
            if re.search(r'logo|brand|emblem', src + alt + cls + iid, re.IGNORECASE):
                logo = src
                break
    # 3. Favicon as fallback
    if not logo:
        icon = soup.find("link", rel=lambda r: r and "icon" in " ".join(r).lower())
        if icon and icon.get("href"):
            logo = icon["href"]

    # Resolve relative logo URL
    if logo:
        if logo.startswith("//"):
            logo = "https:" + logo
        elif logo.startswith("/"):
            logo = base_url + logo
        elif not logo.startswith("http"):
            logo = base_url + "/" + logo
    result["logo"] = logo

    # ── Description / About ───────────────────────────────────────────────────
    # Try meta description first
    meta_desc = soup.find("meta", attrs={"name": "description"}) or \
                soup.find("meta", property="og:description")
    if meta_desc and meta_desc.get("content") and len(meta_desc["content"]) > 30:
        result["description"] = meta_desc["content"][:500]
    else:
        # Look for about/mission sections
        for tag in soup.find_all(["p", "div", "section"], limit=50):
            txt = tag.get_text(" ", strip=True)
            if len(txt) > 80 and re.search(
                r'mission|vision|about us|who we are|established|founded|our work|we are',
                txt, re.IGNORECASE
            ):
                result["description"] = txt[:500]
                break

    # ── Founded year ──────────────────────────────────────────────────────────
    text_body = soup.get_text(" ")
    m = re.search(r'(?:established|founded|since|incorporated)\s+(?:in\s+)?(\d{4})', text_body, re.IGNORECASE)
    if m:
        year = int(m.group(1))
        if 1900 <= year <= datetime.now().year:
            result["founded"] = m.group(1)

    # ── Address ───────────────────────────────────────────────────────────────
    for pat in [
        r'(No\.?\s*\d+[^\n,]{5,60}(?:Street|Road|Avenue|Lane|Township|Yangon|Mandalay|Myanmar)[^\n,]{0,50})',
        r'(\d+[,\s]+[A-Z][^\n]{15,100}(?:Street|Road|Township|Yangon|Myanmar))',
        r'(?:address|located at|office)[:\s]+([^\n]{15,150})',
    ]:
        m2 = re.search(pat, text_body, re.IGNORECASE)
        if m2:
            result["address"] = m2.group(1).strip()
            break

    # ── Company type hints ────────────────────────────────────────────────────
    result["company_type"] = detect_company_type(text_body)

    return result

# =============================================================================
#  TEXT FIELD EXTRACTORS
# =============================================================================

def search_pattern(text: str, patterns: list) -> str:
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return m.group(1).strip(" :.,\t-")
    return ""

def search_block(text: str, header_patterns: list, max_lines: int = 8) -> str:
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
    decoded = text.replace("%40", "@").replace("%2E", ".").replace("%2F", "/")
    emails = re.findall(r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}', decoded)
    skip = ['example.com', 'domain.com', 'email.com', 'yourmail', 'sentry', 'noreply']
    for email in emails:
        if not any(sk in email for sk in skip):
            return email
    return ""

def extract_apply_url(text: str) -> str:
    urls = re.findall(r'https?://[^\s\'"<>)]+', text)
    # Prefer known ATS / apply links
    for url in urls:
        u = url.rstrip('.,)')
        if re.search(r'apply|career|recruit|workday|bamboo|greenhouse|lever|smartrecruiters|smrtr|hr-manager|myworkday', u, re.IGNORECASE):
            return u
    return ""

def extract_application(text: str, existing: str) -> str:
    # 1. Good existing URL
    if existing and existing.startswith("http") and not existing.endswith("/"):
        if re.search(r'apply|career|recruit|workday|bamboo|smrtr|hr-manager|myworkday|forms', existing, re.IGNORECASE):
            return existing

    decoded = text.replace("%40", "@").replace("%2E", ".").replace("%2F", "/")

    # 2. Email near application instructions
    app_section = ""
    for pat in [
        r'(?:to apply|how to apply|application process|submit|send cv|send resume)[^\n]{0,300}',
        r'(?:interested candidates?)[^\n]{0,300}',
        r'(?:please send|please submit|kindly send|applications? (?:should be )?sent)[^\n]{0,300}',
        r'(?:contact us|for more information)[^\n]{0,300}',
    ]:
        m = re.search(pat, decoded, re.IGNORECASE)
        if m:
            app_section += m.group(0) + " "

    email = extract_emails(app_section) or extract_emails(decoded)
    if email:
        return email

    # 3. ATS / apply URL
    url = extract_apply_url(app_section) or extract_apply_url(decoded)
    if url:
        return url

    # 4. Any existing value decoded
    if existing:
        decoded_ex = existing.replace("%40", "@").replace("%2E", ".").replace("%2F", "/").lstrip("/")
        if "@" in decoded_ex or decoded_ex.startswith("http"):
            return decoded_ex

    return ""

def extract_website(text: str, existing: str) -> str:
    urls = re.findall(r'https?://[^\s\'"<>)]+', text)
    candidates = []
    for url in urls:
        u = url.rstrip('.,)')
        if not any(d in u for d in APP_PLATFORM_DOMAINS):
            candidates.append(u)
    # Prefer short-path (homepage) URLs
    for url in candidates:
        path = re.sub(r'https?://[^/]+', '', url)
        if len(path) < 25:
            return url
    if candidates:
        return candidates[0]
    return s(existing)

def extract_address_from_text(text: str) -> str:
    for pat in [
        r'(?:address|office location|our office)[:\s]+([^\n]{15,150})',
        r'(?:located at|based at|head office)[:\s]+([^\n]{15,150})',
        r'(No\.?\s*\d+[^\n,]{5,60}(?:Street|Road|Avenue|Lane|Township|Yangon|Mandalay|Myanmar)[^\n,]{0,50})',
        r'(\d+[,\s]+[A-Z][^\n]{15,100}(?:Street|Road|Township|Yangon|Myanmar))',
    ]:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return ""

def detect_company_founded(text: str) -> str:
    m = re.search(r'(?:established|founded|since|incorporated)\s+(?:in\s+)?(\d{4})', text, re.IGNORECASE)
    if m:
        year = int(m.group(1))
        if 1900 <= year <= datetime.now().year:
            return m.group(1)
    return ""

def extract_company_details(text: str, org: str) -> str:
    org_esc = re.escape(org[:15]) if org else ""
    pats = []
    if org_esc:
        pats += [rf'about\s+{org_esc}', rf'{org_esc}[^\n]{{0,30}}\nis\s+(?:a|an|the)']
    pats += [r'about us', r'background', r'who we are', r'organization overview',
             r'about the organization', r'introduction']
    block = search_block(text, pats, 10)
    if block and len(block) > 40:
        return block[:500]
    return ""

# =============================================================================
#  VERBOSE PRINTER
# =============================================================================

def print_extracted(record: dict, pdf_text: str):
    pad = "      "
    div = pad + "-" * 58
    print(div)
    show = [
        ("Job Title", 70), ("Job Type", 25), ("Job Field", 35),
        ("Job Location", 70), ("Date Posted", 20), ("Deadline", 25),
        ("Salary Range", 70), ("Job Qualifications", 40), ("Job Experience", 25),
        ("Application", 100), ("Company Name", 60), ("Company Logo", 100),
        ("Company Type", 25), ("Company Founded", 10),
        ("Company Website", 100), ("Company Address", 100),
    ]
    for field, maxlen in show:
        val = record.get(field, "")
        if val:
            label = (field + ":").ljust(22)
            display = val[:maxlen] + ("…" if len(val) > maxlen else "")
            print(f"{pad}{label} {display}")
    for field in ("Job Description", "Company Details"):
        val = record.get(field, "")
        if val:
            print(f"{pad}{(field+':').ljust(22)} {val[:160]}{'…' if len(val)>160 else ''}")
    if pdf_text:
        snippet = " ".join(pdf_text[:500].split())
        print(f"\n{pad}--- PDF SNIPPET ---")
        print(f"{pad}{snippet[:500]}")
    print(div)

# =============================================================================
#  MAIN FIELD PARSER
# =============================================================================

def parse_pdf_fields(text: str, row: dict, website_cache: dict) -> dict:
    """Parse all fields from PDF text + sheet row. Fills gaps from org website."""

    title    = s(row.get("Job Title", "")) or search_pattern(text, [
        r'position[:\s]+([^\n]{3,80})', r'job title[:\s]+([^\n]{3,80})',
    ])
    location = s(row.get("Job Location", "")) or search_pattern(text, [
        r'location[:\s]+([^\n]{3,80})', r'duty station[:\s]+([^\n]{3,80})',
        r'place of work[:\s]+([^\n]{3,80})', r'based in[:\s]+([^\n]{3,60})',
    ])
    deadline = s(row.get("Deadline", "")) or search_pattern(text, [
        r'closing date[:\s]+([^\n]{3,40})',
        r'application deadline[:\s]+([^\n]{3,40})',
        r'deadline[:\s]+([^\n]{3,40})',
        r'submit (?:by|before)[:\s]+([^\n]{3,40})',
    ])

    # ── Qualifications ────────────────────────────────────────────────────────
    quals_raw = search_block(text, [
        r'qualif', r'academic requirement', r'minimum requirement',
        r'education requirement', r'degree required',
    ], 6)
    if not quals_raw:
        quals_raw = search_pattern(text, [
            r'(?:qualifications?|educational? requirements?)[:\s]+([^\n]{10,250})',
            r"(bachelor'?s?|master'?s?|phd|diploma|degree|b\.med|mbbs)[^\n]{0,120}",
            r'(minimum\s+(?:diploma|degree|bachelor|high school)[^\n]{0,100})',
        ])
    if not quals_raw:
        quals_raw = s(row.get("Job Qualifications", ""))
    quals_standard = extract_qualification(quals_raw) or extract_qualification(title)
    quals_out = quals_standard or quals_raw[:300]

    # ── Experience ────────────────────────────────────────────────────────────
    exp_raw = search_pattern(text, [
        r'(\d+\+?\s*(?:to\s*\d+\s*)?years?\s+(?:of\s+)?(?:relevant\s+)?(?:work\s+)?experience[^\n]{0,80})',
        r'(minimum\s+(?:of\s+)?\d+\s+years?[^\n]{0,80})',
        r'(at least\s+\d+\s+years?[^\n]{0,80})',
        r'experience[:\s]+([^\n]{5,100})',
    ])
    if not exp_raw:
        exp_raw = s(row.get("Job Experience", ""))
    exp_out = extract_experience(exp_raw + " " + text[:2000]) or exp_raw[:200]

    # ── Description ───────────────────────────────────────────────────────────
    desc = search_block(text, [
        r'key responsibilit', r'main responsibilit', r'responsibilit',
        r'duties and responsibilit', r'key tasks', r'scope of work',
        r'main duties', r'job purpose', r'objective', r'role summary',
        r'key deliverables',
    ], 12)
    if not desc:
        long_lines = [l.strip() for l in text.split("\n") if len(l.strip()) > 50]
        desc = " ".join(long_lines[:5])
    if not desc:
        desc = s(row.get("Job Description", ""))
    desc = str(desc or "")[:600]

    # ── Salary ────────────────────────────────────────────────────────────────
    salary = search_pattern(text, [
        r'salary[:\s]+([^\n]{5,80})',
        r'remuneration[:\s]+([^\n]{5,80})',
        r'compensation[:\s]+([^\n]{5,80})',
        r'([\d,]+\s*(?:MMK|USD|Ks|Kyats?)[^\n]{0,40})',
        r'(\$[\d,]+[^\n]{0,30})',
        r'(competitive salary[^\n]{0,60})',
        r'package[:\s]+([^\n]{5,60})',
    ]) or s(row.get("Salary Range", ""))

    # ── Application ───────────────────────────────────────────────────────────
    app_out = extract_application(text, s(row.get("Application", "")))

    # ── Website ───────────────────────────────────────────────────────────────
    existing_web = s(row.get("Company Website", "")) or s(row.get("Company URL", ""))
    website = extract_website(text, existing_web)

    # ── Address from PDF ──────────────────────────────────────────────────────
    address = extract_address_from_text(text) or s(row.get("Company Address", ""))

    # ── Company details from PDF ──────────────────────────────────────────────
    org = s(row.get("Company Name", ""))
    details = extract_company_details(text, org) or s(row.get("Company Details", ""))

    # ── Founded from PDF ──────────────────────────────────────────────────────
    founded = detect_company_founded(text)

    # ── Job field & type ──────────────────────────────────────────────────────
    job_field = infer_job_field(title, text[:3000])
    job_type  = detect_job_type(text, title)
    comp_type = detect_company_type(text) or s(row.get("Company Type", ""))

    logo = ""

    # ── Fill gaps from website ────────────────────────────────────────────────
    needs_web = (not details or len(details) < 50 or not address or
                 not founded or not comp_type or not logo)

    if website and needs_web:
        cached = website_cache.get(website)
        if cached is None:
            print(f"      🌐 Scraping website: {website}")
            cached = scrape_website(website)
            website_cache[website] = cached
            time.sleep(0.5)
        else:
            print(f"      🌐 Using cached website data")

        if not details or len(details) < 50:
            details = cached.get("description", "") or details
        if not address:
            address = cached.get("address", "")
        if not founded:
            founded = cached.get("founded", "")
        if not comp_type:
            comp_type = cached.get("company_type", "")
        logo = cached.get("logo", "")

    return {
        "Job Title":          str(title or ""),
        "Job Type":           str(job_type or ""),
        "Job Qualifications": str(quals_out or ""),
        "Job Experience":     str(exp_out or ""),
        "Job Location":       str(location or ""),
        "Job Field":          str(job_field or ""),
        "Date Posted":        str(s(row.get("Date Posted", ""))),
        "Deadline":           str(deadline or ""),
        "Job Description":    str(desc or ""),
        "Application":        str(app_out or ""),
        "Company URL":        str(website or ""),
        "Company Name":       str(org or ""),
        "Company Logo":       str(logo or ""),
        "Company Industry":   str(job_field or ""),
        "Company Founded":    str(founded or ""),
        "Company Type":       str(comp_type or ""),
        "Company Website":    str(website or ""),
        "Company Address":    str(address or ""),
        "Company Details":    str(details or "")[:500],
        "Job URL":            str(s(row.get("Job URL", ""))),
        "Estimated Deadline": str(deadline or ""),
        "Salary Range":       str(salary or ""),
    }

# =============================================================================
#  MAIN
# =============================================================================

def main():
    print("=" * 62)
    print("  MIMU Jobs PDF Extractor")
    print(f"  Started : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 62)

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

    # 2. Process
    print(f"\n[2/3] Downloading PDFs and extracting fields …")
    records       = []
    website_cache = {}   # avoid re-scraping same org site
    total         = len(df)
    pdf_ok        = 0
    pdf_skip      = 0

    for idx, row in df.iterrows():
        pdf_url = s(row.get("PDF URL", ""))
        title   = s(row.get("Job Title", ""))
        num     = idx + 1

        print(f"\n  [{num}/{total}] {title}")

        if pdf_url and pdf_url.startswith("http"):
            print(f"      PDF : {pdf_url}")
            pdf_text = fetch_pdf_text(pdf_url)
            if pdf_text:
                print(f"      ✓ {len(pdf_text):,} chars extracted")
                pdf_ok += 1
            else:
                print(f"      ⚠ No text from PDF")
                pdf_skip += 1
        else:
            print(f"      ⚠ No PDF URL — sheet data only")
            pdf_text = ""
            pdf_skip += 1

        enriched = parse_pdf_fields(pdf_text, row.to_dict(), website_cache)
        print_extracted(enriched, pdf_text)
        records.append(enriched)

        time.sleep(0.2)

    # 3. Save
    print(f"\n[3/3] Saving output files …")
    out_df = pd.DataFrame(records, columns=OUTPUT_COLUMNS)

    out_df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
    print(f"      ✓ {OUTPUT_CSV}")

    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False, sheet_name="MIMU Jobs")
        ws = writer.sheets["MIMU Jobs"]
        for col in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)
        ws.freeze_panes = "A2"
    print(f"      ✓ {OUTPUT_XLSX}")

    print(f"\n{'=' * 62}")
    print(f"  ✅ COMPLETE")
    print(f"     Total jobs    : {total}")
    print(f"     PDF success   : {pdf_ok}")
    print(f"     Sheet-only    : {pdf_skip}")
    print(f"     Websites hit  : {len(website_cache)}")
    print(f"     Output        : {OUTPUT_CSV}, {OUTPUT_XLSX}")
    print(f"     Finished      : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 62)


if __name__ == "__main__":
    main()
