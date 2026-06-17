"""
MIMU Jobs PDF Extractor — v4
==============================
Fixes vs v3:
  - Logo extraction: pulls the org logo directly from the embedded images
    in the PDF (page 1, top 35% of page, logo-shaped dimensions).
    Website logo is used only as fallback if PDF has no suitable image.
    Logo is stored as a data URI (data:image/png;base64,...) so it works
    without any external URL.
  - fetch_pdf() replaces fetch_pdf_text() — downloads once, returns both
    text (via pdfplumber) and raw bytes (for PyMuPDF image extraction).
  - Requires: pip install requests pdfplumber pymupdf pandas openpyxl beautifulsoup4

REQUIREMENTS:
    pip install requests pdfplumber pymupdf pandas openpyxl beautifulsoup4

USAGE:
    python mimu_jobs_extractor.py
"""

import requests
import pdfplumber
import fitz          # PyMuPDF — for embedded image extraction
import pandas as pd
import re
import io
import base64
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

# Domains that are ATS/application platforms, not org websites
APP_PLATFORM_DOMAINS = [
    'smrtr.io', 'workday', 'myworkday', 'bamboohr', 'greenhouse.io',
    'lever.co', 'forms.office', 'google.com/forms', 'hr-manager',
    'smartrecruiters', 'themimu.info', 'candidate.', 'apply.',
    'recruiting.', 'jobs.tdh.org', 'theirc.wd1', 'worldvision.wd1',
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
      "budget officer","financial management","head, finance","head of finance"],
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
      "growth hacker","email marketing","campaign manager","communications officer","gedsi"],
     ["marketing","branding","advertising","social media","content","campaign","analytics",
      "google ads","facebook ads","influencer","public relations","gender equality"]),

    ("Human Resources",
     ["hr manager","human resources","recruiter","talent acquisition","hr business partner",
      "hrbp","hr officer","compensation","benefits manager","organisational development",
      "learning and development","l&d","hr generalist","payroll manager","hr coordinator",
      "hr assistant","hr & admin","hr and admin"],
     ["recruitment","onboarding","performance management","employee relations","hr","workforce",
      "personnel","staffing","hiring"]),

    ("Healthcare & Medicine",
     ["doctor","physician","nurse","pharmacist","medical officer","surgeon","anaesthetist",
      "physiotherapist","radiographer","lab technician","clinical officer","healthcare manager",
      "occupational therapist","dentist","midwife","health officer","nurse counsellor",
      "medical assistant","community health","health specialist","nutrition specialist",
      "x-ray technician","cxr technician","health and nutrition","wash officer",
      "wash expert","emergency wash","foot unit worker","domestic health",
      "team leader","community mobilizer","demand generator","program coordinator"],
     ["hospital","clinic","patient","medical","health","pharmaceutical","diagnosis","treatment",
      "tb","hiv","malaria","nutrition","wash","sanitation","hygiene","epidemic",
      "reproductive health","harm reduction","community health"]),

    ("Protection & Social Work",
     ["protection officer","protection assistant","gbv officer","social worker","case manager",
      "community mobilizer","community development officer","welfare officer","safeguarding",
      "child protection","field officer"],
     ["protection","gbv","gender","child protection","safeguarding","community","welfare",
      "beneficiary","case management","psychosocial","displacement","idp"]),

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
      "project officer","field officer","project support","branch manager","admin and hr"],
     ["administration","operations","office","coordination","scheduling","reporting","clerical",
      "planning","implementation","monitoring","reporting"]),

    ("Legal",
     ["lawyer","attorney","legal counsel","paralegal","compliance officer","legal advisor",
      "solicitor","barrister","corporate counsel","legal manager","contract manager"],
     ["legal","law","contracts","litigation","regulatory","compliance","gdpr","policy"]),

    ("Research & Data",
     ["research scientist","data scientist","lab researcher","research analyst",
      "clinical researcher","environmental scientist","chemist","biologist","statistician",
      "data analyst","data assistant","m&e","monitoring and evaluation","data analytics",
      "assistant m&e manager"],
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
#  STANDARDISED QUALIFICATIONS — strict mapping only
# =============================================================================

QUALIFICATION_TIERS = [
    ("PhD / Doctorate",
     ["phd","ph.d","doctorate","doctoral","doctor of philosophy"]),
    ("Master's Degree",
     ["master","msc","m.sc","mba","m.b.a","meng","m.eng","mphil","postgraduate",
      "post-graduate","post graduate","master of"]),
    ("Bachelor's Degree",
     ["bachelor","bsc","b.sc","b.a ","beng","b.eng","bcom","b.com","bba","llb",
      "degree in","undergraduate","honours","hons","b.med","mbbs","m.b.,b.s",
      "b.med.tech","any graduate","be a graduate","be graduate"]),
    ("Higher National Diploma",
     ["hnd","hnc","higher national diploma","higher national certificate",
      "higher diploma","advanced diploma"]),
    ("Diploma",
     ["diploma","dip ","dip.","associate degree","foundation degree","lcci"]),
    ("Professional Certification",
     ["acca","cpa","cfa","cima","pmp","prince2","cissp","aws certified","comptia",
      "cisco","ccna","ccnp","shrm","cipd","chartered","certified public",
      "certified financial","certified project","professional certification",
      "professional certificate"]),
    ("A-Levels / High School",
     ["a-level","a level","hsc","higher school certificate","ib diploma",
      "international baccalaureate","gce advanced","high school","secondary school",
      "matric","matriculation","grade 10","grade 12","tenth standard","passed 10"]),
    ("No Formal Qualification Required",
     ["no qualification","no degree","no formal","school leaver","entry level",
      "no experience required","training provided","will train","primary school",
      "minimum high school","any education"]),
]

def extract_qualification(text: str) -> str:
    """Return standard tier label or empty string — never raw prose."""
    if not text:
        return ""
    lower = text.lower()
    for label, keywords in QUALIFICATION_TIERS:
        if any(k in lower for k in keywords):
            return label
    return ""

# =============================================================================
#  STANDARDISED EXPERIENCE — strict band-map only
# =============================================================================

NO_EXP_KW = [
    "no experience","no prior experience","fresh graduate","freshers","entry level",
    "entry-level","0 years","zero experience","training provided","will train",
    "no experience required","open to fresh",
]
LESS1_KW = [
    "less than 1 year","under 1 year","6 months","less than a year",
    "some experience","minimal experience","at least 6",
]

def years_to_band(n: int) -> str:
    if n <= 0:  return "No Experience Required"
    if n <= 2:  return "1 - 2 Years"
    if n <= 5:  return "3 - 5 Years"
    if n <= 10: return "6 - 10 Years"
    return "10+ Years"

def extract_experience(text: str) -> str:
    """Return standard band or empty string — never raw prose."""
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
        r"(\d+)\s*years?\s*(?:in|of)",
    ]
    for p in patterns:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            raw = int(m.group(1))
            if 0 < raw <= 25:
                return years_to_band(raw)
    return ""

# =============================================================================
#  STANDARDISED JOB TYPE — tighter matching
# =============================================================================

def detect_job_type(text: str, title: str) -> str:
    combined = (text + " " + title).lower()
    # Internship must be explicit — don't match "intern" inside other words
    if re.search(r'\bintern\b|\binternship\b', combined):
        return "Internship"
    if re.search(r'\bpart[-\s]time\b', combined):
        return "Part-time"
    # Volunteer: only count it if "volunteer" refers to the POSITION, not the org
    # e.g. "volunteer based organization" should NOT make the job type Volunteer
    if re.search(r'\bvolunteer\b', combined):
        # Check if it's describing the org ("volunteer based", "volunteers")
        org_vol = re.search(r'volunteer\s+(?:based|organization|member|network|corps)', combined)
        pos_vol = re.search(r'(?:position|role|post|contract|type)[^\n]{0,50}volunteer|volunteer\s+(?:position|role|post|opportunity)', combined)
        if pos_vol and not org_vol:
            return "Volunteer"
        if not org_vol:
            return "Volunteer"
    if re.search(r'\bconsultant\b|\bconsultancy\b|\bterms of reference\b|\btor\b|\bservice provider\b|\brfa\b|\brfp\b', combined):
        return "Consultancy / Contract"
    if re.search(r'\bcontract\b|\bfixed[-\s]term\b|\btemporary\b|\bservice agreement\b', combined):
        return "Consultancy / Contract"
    return "Full-time"

# =============================================================================
#  SALARY — extract numeric/currency amounts only
# =============================================================================

CURRENCY_PATTERNS = [
    # USD / GBP / EUR explicit
    r'USD\s*[\d,]+(?:\s*[-–]\s*[\d,]+)?(?:\s*/\s*\w+)?',
    r'GBP\s*[\d,]+(?:\s*[-–]\s*[\d,]+)?(?:\s*/\s*\w+)?',
    r'EUR\s*[\d,]+(?:\s*[-–]\s*[\d,]+)?(?:\s*/\s*\w+)?',
    r'\$\s*[\d,]+(?:\s*[-–]\s*\$?\s*[\d,]+)?(?:\s*/\s*\w+)?',
    r'£\s*[\d,]+(?:\s*[-–]\s*£?\s*[\d,]+)?(?:\s*/\s*\w+)?',
    # MMK / Kyat
    r'[\d,]+(?:\s*[-–]\s*[\d,]+)?\s*(?:MMK|Ks\.?|Kyats?)\b',
    r'MMK\s*[\d,]+(?:\s*[-–]\s*[\d,]+)?',
    # Range with "per month / year"
    r'[\d,]+(?:\s*[-–]\s*[\d,]+)?\s*/\s*(?:month|year|day|hour)',
    # Consultancy fee
    r'(?:fee|total fee)\s+(?:of\s+)?(?:USD|GBP|EUR|\$|£)?\s*[\d,]+',
]

def extract_salary(text: str, sheet_salary: str) -> str:
    """Return only numeric/currency amounts, never prose sentences."""
    for src in [text[:3000], sheet_salary]:
        if not src:
            continue
        for pat in CURRENCY_PATTERNS:
            m = re.search(pat, src, re.IGNORECASE)
            if m:
                val = m.group(0).strip().rstrip('.,')
                # Sanity: must contain a digit
                if re.search(r'\d', val):
                    return val
    return ""

# =============================================================================
#  COMPANY ADDRESS — strict address patterns only
# =============================================================================

# Words that indicate a value is prose, not an address — reject immediately
_PROSE_WORDS = re.compile(
    r'\b(?:please|ensure|must|will|should|have|been|with|through|across|'
    r'areas|regions|branches|townships|countries|programs|projects|staff|'
    r'services|sector|parent|duty|based|grade|report|department|during|'
    r'providing|working|seeking|implement|support|assist|manage)\b',
    re.IGNORECASE
)

# A valid address MUST contain one of these structural markers
_ADDR_MARKERS = re.compile(
    r'\b(?:Street|Road|Avenue|Lane|Quarter|Ward|Township|Yangon|Mandalay|'
    r'Nay\s*Pyi\s*Taw|Mawlamyine|Myanmar)\b',
    re.IGNORECASE
)

# Patterns that look like real addresses — all must also pass _ADDR_MARKERS
ADDRESS_PATTERNS = [
    # "No. X(B), Street Name, Ward/Quarter, Township"
    r'No\.?\s*\(?[A-Z0-9][A-Z0-9\-]*\)?\s*[,\s]+[^\n]{5,120}(?:Street|Road|Avenue|Lane)',
    # Hash-style: "#123, Street..."
    r'#\s*\d+[,\s]+[^\n]{5,100}(?:Street|Road|Avenue|Lane|Township|Yangon|Mandalay)',
    # "Address:" / "Office:" explicit label
    r'(?:^|\n)\s*(?:address|office address|head office|office)[:\s]+([A-Z0-9#No\.][^\n]{15,150})',
    # "Ward X, Township Y, City"  — only when it has at least two components
    r'Ward\s+(?:No\.?\s*)?\(?\w[\w\s]*\)?\s*,\s*[^\n]{5,80}(?:Township|Yangon|Mandalay|Myanmar)',
]

def extract_address(text: str) -> str:
    """
    Return a clean address string or empty string.
    Rules:
      - Must contain a structural address marker (Street/Road/Ward/Township/city)
      - Must contain at least one digit (building/house number)
      - Must NOT contain prose verbs or operational words
      - Must be under 180 chars
      - Must not be a lone number or single word
    """
    if not text:
        return ""
    for pat in ADDRESS_PATTERNS:
        for m in re.finditer(pat, text, re.IGNORECASE | re.MULTILINE):
            val = (m.group(1) if m.lastindex else m.group(0)).strip().strip('.,\n')
            # Clean up internal newlines
            val = re.sub(r'\s*\n\s*', ', ', val).strip()
            # Must have a digit (house/building number)
            if not re.search(r'\d', val):
                continue
            # Must have an address structural word
            if not _ADDR_MARKERS.search(val):
                continue
            # Reject prose sentences
            if _PROSE_WORDS.search(val):
                continue
            # Reject if too short (single number) or too long (paragraph)
            if len(val) < 10 or len(val) > 180:
                continue
            # Reject if it starts with just a number and nothing meaningful after
            if re.match(r'^\d+\s*$', val):
                continue
            return val
    return ""

# =============================================================================
#  COMPANY TYPE
# =============================================================================

def detect_company_type(text: str) -> str:
    tl = text.lower()
    if re.search(r'\bundp\b|\bunicef\b|\bwfp\b|\bunhcr\b|\bwho\b|\bilo\b|united nations|un agency|\biom\b|\bunesco\b|\bundss\b', tl):
        return "UN Agency"
    if re.search(r'\bingo\b|international ngo|international non-governmental', tl):
        return "INGO"
    if re.search(r'\bngo\b|non.governmental|nonprofit|non-profit|non governmental', tl):
        return "NGO / Non-Profit"
    if re.search(r'\bgovernment\b|ministry of|department of|\bgovernmental\b', tl):
        return "Government"
    if re.search(r'\bprivate\b|\bltd\b|\blimited\b|\bcorporation\b|\binc\b|\bplc\b|\bco\.\b', tl):
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

def get_html(url: str, timeout: int = 15) -> str:
    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout, allow_redirects=True)
        if resp.status_code == 200:
            return resp.text
    except Exception:
        pass
    return ""

def fetch_pdf(pdf_url: str):
    """
    Download a PDF and return (text, raw_bytes).
    text      — full extracted text (empty string on failure)
    raw_bytes — raw PDF bytes (None on failure), used for logo extraction
    """
    if not pdf_url or not pdf_url.startswith("http"):
        return "", None
    try:
        resp = requests.get(pdf_url, headers=HEADERS, timeout=30)
        if resp.status_code != 200:
            print(f"      ✗ HTTP {resp.status_code}")
            return "", None
        raw = resp.content
        with pdfplumber.open(io.BytesIO(raw)) as pdf:
            pages = [p.extract_text() for p in pdf.pages if p.extract_text()]
            text = "\n".join(pages)
        return text, raw
    except Exception as e:
        print(f"      ✗ PDF error: {e}")
        return "", None

# =============================================================================
#  PDF LOGO EXTRACTOR — pulls the org logo embedded in the PDF
# =============================================================================

def extract_logo_from_pdf(pdf_bytes: bytes) -> str:
    """
    Extract the most likely org logo from the first page of a PDF.

    Strategy:
      1. Open with PyMuPDF (fitz).
      2. Scan page 1 (page 2 as fallback) for embedded images.
      3. Match each image's pixel data (via get_images xref) with its
         displayed position/size (via get_image_info bbox).
         Matching is by xref when available; falls back to index order.
      4. Filter to logo candidates:
         - In the top 35% of the page
         - Display width 20–380px, height 12–220px
         - Not a full-page background image
         - Pixel dimensions at least 30x10
      5. Pick the best candidate: leftmost wins (org logo is usually left;
         MIMU header / page-number logos are on the right).
      6. Return as data URI: "data:image/png;base64,..."
         (works without any external URL, stores in CSV/XLSX directly)

    Returns empty string if no suitable image found.
    """
    if not pdf_bytes:
        return ""
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception:
        return ""

    for page_num in range(min(2, len(doc))):
        page    = doc[page_num]
        page_w  = page.rect.width
        page_h  = page.rect.height

        imgs   = page.get_images(full=True)   # [{xref, smask, w, h, ...}]
        infos  = page.get_image_info()        # [{bbox, xref?, width, height, ...}]

        if not imgs:
            continue

        # Build xref→info lookup; fall back to index-order pairing
        xref_to_info = {}
        for info in infos:
            x = info.get("xref")
            if x:
                xref_to_info[x] = info

        # Pair each img entry with its display info
        paired = []
        for idx, img in enumerate(imgs):
            xref = img[0]
            info = xref_to_info.get(xref)
            if info is None and idx < len(infos):
                info = infos[idx]   # index-order fallback
            if info is None:
                continue
            paired.append((xref, img, info))

        candidates = []
        for xref, img, info in paired:
            bbox    = info["bbox"]
            disp_w  = bbox[2] - bbox[0]
            disp_h  = bbox[3] - bbox[1]
            y_top   = bbox[1]
            x_left  = bbox[0]
            pct_top = y_top / page_h if page_h else 1

            pix_w = img[2]
            pix_h = img[3]

            # ── Filters ───────────────────────────────────────────────────────
            if pct_top > 0.35:                          continue  # below header area
            if disp_w < 20  or disp_w > 380:           continue  # too narrow or too wide
            if disp_h < 12  or disp_h > 220:           continue  # too thin or too tall
            if pix_w  < 30  or pix_h < 10:             continue  # pixel-tiny icons
            if disp_w > page_w * 0.85:                 continue  # full-width banner
            if disp_h > page_h * 0.45:                 continue  # full-height image

            candidates.append({
                "xref":    xref,
                "x_left":  x_left,
                "y_top":   y_top,
                "disp_w":  disp_w,
                "disp_h":  disp_h,
                "pct_top": pct_top,
            })

        if not candidates:
            continue

        # Prefer leftmost image in the header (org logo is almost always left)
        best = min(candidates, key=lambda c: (c["x_left"], c["pct_top"]))

        try:
            base_img = doc.extract_image(best["xref"])
            ext      = base_img.get("ext", "png")
            imgdata  = base_img["image"]
            if len(imgdata) < 100:
                continue
            b64 = base64.b64encode(imgdata).decode()
            return f"data:image/{ext};base64,{b64}"
        except Exception:
            continue

    return ""

# =============================================================================
#  WEBSITE SCRAPER — visit About page for company details
# =============================================================================

ABOUT_SLUGS = [
    "/about", "/about-us", "/about_us", "/who-we-are", "/our-story",
    "/organisation", "/organization", "/mission", "/overview",
    "/about/who-we-are", "/en/about", "/en/about-us",
]

def get_about_text(base_url: str, soup: BeautifulSoup, html: str) -> str:
    """Try to find a rich about/mission paragraph from homepage or about page."""
    # 1. Look for about section on homepage
    for tag in soup.find_all(["section", "div", "article"], limit=60):
        cid  = " ".join(tag.get("class", [])) + " " + str(tag.get("id", ""))
        if re.search(r'about|mission|vision|who.we.are|our.story|overview', cid, re.IGNORECASE):
            txt = tag.get_text(" ", strip=True)
            if len(txt) > 100:
                return txt[:1000]

    # 2. Meta description
    meta = soup.find("meta", attrs={"name": "description"}) or \
           soup.find("meta", property="og:description")
    if meta and meta.get("content") and len(meta["content"]) > 60:
        meta_desc = meta["content"][:800]
    else:
        meta_desc = ""

    # 3. Try /about page
    for slug in ABOUT_SLUGS:
        about_url = base_url.rstrip("/") + slug
        about_html = get_html(about_url, timeout=12)
        if not about_html:
            continue
        about_soup = BeautifulSoup(about_html, "html.parser")
        # Remove nav/header/footer noise
        for tag in about_soup.find_all(["nav", "header", "footer", "script", "style"]):
            tag.decompose()
        # Grab first large text block
        for tag in about_soup.find_all(["p", "div", "section"], limit=80):
            txt = tag.get_text(" ", strip=True)
            if len(txt) > 120 and re.search(
                r'mission|vision|about|who we are|established|founded|our work|we are|organisation|organization',
                txt, re.IGNORECASE
            ):
                return txt[:1000]
        # Fallback: longest paragraph
        paras = [t.get_text(" ", strip=True) for t in about_soup.find_all("p")]
        paras = [p for p in paras if len(p) > 80]
        if paras:
            return sorted(paras, key=len, reverse=True)[0][:1000]

    return meta_desc

def scrape_website(url: str) -> dict:
    """
    Visit the org website and extract:
      - description / about text (from about page)
      - logo URL
      - founded year
      - address
      - company type hints
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
    og = soup.find("meta", property="og:image") or soup.find("meta", attrs={"name": "og:image"})
    if og and og.get("content"):
        logo = og["content"]
    if not logo:
        for img in soup.find_all("img"):
            src = img.get("src", "")
            alt = img.get("alt", "")
            cls = " ".join(img.get("class", []))
            iid = img.get("id", "")
            if re.search(r'logo|brand|emblem', src + alt + cls + iid, re.IGNORECASE):
                logo = src
                break
    if not logo:
        icon = soup.find("link", rel=lambda r: r and "icon" in " ".join(r).lower())
        if icon and icon.get("href"):
            logo = icon["href"]
    if logo:
        if logo.startswith("//"):
            logo = "https:" + logo
        elif logo.startswith("/"):
            logo = base_url + logo
        elif not logo.startswith("http"):
            logo = base_url + "/" + logo
    result["logo"] = logo

    # ── About / Description — from About page ─────────────────────────────────
    result["description"] = get_about_text(base_url, soup, html)

    # ── Founded year ──────────────────────────────────────────────────────────
    text_body = soup.get_text(" ")
    m = re.search(r'(?:established|founded|since|incorporated)\s+(?:in\s+)?(\d{4})', text_body, re.IGNORECASE)
    if m:
        year = int(m.group(1))
        if 1900 <= year <= datetime.now().year:
            result["founded"] = m.group(1)

    # ── Address — strict patterns only ────────────────────────────────────────
    result["address"] = extract_address(text_body)

    # ── Company type ──────────────────────────────────────────────────────────
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

def search_block(text: str, header_patterns: list, max_lines: int = 20) -> str:
    """Extract a block of text following a section header."""
    lines = text.split("\n")
    for i, line in enumerate(lines):
        for pat in header_patterns:
            if re.search(pat, line, re.IGNORECASE):
                block = []
                for j in range(i + 1, min(i + 1 + max_lines, len(lines))):
                    l = lines[j].strip()
                    if not l:
                        continue
                    # Stop at next section header
                    if re.match(r'^[A-Z][A-Z\s/&]{4,}:?\s*$', l) and len(l) < 60:
                        break
                    if l.endswith(":") and len(l) < 50 and l == l.upper():
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
    for url in urls:
        u = url.rstrip('.,)')
        if re.search(r'apply|career|recruit|workday|bamboo|greenhouse|lever|smartrecruiters|smrtr|hr-manager|myworkday|forms\.office|tdh\.org/en-GB', u, re.IGNORECASE):
            return u
    return ""

def extract_application(text: str, existing: str) -> str:
    if existing and existing.startswith("http") and not existing.endswith("/"):
        if re.search(r'apply|career|recruit|workday|bamboo|smrtr|hr-manager|myworkday|forms|tdh\.org', existing, re.IGNORECASE):
            return existing

    decoded = text.replace("%40", "@").replace("%2E", ".").replace("%2F", "/")

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

    url = extract_apply_url(app_section) or extract_apply_url(decoded)
    if url:
        return url

    if existing:
        decoded_ex = existing.replace("%40", "@").replace("%2E", ".").replace("%2F", "/").lstrip("/")
        if "@" in decoded_ex or decoded_ex.startswith("http"):
            return decoded_ex

    return ""

def extract_website(text: str, existing: str) -> str:
    """Find the org's own website, not an ATS platform."""
    urls = re.findall(r'https?://[^\s\'"<>)]+', text)
    candidates = []
    for url in urls:
        u = url.rstrip('.,)')
        if not any(d in u for d in APP_PLATFORM_DOMAINS):
            candidates.append(u)
    for url in candidates:
        path = re.sub(r'https?://[^/]+', '', url)
        if len(path) < 25:
            return url
    if candidates:
        return candidates[0]
    # Fall back to existing if it's not an ATS platform
    ex = s(existing)
    if ex and not any(d in ex for d in APP_PLATFORM_DOMAINS):
        return ex
    return ""

# =============================================================================
#  FULL JOB DESCRIPTION EXTRACTOR
# =============================================================================

DESCRIPTION_HEADERS = [
    r'key responsibilit', r'main responsibilit', r'responsibilit',
    r'duties and responsibilit', r'key tasks', r'scope of work',
    r'main duties', r'job purpose', r'objective', r'role summary',
    r'key deliverables', r'position summary', r'job summary',
    r'about the role', r'role description', r'what you will do',
    r'your role', r'the role', r'tasks and responsibilit',
    r'description of duties', r'overview of the role',
]

def extract_description(text: str, sheet_desc: str) -> str:
    """Extract a full, rich job description (up to 1500 chars)."""
    # Try each header — pick the longest block found
    best = ""
    lines = text.split("\n")
    for i, line in enumerate(lines):
        for pat in DESCRIPTION_HEADERS:
            if re.search(pat, line, re.IGNORECASE):
                block_lines = []
                for j in range(i + 1, min(i + 35, len(lines))):
                    l = lines[j].strip()
                    if not l:
                        continue
                    # Stop at next ALL-CAPS section header
                    if re.match(r'^[A-Z][A-Z\s/&]{5,}:?\s*$', l) and len(l) < 70:
                        break
                    block_lines.append(l)
                candidate = " ".join(block_lines)
                if len(candidate) > len(best):
                    best = candidate

    if best and len(best) > 100:
        return best[:1500]

    # Fallback: gather all bullet-point lines (• ❖ - *)
    bullets = []
    for line in lines:
        stripped = line.strip()
        if stripped and re.match(r'^[•❖\-\*►▪]', stripped) and len(stripped) > 20:
            bullets.append(stripped)
    if bullets:
        return " ".join(bullets[:25])[:1500]

    # Last fallback: long sentences
    long_lines = [l.strip() for l in lines if len(l.strip()) > 60]
    candidate = " ".join(long_lines[:15])
    if candidate:
        return candidate[:1500]

    return s(sheet_desc)[:1500]

# =============================================================================
#  COMPANY DETAILS FROM PDF — extract "About the organisation" section
# =============================================================================

# Section headers that introduce org background in PDFs
ABOUT_ORG_HEADERS = [
    r'about\s+(?:the\s+)?(?:organization|organisation|us|our\s+org)',
    r'who\s+we\s+are',
    r'background(?:\s+of\s+(?:the\s+)?(?:organization|organisation))?',
    r'presentation\s+of\s+the\s+organization',
    r'introduction(?:\s+to\s+(?:the\s+)?(?:organization|organisation))?',
    r'about\s+[A-Z]{2,}',          # "About CARE", "About DRC" etc.
    r'overview\s+of\s+(?:the\s+)?(?:organization|organisation)',
    r'organisation\s+background',
    r'(?:the\s+)?(?:organization|organisation)\s+overview',
]

def extract_company_details_from_pdf(text: str, org_name: str) -> str:
    """
    Extract an 'About the organisation' paragraph from the PDF text.
    Returns up to 1000 chars of clean prose.
    """
    if not text:
        return ""

    lines = text.split("\n")

    # 1. Try known section headers
    for i, line in enumerate(lines):
        for pat in ABOUT_ORG_HEADERS:
            if re.search(pat, line, re.IGNORECASE):
                block = []
                for j in range(i + 1, min(i + 20, len(lines))):
                    l = lines[j].strip()
                    if not l:
                        if block:   # stop at first blank line after content
                            break
                        continue
                    # Stop at next section header (ALL-CAPS or ends with colon)
                    if re.match(r'^[A-Z][A-Z\s/&]{5,}:?\s*$', l) and len(l) < 80:
                        break
                    block.append(l)
                candidate = " ".join(block).strip()
                if len(candidate) > 80:
                    return candidate[:1000]

    # 2. Look for a paragraph that mentions the org name and sounds like a description
    if org_name and len(org_name) > 2:
        org_short = org_name[:20]
        for i, line in enumerate(lines):
            if org_short.lower() in line.lower() and len(line) > 80:
                # Grab this line + a few after it
                block = [line.strip()]
                for j in range(i + 1, min(i + 6, len(lines))):
                    l = lines[j].strip()
                    if not l or re.match(r'^[A-Z][A-Z\s/&]{5,}:?\s*$', l):
                        break
                    block.append(l)
                candidate = " ".join(block).strip()
                if len(candidate) > 100:
                    return candidate[:1000]

    # 3. Find the first paragraph > 100 chars that sounds like an org description
    for line in lines:
        l = line.strip()
        if len(l) > 100 and re.search(
            r'(?:is\s+(?:a|an|the)|was\s+(?:established|founded)|has\s+been|'
            r'organization|organisation|non.profit|humanitarian|ngo|ingo|'
            r'international|development|working\s+in|operating\s+in)',
            l, re.IGNORECASE
        ):
            return l[:1000]

    return ""

def detect_company_founded(text: str) -> str:
    m = re.search(
        r'(?:established|founded|since|incorporated|organisation\s+in|organization\s+in)\s+(?:in\s+)?(\d{4})',
        text, re.IGNORECASE
    )
    if m:
        year = int(m.group(1))
        if 1900 <= year <= datetime.now().year:
            return m.group(1)
    return ""

# =============================================================================
#  VERBOSE PRINTER
# =============================================================================

def print_extracted(record: dict, pdf_text: str):
    pad = "      "
    div = pad + "-" * 60
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
            print(f"{pad}{(field+':').ljust(22)} {val[:200]}{'…' if len(val)>200 else ''}")
    if pdf_text:
        snippet = " ".join(pdf_text[:500].split())
        print(f"\n{pad}--- PDF SNIPPET ---")
        print(f"{pad}{snippet[:500]}")
    print(div)

# =============================================================================
#  MAIN FIELD PARSER
# =============================================================================

def parse_pdf_fields(text: str, pdf_bytes, row: dict, website_cache: dict) -> dict:
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

    # ── Qualifications — strict tier mapping ──────────────────────────────────
    # Search specific qualification sections first
    quals_section = search_block(text, [
        r'qualif', r'academic requirement', r'minimum requirement',
        r'education requirement', r'degree required', r'education background',
    ], 8) or search_pattern(text, [
        r"(bachelor'?s?|master'?s?|phd|diploma|degree|b\.med|mbbs|b\.med\.tech|m\.b\.,b\.s)[^\n]{0,180}",
        r'(minimum\s+(?:diploma|degree|bachelor|high school)[^\n]{0,100})',
        r'((?:any\s+)?graduate[^\n]{0,80})',
        r'qualifications?[:\s]+([^\n]{10,200})',
    ])
    quals_standard = extract_qualification(quals_section) or extract_qualification(title)
    # If still empty, scan full PDF text
    if not quals_standard:
        quals_standard = extract_qualification(text[:4000])
    quals_out = quals_standard  # only standard label — never raw prose

    # ── Experience — strict band mapping ─────────────────────────────────────
    exp_section = search_pattern(text, [
        r'(\d+\+?\s*(?:to\s*\d+\s*)?years?\s+(?:of\s+)?(?:relevant\s+)?(?:work\s+)?experience[^\n]{0,100})',
        r'(minimum\s+(?:of\s+)?\d+\s+years?[^\n]{0,100})',
        r'(at least\s+\d+\s+years?[^\n]{0,100})',
        r'(at least\s+\d+\s+months?[^\n]{0,80})',
        r'experience[:\s]+([^\n]{5,120})',
    ])
    exp_out = extract_experience((exp_section or "") + " " + text[:3000])
    # Only standard band — never raw prose

    # ── Description — full ────────────────────────────────────────────────────
    desc = extract_description(text, s(row.get("Job Description", "")))

    # ── Salary — numeric/currency only ────────────────────────────────────────
    salary = extract_salary(text, s(row.get("Salary Range", "")))

    # ── Application ───────────────────────────────────────────────────────────
    app_out = extract_application(text, s(row.get("Application", "")))

    # ── Website ───────────────────────────────────────────────────────────────
    existing_web = s(row.get("Company Website", "")) or s(row.get("Company URL", ""))
    website = extract_website(text, existing_web)

    # ── Job field & type ──────────────────────────────────────────────────────
    job_field = infer_job_field(title, text[:4000])
    job_type  = detect_job_type(text[:2000], title)
    comp_type = detect_company_type(text[:3000]) or s(row.get("Company Type", ""))

    # ── Founded from PDF ──────────────────────────────────────────────────────
    founded = detect_company_founded(text)

    # ── Address from PDF — strict ─────────────────────────────────────────────
    address = extract_address(text)

    # ── Company info placeholders ─────────────────────────────────────────────
    org      = s(row.get("Company Name", ""))
    logo     = ""
    details  = ""

    # ── Extract logo from PDF first (most reliable source) ───────────────────
    if pdf_bytes:
        pdf_logo = extract_logo_from_pdf(pdf_bytes)
        if pdf_logo:
            logo = pdf_logo
            print(f"      🖼  Logo extracted from PDF ({len(pdf_logo)} chars b64)")

    # ── Fill company info from website ────────────────────────────────────────
    if website:
        cached = website_cache.get(website)
        if cached is None:
            print(f"      🌐 Scraping website: {website}")
            cached = scrape_website(website)
            website_cache[website] = cached
            time.sleep(0.8)
        else:
            print(f"      🌐 Using cached: {website}")

        details   = cached.get("description", "")
        # Only use website logo if PDF gave nothing
        if not logo:
            logo = cached.get("logo", "")
        if not address:
            address  = cached.get("address", "")
        if not founded:
            founded  = cached.get("founded", "")
        if not comp_type:
            comp_type = cached.get("company_type", "")

    # ── Fallback: extract company details from PDF if website gave nothing ────
    if not details and text:
        details = extract_company_details_from_pdf(text, org)

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
        "Company Details":    str(details or "")[:1000],
        "Job URL":            str(s(row.get("Job URL", ""))),
        "Estimated Deadline": str(deadline or ""),
        "Salary Range":       str(salary or ""),
    }

# =============================================================================
#  MAIN
# =============================================================================

def main():
    print("=" * 64)
    print("  MIMU Jobs PDF Extractor v4")
    print(f"  Started : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 64)

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
    website_cache = {}
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
            pdf_text, pdf_bytes = fetch_pdf(pdf_url)
            if pdf_text:
                print(f"      ✓ {len(pdf_text):,} chars extracted")
                pdf_ok += 1
            else:
                print(f"      ⚠ No text from PDF")
                pdf_skip += 1
        else:
            print(f"      ⚠ No PDF URL — sheet data only")
            pdf_text  = ""
            pdf_bytes = None
            pdf_skip += 1

        enriched = parse_pdf_fields(pdf_text, pdf_bytes, row.to_dict(), website_cache)
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

    print(f"\n{'=' * 64}")
    print(f"  ✅ COMPLETE")
    print(f"     Total jobs    : {total}")
    print(f"     PDF success   : {pdf_ok}")
    print(f"     Sheet-only    : {pdf_skip}")
    print(f"     Websites hit  : {len(website_cache)}")
    print(f"     Output        : {OUTPUT_CSV}, {OUTPUT_XLSX}")
    print(f"     Finished      : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 64)


if __name__ == "__main__":
    main()
