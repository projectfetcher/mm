"""
MIMU Jobs PDF Extractor — v5
==============================
New in v5 (on top of v4):
  - Mistral paraphrasing: job titles, descriptions, company details, taglines.
    ENABLE_PARAPHRASE=True triggers up to 4 attempts per title and 3 per paragraph,
    with similarity gating, word-count checks, and verbose per-attempt logging.
  - WordPress REST API posting: creates/updates job-listings and companies via
    WP Job Manager endpoints. Uploads logos as WP Media attachments (data-URI or URL).
  - Duplicate tracker (processed_ids.csv): records job IDs so re-runs are idempotent.
  - All config via environment variables (WP_BASE_URL, WP_USERNAME, WP_APP_PASSWORD,
    MISTRAL_API_KEY) — no hardcoded secrets.

REQUIREMENTS:
    pip install requests pdfplumber pymupdf pandas openpyxl beautifulsoup4 \
    language-tool-python==2.7.1

USAGE:
    export WP_BASE_URL="https://your-site.com/wp-json/wp/v2"
    export WP_USERNAME="admin"
    export WP_APP_PASSWORD="xxxx xxxx xxxx xxxx"
    export MISTRAL_API_KEY="sk-..."
    python mimu_jobs_extractor_v5.py
"""

import requests
import pdfplumber
import fitz          # PyMuPDF
import pandas as pd
import re
import io
import os
import base64
import time
import sys
import math
import hashlib
import logging
import warnings
from datetime import datetime
from bs4 import BeautifulSoup

import language_tool_python

warnings.filterwarnings("ignore")


# =============================================================================
#  LOGGING
# =============================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# =============================================================================
#  CONFIG — all secrets via env vars
# =============================================================================

SHEET_CSV_URL = (
    "https://docs.google.com/spreadsheets/d/e/"
    "2PACX-1vS9_Zza6If2T_LT-iVvQjBTIFTeqt_OVBws70v_s3NJavT-ZosZ28qtE7xds7iS5rLmU2UbhzxWnOsY"
    "/pub?gid=964760760&single=true&output=csv"
)

OUTPUT_CSV         = "mimu_jobs.csv"
OUTPUT_XLSX        = "mimu_jobs.xlsx"
PROCESSED_IDS_FILE = "processed_ids.csv"

# ── WordPress ─────────────────────────────────────────────────────────────────
WP_URL      = os.environ.get("WP_BASE_URL", "")          # e.g. https://site.com/wp-json/wp/v2
WP_USER     = os.environ.get("WP_USERNAME", "")
WP_PASSWORD = os.environ.get("WP_APP_PASSWORD", "")
WP_BASE        = WP_URL.rstrip("/")
WP_JOBS_URL    = f"{WP_BASE}/job-listings"
WP_COMPANY_URL = f"{WP_BASE}/companies"
WP_MEDIA_URL   = f"{WP_BASE}/media"

# ── Mistral ───────────────────────────────────────────────────────────────────
MISTRAL_API_KEY = os.environ.get("MISTRAL_API_KEY", "")
MISTRAL_MODEL   = "mistral-small-latest"
MISTRAL_URL     = "https://api.mistral.ai/v1/chat/completions"
ENABLE_PARAPHRASE = True

# ── Startup warnings ──────────────────────────────────────────────────────────
for _var, _val, _feature in [
    ("MISTRAL_API_KEY", MISTRAL_API_KEY, "paraphrasing"),
    ("WP_USERNAME",     WP_USER,         "WordPress posting"),
    ("WP_APP_PASSWORD", WP_PASSWORD,     "WordPress posting"),
    ("WP_BASE_URL",     WP_URL,          "WordPress posting"),
]:
    if not _val:
        log.warning(f"Env var {_var} not set — {_feature} will be disabled/skipped.")

# =============================================================================
#  GRAMMAR MODEL
#  SentenceTransformer removed — similarity is computed via Mistral API instead,
#  avoiding HuggingFace rate-limit (429) errors on GitHub Actions runners.
# =============================================================================

print("⏳ Loading LanguageTool grammar checker…")
try:
    _grammar_tool = language_tool_python.LanguageTool(
        "en-US", remote_server="https://api.languagetool.org"
    )
    _GRAMMAR_ENABLED = True
    print("✅ LanguageTool ready.")
except Exception as _lt_err:
    log.warning(f"LanguageTool unavailable ({_lt_err}) — grammar correction disabled.")
    _grammar_tool    = None
    _GRAMMAR_ENABLED = False

# =============================================================================
#  OUTPUT COLUMNS
# =============================================================================

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

APP_PLATFORM_DOMAINS = [
    'smrtr.io', 'workday', 'myworkday', 'bamboohr', 'greenhouse.io',
    'lever.co', 'forms.office', 'google.com/forms', 'hr-manager',
    'smartrecruiters', 'themimu.info', 'candidate.', 'apply.',
    'recruiting.', 'jobs.tdh.org', 'theirc.wd1', 'worldvision.wd1',
]

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
        score  = sum(3 for k in high_keys  if k in combined)
        score += sum(1 for k in supporting if k in combined)
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
    "some experience","minimal experience","at least 6",
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
#  JOB TYPE
# =============================================================================

def detect_job_type(text: str, title: str) -> str:
    combined = (text + " " + title).lower()
    if re.search(r'\bintern\b|\binternship\b', combined):
        return "Internship"
    if re.search(r'\bpart[-\s]time\b', combined):
        return "Part-time"
    if re.search(r'\bvolunteer\b', combined):
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
#  SALARY
# =============================================================================

CURRENCY_PATTERNS = [
    r'USD\s*[\d,]+(?:\s*[-–]\s*[\d,]+)?(?:\s*/\s*\w+)?',
    r'GBP\s*[\d,]+(?:\s*[-–]\s*[\d,]+)?(?:\s*/\s*\w+)?',
    r'EUR\s*[\d,]+(?:\s*[-–]\s*[\d,]+)?(?:\s*/\s*\w+)?',
    r'\$\s*[\d,]+(?:\s*[-–]\s*\$?\s*[\d,]+)?(?:\s*/\s*\w+)?',
    r'£\s*[\d,]+(?:\s*[-–]\s*£?\s*[\d,]+)?(?:\s*/\s*\w+)?',
    r'[\d,]+(?:\s*[-–]\s*[\d,]+)?\s*(?:MMK|Ks\.?|Kyats?)\b',
    r'MMK\s*[\d,]+(?:\s*[-–]\s*[\d,]+)?',
    r'[\d,]+(?:\s*[-–]\s*[\d,]+)?\s*/\s*(?:month|year|day|hour)',
    r'(?:fee|total fee)\s+(?:of\s+)?(?:USD|GBP|EUR|\$|£)?\s*[\d,]+',
]

def extract_salary(text: str, sheet_salary: str) -> str:
    for src in [text[:3000], sheet_salary]:
        if not src:
            continue
        for pat in CURRENCY_PATTERNS:
            m = re.search(pat, src, re.IGNORECASE)
            if m:
                val = m.group(0).strip().rstrip('.,')
                if re.search(r'\d', val):
                    return val
    return ""

# =============================================================================
#  ADDRESS
# =============================================================================

_PROSE_WORDS = re.compile(
    r'\b(?:please|ensure|must|will|should|have|been|with|through|across|'
    r'areas|regions|branches|townships|countries|programs|projects|staff|'
    r'services|sector|parent|duty|based|grade|report|department|during|'
    r'providing|working|seeking|implement|support|assist|manage)\b',
    re.IGNORECASE
)
_ADDR_MARKERS = re.compile(
    r'\b(?:Street|Road|Avenue|Lane|Quarter|Ward|Township|Yangon|Mandalay|'
    r'Nay\s*Pyi\s*Taw|Mawlamyine|Myanmar)\b',
    re.IGNORECASE
)
ADDRESS_PATTERNS = [
    r'No\.?\s*\(?[A-Z0-9][A-Z0-9\-]*\)?\s*[,\s]+[^\n]{5,120}(?:Street|Road|Avenue|Lane)',
    r'#\s*\d+[,\s]+[^\n]{5,100}(?:Street|Road|Avenue|Lane|Township|Yangon|Mandalay)',
    r'(?:^|\n)\s*(?:address|office address|head office|office)[:\s]+([A-Z0-9#No\.][^\n]{15,150})',
    r'Ward\s+(?:No\.?\s*)?\(?\w[\w\s]*\)?\s*,\s*[^\n]{5,80}(?:Township|Yangon|Mandalay|Myanmar)',
]

def extract_address(text: str) -> str:
    if not text:
        return ""
    for pat in ADDRESS_PATTERNS:
        for m in re.finditer(pat, text, re.IGNORECASE | re.MULTILINE):
            val = (m.group(1) if m.lastindex else m.group(0)).strip().strip('.,\n')
            val = re.sub(r'\s*\n\s*', ', ', val).strip()
            if not re.search(r'\d', val):
                continue
            if not _ADDR_MARKERS.search(val):
                continue
            if _PROSE_WORDS.search(val):
                continue
            if len(val) < 10 or len(val) > 180:
                continue
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
    """Download PDF, return (text, raw_bytes)."""
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
#  PDF LOGO EXTRACTOR
# =============================================================================

def extract_logo_from_pdf(pdf_bytes: bytes) -> str:
    if not pdf_bytes:
        return ""
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception:
        return ""

    for page_num in range(min(2, len(doc))):
        page   = doc[page_num]
        page_w = page.rect.width
        page_h = page.rect.height
        imgs   = page.get_images(full=True)
        infos  = page.get_image_info()

        if not imgs:
            continue

        xref_to_info = {}
        for info in infos:
            x = info.get("xref")
            if x:
                xref_to_info[x] = info

        paired = []
        for idx, img in enumerate(imgs):
            xref = img[0]
            info = xref_to_info.get(xref)
            if info is None and idx < len(infos):
                info = infos[idx]
            if info is None:
                continue
            paired.append((xref, img, info))

        candidates = []
        for xref, img, info in paired:
            bbox   = info["bbox"]
            disp_w = bbox[2] - bbox[0]
            disp_h = bbox[3] - bbox[1]
            y_top  = bbox[1]
            x_left = bbox[0]
            pct_top = y_top / page_h if page_h else 1
            pix_w  = img[2]
            pix_h  = img[3]

            if pct_top > 0.35:               continue
            if disp_w < 20 or disp_w > 380:  continue
            if disp_h < 12 or disp_h > 220:  continue
            if pix_w  < 30 or pix_h < 10:    continue
            if disp_w > page_w * 0.85:        continue
            if disp_h > page_h * 0.45:        continue

            candidates.append({"xref": xref, "x_left": x_left,
                                "y_top": y_top, "disp_w": disp_w,
                                "disp_h": disp_h, "pct_top": pct_top})

        if not candidates:
            continue

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
#  WEBSITE SCRAPER
# =============================================================================

ABOUT_SLUGS = [
    "/about", "/about-us", "/about_us", "/who-we-are", "/our-story",
    "/organisation", "/organization", "/mission", "/overview",
    "/about/who-we-are", "/en/about", "/en/about-us",
]

def get_about_text(base_url: str, soup: BeautifulSoup, html: str) -> str:
    for tag in soup.find_all(["section", "div", "article"], limit=60):
        cid = " ".join(tag.get("class", [])) + " " + str(tag.get("id", ""))
        if re.search(r'about|mission|vision|who.we.are|our.story|overview', cid, re.IGNORECASE):
            txt = tag.get_text(" ", strip=True)
            if len(txt) > 100:
                return txt[:1000]

    meta = (soup.find("meta", attrs={"name": "description"}) or
            soup.find("meta", property="og:description"))
    meta_desc = ""
    if meta and meta.get("content") and len(meta["content"]) > 60:
        meta_desc = meta["content"][:800]

    for slug in ABOUT_SLUGS:
        about_url  = base_url.rstrip("/") + slug
        about_html = get_html(about_url, timeout=12)
        if not about_html:
            continue
        about_soup = BeautifulSoup(about_html, "html.parser")
        for tag in about_soup.find_all(["nav", "header", "footer", "script", "style"]):
            tag.decompose()
        for tag in about_soup.find_all(["p", "div", "section"], limit=80):
            txt = tag.get_text(" ", strip=True)
            if len(txt) > 120 and re.search(
                r'mission|vision|about|who we are|established|founded|our work|we are|organisation|organization',
                txt, re.IGNORECASE
            ):
                return txt[:1000]
        paras = [t.get_text(" ", strip=True) for t in about_soup.find_all("p")]
        paras = [p for p in paras if len(p) > 80]
        if paras:
            return sorted(paras, key=len, reverse=True)[0][:1000]

    return meta_desc

def scrape_website(url: str) -> dict:
    result = {"description": "", "logo": "", "founded": "", "address": "", "company_type": ""}
    if not url or not url.startswith("http"):
        return result

    html = get_html(url)
    if not html:
        return result

    soup     = BeautifulSoup(html, "html.parser")
    base     = re.match(r'(https?://[^/]+)', url)
    base_url = base.group(1) if base else ""

    logo = ""
    og   = soup.find("meta", property="og:image") or soup.find("meta", attrs={"name": "og:image"})
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

    result["description"] = get_about_text(base_url, soup, html)

    text_body = soup.get_text(" ")
    m = re.search(r'(?:established|founded|since|incorporated)\s+(?:in\s+)?(\d{4})', text_body, re.IGNORECASE)
    if m:
        year = int(m.group(1))
        if 1900 <= year <= datetime.now().year:
            result["founded"] = m.group(1)

    result["address"]      = extract_address(text_body)
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
    lines = text.split("\n")
    for i, line in enumerate(lines):
        for pat in header_patterns:
            if re.search(pat, line, re.IGNORECASE):
                block = []
                for j in range(i + 1, min(i + 1 + max_lines, len(lines))):
                    l = lines[j].strip()
                    if not l:
                        continue
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
    emails  = re.findall(r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}', decoded)
    skip    = ['example.com', 'domain.com', 'email.com', 'yourmail', 'sentry', 'noreply']
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

    decoded     = text.replace("%40", "@").replace("%2E", ".").replace("%2F", "/")
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
    urls       = re.findall(r'https?://[^\s\'"<>)]+', text)
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
    ex = s(existing)
    if ex and not any(d in ex for d in APP_PLATFORM_DOMAINS):
        return ex
    return ""

# =============================================================================
#  DESCRIPTION EXTRACTOR
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
    best  = ""
    lines = text.split("\n")
    for i, line in enumerate(lines):
        for pat in DESCRIPTION_HEADERS:
            if re.search(pat, line, re.IGNORECASE):
                block_lines = []
                for j in range(i + 1, min(i + 35, len(lines))):
                    l = lines[j].strip()
                    if not l:
                        continue
                    if re.match(r'^[A-Z][A-Z\s/&]{5,}:?\s*$', l) and len(l) < 70:
                        break
                    block_lines.append(l)
                candidate = " ".join(block_lines)
                if len(candidate) > len(best):
                    best = candidate

    if best and len(best) > 100:
        return best[:1500]

    bullets = []
    for line in lines:
        stripped = line.strip()
        if stripped and re.match(r'^[•❖\-\*►▪]', stripped) and len(stripped) > 20:
            bullets.append(stripped)
    if bullets:
        return " ".join(bullets[:25])[:1500]

    long_lines = [l.strip() for l in lines if len(l.strip()) > 60]
    candidate  = " ".join(long_lines[:15])
    if candidate:
        return candidate[:1500]

    return s(sheet_desc)[:1500]

# =============================================================================
#  COMPANY DETAILS FROM PDF
# =============================================================================

ABOUT_ORG_HEADERS = [
    r'about\s+(?:the\s+)?(?:organization|organisation|us|our\s+org)',
    r'who\s+we\s+are',
    r'background(?:\s+of\s+(?:the\s+)?(?:organization|organisation))?',
    r'presentation\s+of\s+the\s+organization',
    r'introduction(?:\s+to\s+(?:the\s+)?(?:organization|organisation))?',
    r'about\s+[A-Z]{2,}',
    r'overview\s+of\s+(?:the\s+)?(?:organization|organisation)',
    r'organisation\s+background',
    r'(?:the\s+)?(?:organization|organisation)\s+overview',
]

def extract_company_details_from_pdf(text: str, org_name: str) -> str:
    if not text:
        return ""
    lines = text.split("\n")

    for i, line in enumerate(lines):
        for pat in ABOUT_ORG_HEADERS:
            if re.search(pat, line, re.IGNORECASE):
                block = []
                for j in range(i + 1, min(i + 20, len(lines))):
                    l = lines[j].strip()
                    if not l:
                        if block:
                            break
                        continue
                    if re.match(r'^[A-Z][A-Z\s/&]{5,}:?\s*$', l) and len(l) < 80:
                        break
                    block.append(l)
                candidate = " ".join(block).strip()
                if len(candidate) > 80:
                    return candidate[:1000]

    if org_name and len(org_name) > 2:
        org_short = org_name[:20]
        for i, line in enumerate(lines):
            if org_short.lower() in line.lower() and len(line) > 80:
                block = [line.strip()]
                for j in range(i + 1, min(i + 6, len(lines))):
                    l = lines[j].strip()
                    if not l or re.match(r'^[A-Z][A-Z\s/&]{5,}:?\s*$', l):
                        break
                    block.append(l)
                candidate = " ".join(block).strip()
                if len(candidate) > 100:
                    return candidate[:1000]

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
#  MISTRAL API
# =============================================================================

def mistral_generate(prompt: str, max_tokens: int = 400, temperature: float = 0.7) -> str:
    if not MISTRAL_API_KEY:
        log.warning("MISTRAL_API_KEY not set — skipping paraphrase")
        return ""
    try:
        response = requests.post(
            MISTRAL_URL,
            headers={
                "Authorization": f"Bearer {MISTRAL_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": MISTRAL_MODEL,
                "messages": [{"role": "user", "content": prompt}],
                "max_tokens": max_tokens,
                "temperature": temperature,
            },
            timeout=30,
        )
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"].strip()
    except Exception as e:
        log.error(f"Mistral API error: {e}")
        return ""

# =============================================================================
#  TEXT HELPERS FOR PARAPHRASE
# =============================================================================

# Common mojibake sequences produced by mis-decoded UTF-8
_MOJIBAKE = [
    ("Â", ""), ("â€™", "'"), ("â€œ", '"'), ("â€\x9d", '"'), ("â€", '"'),
    ("â€¢", "•"), ("â„¢", "™"), ("\u00a0", " "), ("\u200b", ""), ("\ufeff", ""),
]

def _fix_mojibake(text: str) -> str:
    for pattern, replacement in _MOJIBAKE:
        text = text.replace(pattern, replacement)
    # Strip non-printable control characters (keep newline \n = 0x0A)
    text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", text)
    return text

def sanitize_text(text: str, is_url: bool = False, is_email: bool = False) -> str:
    """
    Clean and normalise text before sending to Mistral or WordPress.
    - Fixes mojibake / encoding artefacts
    - Strips HTML tags
    - Collapses excess whitespace
    - Removes markdown artefacts (##, **)
    """
    if not isinstance(text, str):
        text = str(text) if text is not None else ""
    text = text.strip()
    if text in ("nan", "None", "NaN", "", "N/A", "n/a", "NA", "na"):
        return ""
    text = _fix_mojibake(text)
    # For URLs/emails: just collapse whitespace, no further stripping
    if is_url or is_email:
        return re.sub(r"[ \t\r\n\f\v]+", " ", text).strip()
    # Strip HTML tags
    text = re.sub(r"<[^>]+>", " ", text)
    # Remove markdown artefacts
    text = re.sub(r"#+\s*", "", text)
    text = re.sub(r"\*\*", "", text)
    # Keep only printable Latin + extended Latin + common punctuation
    text = re.sub(
        r"[^\x20-\x7E\n\u00C0-\u017F\u2013\u2014\u2018-\u201D\u2022]", "", text
    )
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()

def _grammar_correct(text: str) -> str:
    """Apply LanguageTool grammar correction if available."""
    if not _GRAMMAR_ENABLED or not _grammar_tool:
        return text
    try:
        return language_tool_python.utils.correct(text, _grammar_tool.check(text))
    except Exception:
        return text

def clean_output(raw: str) -> str:
    """
    Strip model artefacts from Mistral output, fix encoding, apply grammar correction.
    Handles: markdown fences, [INST] tags, leading labels, excess whitespace.
    """
    if not raw:
        return ""
    raw = _fix_mojibake(raw)
    # Remove markdown code fences
    raw = re.sub(r"```[a-z]*", "", raw).replace("```", "")
    # Remove model instruction tags
    raw = re.sub(r"\[/?INST\]|</?s>", "", raw)
    # Strip leading "Rewritten:", "Output:", etc.
    raw = re.sub(
        r"^(?:rewritten?|rephrased?|output|paraphrase[d]?)[:\s]+",
        "", raw, flags=re.IGNORECASE,
    )
    # Strip markdown bold / headers / dividers
    raw = re.sub(r"\*\*|###|---", "", raw)
    raw = re.sub(r"[ \t]+", " ", raw)
    raw = re.sub(r"\n{3,}", "\n\n", raw)
    return _grammar_correct(raw.strip())

def similarity_score(a: str, b: str) -> float:
    """
    Semantic similarity scored by Mistral (no HuggingFace download needed).
    Asks the model to rate how semantically similar two texts are on 0-10,
    then normalises to [0, 1].  Falls back to Jaccard on any API failure.
    """
    if not a or not b:
        return 0.0
    if not MISTRAL_API_KEY:
        # Fallback: Jaccard word overlap
        sa = set(a.lower().split())
        sb = set(b.lower().split())
        return len(sa & sb) / len(sa | sb) if (sa or sb) else 0.0
    try:
        prompt = (
            "Rate the semantic similarity of these two texts on a scale of 0 to 10.\n"
            "0 = completely unrelated, 10 = identical meaning.\n"
            "Reply with ONLY a single integer (0-10), nothing else.\n\n"
            f"Text A: {a[:300]}\n\nText B: {b[:300]}"
        )
        raw = mistral_generate(prompt, max_tokens=5, temperature=0.0)
        score = float(re.search(r"\d+", raw).group()) if re.search(r"\d+", raw) else 5.0
        return min(max(score / 10.0, 0.0), 1.0)
    except Exception:
        # Fallback: Jaccard word overlap
        sa = set(a.lower().split())
        sb = set(b.lower().split())
        return len(sa & sb) / len(sa | sb) if (sa or sb) else 0.0

def _print_wrapped(text: str, prefix: str = "   ", width: int = 100):
    words = text.split()
    line  = []
    for w in words:
        line.append(w)
        if len(" ".join(line)) >= width:
            print(f"{prefix}{' '.join(line)}")
            line = []
    if line:
        print(f"{prefix}{' '.join(line)}")

# =============================================================================
#  PARAPHRASE FUNCTIONS
# =============================================================================

def paraphrase_title(title: str) -> str:
    if not ENABLE_PARAPHRASE:
        return title
    clean = sanitize_text(title)
    if not clean:
        return title

    print(f"\n ┌─ TITLE PARAPHRASE {'─'*45}")
    print(f" │ Original : \"{clean}\"")
    print(f" │ {'─'*60}")

    best_result = None
    best_sim    = 0.0

    for attempt in range(4):
        temp = round(0.68 + attempt * 0.06, 2)
        print(f" │ Attempt {attempt+1} (temp={temp}):")

        prompt = (
            f"Rewrite this job title professionally using different words. "
            f"Output ONLY the rewritten title, nothing else. "
            f"Keep it between 4 and 12 words.\n\nJob title: {clean}"
        )

        raw    = mistral_generate(prompt, max_tokens=50, temperature=temp)
        result = clean_output(raw).split("\n")[0].strip().strip('"').strip("'")

        wc     = len(result.split()) if result else 0
        sim    = similarity_score(clean, result) if result else 0.0
        is_dup = result.lower().strip() == clean.lower().strip()

        print(f" │    Output  : \"{result}\"")
        print(f" │    Words   : {wc} | Similarity: {sim:.3f} | Duplicate: {'Yes ⚠️' if is_dup else 'No'}")

        valid = bool(result) and 4 <= wc <= 14 and sim >= 0.55 and not is_dup

        if not valid:
            reasons = []
            if not result:  reasons.append("empty output")
            if wc < 4:      reasons.append(f"too short ({wc} words, min=4)")
            if wc > 14:     reasons.append(f"too long ({wc} words, max=14)")
            if sim < 0.55:  reasons.append(f"sim={sim:.3f} < 0.55")
            if is_dup:      reasons.append("identical to original")
            print(f" │    → ❌ REJECTED — {', '.join(reasons)}")
        else:
            if sim > best_sim:
                best_sim    = sim
                best_result = result
                print(f" │    → ✅ ACCEPTED — new best candidate (sim={sim:.3f})")
            else:
                print(f" │    → ✅ VALID but not better than current best (best sim={best_sim:.3f})")

        print(f" │ {'─'*60}")
        time.sleep(1)

    if best_result:
        print(f" │ 🏆 FINAL SELECTED : \"{best_result}\"")
        print(f" │    Similarity     : {best_sim:.3f}")
        print(f" └{'─'*65}")
        return best_result
    else:
        print(f" │ ⚠️  No valid paraphrase found → Keeping original: \"{clean}\"")
        print(f" └{'─'*65}")
        return clean


def paraphrase_description(text: str) -> str:
    if not ENABLE_PARAPHRASE:
        return text
    clean = sanitize_text(text)
    if not clean:
        return text

    paragraphs    = [p.strip() for p in clean.split("\n") if p.strip()]
    rewritten     = []
    success_count = 0

    print(f"\n ┌─ DESCRIPTION PARAPHRASE ({len(paragraphs)} paragraphs) {'─'*25}")

    for i, para in enumerate(paragraphs):
        orig_wc = len(para.split())
        print(f"\n │ ┌─ Paragraph {i+1}/{len(paragraphs)} {'─'*50}")
        print(f" │ │ ORIGINAL ({orig_wc} words):")
        _print_wrapped(para, prefix=" │ │    ")
        print(f" │ │ {'─'*60}")

        prompt = (
            f"Rewrite this job description paragraph professionally in English. "
            f"Keep ALL facts, requirements, and responsibilities. "
            f"Use different sentence structure and vocabulary. "
            f"Output ONLY the rewritten paragraph — no labels, no explanation.\n\n"
            f"Original:\n{para}"
        )

        best_result   = None
        best_sim      = 0.0
        accepted_text = None

        for attempt in range(3):
            temp   = round(0.65 + attempt * 0.08, 2)
            print(f" │ │ Attempt {attempt+1}/3 (temp={temp}):")

            raw    = mistral_generate(prompt, max_tokens=500, temperature=temp)
            result = clean_output(raw).strip()

            rw  = len(result.split()) if result else 0
            sim = similarity_score(para, result) if result and rw >= 5 else 0.0

            if result:
                print(f" │ │    Paraphrased ({rw} words, sim={sim:.3f}):")
                _print_wrapped(result, prefix=" │ │       ")
            else:
                print(f" │ │    Paraphrased : (no output from model)")

            valid = bool(result) and rw >= 8 and sim >= 0.48

            if not valid:
                reasons = []
                if not result: reasons.append("empty output")
                if rw < 8:     reasons.append(f"too short ({rw} words, min=8)")
                if sim < 0.48: reasons.append(f"sim={sim:.3f} < 0.48")
                print(f" │ │    → ❌ REJECTED — {', '.join(reasons)}")
                if result and sim > best_sim:
                    best_sim    = sim
                    best_result = result
                    print(f" │ │       (stored as best fallback, sim={sim:.3f})")
            else:
                print(f" │ │    → ✅ ACCEPTED on attempt {attempt+1}")
                rewritten.append(result)
                success_count += 1
                accepted_text  = result
                break

            print(f" │ │ {'─'*60}")
            time.sleep(1)

        if accepted_text is None:
            print(f" │ │ {'─'*60}")
            if best_result and best_sim >= 0.40:
                print(f" │ │ 🔁 FALLBACK — Using best attempt (sim={best_sim:.3f}):")
                _print_wrapped(best_result, prefix=" │ │    ")
                rewritten.append(best_result)
                success_count += 1
            else:
                print(f" │ │ ⚠️  KEPT ORIGINAL — no acceptable paraphrase (best sim={best_sim:.3f})")
                rewritten.append(para)

        print(f" │ └{'─'*62}")

    print(f"\n │ SUMMARY: {success_count}/{len(paragraphs)} paragraphs successfully paraphrased")
    print(f" └{'─'*80}\n")
    return "\n\n".join(rewritten)


def paraphrase_company(text: str) -> str:
    if not ENABLE_PARAPHRASE:
        return text
    clean = sanitize_text(text)
    if not clean:
        return text

    print(f"\n ┌─ COMPANY PARAPHRASE {'─'*43}")
    orig_wc = len(clean.split())
    print(f" │ Original ({orig_wc} words):")
    _print_wrapped(clean, prefix=" │    ")
    print(f" │ {'─'*60}")

    prompt = (
        f"Rewrite this company description professionally in English. "
        f"Preserve all facts. Use different wording. "
        f"Output ONLY the rewritten description.\n\nOriginal:\n{clean}"
    )

    raw    = mistral_generate(prompt, max_tokens=600, temperature=0.68)
    result = clean_output(raw)
    rw     = len(result.split()) if result else 0
    sim    = similarity_score(clean, result) if result and rw >= 10 else 0.0

    if result and rw >= 10:
        print(f" │ Paraphrased ({rw} words, sim={sim:.3f}):")
        _print_wrapped(result, prefix=" │    ")
        print(f" │ → ✅ ACCEPTED")
        print(f" └{'─'*65}")
        time.sleep(1)
        return result
    else:
        reasons = []
        if not result: reasons.append("empty output")
        if rw < 10:    reasons.append(f"too short ({rw} words, min=10)")
        print(f" │ → ❌ REJECTED — {', '.join(reasons)} — keeping original")
        print(f" └{'─'*65}")
        time.sleep(1)
        return clean


def paraphrase_tagline(text: str) -> str:
    if not ENABLE_PARAPHRASE:
        return text
    clean = sanitize_text(text[:300])
    if not clean:
        return text

    print(f"\n ┌─ TAGLINE PARAPHRASE {'─'*43}")
    print(f" │ Original : \"{clean}\"")
    print(f" │ {'─'*60}")

    prompt = (
        f"Rewrite this company tagline as a crisp, professional English phrase. "
        f"Output ONLY the rewritten tagline (5–12 words). No explanation.\n\n"
        f"Original: {clean}"
    )

    raw    = mistral_generate(prompt, max_tokens=35, temperature=0.75)
    result = clean_output(raw).split("\n")[0].strip().strip('"').strip("'")
    wc     = len(result.split()) if result else 0

    print(f" │ Paraphrased : \"{result}\"")
    print(f" │ Words: {wc} | Similarity: {similarity_score(clean, result) if result else 0.0:.3f}")

    if result and 3 <= wc <= 15:
        print(f" │ → ✅ ACCEPTED")
        print(f" └{'─'*65}")
        time.sleep(1)
        return result
    else:
        reasons = []
        if not result: reasons.append("empty output")
        if wc < 3:     reasons.append(f"too short ({wc} words, min=3)")
        if wc > 15:    reasons.append(f"too long ({wc} words, max=15)")
        print(f" │ → ❌ REJECTED — {', '.join(reasons)} — keeping original")
        print(f" └{'─'*65}")
        time.sleep(1)
        return clean

# =============================================================================
#  DUPLICATE TRACKER
# =============================================================================

def _init_tracker():
    if not os.path.exists(PROCESSED_IDS_FILE):
        pd.DataFrame(columns=[
            "Job ID", "Job URL", "Job Title", "Company Name",
            "Status", "Timestamp", "WP ID",
        ]).to_csv(PROCESSED_IDS_FILE, index=False)

def load_processed_ids() -> tuple:
    _init_tracker()
    df = pd.read_csv(PROCESSED_IDS_FILE)
    return (
        set(df["Job ID"].fillna("").astype(str)),
        set(df.get("Job URL", pd.Series()).fillna("").astype(str)),
    )

def _upsert_row(job_id: str, updates: dict):
    _init_tracker()
    df   = pd.read_csv(PROCESSED_IDS_FILE)
    mask = df["Job ID"].astype(str) == str(job_id)
    if mask.any():
        for col, val in updates.items():
            if col in df.columns:
                df.loc[mask, col] = val
        df.loc[mask, "Timestamp"] = datetime.now().isoformat()
    else:
        row = {"Job ID": job_id, "Timestamp": datetime.now().isoformat()}
        row.update(updates)
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    df.to_csv(PROCESSED_IDS_FILE, index=False)

def make_job_id(job_url: str, title: str = "", company: str = "", idx: int = 0) -> str:
    if job_url:
        return hashlib.md5(job_url.encode()).hexdigest()[:16]
    seed = f"{title}{company}{idx}"
    return hashlib.md5(seed.encode()).hexdigest()[:16]

def mark_scraped(job_id, job_url, title, company):
    _upsert_row(job_id, {"Job URL": job_url, "Job Title": title,
                          "Company Name": company, "Status": "scraped"})

def mark_paraphrased(job_id):
    _upsert_row(job_id, {"Status": "paraphrased"})

def mark_posted(job_id, wp_id, wp_url=""):
    _upsert_row(job_id, {"Status": "posted", "WP ID": wp_id})

def mark_failed(job_id, reason):
    _upsert_row(job_id, {"Status": f"failed|{reason}"})

# =============================================================================
#  WORDPRESS REST API
# =============================================================================

def _wp_auth() -> tuple:
    return (WP_USER, WP_PASSWORD)

def wp_upload_logo(logo: str, company_name: str) -> int:
    """
    Upload a logo to WP Media library.
    Accepts either:
      - data URI  → "data:image/png;base64,..."
      - HTTP URL  → download then upload
    Returns the WP attachment ID, or 0 on failure.
    """
    if not WP_USER or not WP_PASSWORD or not WP_MEDIA_URL:
        return 0
    if not logo:
        return 0

    try:
        # Resolve bytes + mime from data URI or URL
        if logo.startswith("data:"):
            m = re.match(r'data:(image/[^;]+);base64,(.+)', logo, re.DOTALL)
            if not m:
                return 0
            mime     = m.group(1)
            img_data = base64.b64decode(m.group(2))
        elif logo.startswith("http"):
            resp = requests.get(logo, timeout=15)
            if resp.status_code != 200:
                return 0
            img_data = resp.content
            ct       = resp.headers.get("Content-Type", "image/png")
            mime     = ct.split(";")[0].strip()
        else:
            return 0

        ext      = mime.split("/")[-1].replace("jpeg", "jpg")
        filename = re.sub(r'[^a-z0-9]+', '-', company_name.lower())[:40] + f"-logo.{ext}"

        resp = requests.post(
            WP_MEDIA_URL,
            auth=_wp_auth(),
            headers={
                "Content-Disposition": f'attachment; filename="{filename}"',
                "Content-Type": mime,
            },
            data=img_data,
            timeout=30,
        )
        if resp.status_code in (200, 201):
            attach_id = resp.json().get("id", 0)
            log.info(f"      🖼  Logo uploaded to WP media: ID={attach_id}")
            return attach_id
        else:
            log.warning(f"      ⚠ Logo upload failed: HTTP {resp.status_code} — {resp.text[:200]}")
            return 0
    except Exception as e:
        log.error(f"      ✗ Logo upload error: {e}")
        return 0


def wp_get_or_create_company(record: dict) -> int:
    """
    Find existing WP company by name or create a new one.
    Returns the WP company post ID (or 0 on failure).

    This assumes WP Job Manager with a 'companies' custom post type.
    Adjust endpoint / meta keys to match your theme's CPT slug.
    """
    if not WP_USER or not WP_PASSWORD or not WP_COMPANY_URL:
        return 0

    company_name = record.get("Company Name", "").strip()
    if not company_name:
        return 0

    # ── Search for existing company ───────────────────────────────────────────
    try:
        search_resp = requests.get(
            WP_COMPANY_URL,
            auth=_wp_auth(),
            params={"search": company_name, "per_page": 5},
            timeout=15,
        )
        if search_resp.status_code == 200:
            results = search_resp.json()
            for item in results:
                if item.get("title", {}).get("rendered", "").lower() == company_name.lower():
                    log.info(f"      🏢 Company already exists: ID={item['id']}")
                    return item["id"]
    except Exception as e:
        log.warning(f"      ⚠ Company search error: {e}")

    # ── Upload logo ───────────────────────────────────────────────────────────
    logo_id = wp_upload_logo(record.get("Company Logo", ""), company_name)

    # ── Create company ────────────────────────────────────────────────────────
    payload = {
        "title":   company_name,
        "status":  "publish",
        "content": record.get("Company Details", ""),
        "meta": {
            "_company_website":  record.get("Company Website", ""),
            "_company_tagline":  record.get("Company Industry", ""),
            "_company_twitter":  "",
            "_company_linkedin": "",
        },
    }
    if logo_id:
        payload["featured_media"] = logo_id

    try:
        resp = requests.post(
            WP_COMPANY_URL,
            auth=_wp_auth(),
            json=payload,
            timeout=20,
        )
        if resp.status_code in (200, 201):
            company_id = resp.json().get("id", 0)
            log.info(f"      🏢 Company created: ID={company_id}")
            return company_id
        else:
            log.warning(f"      ⚠ Company create failed: HTTP {resp.status_code} — {resp.text[:300]}")
            return 0
    except Exception as e:
        log.error(f"      ✗ Company create error: {e}")
        return 0


def wp_post_job(record: dict, company_id: int) -> tuple:
    """
    Create a WP Job Manager job listing via REST API.
    Returns (wp_post_id, wp_post_url) or (0, "").

    Meta keys follow WP Job Manager conventions; adjust if your theme differs.
    """
    if not WP_USER or not WP_PASSWORD or not WP_JOBS_URL:
        return 0, ""

    # Build the post content (description)
    content = record.get("Job Description", "")

    # Add How to Apply section
    application = record.get("Application", "")
    if application:
        how = f"\n\n<strong>How to Apply:</strong> {application}"
        content += how

    payload = {
        "title":   record.get("Job Title", "Untitled"),
        "status":  "publish",
        "content": content,
        "meta": {
            # WP Job Manager core meta
            "_job_location":    record.get("Job Location", ""),
            "_application":     application,
            "_company_name":    record.get("Company Name", ""),
            "_company_website": record.get("Company Website", ""),
            "_job_expires":     record.get("Deadline", ""),
            # Extended / custom meta (adjust keys to your theme)
            "_job_salary":      record.get("Salary Range", ""),
            "_job_type":        record.get("Job Type", "Full-time"),
            "_job_field":       record.get("Job Field", ""),
            "_job_experience":  record.get("Job Experience", ""),
            "_job_qualification": record.get("Job Qualifications", ""),
            "_date_posted":     record.get("Date Posted", ""),
            "_company_id":      str(company_id) if company_id else "",
            "_company_type":    record.get("Company Type", ""),
            "_company_founded": record.get("Company Founded", ""),
            "_job_source_url":  record.get("Job URL", ""),
        },
    }

    # Attach job type taxonomy if supported
    job_type_slug = {
        "Full-time":             "full-time",
        "Part-time":             "part-time",
        "Internship":            "internship",
        "Volunteer":             "volunteer",
        "Consultancy / Contract":"contract",
    }.get(record.get("Job Type", ""), "full-time")
    payload["job_listing_type"] = [job_type_slug]

    try:
        resp = requests.post(
            WP_JOBS_URL,
            auth=_wp_auth(),
            json=payload,
            timeout=30,
        )
        if resp.status_code in (200, 201):
            data   = resp.json()
            wp_id  = data.get("id", 0)
            wp_url = data.get("link", "")
            log.info(f"      ✅ Job posted: ID={wp_id}  {wp_url}")
            return wp_id, wp_url
        else:
            log.warning(f"      ⚠ Job post failed: HTTP {resp.status_code} — {resp.text[:400]}")
            return 0, ""
    except Exception as e:
        log.error(f"      ✗ Job post error: {e}")
        return 0, ""

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
            label   = (field + ":").ljust(22)
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
    if not quals_standard:
        quals_standard = extract_qualification(text[:4000])
    quals_out = quals_standard

    exp_section = search_pattern(text, [
        r'(\d+\+?\s*(?:to\s*\d+\s*)?years?\s+(?:of\s+)?(?:relevant\s+)?(?:work\s+)?experience[^\n]{0,100})',
        r'(minimum\s+(?:of\s+)?\d+\s+years?[^\n]{0,100})',
        r'(at least\s+\d+\s+years?[^\n]{0,100})',
        r'(at least\s+\d+\s+months?[^\n]{0,80})',
        r'experience[:\s]+([^\n]{5,120})',
    ])
    exp_out = extract_experience((exp_section or "") + " " + text[:3000])

    desc   = extract_description(text, s(row.get("Job Description", "")))
    salary = extract_salary(text, s(row.get("Salary Range", "")))
    app_out = extract_application(text, s(row.get("Application", "")))

    existing_web = s(row.get("Company Website", "")) or s(row.get("Company URL", ""))
    website      = extract_website(text, existing_web)

    job_field  = infer_job_field(title, text[:4000])
    job_type   = detect_job_type(text[:2000], title)
    comp_type  = detect_company_type(text[:3000]) or s(row.get("Company Type", ""))
    founded    = detect_company_founded(text)
    address    = extract_address(text)

    org     = s(row.get("Company Name", ""))
    logo    = ""
    details = ""

    if pdf_bytes:
        pdf_logo = extract_logo_from_pdf(pdf_bytes)
        if pdf_logo:
            logo = pdf_logo
            print(f"      🖼  Logo extracted from PDF ({len(pdf_logo)} chars b64)")

    if website:
        cached = website_cache.get(website)
        if cached is None:
            print(f"      🌐 Scraping website: {website}")
            cached = scrape_website(website)
            website_cache[website] = cached
            time.sleep(0.8)
        else:
            print(f"      🌐 Using cached: {website}")

        details = cached.get("description", "")
        if not logo:
            logo = cached.get("logo", "")
        if not address:
            address  = cached.get("address", "")
        if not founded:
            founded  = cached.get("founded", "")
        if not comp_type:
            comp_type = cached.get("company_type", "")

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
    print("  MIMU Jobs PDF Extractor v5")
    print(f"  Started : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Paraphrase : {'ENABLED' if ENABLE_PARAPHRASE and MISTRAL_API_KEY else 'DISABLED'}")
    print(f"  WP Posting : {'ENABLED' if WP_USER and WP_PASSWORD and WP_URL else 'DISABLED'}")
    print("=" * 64)

    # 1. Load sheet
    print(f"\n[1/4] Fetching Google Sheet CSV …")
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

    # Load duplicate tracker
    processed_ids, processed_urls = load_processed_ids()
    print(f"      Already processed: {len(processed_ids)} jobs")

    # 2. Extract
    print(f"\n[2/4] Downloading PDFs and extracting fields …")
    records       = []
    website_cache = {}
    total         = len(df)
    pdf_ok        = 0
    pdf_skip      = 0

    for idx, row in df.iterrows():
        pdf_url = s(row.get("PDF URL", ""))
        title   = s(row.get("Job Title", ""))
        num     = idx + 1

        # ── Duplicate check ───────────────────────────────────────────────────
        job_id = make_job_id(pdf_url, title, s(row.get("Company Name", "")), idx)
        if job_id in processed_ids or (pdf_url and pdf_url in processed_urls):
            print(f"\n  [{num}/{total}] ⏭  SKIP (already processed): {title}")
            continue

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
            pdf_skip  += 1

        enriched = parse_pdf_fields(pdf_text, pdf_bytes, row.to_dict(), website_cache)
        print_extracted(enriched, pdf_text)
        mark_scraped(job_id, pdf_url, title, enriched.get("Company Name", ""))
        records.append((job_id, enriched))
        time.sleep(0.2)

    # 3. Paraphrase
    print(f"\n[3/4] Paraphrasing …")
    paraphrased_records = []
    for job_id, rec in records:
        print(f"\n  ── {rec.get('Job Title', '?')} ──")
        rec["Job Title"]       = paraphrase_title(rec.get("Job Title", ""))
        rec["Job Description"] = paraphrase_description(rec.get("Job Description", ""))
        rec["Company Details"] = paraphrase_company(rec.get("Company Details", ""))
        mark_paraphrased(job_id)
        paraphrased_records.append((job_id, rec))

    # 4. Save CSV/XLSX + post to WordPress
    print(f"\n[4/4] Saving outputs and posting to WordPress …")
    out_records = [rec for _, rec in paraphrased_records]
    out_df      = pd.DataFrame(out_records, columns=OUTPUT_COLUMNS)

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

    # WordPress posting
    wp_success = 0
    wp_failed  = 0
    if WP_USER and WP_PASSWORD and WP_URL:
        for job_id, rec in paraphrased_records:
            print(f"\n  📤 Posting: {rec.get('Job Title', '?')}")
            company_id = wp_get_or_create_company(rec)
            wp_id, wp_url = wp_post_job(rec, company_id)
            if wp_id:
                mark_posted(job_id, wp_id, wp_url)
                wp_success += 1
            else:
                mark_failed(job_id, "wp_post_failed")
                wp_failed += 1
            time.sleep(0.5)
    else:
        print("      ⏭  WordPress posting skipped (env vars not set).")

    # Summary
    total_processed = len(records)
    print(f"\n{'=' * 64}")
    print(f"  ✅ COMPLETE")
    print(f"     Total rows in sheet : {total}")
    print(f"     New jobs processed  : {total_processed}")
    print(f"     PDF success         : {pdf_ok}")
    print(f"     Sheet-only          : {pdf_skip}")
    print(f"     Websites hit        : {len(website_cache)}")
    print(f"     WP posted OK        : {wp_success}")
    print(f"     WP failed           : {wp_failed}")
    print(f"     Output              : {OUTPUT_CSV}, {OUTPUT_XLSX}")
    print(f"     Finished            : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 64)


if __name__ == "__main__":
    main()
