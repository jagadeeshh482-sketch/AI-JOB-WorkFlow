"""
Job Hunt Automation
Runs every 12 hours via GitHub Actions (or cron).
Scrapes company career pages → filters by visa mode (H1B / OPT-CPT / STEM-OPT / all)
→ ATS scores vs resume → saves 3-sheet Excel (Tracker + Dashboard + Legend) → emails top matches.

SETUP — set these GitHub Secrets (or environment variables):
  YOUR_NAME         Your full name
  YOUR_EMAIL        Your Gmail address
  GMAIL_APP_PASS    Gmail App Password (not your real password)
  RESUME_URL        Public URL to your resume PDF (optional)
  YOUR_TITLE        Your professional title for email subject  e.g. "Senior Software Engineer"
  YOUR_PHONE        Your phone number for email signature
  YOUR_LINKEDIN     Your LinkedIn profile URL
  RESUME_SUMMARY    Plain-text resume summary for ATS keyword scoring (optional)
  VISA_MODE         "h1b" | "opt_cpt" | "stem_opt" | "all"  (default: "all")
  JOB_SEARCH_QUERY  Keywords to search  (default: "software engineer")
"""

import os, re, json, time, logging, smtplib, hashlib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# CONFIGURATION — set via GitHub Secrets or env vars (no hardcoded values)
# ─────────────────────────────────────────────

YOUR_NAME      = os.getenv("YOUR_NAME",  "")           # e.g. "Jane Smith"
YOUR_EMAIL     = os.getenv("YOUR_EMAIL", "")           # e.g. "jane@gmail.com"
YOUR_TITLE     = os.getenv("YOUR_TITLE", "Software Engineer")  # shown in email subject line
YOUR_PHONE     = os.getenv("YOUR_PHONE", "")           # e.g. "+1 555-123-4567"
YOUR_LINKEDIN  = os.getenv("YOUR_LINKEDIN", "")        # e.g. "https://linkedin.com/in/yourname"
GMAIL_APP_PASS = os.getenv("GMAIL_APP_PASS", "")       # Gmail App Password (not your real password)
EXCEL_PATH     = Path("output/job_tracker.xlsx")
SEEN_JOBS_FILE = Path("output/seen_jobs.json")
MIN_ATS_SCORE  = int(os.getenv("MIN_ATS_SCORE", "70"))     # only shortlist jobs scoring >= this
MAX_APPLY_PER_RUN = int(os.getenv("MAX_APPLY_PER_RUN", "5"))  # cap applications per run

# Job search keyword — customise per role/field
JOB_SEARCH_QUERY = os.getenv("JOB_SEARCH_QUERY", "software engineer")

# ── Visa mode ─────────────────────────────────────────────────────────────────
# Set VISA_MODE env var to control which jobs are surfaced:
#   "h1b"       → only H1B sponsoring companies
#   "opt_cpt"   → companies accepting OPT / CPT / STEM-OPT
#   "stem_opt"  → STEM-OPT specific
#   "all"       → all three modes combined (widest net, default)
VISA_MODE = os.getenv("VISA_MODE", "all").lower()

# ── Resume summary for ATS keyword scoring ───────────────────────────────────
# Option A (recommended): set the RESUME_SUMMARY env var / GitHub Secret with
#   a plain-text paste of your resume.
# Option B: edit the fallback string below.
_RESUME_SUMMARY_DEFAULT = """
Software engineer with experience in backend and full-stack development.
Skills: Python, Java, JavaScript, TypeScript, Node.js, React, REST APIs,
SQL, PostgreSQL, MongoDB, Docker, Kubernetes, AWS, CI/CD, Git, Agile.
Authorized to work in the US (requires visa sponsorship).
"""
RESUME_SUMMARY = os.getenv("RESUME_SUMMARY", _RESUME_SUMMARY_DEFAULT)

# Target companies: (display name, careers page URL, optional keyword to find job links)
# ── Direct company career pages ──────────────────────────────────────────────
# Target companies — search URLs use JOB_SEARCH_QUERY so they adapt to your role
# ── Direct company career pages ──────────────────────────────────────────────
_Q = JOB_SEARCH_QUERY.replace(" ", "+")
_Q_ENC = JOB_SEARCH_QUERY.replace(" ", "%20")
COMPANIES = [
    ("Google",         f"https://careers.google.com/jobs/results/?q={_Q}&employment_type=FULL_TIME",                           None),
    ("Meta",           f"https://www.metacareers.com/jobs?q={_Q}&teams[0]=Engineering",                                        None),
    ("Amazon",         f"https://www.amazon.jobs/en/search?base_query={_Q}&loc_query=",                                        None),
    ("Microsoft",      f"https://jobs.microsoft.com/us/en/search#q={_Q_ENC}&p=1",                                              None),
    ("Salesforce",     f"https://careers.salesforce.com/en/jobs/?search={_Q}&department=Software+Engineering",                 None),
    ("JPMorgan Chase", f"https://jpmc.fa.oraclecloud.com/hcmUI/CandidateExperience/en/sites/CX_1001/requisitions?keyword={_Q}",None),
    ("Apple",          f"https://jobs.apple.com/en-us/search?search={_Q}&sort=relevance",                                      None),
    ("Netflix",        f"https://jobs.netflix.com/search?q={_Q_ENC}",                                                         None),
    ("Stripe",         f"https://stripe.com/jobs/search?query={_Q}",                                                          None),
    ("Uber",           f"https://www.uber.com/us/en/careers/list/?query={_Q}",                                                 None),
    ("Twilio",          "https://boards.greenhouse.io/twilio",                                                                  "engineer"),
    ("Atlassian",      f"https://www.atlassian.com/company/careers/all-jobs?team=Engineering&search={_Q}",                     None),
    ("ServiceNow",     f"https://careers.servicenow.com/careers/jobs?keywords={_Q}",                                           None),
    ("Workday",        f"https://www.workday.com/en-us/company/careers/open-positions.html?q={_Q}",                            None),
    ("Adobe",          f"https://careers.adobe.com/us/en/search-results?keywords={_Q}",                                        None),
    ("Intuit",         f"https://jobs.intuit.com/search-jobs?keyword={_Q}",                                                    None),
    ("PayPal",         f"https://careers.pypl.com/jobs/?keyword={_Q}",                                                         None),
    ("Cisco",          f"https://jobs.cisco.com/jobs/SearchJobs/{_Q_ENC}",                                                     None),
    ("Oracle",         f"https://careers.oracle.com/jobs/#en/sites/jobsearch/jobs?keyword={_Q}",                               None),
    ("IBM",            f"https://www.ibm.com/employment/#jobs?job-search={_Q}",                                                None),
]

# ── Job portals — parsed with dedicated scrapers below ───────────────────────
# Each entry: (portal_name, search_url, parser_key)
JOB_PORTALS = [
    ("LinkedIn",      f"https://www.linkedin.com/jobs/search/?keywords={_Q}+visa+sponsorship&location=United+States&f_WT=2&f_JT=F",  "linkedin"),
    ("Indeed",        f"https://www.indeed.com/jobs?q={_Q}+%22visa+sponsorship%22&l=United+States&jt=fulltime",                       "indeed"),
    ("Dice",          f"https://www.dice.com/jobs?q={_Q}&location=United+States&filters.workplaceTypes=Remote&filters.employmentType=FULLTIME", "dice"),
    ("Monster",       f"https://www.monster.com/jobs/search?q={_Q}&where=United+States&jobtype=fulltime",                             "monster"),
    ("ZipRecruiter",  f"https://www.ziprecruiter.com/Jobs/{_Q.replace('+', '-')}?radius=25&days=3",                                   "generic"),
    ("SimplyHired",   f"https://www.simplyhired.com/search?q={_Q}+visa+sponsorship&l=United+States",                                  "simplyhired"),
    ("Glassdoor",     f"https://www.glassdoor.com/Job/united-states-{_Q.replace('+', '-')}-jobs-SRCH_IL.0,13_IN1.htm",               "glassdoor"),
    ("CareerBuilder", f"https://www.careerbuilder.com/jobs?keywords={_Q}+visa+sponsorship&location=United+States&emp=jtft",           "generic"),
    ("Wellfound",      "https://wellfound.com/jobs?role=software-engineer&remote=true",                                               "generic"),
    ("Greenhouse",     "https://boards.greenhouse.io/embed/job_board?for=",                                                           "greenhouse"),
    ("Lever",          "https://jobs.lever.co/",                                                                                      "lever"),
    ("Built In",      f"https://builtin.com/jobs/dev-engineer?title={_Q}",                                                           "generic"),
    ("Remotive",       "https://remotive.com/remote-jobs/software-dev",                                                              "remotive"),
    ("WeWorkRemotely",f"https://weworkremotely.com/remote-jobs/search?term={_Q}",                                                     "generic"),
    ("MyVisaJobs",    f"https://www.myvisajobs.com/{_Q.replace('+', '-')}_JT.htm",                                                   "generic"),
    ("H1BGrader",     f"https://h1bgrader.com/job-openings?q={_Q}",                                                                  "generic"),
    ("OPTNation",     f"https://www.optnation.com/opt-jobs-for-international-students?job={_Q}",                                      "generic"),
    ("F1Hire",        f"https://www.f1hire.com/jobs?q={_Q}",                                                                         "generic"),
    ("RippleMatch",   f"https://ripplematch.com/jobs?q={_Q}&workAuth=OPT",                                                           "generic"),
]

# ── Visa / sponsorship keyword lists ─────────────────────────────────────────

H1B_POSITIVE = [
    "visa sponsorship", "sponsor visa", "h1b", "h-1b", "will sponsor",
    "work authorization provided", "we sponsor", "sponsorship available",
    "immigration sponsorship", "visa support", "relocation and visa",
    "sponsor work visa", "work visa sponsorship", "visa assistance",
    "tn visa", "o-1 visa", "employment authorization", "sponsor employment",
    "sponsor immigration", "visa transfer", "h1b transfer"
]
H1B_NEGATIVE = [
    "no sponsorship", "not sponsor", "must be authorized", "must be legally authorized",
    "no visa", "citizens only", "us citizens and permanent residents only",
    "must have authorization to work", "no h1b", "no work visa",
    "permanent resident only", "green card only", "no opt", "no cpt",
    "no student visa", "not eligible for sponsorship"
]

# OPT / CPT / STEM-OPT — companies that explicitly welcome F-1 students
OPT_CPT_POSITIVE = [
    "opt", "cpt", "stem opt", "stem extension", "f-1", "f1 visa",
    "student visa", "curricular practical training", "optional practical training",
    "welcome opt", "accept opt", "f-1 students", "international students welcome",
    "open to opt", "open to cpt", "internship opt", "work authorization",
    "eligible to work", "authorized to work", "employment eligibility",
]
OPT_CPT_NEGATIVE = [
    "no opt", "no cpt", "no f-1", "no student visa", "no international",
    "citizens only", "permanent resident only", "green card only",
    "no sponsorship", "no visa", "not eligible for sponsorship",
    "us citizens and permanent residents only",
]
USA_PATTERNS = [
    r"\b(?:united states|united states of america|usa|u\.s\.a|u\.s\.|us|america)\b",
    r"\b(?:new york|california|texas|florida|illinois|washington|seattle|san francisco|ny|ca|tx)\b",
]

INDIA_PATTERNS = [
    r"\b(?:india|indian)\b",
    r"\b(?:chennai|bangalore|bangaluru|mumbai|delhi|hyderabad|pune|kolkata|gurgaon|noida)\b",
]

# allow remote only when explicitly mentioning US/India
REMOTE_OK = [r"\bremote\b.*\b(?:usa|us|united states|india|indian)\b", r"\b(?:usa|india)\b.*\bremote\b"]

ROLE_POSITIVE = [
    # QA / Test (original)
    "sdet", "software engineer in test", "qa automation", "test automation",
    "quality engineer", "automation engineer", "qa engineer", "quality assurance",
    "test engineer",
    # Software Engineering (new)
    "software engineer", "software developer", "backend engineer", "backend developer",
    "frontend engineer", "frontend developer", "full stack engineer", "full stack developer",
    "full-stack engineer", "full-stack developer", "web developer", "web engineer",
    "application developer", "application engineer", "platform engineer",
    "infrastructure engineer", "devops engineer", "site reliability engineer", "sre",
    "cloud engineer", "systems engineer", "data engineer", "ml engineer",
    "machine learning engineer", "api engineer", "java developer", "python developer",
    "nodejs developer", "node.js developer",
]

ROLE_NEGATIVE = [
    "data scientist", "sales", "recruiter", "marketing", "human resources",
    "hr manager", "accountant", "business analyst", "product manager",
    "project manager", "content writer", "graphic designer", "customer support",
    "customer service", "operations manager", "finance manager",
]

def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip().lower())


def location_allowed(job: dict) -> bool:
    text = " ".join([
        job.get("title", ""), job.get("company", ""), job.get("snippet", ""),
        job.get("source", ""), job.get("url", ""), job.get("location", "")
    ])
    text = normalize_text(text)

    # explicit country match
    if any(re.search(p, text, re.I) for p in USA_PATTERNS):
        return True
    if any(re.search(p, text, re.I) for p in INDIA_PATTERNS):
        return True

    # remote is allowed only if it explicitly mentions US/India
    if any(re.search(p, text, re.I) for p in REMOTE_OK):
        return True

    return False


def role_allowed(job: dict) -> bool:
    """Return True if job title/snippet matches target roles (QA/SDET/test automation).
    Also excludes clearly unrelated roles via ROLE_NEGATIVE."""
    text = " ".join([job.get("title", ""), job.get("snippet", ""), job.get("company", "")])
    text = normalize_text(text)
    if any(neg in text for neg in ROLE_NEGATIVE):
        return False
    return any(p in text for p in ROLE_POSITIVE)
# ─────────────────────────────────────────────
# UTILITIES
# ─────────────────────────────────────────────

def job_id(title: str, company: str) -> str:
    return hashlib.md5(f"{title.lower().strip()}{company.lower().strip()}".encode()).hexdigest()[:12]

def load_seen() -> set:
    if SEEN_JOBS_FILE.exists():
        return set(json.loads(SEEN_JOBS_FILE.read_text()))
    return set()

def save_seen(seen: set):
    SEEN_JOBS_FILE.write_text(json.dumps(list(seen)))

def h1b_status(text: str) -> str:
    """Returns 'yes', 'no', or 'unknown' based on H1B sponsorship signals."""
    t = text.lower()
    if any(p in t for p in H1B_NEGATIVE):
        return "no"
    if any(p in t for p in H1B_POSITIVE):
        return "yes"
    return "unknown"


def opt_cpt_status(text: str) -> str:
    """Returns 'yes', 'no', or 'unknown' based on OPT/CPT/STEM-OPT signals."""
    t = text.lower()
    if any(p in t for p in OPT_CPT_NEGATIVE):
        return "no"
    if any(p in t for p in OPT_CPT_POSITIVE):
        return "yes"
    return "unknown"


def annotate_job(job: dict) -> dict:
    """Add derived fields: H1B sponsorship and OPT/CPT eligibility."""
    text = " ".join([
        job.get("title", ""), job.get("snippet", ""), job.get("company", ""),
        job.get("source", ""), job.get("url", "")
    ])
    job["h1b_likely"]     = h1b_status(text)
    job["opt_cpt_likely"] = opt_cpt_status(text)
    return job

# ─────────────────────────────────────────────
# SCRAPING
# ─────────────────────────────────────────────

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9",
}

def fetch_page(url: str, timeout=15) -> str | None:
    try:
        r = requests.get(url, headers=HEADERS, timeout=timeout)
        r.raise_for_status()
        return r.text
    except Exception as e:
        log.warning(f"Fetch failed {url}: {e}")
        return None

def extract_jobs_from_html(html: str, company: str) -> list[dict]:
    """
    Generic extractor — looks for common job listing patterns.
    Returns list of {title, url, snippet} dicts.
    """
    soup = BeautifulSoup(html, "html.parser")
    jobs = []

    # Remove nav/footer noise
    for tag in soup(["nav", "footer", "header", "script", "style"]):
        tag.decompose()

    # Strategy 1: look for <a> tags containing job-like text near role keywords
    role_keywords = r"(engineer|developer|sdet|qa|quality|automation|test|analyst|devops|backend|frontend|fullstack|full.stack|platform|sre|cloud|java|python)"
    seen_titles = set()

    for a in soup.find_all("a", href=True):
        title = a.get_text(strip=True)
        if not title or len(title) < 6 or len(title) > 120:
            continue
        if not re.search(role_keywords, title, re.I):
            continue
        if title in seen_titles:
            continue
        seen_titles.add(title)
        href = a["href"]
        if href.startswith("/"):
            href = f"https://{requests.utils.urlparse(a.base_url if hasattr(a,'base_url') else '').netloc}{href}" if False else href
        jobs.append({"title": title, "url": href, "company": company, "snippet": ""})

    # Strategy 2: JSON-LD structured data (many modern career sites use this)
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(script.string or "")
            if isinstance(data, list):
                items = data
            elif isinstance(data, dict):
                items = data.get("itemListElement", [data])
            else:
                continue
            for item in items:
                if item.get("@type") in ("JobPosting", "ListItem"):
                    title = item.get("title") or item.get("name", "")
                    url   = item.get("url", "")
                    desc  = item.get("description", "")[:500]
                    if title and re.search(role_keywords, title, re.I):
                        jobs.append({"title": title, "url": url, "company": company, "snippet": desc})
        except Exception:
            pass

    return jobs[:30]  # cap per company


# ─────────────────────────────────────────────
# PORTAL-SPECIFIC SCRAPERS
# ─────────────────────────────────────────────

def scrape_linkedin(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    jobs = []
    for card in soup.select("div.base-card, li.jobs-search__results-list > div, div.job-search-card"):
        title_el = card.select_one("h3.base-search-card__title, h3, .job-title")
        co_el    = card.select_one("h4.base-search-card__subtitle, h4, .job-listing-name")
        link_el  = card.select_one("a.base-card__full-link, a[href*='/jobs/view/']")
        if not title_el: continue
        jobs.append({
            "title":   title_el.get_text(strip=True),
            "company": co_el.get_text(strip=True) if co_el else "LinkedIn listing",
            "url":     link_el["href"] if link_el else "",
            "snippet": "",
            "source":  "LinkedIn",
        })
    return jobs[:40]

def scrape_indeed(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    jobs = []
    # Indeed uses data-jk attributes and mosaic components
    for card in soup.select("div.job_seen_beacon, div.jobsearch-SerpJobCard, li[data-jk]"):
        title_el = card.select_one("h2.jobTitle span, a.jobtitle, [data-testid='job-title']")
        co_el    = card.select_one("span.companyName, .company, [data-testid='company-name']")
        link_el  = card.select_one("a[id^='job_'], a[href*='/viewjob'], a[href*='/rc/clk']")
        snippet_el = card.select_one("div.job-snippet, .summary")
        if not title_el: continue
        href = ""
        if link_el:
            href = link_el.get("href","")
            if href.startswith("/"):
                href = "https://www.indeed.com" + href
        jobs.append({
            "title":   title_el.get_text(strip=True),
            "company": co_el.get_text(strip=True) if co_el else "Indeed listing",
            "url":     href,
            "snippet": snippet_el.get_text(strip=True)[:400] if snippet_el else "",
            "source":  "Indeed",
        })
    # Also try JSON embedded data
    for script in soup.find_all("script", type="application/json"):
        try:
            data = json.loads(script.string or "")
            hits = data.get("props",{}).get("pageProps",{}).get("jobResults",{}).get("results",[])
            for h in hits[:20]:
                t = h.get("displayTitle") or h.get("title","")
                if t:
                    jobs.append({"title":t,"company":h.get("company",""),"url":h.get("viewJobLink",""),"snippet":h.get("snippet","")[:400],"source":"Indeed"})
        except Exception:
            pass
    return jobs[:40]

def scrape_dice(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    jobs = []
    # Dice embeds job data as JSON in a __NEXT_DATA__ script
    script = soup.find("script", id="__NEXT_DATA__")
    if script:
        try:
            data = json.loads(script.string)
            results = (data.get("props",{}).get("pageProps",{})
                          .get("initialState",{}).get("search",{})
                          .get("results",[]))
            for r in results:
                jobs.append({
                    "title":   r.get("title",""),
                    "company": r.get("employerDisplayName", r.get("companyPageUrl","Dice listing")),
                    "url":     "https://www.dice.com/jobs/" + r.get("id","") if r.get("id") else r.get("applyUrls",[""])[0],
                    "snippet": r.get("jobDescription","")[:400],
                    "source":  "Dice",
                })
        except Exception as e:
            log.debug(f"Dice JSON parse: {e}")
    # Fallback: HTML cards
    if not jobs:
        for card in soup.select("dhi-search-result, div[data-cy='search-result']"):
            title_el = card.select_one("a.card-title-link, h5")
            co_el    = card.select_one("span.employer-name, .company-name")
            if title_el:
                jobs.append({"title": title_el.get_text(strip=True),
                             "company": co_el.get_text(strip=True) if co_el else "Dice listing",
                             "url": title_el.get("href",""), "snippet":"","source":"Dice"})
    return jobs[:40]

def scrape_monster(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    jobs = []
    for card in soup.select("section.card-content, div.summary, article[data-jobid]"):
        title_el = card.select_one("h2.title a, a.job-title, h3.title")
        co_el    = card.select_one("div.company, span.name, div.company-name")
        snippet_el = card.select_one("div.job-description, p.summary-text")
        if not title_el: continue
        href = title_el.get("href","")
        if href.startswith("/"):
            href = "https://www.monster.com" + href
        jobs.append({
            "title":   title_el.get_text(strip=True),
            "company": co_el.get_text(strip=True) if co_el else "Monster listing",
            "url":     href,
            "snippet": snippet_el.get_text(strip=True)[:400] if snippet_el else "",
            "source":  "Monster",
        })
    return jobs[:40]

def scrape_simplyhired(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    jobs = []
    for card in soup.select("article.SerpJob, div[data-testid='JobCard']"):
        title_el = card.select_one("h2 a, a.jobposting-title, h3 a")
        co_el    = card.select_one("span[data-testid='company'], .company, b")
        snippet_el = card.select_one("p.jobposting-snippet, div.SerpJob-jobExcerpt")
        if not title_el: continue
        href = title_el.get("href","")
        if href.startswith("/"):
            href = "https://www.simplyhired.com" + href
        jobs.append({
            "title":   title_el.get_text(strip=True),
            "company": co_el.get_text(strip=True) if co_el else "SimplyHired listing",
            "url":     href,
            "snippet": snippet_el.get_text(strip=True)[:400] if snippet_el else "",
            "source":  "SimplyHired",
        })
    return jobs[:40]

def scrape_glassdoor(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    jobs = []
    for card in soup.select("li.react-job-listing, div.jobCard, article"):
        title_el = card.select_one("a[data-test='job-title'], span.job-title, h2")
        co_el    = card.select_one("div.employer-name, span.employer-short-name, .companyName")
        if not title_el: continue
        href = title_el.get("href","") if title_el.name == "a" else ""
        if href.startswith("/"):
            href = "https://www.glassdoor.com" + href
        jobs.append({
            "title":   title_el.get_text(strip=True),
            "company": co_el.get_text(strip=True) if co_el else "Glassdoor listing",
            "url":     href,
            "snippet": "",
            "source":  "Glassdoor",
        })
    return jobs[:40]

def scrape_remotive(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    jobs = []
    for card in soup.select("li.job, div.job-card, section[itemtype*='JobPosting']"):
        title_el = card.select_one("h2, h3, span[itemprop='title'], a.job-title")
        co_el    = card.select_one("span[itemprop='name'], .company-name, strong")
        link_el  = card.select_one("a[href*='/remote-jobs/']")
        if not title_el: continue
        href = link_el["href"] if link_el else ""
        if href and not href.startswith("http"):
            href = "https://remotive.com" + href
        jobs.append({
            "title":   title_el.get_text(strip=True),
            "company": co_el.get_text(strip=True) if co_el else "Remotive listing",
            "url":     href,
            "snippet": "",
            "source":  "Remotive",
        })
    return jobs[:40]

PORTAL_PARSERS = {
    "linkedin":    scrape_linkedin,
    "indeed":      scrape_indeed,
    "dice":        scrape_dice,
    "monster":     scrape_monster,
    "simplyhired": scrape_simplyhired,
    "glassdoor":   scrape_glassdoor,
    "remotive":    scrape_remotive,
    "generic":     lambda html: [],   # falls back to generic extractor
    "greenhouse":  lambda html: [],
    "lever":       lambda html: [],
}

def scrape_portals() -> list[dict]:
    all_jobs = []
    for portal_name, url, parser_key in JOB_PORTALS:
        log.info(f"Scraping portal: {portal_name}...")
        html = fetch_page(url)
        if not html:
            log.warning(f"  Could not fetch {portal_name}")
            continue
        parser = PORTAL_PARSERS.get(parser_key, PORTAL_PARSERS["generic"])
        jobs = parser(html)
        # Fallback to generic extractor if portal parser found nothing
        if not jobs:
            jobs = extract_jobs_from_html(html, portal_name)
        # Tag source portal
        for j in jobs:
            j.setdefault("source", portal_name)
            j.setdefault("company", portal_name + " listing")
        log.info(f"  → {len(jobs)} listings from {portal_name}")
        all_jobs.extend(jobs)
        time.sleep(3)
    return all_jobs


def scrape_all_companies() -> list[dict]:
    all_jobs = []
    # 1. Direct company pages
    log.info("── Scraping direct company career pages ──")
    for company, url, keyword in COMPANIES:
        log.info(f"Scraping {company}...")
        html = fetch_page(url)
        if not html:
            continue
        jobs = extract_jobs_from_html(html, company)
        for j in jobs:
            j.setdefault("source", "Company Site")
        log.info(f"  → found {len(jobs)} potential listings")
        all_jobs.extend(jobs)
        time.sleep(2)
    # 2. Job portals
    log.info("── Scraping job portals ──")
    all_jobs.extend(scrape_portals())
    log.info(f"Total raw listings: {len(all_jobs)}")
    return all_jobs


# ─────────────────────────────────────────────
# ATS SCORING — 100% FREE, local keyword match
# ─────────────────────────────────────────────

SKILL_WEIGHTS = [
    # QA / Test skills
    ("playwright",        12), ("selenium",          12), ("sdet",              10),
    ("qa automation",     10), ("test automation",    9),  ("restassured",        5),
    ("cucumber",           5), ("bdd",                5),  ("testng",             4),
    ("junit",              4), ("pytest",             4),  ("api testing",        6),
    # General engineering skills
    ("java",               7), ("python",             6),  ("javascript",         5),
    ("typescript",         5), ("node.js",            5),  ("nodejs",             5),
    ("react",              4), ("spring boot",        5),  ("microservices",      5),
    ("rest api",           5), ("graphql",            4),  ("sql",                4),
    ("mysql",              3), ("postgresql",         3),  ("mongodb",            3),
    # DevOps / Cloud
    ("jenkins",            4), ("github actions",     4),  ("ci/cd",              4),
    ("ci cd",              4), ("aws",                5),  ("azure",              4),
    ("gcp",                4), ("docker",             4),  ("kubernetes",         5),
    # Other
    ("salesforce",         4), ("agile",              3),  ("scrum",              2),
    ("jira",               2), ("senior",             3),  ("lead",               3),
    ("staff",              2), ("principal",          2),
]
JUNIOR_SIGNALS = ["junior", "associate", "entry level", "entry-level", "intern", "0-2 years", "1-2 years"]
TITLE_BOOSTS   = {
    # QA
    "sdet": 15, "qa automation": 12, "test automation": 12, "quality engineer": 8,
    "automation engineer": 8, "qa lead": 10, "qa engineer": 7,
    "software engineer in test": 10,
    # Engineering
    "software engineer": 8, "backend engineer": 8, "full stack engineer": 8,
    "full-stack engineer": 8, "senior software engineer": 12, "staff engineer": 10,
    "platform engineer": 8, "devops engineer": 8, "site reliability engineer": 8,
}
MAX_POSSIBLE   = sum(w for _, w in SKILL_WEIGHTS) + 15

def ats_score_job(job: dict) -> dict:
    text = (job.get("title","") + " " + job.get("snippet","")).lower()
    matched, raw = [], 0
    for keyword, weight in SKILL_WEIGHTS:
        if keyword in text:
            matched.append(keyword); raw += weight
    for phrase, bonus in TITLE_BOOSTS.items():
        if phrase in job.get("title","").lower():
            raw += bonus; break
    if any(s in text for s in JUNIOR_SIGNALS):
        raw = max(0, raw - 20)
    score  = min(100, int((raw / MAX_POSSIBLE) * 100))
    important = ["playwright","selenium","java","python","ci/cd","api testing"]
    gaps   = [k for k in important if k not in matched]
    job["score"]         = score
    job["match_reasons"] = ", ".join(matched[:6])
    job["gaps"]          = ", ".join(gaps)
    job["ai_summary"]    = f"Matched {len(matched)} skills. " + ("Good seniority match." if any(s in text for s in ["senior","lead","staff"]) else "Check seniority level.")
    return job

def ats_score_batch(jobs: list[dict]) -> list[dict]:
    """Score all jobs locally — 100% free, no API calls."""
    for job in jobs:
        ats_score_job(job)
    return jobs


# ─────────────────────────────────────────────
# EXCEL EXPORT  — multi-sheet workbook
# Sheet 1: Job Tracker  (all scraped jobs)
# Sheet 2: Dashboard    (summary stats)
# Sheet 3: Legend       (column + status guide)
# ─────────────────────────────────────────────

EXCEL_COLS = [
    "Date Found",       # A  1
    "Source",           # B  2
    "Company",          # C  3
    "Job Title",        # D  4
    "ATS Score",        # E  5
    "Visa Type",        # F  6  ← NEW: H1B / OPT-CPT / STEM-OPT / UNKNOWN / NO SPONSOR
    "H1B Sponsor",      # G  7
    "OPT/CPT Ok",       # H  8  ← NEW
    "Match Skills",     # I  9
    "Gaps",             # J  10
    "AI Summary",       # K  11
    "Apply URL",        # L  12
    "Status",           # M  13 ← tracker dropdown values below
    "Interview Date",   # N  14 ← NEW
    "Follow-Up Date",   # O  15 ← NEW
    "Offer Details",    # P  16 ← NEW
    "Notes",            # Q  17
]

# ── Application status values & their colors ─────────────────────────────────
STATUS_COLORS = {
    "New":              "DDEEFF",   # soft blue
    "Saved":            "EEF0F1",   # light grey
    "Applied":          "BDE0FE",   # blue
    "Phone Screen":     "FFF1A8",   # yellow
    "Interview":        "FFD6A5",   # orange
    "Take-Home Test":   "FDFFB6",   # pale yellow
    "Final Round":      "CAFFBF",   # light green
    "Offer Received":   "52B788",   # green  (bold white text)
    "Offer Accepted":   "1B4332",   # dark green (bold white text)
    "Offer Declined":   "E9C46A",   # amber
    "Rejected":         "FFADAD",   # red
    "Ghosted":          "D3D3D3",   # grey
    "Withdrawn":        "F4A261",   # warm orange
    "On Hold":          "CDB4DB",   # lavender
}
STATUS_WHITE_TEXT = {"Offer Received", "Offer Accepted"}

# ── Visa type label → display color ──────────────────────────────────────────
VISA_TYPE_COLORS = {
    "H1B":       "C8F7C5",   # green
    "OPT-CPT":   "BDE0FE",   # blue
    "STEM-OPT":  "A8DADC",   # teal
    "POSSIBLE":  "FFF3CD",   # yellow  (unknown/unconfirmed)
    "NO SPONSOR":"F8D7DA",   # red
}

COL_WIDTHS = [16, 14, 22, 42, 10, 13, 12, 12, 35, 25, 45, 52, 16, 15, 15, 22, 22]


def _visa_type_label(job: dict) -> str:
    """Derive a single display label for the Visa Type column."""
    h1b = job.get("h1b_likely", "unknown")
    opt = job.get("opt_cpt_likely", "unknown")
    snippet = (job.get("snippet", "") + " " + job.get("title", "")).lower()

    # Explicit STEM-OPT mention
    if "stem opt" in snippet or "stem extension" in snippet:
        return "STEM-OPT"
    if opt == "yes" and h1b == "yes":
        return "H1B"          # H1B takes priority if both confirmed
    if h1b == "yes":
        return "H1B"
    if opt == "yes":
        return "OPT-CPT"
    if h1b == "no" and opt == "no":
        return "NO SPONSOR"
    return "POSSIBLE"         # neither confirmed nor denied


def _visa_eligible(job: dict) -> bool:
    """Return True if job passes the current VISA_MODE filter."""
    label = _visa_type_label(job)
    if label == "NO SPONSOR":
        return False
    if VISA_MODE == "h1b":
        return label in ("H1B", "POSSIBLE")
    if VISA_MODE == "opt_cpt":
        return label in ("OPT-CPT", "STEM-OPT", "POSSIBLE")
    if VISA_MODE == "stem_opt":
        return label in ("STEM-OPT", "POSSIBLE")
    # "all" — include everything except explicit no-sponsorship
    return True


def init_excel() -> openpyxl.Workbook:
    EXCEL_PATH.parent.mkdir(exist_ok=True)
    if EXCEL_PATH.exists():
        return openpyxl.load_workbook(EXCEL_PATH)

    wb = openpyxl.Workbook()

    # ── Sheet 1: Job Tracker ──────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Job Tracker"

    hdr_fill = PatternFill("solid", fgColor="1D3557")
    hdr_font = Font(bold=True, color="FFFFFF", size=11)
    thin     = Border(bottom=Side(style="thin", color="CCCCCC"))

    for ci, name in enumerate(EXCEL_COLS, 1):
        c = ws.cell(row=1, column=ci, value=name)
        c.fill = hdr_fill; c.font = hdr_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin

    for i, w in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 24
    ws.freeze_panes = "A2"

    # ── Sheet 2: Dashboard ────────────────────────────────────────────────────
    dash = wb.create_sheet("Dashboard")
    _build_dashboard(dash, [], is_init=True)

    # ── Sheet 3: Legend ───────────────────────────────────────────────────────
    leg = wb.create_sheet("Legend")
    _build_legend(leg)

    wb.save(EXCEL_PATH)
    return wb


def _build_dashboard(ws, jobs: list[dict], is_init=False):
    """Write summary stats to the Dashboard sheet."""
    ws.delete_rows(1, ws.max_row + 1)   # clear existing

    title_font  = Font(bold=True, size=14, color="1D3557")
    head_fill   = PatternFill("solid", fgColor="1D3557")
    head_font   = Font(bold=True, color="FFFFFF", size=11)
    sub_fill    = PatternFill("solid", fgColor="E8F4FD")
    sub_font    = Font(bold=True, size=11)

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 18

    # Title
    ws["A1"] = f"📊  Job Hunt Dashboard — {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws["A1"].font = title_font
    ws.row_dimensions[1].height = 28

    if is_init or not jobs:
        ws["A3"] = "No data yet — run the pipeline to populate."
        ws["A3"].font = Font(italic=True, color="888888")
        return

    # ── Visa type breakdown ───────────────────────────────────────────────────
    ws["A3"] = "Visa Type Breakdown"
    ws["A3"].font = head_font; ws["A3"].fill = head_fill
    ws["B3"].fill = head_fill; ws["C3"].fill = head_fill
    ws["B3"].font = head_font; ws["C3"].font = head_font
    ws["B3"] = "Count"; ws["C3"] = "% of Total"

    from collections import Counter
    visa_counts = Counter(_visa_type_label(j) for j in jobs)
    total = len(jobs)
    row = 4
    for label, color in VISA_TYPE_COLORS.items():
        count = visa_counts.get(label, 0)
        ws.cell(row, 1, label).font = Font(bold=True)
        ws.cell(row, 1).fill = PatternFill("solid", fgColor=color)
        ws.cell(row, 2, count)
        ws.cell(row, 3, f"{count/total*100:.1f}%" if total else "0%")
        row += 1

    # ── Status breakdown ─────────────────────────────────────────────────────
    row += 1
    ws.cell(row, 1, "Application Status").font = head_font
    ws.cell(row, 1).fill = head_fill
    ws.cell(row, 2, "Count").font = head_font; ws.cell(row, 2).fill = head_fill
    status_col = EXCEL_COLS.index("Status") + 1
    # Count statuses from jobs list
    status_counts = Counter(j.get("status", "New") for j in jobs)
    row += 1
    for status, color in STATUS_COLORS.items():
        count = status_counts.get(status, 0)
        ws.cell(row, 1, status).font = Font(bold=True)
        ws.cell(row, 1).fill = PatternFill("solid", fgColor=color)
        if status in STATUS_WHITE_TEXT:
            ws.cell(row, 1).font = Font(bold=True, color="FFFFFF")
        ws.cell(row, 2, count)
        row += 1

    # ── Top companies ─────────────────────────────────────────────────────────
    row += 1
    ws.cell(row, 1, "Top Companies (by job count)").font = head_font
    ws.cell(row, 1).fill = head_fill
    ws.cell(row, 2, "Jobs").font = head_font; ws.cell(row, 2).fill = head_fill
    ws.cell(row, 3, "Avg ATS").font = head_font; ws.cell(row, 3).fill = head_fill
    row += 1
    from itertools import groupby
    by_company: dict = {}
    for j in jobs:
        co = j.get("company", "Unknown")
        by_company.setdefault(co, []).append(j.get("score", 0))
    top = sorted(by_company.items(), key=lambda x: len(x[1]), reverse=True)[:10]
    for co, scores in top:
        ws.cell(row, 1, co); ws.cell(row, 2, len(scores))
        ws.cell(row, 3, f"{sum(scores)/len(scores):.0f}%")
        ws.cell(row, 1).fill = sub_fill
        row += 1


def _build_legend(ws):
    """Populate the Legend sheet with column explanations and status guide."""
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 55

    hdr_fill = PatternFill("solid", fgColor="1D3557")
    hdr_font = Font(bold=True, color="FFFFFF", size=11)
    sub_font = Font(bold=True, size=11, color="1D3557")

    def hrow(row, a, b=""):
        ws.cell(row, 1, a).fill = hdr_fill; ws.cell(row, 1).font = hdr_font
        ws.cell(row, 2, b).fill = hdr_fill; ws.cell(row, 2).font = hdr_font

    row = 1
    hrow(row, "Column", "Description"); row += 1
    for name in EXCEL_COLS:
        descs = {
            "Date Found":     "Timestamp the job was scraped",
            "Source":         "Portal or company site it came from",
            "Company":        "Employer name",
            "Job Title":      "Role title as listed",
            "ATS Score":      "Keyword match score vs your resume (0-100%)",
            "Visa Type":      "H1B / OPT-CPT / STEM-OPT / POSSIBLE / NO SPONSOR",
            "H1B Sponsor":    "YES=confirmed H1B, UNKNOWN=not stated, NO=explicitly rejected",
            "OPT/CPT Ok":     "YES=OPT/CPT signals detected, UNKNOWN=not stated, NO=rejected",
            "Match Skills":   "Top skills found in the job description matching your resume",
            "Gaps":           "Key skills mentioned in JD but not in your resume",
            "AI Summary":     "Auto-generated snippet about the match quality",
            "Apply URL":      "Direct link to the application page",
            "Status":         "Your application status — update manually (see status guide below)",
            "Interview Date": "Date of scheduled interview — fill in manually",
            "Follow-Up Date": "Date to follow up if no response — fill in manually",
            "Offer Details":  "Salary / benefits / notes if offer received",
            "Notes":          "Any free-form notes",
        }
        ws.cell(row, 1, name).font = Font(bold=True)
        ws.cell(row, 2, descs.get(name, ""))
        row += 1

    row += 1
    hrow(row, "Status Value", "Meaning / Color"); row += 1
    for status, color in STATUS_COLORS.items():
        c1 = ws.cell(row, 1, status)
        c1.fill = PatternFill("solid", fgColor=color)
        c1.font = Font(bold=True, color="FFFFFF" if status in STATUS_WHITE_TEXT else "000000")
        meanings = {
            "New":             "Just scraped — not yet reviewed",
            "Saved":           "Bookmarked to apply later",
            "Applied":         "Application submitted",
            "Phone Screen":    "Recruiter phone call scheduled/done",
            "Interview":       "Technical / behavioral interview scheduled",
            "Take-Home Test":  "Coding challenge / assignment received",
            "Final Round":     "In final interview stage",
            "Offer Received":  "Written offer in hand",
            "Offer Accepted":  "Offer accepted — congrats!",
            "Offer Declined":  "Offer declined",
            "Rejected":        "Application rejected",
            "Ghosted":         "No response after follow-up",
            "Withdrawn":       "You withdrew the application",
            "On Hold":         "Company paused hiring",
        }
        ws.cell(row, 2, meanings.get(status, ""))
        row += 1

    row += 1
    hrow(row, "Visa Type", "Meaning"); row += 1
    visa_meanings = {
        "H1B":       "Company confirmed H1B visa sponsorship",
        "OPT-CPT":   "Company accepts OPT / CPT work authorization",
        "STEM-OPT":  "Company explicitly mentions STEM-OPT extension",
        "POSSIBLE":  "No explicit mention — verify before applying",
        "NO SPONSOR":"Explicitly states no sponsorship / citizens only",
    }
    for label, meaning in visa_meanings.items():
        color = VISA_TYPE_COLORS[label]
        c1 = ws.cell(row, 1, label)
        c1.fill = PatternFill("solid", fgColor=color)
        c1.font = Font(bold=True)
        ws.cell(row, 2, meaning)
        row += 1


def append_jobs_to_excel(jobs: list[dict]):
    wb    = init_excel()
    ws    = wb["Job Tracker"]
    dash  = wb["Dashboard"]

    next_row = ws.max_row + 1

    score_colors = {"high": "C8F7C5", "medium": "FFF3CD", "low": "F8D7DA"}

    for job in jobs:
        score      = job.get("score", 0)
        tier       = "high" if score >= 80 else "medium" if score >= 65 else "low"
        h1b        = job.get("h1b_likely", "unknown")
        opt        = job.get("opt_cpt_likely", "unknown")
        visa_label = _visa_type_label(job)
        visa_color = VISA_TYPE_COLORS.get(visa_label, "FFFFFF")

        # H1B column color
        h1b_color = "C8F7C5" if h1b == "yes" else "FFF3CD" if h1b == "unknown" else "F8D7DA"
        # OPT/CPT column color
        opt_color = "BDE0FE" if opt == "yes" else "FFF3CD" if opt == "unknown" else "F8D7DA"

        row_data = [
            datetime.now().strftime("%Y-%m-%d %H:%M"),  # A Date Found
            job.get("source", "Company Site"),           # B Source
            job.get("company", ""),                      # C Company
            job.get("title", ""),                        # D Job Title
            score,                                       # E ATS Score
            visa_label,                                  # F Visa Type
            h1b.upper(),                                 # G H1B Sponsor
            opt.upper(),                                 # H OPT/CPT Ok
            job.get("match_reasons", ""),                # I Match Skills
            job.get("gaps", ""),                         # J Gaps
            job.get("ai_summary", ""),                   # K AI Summary
            job.get("url", ""),                          # L Apply URL
            "New",                                       # M Status
            "",                                          # N Interview Date
            "",                                          # O Follow-Up Date
            "",                                          # P Offer Details
            "",                                          # Q Notes
        ]

        for ci, value in enumerate(row_data, 1):
            cell = ws.cell(row=next_row, column=ci, value=value)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if ci == 5:   # ATS Score
                cell.fill = PatternFill("solid", fgColor=score_colors[tier])
                cell.font = Font(bold=True)
            elif ci == 6: # Visa Type
                cell.fill = PatternFill("solid", fgColor=visa_color)
                cell.font = Font(bold=True)
            elif ci == 7: # H1B Sponsor
                cell.fill = PatternFill("solid", fgColor=h1b_color)
                cell.font = Font(bold=True)
            elif ci == 8: # OPT/CPT Ok
                cell.fill = PatternFill("solid", fgColor=opt_color)
                cell.font = Font(bold=True)
            elif ci == 13: # Status
                status_color = STATUS_COLORS.get("New", "DDEEFF")
                cell.fill = PatternFill("solid", fgColor=status_color)

        ws.row_dimensions[next_row].height = 38
        next_row += 1

    # Refresh dashboard with all rows
    all_jobs_from_sheet = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]:  # Date Found not empty
            all_jobs_from_sheet.append({
                "h1b_likely":     (row[6] or "unknown").lower(),
                "opt_cpt_likely": (row[7] or "unknown").lower(),
                "score":          row[4] or 0,
                "company":        row[2] or "",
                "snippet":        row[10] or "",
                "title":          row[3] or "",
                "status":         row[12] or "New",
            })
    _build_dashboard(dash, all_jobs_from_sheet)

    wb.save(EXCEL_PATH)
    log.info(f"Saved {len(jobs)} jobs to {EXCEL_PATH}")


# ─────────────────────────────────────────────
# EMAIL APPLICATION
# ─────────────────────────────────────────────

EMAIL_TEMPLATE = """\
Dear Hiring Team at {company},

I am writing to express my strong interest in the {job_title} position at {company}.

{resume_highlight}

I require visa sponsorship to work in the US and am actively seeking an employer
who can support that process. I would welcome the opportunity to discuss how my
background can strengthen {company}'s team.

Apply link: {apply_url}

Best regards,
{your_name}
{your_email}{phone_line}{linkedin_line}
"""

def send_application_email(job: dict, to_email: str | None = None):
    """Send an application email. Defaults to emailing yourself if no recipient given."""
    if not GMAIL_APP_PASS:
        log.warning("GMAIL_APP_PASS not set — skipping email send.")
        return False

    # Build optional signature lines only when the values are set
    phone_line   = f"\n{YOUR_PHONE}"    if YOUR_PHONE    else ""
    linkedin_line = f"\n{YOUR_LINKEDIN}" if YOUR_LINKEDIN else ""

    # Resume highlight: first non-empty line of RESUME_SUMMARY, or a generic fallback
    highlight_lines = [l.strip() for l in RESUME_SUMMARY.strip().splitlines() if l.strip()]
    resume_highlight = highlight_lines[0] if highlight_lines else \
        "I bring strong technical skills and a track record of delivering quality software."

    body = EMAIL_TEMPLATE.format(
        job_title        = job["title"],
        company          = job["company"],
        your_name        = YOUR_NAME,
        your_email       = YOUR_EMAIL,
        apply_url        = job.get("url", "See attachment"),
        resume_highlight = resume_highlight,
        phone_line       = phone_line,
        linkedin_line    = linkedin_line,
    )

    recipient = to_email or YOUR_EMAIL  # default: email yourself the digest

    msg = MIMEMultipart()
    msg["From"]    = YOUR_EMAIL
    msg["To"]      = recipient
    msg["Subject"] = f"Application: {job['title']} @ {job['company']} | {YOUR_TITLE} | ATS {job.get('score',0)}%"

    msg.attach(MIMEText(body, "plain"))

    # Attach resume — download from RESUME_URL secret or use local file
    resume_path = Path("config/resume.pdf")
    resume_url  = os.getenv("RESUME_URL", "")
    if not resume_path.exists() and resume_url:
        try:
            r = requests.get(resume_url, timeout=20)
            r.raise_for_status()
            resume_path.parent.mkdir(exist_ok=True)
            resume_path.write_bytes(r.content)
            log.info("Resume downloaded from RESUME_URL.")
        except Exception as e:
            log.warning(f"Could not download resume: {e}")
    if resume_path.exists():
        with open(resume_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            resume_filename = f"{YOUR_NAME.replace(' ', '_')}_Resume.pdf" if YOUR_NAME else "Resume.pdf"
            part.add_header("Content-Disposition",
                            f'attachment; filename="{resume_filename}"')
            msg.attach(part)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(YOUR_EMAIL, GMAIL_APP_PASS)
            smtp.sendmail(YOUR_EMAIL, recipient, msg.as_string())
        log.info(f"Email sent for: {job['title']} @ {job['company']}")
        return True
    except Exception as e:
        log.error(f"Email failed: {e}")
        return False


def send_digest_email(shortlisted: list[dict]):
    """Send yourself a summary digest grouped by visa tier."""
    if not GMAIL_APP_PASS or not shortlisted:
        return

    t1 = [j for j in shortlisted if _visa_type_label(j) == "H1B"]
    t2 = [j for j in shortlisted if _visa_type_label(j) in ("OPT-CPT", "STEM-OPT")]
    t3 = [j for j in shortlisted if _visa_type_label(j) == "POSSIBLE"]

    def fmt(jobs):
        return "\n".join(
            f"  [{j['score']}%] {j['title']} @ {j['company']}  [{_visa_type_label(j)}]\n"
            f"         {j.get('url','no url')}\n"
            f"         Skills: {j.get('match_reasons','')}"
            for j in jobs
        ) or "  (none)"

    body = f"""Job Hunt Digest — {datetime.now().strftime('%Y-%m-%d %H:%M')}   [Mode: {VISA_MODE.upper()}]
{'='*65}

{len(shortlisted)} shortlisted jobs (ATS ≥ {MIN_ATS_SCORE}%):

✅  CONFIRMED H1B SPONSORSHIP ({len(t1)})
{fmt(t1)}

🎓  OPT / CPT / STEM-OPT ACCEPTED ({len(t2)})
{fmt(t2)}

🔍  POSSIBLE — VERIFY BEFORE APPLYING ({len(t3)})
{fmt(t3)}

Full tracker: job_tracker.xlsx  (3 sheets: Job Tracker · Dashboard · Legend)
"""
    msg = MIMEMultipart()
    msg["From"]    = YOUR_EMAIL
    msg["To"]      = YOUR_EMAIL
    msg["Subject"] = f"[Job Hunt] {len(shortlisted)} new matches — {datetime.now().strftime('%b %d')}"
    msg.attach(MIMEText(body, "plain"))

    # Attach the Excel file
    if EXCEL_PATH.exists():
        with open(EXCEL_PATH, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", 'attachment; filename="job_tracker.xlsx"')
            msg.attach(part)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(YOUR_EMAIL, GMAIL_APP_PASS)
            smtp.sendmail(YOUR_EMAIL, YOUR_EMAIL, msg.as_string())
        log.info("Digest email sent.")
    except Exception as e:
        log.error(f"Digest email failed: {e}")


# ─────────────────────────────────────────────
# MAIN PIPELINE
# ─────────────────────────────────────────────

def run_pipeline():
    log.info("=" * 55)
    log.info(f"Job Hunt Pipeline [{VISA_MODE.upper()} mode] — {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    log.info("=" * 55)

    seen = load_seen()

    # 1. Scrape
    log.info("STEP 1/5 — Scraping career pages...")
    raw_jobs = [annotate_job(job) for job in scrape_all_companies()]
    log.info(f"  Total raw listings: {len(raw_jobs)}")

    # 2. De-duplicate
    log.info("STEP 2/5 — De-duplicating...")
    new_jobs = []
    for job in raw_jobs:
        jid = job_id(job["title"], job["company"])
        if jid not in seen:
            job["job_id"] = jid
            new_jobs.append(job)
    log.info(f"  New (unseen) listings: {len(new_jobs)}")

    if not new_jobs:
        log.info("No new jobs found. Pipeline done.")
        return

    # 3. ATS Score
    log.info("STEP 3/5 — ATS scoring with Claude...")
    scored_jobs = ats_score_batch(new_jobs)

    # 4. Filter: USA / India + role match + visa mode
    log.info("STEP 4/5 — Filtering & shortlisting...")
    location_filtered = [j for j in scored_jobs if location_allowed(j)]
    log.info(f"  Location filtered (USA/India): {len(location_filtered)}")
    role_filtered = [j for j in location_filtered if role_allowed(j)]
    log.info(f"  Role filtered (all engineering roles): {len(role_filtered)}")

    # Apply visa mode filter (respects VISA_MODE env var)
    visa_eligible = [j for j in role_filtered if _visa_eligible(j)]
    log.info(f"  Visa-eligible (mode={VISA_MODE}): {len(visa_eligible)}")

    # ── Tier 1: Confirmed H1B ────────────────────────────────────────────────
    confirmed_h1b = sorted(
        [j for j in visa_eligible if j.get("score", 0) >= MIN_ATS_SCORE
         and _visa_type_label(j) == "H1B"],
        key=lambda x: x.get("score", 0), reverse=True
    )
    # ── Tier 2: OPT / CPT / STEM-OPT confirmed ───────────────────────────────
    confirmed_opt = sorted(
        [j for j in visa_eligible if j.get("score", 0) >= MIN_ATS_SCORE
         and _visa_type_label(j) in ("OPT-CPT", "STEM-OPT")],
        key=lambda x: x.get("score", 0), reverse=True
    )
    # ── Tier 3: Possible (unknown sponsorship) ────────────────────────────────
    possible = sorted(
        [j for j in visa_eligible if j.get("score", 0) >= MIN_ATS_SCORE
         and _visa_type_label(j) == "POSSIBLE"],
        key=lambda x: x.get("score", 0), reverse=True
    )

    log.info(f"  Confirmed H1B ≥{MIN_ATS_SCORE}%: {len(confirmed_h1b)}")
    log.info(f"  Confirmed OPT/CPT/STEM-OPT ≥{MIN_ATS_SCORE}%: {len(confirmed_opt)}")
    log.info(f"  Possible (unconfirmed) ≥{MIN_ATS_SCORE}%: {len(possible)}")

    shortlisted = confirmed_h1b + confirmed_opt + possible   # priority order

    # 5. Save all scored to Excel
    log.info("STEP 5/5 — Saving to Excel & sending emails...")
    append_jobs_to_excel(scored_jobs)

    # Update seen set with all new jobs
    for j in new_jobs:
        seen.add(j["job_id"])
    save_seen(seen)

    # Send individual applications for top shortlisted
    applied_count = 0
    for job in shortlisted[:MAX_APPLY_PER_RUN]:
        send_application_email(job)
        applied_count += 1
        time.sleep(3)

    # Send digest to yourself
    send_digest_email(shortlisted)

    log.info("─" * 55)
    log.info(f"Done. Scraped: {len(raw_jobs)} | New: {len(new_jobs)} | "
             f"H1B: {len(confirmed_h1b)} | OPT/CPT: {len(confirmed_opt)} | "
             f"Possible: {len(possible)} | Applied: {applied_count}")
    log.info("─" * 55)


def validate_config():
    """Warn loudly if required env vars are missing before the pipeline starts."""
    missing = []
    if not YOUR_NAME:   missing.append("YOUR_NAME")
    if not YOUR_EMAIL:  missing.append("YOUR_EMAIL")
    if missing:
        log.error("=" * 55)
        log.error("MISSING REQUIRED CONFIGURATION:")
        for v in missing:
            log.error(f"  ✗  {v}  — set this as a GitHub Secret or env var")
        log.error("See the module docstring at the top of job_hunter.py for setup instructions.")
        log.error("=" * 55)
        raise SystemExit(1)
    log.info(f"Config OK — user: {YOUR_NAME} | query: '{JOB_SEARCH_QUERY}' | visa mode: {VISA_MODE}")


if __name__ == "__main__":
    validate_config()
    run_pipeline()
