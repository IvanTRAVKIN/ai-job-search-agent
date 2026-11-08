"""
The bot searches LinkedIn for jobs and makes a tailored CV for each one. You apply yourself.
 
What the bot does:
  1. Opens Google Chrome (or Edge) with your LinkedIn login (a separate bot profile).
  2. Searches jobs with the phrases from SEARCH_QUERIES (or SEARCH_KEYWORDS in SEARCH_LOCATION),
     only recent ones (DATE_POSTED: 24h / week / month), with or without Easy Apply.
  3. For each new job it reads the description, checks whether it fits you,
     and makes a CV (and a cover letter) in the output/ folder.
  4. Saves everything to output/applications.csv and output/vacancies.html: the job link,
     how to apply (Easy Apply or company site) and the company site link.
Then you open the page, follow the links and apply yourself with the ready CV.
 
The bot never submits anything. For jobs without Easy Apply it only clicks the
"Apply" button to get the company site link, and closes that tab right away.
 
Run:  python find_jobs.py
"""
 
import csv
import datetime
import os
import random
import re
import sys
import time
import webbrowser
from urllib.parse import urlencode
 
import generate_cv as gen
 
try:
    from playwright.sync_api import sync_playwright, Error as PlaywrightError, TimeoutError as PlaywrightTimeout
except ImportError:
    print("The playwright library is missing. Run this command:")
    print("    python -m pip install -r requirements.txt")
    sys.exit(1)
 
PROFILE_DIR = gen.BASE_DIR / "browser_profile"
 
# LinkedIn often changes its layout. If the bot stops finding jobs or their text,
# usually fixing these selectors is enough.
JOB_CARD = "li[data-occludable-job-id]"
JOB_DESCRIPTION = ("#job-details, .jobs-description__content, .jobs-description-content__text, "
                   ".jobs-box__html-content, [class*='jobs-description'], .description__text")
 
# Fallback: find the "About the job" / "À propos de l'offre" heading and take the block below it.
DESCRIPTION_BY_HEADING_JS = """() => {
  const re = /about the job|à propos de l.offre|à propos du poste|about this job|description du poste/i;
  const h = [...document.querySelectorAll('h1,h2,h3,h4,span,strong')].find(x => re.test(x.innerText || ''));
  if (!h) return '';
  let el = h.parentElement;
  for (let i = 0; i < 6 && el && el.innerText.length < 400; i++) el = el.parentElement;
  return el ? el.innerText : '';
}"""
# Apply button: LinkedIn often changes its layout, so we search both by class and by text/label.
APPLY_BUTTON = ("button.jobs-apply-button, a.jobs-apply-button, "
                "button[aria-label*='Apply' i], a[aria-label*='Apply' i], "
                "button[aria-label*='Postuler' i], a[aria-label*='Postuler' i], "
                "button[aria-label*='Candidature' i], a[aria-label*='Candidature' i]")
APPLY_TEXT = re.compile(r"^\s*(easy apply|apply|postuler|candidature simplifiée|candidater)\b", re.IGNORECASE)
CLOSED_TEXT = ("no longer accepting applications", "n’accepte plus de candidatures",
               "n'accepte plus de candidatures", "candidatures ne sont plus acceptées")
 
# LinkedIn "Date posted" filter: seconds for the f_TPR parameter
DATE_FILTERS = {"24h": "r86400", "week": "r604800", "month": "r2592000"}
 
 
def search_urls():
    """LinkedIn search URLs for each job title, newest first."""
    keywords = [k.strip() for k in os.getenv(
        "SEARCH_KEYWORDS", "Business Analyst;Data Analyst;Pricing Analyst").split(";") if k.strip()]
    location = os.getenv("SEARCH_LOCATION", "Paris, Île-de-France, France").strip()
    date_filter = DATE_FILTERS.get(os.getenv("DATE_POSTED", "week").strip().lower(), "r604800")
    urls = []
    for keyword in keywords:
        # A title can have its own period: "Data Analyst|24h"
        keyword, _, own_date = keyword.partition("|")
        keyword = keyword.strip()
        period = DATE_FILTERS.get(own_date.strip().lower(), date_filter)
        # Words to exclude in the LinkedIn search (none by default: plain word search)
        minus = [w.strip() for w in os.getenv("SEARCH_EXCLUDE", "").split(";") if w.strip()]
        query = keyword + "".join(f" NOT {w}" for w in minus)
        params = {"keywords": query, "location": location, "f_TPR": period, "sortBy": "DD"}
        levels = os.getenv("EXPERIENCE_LEVELS", "2,3").strip()  # 1 internship, 2 entry level, 3 associate
        if levels:
            params["f_E"] = levels
        urls.append((keyword, "https://www.linkedin.com/jobs/search/?" + urlencode(params)))
    return urls
 
 
# Titles that clearly don't fit: these jobs are not even opened (visible from the title in the list).
SKIP_TITLE_DEFAULT = ("senior;sénior;lead;head of;principal;manager;director;directeur;directrice;"
                      "data engineer;data scientist;quant;développeur;developer;stage;stagiaire;intern;"
                      "internship;alternance;alternant;apprenti")
 
 
def skip_by_title(title):
    words = os.getenv("SKIP_TITLE_WORDS", SKIP_TITLE_DEFAULT).split(";")
    lower = title.lower()
    for word in (w.strip().lower() for w in words):
        if word and re.search(rf"(?<!\w){re.escape(word)}(?!\w)", lower):
            return word
    return ""
 
 
FRENCH_TITLE = re.compile(r"\b(H\s*/\s*F|F\s*/\s*H|chargée?|analyste|consultante?\s*\(e\)|ingénieure?|"
                          r"responsable|alternance|stagiaire|développeur|données|métier|contrôleur|confirmée?)\b",
                          re.IGNORECASE)
 
 
def french_title(title):
    """The title is clearly French (H/F, Chargé, Analyste...): not opened in English-only mode."""
    # Off by default: some jobs have a French title but an English description.
    if os.getenv("SKIP_FRENCH_TITLES", "no").strip().lower() not in gen.YES:
        return False
    return os.getenv("JOB_LANGUAGE", "any") == "en" and bool(FRENCH_TITLE.search(title))
 
 
def collect_cards(page):
    """Job links and titles from the search page (scrolls the list so LinkedIn shows all cards)."""
    cards = page.locator(JOB_CARD)
    result = []
    for i in range(cards.count()):
        card = cards.nth(i)
        job_id = card.get_attribute("data-occludable-job-id")
        if not job_id:
            continue
        try:
            card.scroll_into_view_if_needed(timeout=3000)
            page.wait_for_timeout(150)
            text = card.inner_text(timeout=3000).strip()
        except PlaywrightError:
            text = ""
        title = text.splitlines()[0].strip() if text else ""
        result.append((f"https://www.linkedin.com/jobs/view/{job_id}/", title))
    return result
 
 
def find_apply_button(page):
    """The first visible "Apply / Easy Apply / Postuler" button or link on the job page."""
    candidates = [page.locator(APPLY_BUTTON),
                  page.get_by_role("button", name=APPLY_TEXT),
                  page.get_by_role("link", name=APPLY_TEXT)]
    for locator in candidates:
        try:
            for i in range(min(locator.count(), 6)):
                item = locator.nth(i)
                if item.is_visible():
                    return item
        except PlaywrightError:
            continue
    return None
 
 
def button_label(button):
    text = (button.inner_text() or "") + " " + (button.get_attribute("aria-label") or "")
    return text.lower()
 
 
def apply_type(page):
    """Easy Apply, company site, or no button."""
    button = find_apply_button(page)
    if button is None:
        main_text = page.locator("main").first.inner_text().lower() if page.locator("main").count() else ""
        if any(t in main_text for t in CLOSED_TEXT):
            return "job closed"
        return "button not found, open the LinkedIn link"
    label = button_label(button)
    if "easy" in label or "simplifiée" in label:
        return "Easy Apply on LinkedIn"
    return "on company site"
 
 
def apply_info(page):
    """How to apply to the job and the company site link (if not Easy Apply)."""
    how = apply_type(page)
    if how != "on company site":
        return how, ""
    button = find_apply_button(page)
    href = button.get_attribute("href") or ""
    if href.startswith("http") and "linkedin.com" not in href:
        return "on company site", href
    # The "Apply / Postuler" button opens the company site in a new tab: take the link and close it.
    try:
        with page.context.expect_page(timeout=10000) as new_tab:
            button.click()
        tab = new_tab.value
        tab.wait_for_load_state("domcontentloaded", timeout=15000)
        url = tab.url
        tab.close()
        return "on company site", url
    except PlaywrightTimeout:
        page.keyboard.press("Escape")
        return "on company site (could not get the link, open the job)", ""
 
 
def pause(min_s=4, max_s=9):
    """A human-like pause, so pages are not opened too fast."""
    time.sleep(random.uniform(min_s, max_s))
 
 
def already_processed():
    """Links of jobs that are already in the application tracker."""
    if not gen.TRACKER.exists():
        return set()
    with open(gen.TRACKER, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f, delimiter=";"))
    if "--recheck" in sys.argv:
        # python find_jobs.py --recheck : check again the jobs that the AI scored 25-49
        def borderline(r):
            score = r.get("Score", "")
            return r.get("Fit") == "no" and score.isdigit() and 25 <= int(score) < 50
        again = [r for r in rows if borderline(r)]
        rows = [r for r in rows if not borderline(r)]
        with open(gen.TRACKER, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=gen.TRACKER_COLUMNS, delimiter=";", extrasaction="ignore")
            writer.writeheader()
            writer.writerows(rows)
        print(f"Will recheck {len(again)} jobs scored 25-49 (if they are still in the search).")
    return {r.get("Link", "") for r in rows if r.get("Link")}
 
 
# Clicks "… more" / "… plus" under the description to show the full job text.
EXPAND_DESCRIPTION_JS = r"""() => {
  const re = /^(…|\.\.\.)?\s*(more|see more|show more|plus|voir plus|afficher plus)$/i;
  const heading = /about the job|à propos de l.offre|à propos du poste|about this job/i;
  const h = [...document.querySelectorAll('h1,h2,h3,h4')].find(x => heading.test(x.innerText || ''));
  let scope = document;
  if (h) {
    let el = h.parentElement;
    for (let i = 0; i < 6 && el && el.innerText.length < 400; i++) el = el.parentElement;
    if (el) scope = el;
  }
  let clicked = 0;
  for (const b of scope.querySelectorAll('button, [role=button]')) {
    const label = (b.innerText || '').trim();
    const aria = (b.getAttribute('aria-label') || '').toLowerCase();
    if (re.test(label) || aria.includes('see more description') || aria.includes('voir plus')) {
      b.click(); clicked++;
    }
  }
  return clicked;
}"""
 
 
def expand_description(page):
    try:
        if page.evaluate(EXPAND_DESCRIPTION_JS):
            time.sleep(1.5)
    except PlaywrightError:
        pass  # no button: the full text is already shown
 
 
def read_description(page):
    """The job text. Tries several ways, because LinkedIn often changes its layout."""
    for _ in range(8):  # up to ~16 seconds while the page loads
        expand_description(page)
        box = page.locator(JOB_DESCRIPTION).first
        if box.count() > 0:
            text = box.inner_text().strip()
            if len(text) > 200:
                return text
        text = page.evaluate(DESCRIPTION_BY_HEADING_JS).strip()
        if len(text) > 200:
            return text
        time.sleep(2)
    # Last resort: the whole page text (Claude will find the description in it).
    text = page.evaluate("() => (document.querySelector('main') || document.body).innerText").strip()
    if len(text) > 500:
        return text[:15000]
    debug = gen.OUTPUT_DIR / "debug_page.html"
    gen.OUTPUT_DIR.mkdir(exist_ok=True)
    debug.write_text(page.content(), encoding="utf-8")
    print(f"  (page saved to {debug.name} for debugging, check it if this happens with every job)")
    return ""
 
 
def goto(page, url):
    """Opens a page. LinkedIn sometimes redirects to another page by itself
    (for example after login): then we wait and try again."""
    for _ in range(4):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=45000)
            return
        except PlaywrightTimeout:
            print("  LinkedIn is slow to respond, waiting 20 seconds and trying again...")
            time.sleep(20)
        except PlaywrightError as e:
            if "interrupted by another navigation" not in str(e) and "ERR_ABORTED" not in str(e):
                raise
            time.sleep(5)
    page.goto(url, wait_until="domcontentloaded", timeout=45000)
 
 
def logged_in(page):
    return not any(word in page.url for word in ("login", "checkpoint", "authwall", "signup", "uas/"))
 
 
def wait_for_login(page):
    goto(page, "https://www.linkedin.com/feed/")
    while not logged_in(page):
        print("\nLog in to LinkedIn in the browser window that opened (needed only once).")
        input("When you see your LinkedIn feed, come back here and press Enter... ")
        time.sleep(3)
        goto(page, "https://www.linkedin.com/feed/")
    print("Logged in to LinkedIn.")
 
 
# Fallback way to find jobs on the page if LinkedIn changed the list layout
JOB_LINKS_JS = r"""() => {
  const out = [];
  const seen = new Set();
  const add = (id, title) => { if (id && !seen.has(id)) { seen.add(id);
    let t = (title || '').split(/\s*\((?:Verified job|Offre vérifiée)\)/i)[0].trim();
    const half = t.length / 2;  // in the new layout the title is repeated twice in a row
    if (t.length % 2 === 0 && t.slice(0, half) === t.slice(half)) t = t.slice(0, half);
    out.push([id, t]); } };
  document.querySelectorAll('[data-occludable-job-id],[data-job-id]').forEach(el =>
    add(el.getAttribute('data-occludable-job-id') || el.getAttribute('data-job-id'),
        (el.innerText || '').split('\n')[0]));
  document.querySelectorAll('a[href*="/jobs/view/"], a[href*="currentJobId="]').forEach(a => {
    const m = a.href.match(/\/jobs\/view\/(\d+)/) || a.href.match(/currentJobId=(\d+)/);
    if (m) add(m[1], (a.innerText || '').split('\n')[0]);
  });
  // New LinkedIn layout: the job id is hidden in attributes like "urn:li:fsd_jobPosting:123"
  document.querySelectorAll('*').forEach(el => {
    for (const at of el.attributes) {
      const m = at.value.match(/jobPosting(?::|%3A|\(|,)+(\d{6,})/) || at.value.match(/job-card-component-ref-(\d{6,})/);
      if (m) { add(m[1], (el.innerText || '').split('\n')[0]); break; }
    }
  });
  return out;
}"""
SEARCH_BOX = ("input[aria-label*='Search by title' i], input[aria-label*='Rechercher par' i], "
              "input[aria-label*='job' i][role='combobox'], input[aria-label*='emploi' i][role='combobox'], "
              "input[placeholder*='job' i], input[placeholder*='emploi' i]")
 
 
def search_queries():
    """Search phrases from .env: SEARCH_QUERIES (as you type in the LinkedIn search box), separated by ;"""
    lang = os.getenv("JOB_LANGUAGE", "any").strip().upper()[:2]
    raw = (os.getenv(f"SEARCH_QUERIES_{lang}", "") or os.getenv("SEARCH_QUERIES", "")).strip()
    return [q.strip() for q in raw.split(";") if q.strip()]
 
 
def to_france(query):
    """'... paris ...' -> '... france ...' for the second search round."""
    return re.sub(r"(?i)\bparis\b", "france", query)
 
 
def type_search(page, query):
    """Types the phrase into the LinkedIn search box and presses Enter, like a human. True if it worked."""
    goto(page, "https://www.linkedin.com/jobs/")
    pause(3, 6)
    box = page.locator(SEARCH_BOX)
    for i in range(min(box.count(), 4)):
        item = box.nth(i)
        try:
            if item.is_visible():
                item.click()
                item.fill("")
                item.type(query, delay=random.randint(40, 90))
                page.keyboard.press("Enter")
                page.wait_for_load_state("domcontentloaded", timeout=20000)
                return True
        except PlaywrightError:
            continue
    return False
 
 
SCROLL_RESULTS_JS = r"""() => {
  // Scrolls all scrollable blocks (the job list) and the page itself down
  const boxes = [...document.querySelectorAll('*')].filter(el => {
    const st = getComputedStyle(el);
    return /(auto|scroll)/.test(st.overflowY) && el.scrollHeight > el.clientHeight + 50;
  });
  boxes.forEach(el => el.scrollBy(0, el.clientHeight * 0.9));
  window.scrollBy(0, window.innerHeight * 0.9);
  return boxes.length;
}"""
 
 
def cards_on_page(page):
    """Job links and titles on the open search page. Scrolls the list so LinkedIn loads all of them."""
    try:
        page.wait_for_function(
            "() => document.querySelector('[data-occludable-job-id],[data-job-id],a[href*=\"/jobs/view/\"],a[href*=\"currentJobId=\"]')",
            timeout=20000)
    except PlaywrightTimeout:
        return []
    page.wait_for_timeout(2000)
    found = {}
    same = 0
    for _ in range(25):  # keep scrolling while new cards appear
        before = len(found)
        for job_id, title in page.evaluate(JOB_LINKS_JS):
            if title or job_id not in found:
                found[job_id] = title or found.get(job_id, "")
        same = same + 1 if len(found) == before else 0
        if same >= 3:
            break
        page.evaluate(SCROLL_RESULTS_JS)
        page.wait_for_timeout(random.randint(700, 1200))
    if len(found) < 5:  # suspiciously few: save the page for debugging
        debug = gen.OUTPUT_DIR / "debug_search.html"
        gen.OUTPUT_DIR.mkdir(exist_ok=True)
        debug.write_text(page.content(), encoding="utf-8")
        print(f"  (few jobs on the page, saved it to {debug} for checking)")
    return [(f"https://www.linkedin.com/jobs/view/{job_id}/", title) for job_id, title in found.items()]
 
 
def with_start(url, start):
    """The same search, but from the next results page."""
    url = re.sub(r"([?&])start=\d+&?", r"\1", url).rstrip("&?")
    return url + ("&" if "?" in url else "?") + f"start={start}"
 
 
def collect(page, label, first_url, seen, taken, typed_query=None):
    """Collects new jobs from one search (up to SEARCH_PAGES pages)."""
    pages = int(os.getenv("SEARCH_PAGES", "3"))
    found, fresh, skipped = 0, [], 0
    if typed_query is not None:
        if not type_search(page, typed_query):
            goto(page, first_url)  # search box not found, opening the search by URL
        base_url = None
    else:
        goto(page, first_url)
        base_url = first_url
    pause()
    for page_no in range(pages):
        if page_no:
            # New search page: "Page 2", "Page 3" buttons at the bottom, click them like a human
            button = page.locator(f"button[aria-label='Page {page_no + 1}'], button[aria-label='Page {page_no + 1} ']")
            if button.count():
                button.first.scroll_into_view_if_needed()
                button.first.click()
                page.wait_for_timeout(3000)
            else:
                goto(page, with_start(base_url, 25 * page_no))
            pause()
        if base_url is None:
            base_url = page.url  # the URL LinkedIn built from the phrase
        cards = cards_on_page(page)
        if not cards:
            if page_no == 0:
                print(f"Search '{label}': no jobs found (or LinkedIn changed the page). URL: {page.url}")
            break
        found += len(cards)
        new_here = 0
        for link, title in cards:
            if link in seen or link in taken or link in fresh:
                continue
            new_here += 1
            if title and (skip_by_title(title) or french_title(title)):
                skipped += 1
                continue
            fresh.append(link)
        if len(cards) < 20 or new_here == 0:
            break  # last page, or only already seen jobs from here on
    print(f"Search '{label}': found {found}, new with a fitting title {len(fresh)}, "
          f"skipped by title {skipped} (senior, lead, manager, internships, etc."
          + (", French titles" if os.getenv("JOB_LANGUAGE", "any") == "en"
             and os.getenv("SKIP_FRENCH_TITLES", "no").strip().lower() in gen.YES else "") + ")")
    return fresh
 
 
def english_keyword_rounds(page, seen):
    """Extra search for en mode: words from SEARCH_KEYWORDS_EN + English phrases in the description."""
    en_keywords = [k.strip() for k in os.getenv("SEARCH_KEYWORDS_EN", "").split(";") if k.strip()]
    if os.getenv("JOB_LANGUAGE", "any") == "en" and en_keywords:
        # English mode: classic LinkedIn keyword search + English phrases that only appear
        # in English descriptions. This way LinkedIn returns mostly English postings.
        marker = os.getenv("ENGLISH_MARKER",
                           '("you will" OR "we are looking for" OR "your role" OR "you have" OR "about you")')
        period = DATE_FILTERS.get(os.getenv("DATE_POSTED", "week").strip().lower(), "r604800")
        locations = [("Paris", os.getenv("SEARCH_LOCATION", "Paris, Île-de-France, France").strip())]
        if os.getenv("FRANCE_FALLBACK", "yes").strip().lower() in gen.YES:
            locations.append(("all of France", "France"))
        for name, location in locations:
            print(f"\n=== Search (English jobs): {name} ===")
            taken = []
            for keyword in en_keywords:
                query = f'"{keyword}" AND {marker}' if marker else keyword
                url = "https://www.linkedin.com/jobs/search/?" + urlencode(
                    {"keywords": query, "location": location, "f_TPR": period, "sortBy": "DD"})
                taken += collect(page, f"{keyword} (EN)", url, seen, taken)
            yield name, taken
 
 
def search_rounds(page, seen):
    """Round 1: your phrases (Paris). Round 2: the same phrases for all of France. Yields new links per round."""
    queries = search_queries()
    if queries:
        rounds = [("Paris", queries)]
        france = [to_france(q) for q in queries if to_france(q) != q]
        if france and os.getenv("FRANCE_FALLBACK", "yes").strip().lower() in gen.YES:
            rounds.append(("all of France", france))
        for name, round_queries in rounds:
            print(f"\n=== Search: {name} ===")
            taken = []
            for q in round_queries:
                url = "https://www.linkedin.com/jobs/search/?" + urlencode({"keywords": q})
                taken += collect(page, q, url, seen, taken, typed_query=q)
            yield name, taken
    else:  # old way: SEARCH_KEYWORDS + filters
        taken = []
        for keyword, url in search_urls():
            taken += collect(page, keyword, url, seen, taken)
        yield "search", taken
    # In en mode, after your phrases: an extra keyword search with English phrases in the description
    yield from english_keyword_rounds(page, seen)
 
 
APPLICANTS_RE = re.compile(
    r"(over|plus de)?\s*(\d[\d\s,.\u202f]*)\s*(people clicked apply|applicants|candidats|personnes ont cliqué)",
    re.IGNORECASE)
 
 
def applicants_count(page):
    """How many people already applied (from "29 people clicked apply" / "Over 100 applicants")."""
    try:
        text = page.locator("main").first.inner_text(timeout=5000)[:6000]
    except PlaywrightError:
        return None
    m = APPLICANTS_RE.search(text)
    if not m:
        return None
    number = int(re.sub(r"\D", "", m.group(2)) or 0)
    return number + 1 if m.group(1) else number  # "Over 100" counts as 101
 
 
def excluded_company(company):
    names = [c.strip().lower() for c in os.getenv("EXCLUDE_COMPANIES", "").split(";") if c.strip()]
    lower = company.lower()
    return next((n for n in names if n in lower), "")
 
 
def skip_row(link, title, company, reason):
    print(f"  Skipping without AI check: {reason}.")
    gen.add_to_tracker({"Date": datetime.date.today().isoformat(), "Link": link, "Company": company,
                        "Job title": title, "Fit": "no", "Reason": reason})
 
 
def process_job(page, link, cv_text, settings):
    """Opens the job, checks it and makes a CV. True if a CV was made."""
    goto(page, link)
    pause()
    description = read_description(page)
    if not description:
        print("  Could not read the job description, skipping.")
        return False
    parts = [p.strip() for p in page.title().split("|")]
    title = parts[0]
    company = parts[1] if len(parts) > 2 else ""
    bad = excluded_company(company + " " + description[:300])
    if bad:
        skip_row(link, title, company, f"company is on the exclude list ({bad})")
        return False
    limit = int(os.getenv("MAX_APPLICANTS", "0") or 0)
    count = applicants_count(page) if limit else None
    if count is not None and count >= limit:
        skip_row(link, title, company, f"already {count} applicants (limit {limit})")
        return False
    vacancy_text = f"{link}\n{title}\n\n{description}"
    pdf_path, row = gen.generate_for_text(vacancy_text, link, cv_text, settings,
                                          extra={"How to apply": apply_type(page)})
    if not row.get("CV folder"):
        return False  # not a fit: already written to the tracker
    how, company_url = apply_info(page)
    row.update({"How to apply": how, "Company site": company_url})
    print(f"  Apply: {how} {company_url}")
    gen.add_to_tracker(row)
    return bool(pdf_path)
 
 
LANGUAGE_ARGS = {"en": "en", "eng": "en", "english": "en",
                 "fr": "fr", "french": "fr", "any": "any", "all": "any"}
 
 
def language_from_command():
    """python find_jobs.py en / fr / any: job language for this run (otherwise JOB_LANGUAGE from .env)."""
    for arg in sys.argv[1:]:
        lang = LANGUAGE_ARGS.get(arg.lower().lstrip("-"))
        if lang:
            os.environ["JOB_LANGUAGE"] = lang
    lang = LANGUAGE_ARGS.get(os.getenv("JOB_LANGUAGE", "any").strip().lower(), "any")
    os.environ["JOB_LANGUAGE"] = lang
    return lang
 
 
def main():
    language = language_from_command()
    settings = gen.load_settings()
    settings["track"] = False  # we add the tracker row ourselves, together with the company site link
    browser_channel = os.getenv("BROWSER", "chrome").strip().lower()  # chrome or msedge
    max_jobs = int(os.getenv("MAX_JOBS", "5"))       # how many new CVs to make per run
    max_checks = int(os.getenv("MAX_CHECKS", "30"))  # maximum number of jobs to open per run
 
    print("Reading CVs from the cv/ folder ...")
    cv_text = gen.read_cvs()
    seen = already_processed()
    # Jobs skipped earlier because of their language: skip them only if the language still doesn't match
    for link, lang in gen.language_skips().items():
        if language in ("en", "fr") and lang != language:
            seen.add(link)
    names = {"en": "in English only", "fr": "in French only", "any": "in any language"}
    print(f"Looking for jobs {names.get(language, 'in any language')}.")
 
    made = checked = errors_in_row = 0
    with sync_playwright() as p:
        browser = p.chromium.launch_persistent_context(
            str(PROFILE_DIR), channel=browser_channel, headless=False,
            viewport={"width": 1300, "height": 900})
        page = browser.pages[0] if browser.pages else browser.new_page()
        wait_for_login(page)
 
        for name, links in search_rounds(page, seen):
            print(f"\n{name}: {len(links)} new jobs. Checking until {max_jobs} CVs are made "
                  f"(opened {checked} of {max_checks}).")
            for link in links:
                if made >= max_jobs or checked >= max_checks:
                    break
                checked += 1
                seen.add(link)
                print(f"\n[{checked}/{max_checks}] {link}")
                try:
                    if process_job(page, link, cv_text, settings):
                        made += 1
                    errors_in_row = 0
                except PlaywrightError as e:
                    errors_in_row += 1
                    print(f"  Could not open the job ({str(e).splitlines()[0][:100]}). Skipping.")
                    if errors_in_row >= 3:
                        print("\nLinkedIn did not respond 3 times in a row, it probably wants a break. "
                              "Stopping, run the bot again in an hour.")
                        max_checks = checked  # exit all loops
                        break
                    time.sleep(30)
            if made >= max_jobs or checked >= max_checks:
                break
            print(f"\n{name}: no more new jobs.")
 
        if made >= max_jobs:
            print(f"\nMade {made} CVs, that is MAX_JOBS, done for today.")
        elif checked >= max_checks:
            print(f"\nOpened {checked} jobs, that is MAX_CHECKS, done for today.")
        browser.close()
 
    gen.build_tracker_page()
    print(f"\nDone. New CVs: {made}.")
    print(f"All jobs with links: {gen.TRACKER_PAGE}, opening it in the browser.")
    if gen.TRACKER_PAGE.exists():
        webbrowser.open(gen.TRACKER_PAGE.resolve().as_uri())
 
 
if __name__ == "__main__":
    main()
 
