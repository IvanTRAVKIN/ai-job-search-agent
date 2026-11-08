"""
Бот сам ищет вакансии на LinkedIn и делает под них CV. Откликаешься ты сам.

Что делает бот:
  1. Открывает Google Chrome (или Edge) с твоим входом в LinkedIn (отдельный профиль бота).
  2. Сам ищет вакансии по названиям из SEARCH_KEYWORDS в городе SEARCH_LOCATION,
     только свежие (DATE_POSTED: 24h / week / month), и с Easy Apply, и без.
  3. Для каждой новой вакансии читает описание, проверяет, подходит ли она тебе,
     и делает CV (и письмо) в папку output/.
  4. Записывает всё в output/applications.csv: ссылку на вакансию, как откликаться
     (Easy Apply или сайт компании) и ссылку на сайт компании.
Потом ты открываешь таблицу, переходишь по ссылкам и откликаешься сам с готовым CV.

Бот ничего не отправляет. Для вакансий без Easy Apply он только открывает кнопку
«Apply», чтобы узнать ссылку на сайт компании, и сразу закрывает эту вкладку.

Запуск:  python find_jobs.py
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
    print("Не хватает библиотеки playwright. Выполни команду:")
    print("    python -m pip install -r requirements.txt")
    sys.exit(1)

PROFILE_DIR = gen.BASE_DIR / "browser_profile"

# LinkedIn часто меняет вёрстку. Если бот перестал находить вакансии или текст,
# обычно достаточно поправить эти селекторы.
JOB_CARD = "li[data-occludable-job-id]"
JOB_DESCRIPTION = ("#job-details, .jobs-description__content, .jobs-description-content__text, "
                   ".jobs-box__html-content, [class*='jobs-description'], .description__text")

# Запасной способ: найти заголовок «About the job» / «À propos de l'offre» и взять блок под ним.
DESCRIPTION_BY_HEADING_JS = """() => {
  const re = /about the job|à propos de l.offre|à propos du poste|about this job|description du poste/i;
  const h = [...document.querySelectorAll('h1,h2,h3,h4,span,strong')].find(x => re.test(x.innerText || ''));
  if (!h) return '';
  let el = h.parentElement;
  for (let i = 0; i < 6 && el && el.innerText.length < 400; i++) el = el.parentElement;
  return el ? el.innerText : '';
}"""
# Кнопка отклика: LinkedIn часто меняет вёрстку, поэтому ищем и по классу, и по тексту/подписи.
APPLY_BUTTON = ("button.jobs-apply-button, a.jobs-apply-button, "
                "button[aria-label*='Apply' i], a[aria-label*='Apply' i], "
                "button[aria-label*='Postuler' i], a[aria-label*='Postuler' i], "
                "button[aria-label*='Candidature' i], a[aria-label*='Candidature' i]")
APPLY_TEXT = re.compile(r"^\s*(easy apply|apply|postuler|candidature simplifiée|candidater)\b", re.IGNORECASE)
CLOSED_TEXT = ("no longer accepting applications", "n’accepte plus de candidatures",
               "n'accepte plus de candidatures", "candidatures ne sont plus acceptées")

# Фильтр LinkedIn «Дата публикации»: секунды для параметра f_TPR
DATE_FILTERS = {"24h": "r86400", "week": "r604800", "month": "r2592000"}


def search_urls():
    """Ссылки на поиск LinkedIn для каждого названия вакансии, свежие сверху."""
    keywords = [k.strip() for k in os.getenv(
        "SEARCH_KEYWORDS", "Business Analyst;Data Analyst;Pricing Analyst").split(";") if k.strip()]
    location = os.getenv("SEARCH_LOCATION", "Paris, Île-de-France, France").strip()
    date_filter = DATE_FILTERS.get(os.getenv("DATE_POSTED", "week").strip().lower(), "r604800")
    urls = []
    for keyword in keywords:
        # У отдельной должности можно указать свой срок: "Data Analyst|24h"
        keyword, _, own_date = keyword.partition("|")
        keyword = keyword.strip()
        period = DATE_FILTERS.get(own_date.strip().lower(), date_filter)
        # Минус-слова в поиске LinkedIn (по умолчанию нет: Иван ищет обычными словами)
        minus = [w.strip() for w in os.getenv("SEARCH_EXCLUDE", "").split(";") if w.strip()]
        query = keyword + "".join(f" NOT {w}" for w in minus)
        params = {"keywords": query, "location": location, "f_TPR": period, "sortBy": "DD"}
        levels = os.getenv("EXPERIENCE_LEVELS", "2,3").strip()  # 1 стажировка, 2 начальный, 3 младший
        if levels:
            params["f_E"] = levels
        urls.append((keyword, "https://www.linkedin.com/jobs/search/?" + urlencode(params)))
    return urls


# Должности, которые точно не подходят: такие вакансии даже не открываем (видно по названию в списке).
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
    """Название явно французское (H/F, Chargé, Analyste...) — для режима «только английский» не открываем."""
    # По умолчанию выключено: бывает французское название, но описание на английском.
    if os.getenv("SKIP_FRENCH_TITLES", "no").strip().lower() not in gen.YES:
        return False
    return os.getenv("JOB_LANGUAGE", "any") == "en" and bool(FRENCH_TITLE.search(title))


def collect_cards(page):
    """Ссылки и названия вакансий со страницы поиска (листаем список, чтобы LinkedIn показал все карточки)."""
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
    """Первая видимая кнопка или ссылка «Apply / Easy Apply / Postuler» на странице вакансии."""
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
    """Easy Apply, сайт компании или кнопки нет."""
    button = find_apply_button(page)
    if button is None:
        main_text = page.locator("main").first.inner_text().lower() if page.locator("main").count() else ""
        if any(t in main_text for t in CLOSED_TEXT):
            return "вакансия закрыта"
        return "кнопку не нашёл, открой ссылку на LinkedIn"
    label = button_label(button)
    if "easy" in label or "simplifiée" in label:
        return "Easy Apply на LinkedIn"
    return "на сайте компании"


def apply_info(page):
    """Как откликаться на вакансию и ссылка на сайт компании (если не Easy Apply)."""
    how = apply_type(page)
    if how != "на сайте компании":
        return how, ""
    button = find_apply_button(page)
    href = button.get_attribute("href") or ""
    if href.startswith("http") and "linkedin.com" not in href:
        return "на сайте компании", href
    # Кнопка «Apply / Postuler» открывает сайт компании в новой вкладке: берём ссылку и закрываем.
    try:
        with page.context.expect_page(timeout=10000) as new_tab:
            button.click()
        tab = new_tab.value
        tab.wait_for_load_state("domcontentloaded", timeout=15000)
        url = tab.url
        tab.close()
        return "на сайте компании", url
    except PlaywrightTimeout:
        page.keyboard.press("Escape")
        return "на сайте компании (ссылку не удалось получить, открой вакансию)", ""


def pause(min_s=4, max_s=9):
    """Пауза как у человека, чтобы не открывать страницы слишком быстро."""
    time.sleep(random.uniform(min_s, max_s))


def already_processed():
    """Ссылки на вакансии, которые уже есть в таблице откликов."""
    if not gen.TRACKER.exists():
        return set()
    with open(gen.TRACKER, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f, delimiter=";"))
    if "--recheck" in sys.argv:
        # python find_jobs.py --recheck : ещё раз проверить вакансии, которым ИИ дал 25-49 баллов
        def borderline(r):
            score = r.get("Оценка", "")
            return r.get("Подходит") == "нет" and score.isdigit() and 25 <= int(score) < 50
        again = [r for r in rows if borderline(r)]
        rows = [r for r in rows if not borderline(r)]
        with open(gen.TRACKER, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=gen.TRACKER_COLUMNS, delimiter=";", extrasaction="ignore")
            writer.writeheader()
            writer.writerows(rows)
        print(f"Проверю заново {len(again)} вакансий с оценкой 25-49 (если они ещё есть в поиске).")
    return {r.get("Ссылка", "") for r in rows if r.get("Ссылка")}


# Нажимает «… more» / «… plus» под описанием, чтобы раскрыть весь текст вакансии.
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
        pass  # не нашлось кнопки: значит текст и так показан полностью


def read_description(page):
    """Текст вакансии. Пробует несколько способов, потому что LinkedIn часто меняет вёрстку."""
    for _ in range(8):  # до ~16 секунд, пока страница догружается
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
    # Последний вариант: весь текст страницы (Claude сам найдёт в нём описание).
    text = page.evaluate("() => (document.querySelector('main') || document.body).innerText").strip()
    if len(text) > 500:
        return text[:15000]
    debug = gen.OUTPUT_DIR / "debug_page.html"
    gen.OUTPUT_DIR.mkdir(exist_ok=True)
    debug.write_text(page.content(), encoding="utf-8")
    print(f"  (страницу сохранил в {debug.name}, пришли её мне, если так будет со всеми вакансиями)")
    return ""


def goto(page, url):
    """Открывает страницу. LinkedIn иногда сам перенаправляет на другую страницу
    (например, после входа): тогда ждём и пробуем ещё раз."""
    for _ in range(4):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=45000)
            return
        except PlaywrightTimeout:
            print("  LinkedIn долго не отвечает, жду 20 секунд и пробую ещё раз...")
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
        print("\nВойди в LinkedIn в открывшемся окне браузера (это нужно только один раз).")
        input("Когда увидишь свою ленту LinkedIn, вернись сюда и нажми Enter... ")
        time.sleep(3)
        goto(page, "https://www.linkedin.com/feed/")
    print("Вход в LinkedIn есть.")


# Запасной способ найти вакансии на странице, если LinkedIn поменял вид списка
JOB_LINKS_JS = r"""() => {
  const out = [];
  const seen = new Set();
  const add = (id, title) => { if (id && !seen.has(id)) { seen.add(id);
    let t = (title || '').split(/\s*\((?:Verified job|Offre vérifiée)\)/i)[0].trim();
    const half = t.length / 2;  // в новом виде название повторяется дважды подряд
    if (t.length % 2 === 0 && t.slice(0, half) === t.slice(half)) t = t.slice(0, half);
    out.push([id, t]); } };
  document.querySelectorAll('[data-occludable-job-id],[data-job-id]').forEach(el =>
    add(el.getAttribute('data-occludable-job-id') || el.getAttribute('data-job-id'),
        (el.innerText || '').split('\n')[0]));
  document.querySelectorAll('a[href*="/jobs/view/"], a[href*="currentJobId="]').forEach(a => {
    const m = a.href.match(/\/jobs\/view\/(\d+)/) || a.href.match(/currentJobId=(\d+)/);
    if (m) add(m[1], (a.innerText || '').split('\n')[0]);
  });
  // Новый вид LinkedIn: номер вакансии спрятан в атрибутах вроде "urn:li:fsd_jobPosting:123"
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
    """Фразы для поиска из .env: SEARCH_QUERIES (как пишешь в строке поиска LinkedIn), через ;"""
    lang = os.getenv("JOB_LANGUAGE", "any").strip().upper()[:2]
    raw = (os.getenv(f"SEARCH_QUERIES_{lang}", "") or os.getenv("SEARCH_QUERIES", "")).strip()
    return [q.strip() for q in raw.split(";") if q.strip()]


def to_france(query):
    """«... paris ...» -> «... france ...» для второго круга поиска."""
    return re.sub(r"(?i)\bparis\b", "france", query)


def type_search(page, query):
    """Вводит фразу в строку поиска LinkedIn и жмёт Enter — как человек. True, если получилось."""
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
  // Прокручивает все прокручиваемые блоки (список вакансий) и саму страницу вниз
  const boxes = [...document.querySelectorAll('*')].filter(el => {
    const st = getComputedStyle(el);
    return /(auto|scroll)/.test(st.overflowY) && el.scrollHeight > el.clientHeight + 50;
  });
  boxes.forEach(el => el.scrollBy(0, el.clientHeight * 0.9));
  window.scrollBy(0, window.innerHeight * 0.9);
  return boxes.length;
}"""


def cards_on_page(page):
    """Ссылки и названия вакансий на открытой странице поиска. Листает список, чтобы LinkedIn догрузил все."""
    try:
        page.wait_for_function(
            "() => document.querySelector('[data-occludable-job-id],[data-job-id],a[href*=\"/jobs/view/\"],a[href*=\"currentJobId=\"]')",
            timeout=20000)
    except PlaywrightTimeout:
        return []
    page.wait_for_timeout(2000)
    found = {}
    same = 0
    for _ in range(25):  # листаем, пока карточки добавляются
        before = len(found)
        for job_id, title in page.evaluate(JOB_LINKS_JS):
            if title or job_id not in found:
                found[job_id] = title or found.get(job_id, "")
        same = same + 1 if len(found) == before else 0
        if same >= 3:
            break
        page.evaluate(SCROLL_RESULTS_JS)
        page.wait_for_timeout(random.randint(700, 1200))
    if len(found) < 5:  # подозрительно мало: сохраняю страницу, чтобы разобраться
        debug = gen.OUTPUT_DIR / "debug_search.html"
        gen.OUTPUT_DIR.mkdir(exist_ok=True)
        debug.write_text(page.content(), encoding="utf-8")
        print(f"  (мало вакансий на странице — сохранил её в {debug} для проверки)")
    return [(f"https://www.linkedin.com/jobs/view/{job_id}/", title) for job_id, title in found.items()]


def with_start(url, start):
    """Тот же поиск, но со следующей страницы результатов."""
    url = re.sub(r"([?&])start=\d+&?", r"\1", url).rstrip("&?")
    return url + ("&" if "?" in url else "?") + f"start={start}"


def collect(page, label, first_url, seen, taken, typed_query=None):
    """Собирает новые вакансии по одному поиску (до SEARCH_PAGES страниц)."""
    pages = int(os.getenv("SEARCH_PAGES", "3"))
    found, fresh, skipped = 0, [], 0
    if typed_query is not None:
        if not type_search(page, typed_query):
            goto(page, first_url)  # строку поиска не нашёл — открываю поиск по ссылке
        base_url = None
    else:
        goto(page, first_url)
        base_url = first_url
    pause()
    for page_no in range(pages):
        if page_no:
            # Новая страница поиска: внизу кнопки «Page 2», «Page 3» — жмём их, как человек
            button = page.locator(f"button[aria-label='Page {page_no + 1}'], button[aria-label='Page {page_no + 1} ']")
            if button.count():
                button.first.scroll_into_view_if_needed()
                button.first.click()
                page.wait_for_timeout(3000)
            else:
                goto(page, with_start(base_url, 25 * page_no))
            pause()
        if base_url is None:
            base_url = page.url  # адрес, который LinkedIn сам сделал из фразы
        cards = cards_on_page(page)
        if not cards:
            if page_no == 0:
                print(f"Поиск «{label}»: вакансий не нашёл (или LinkedIn поменял страницу). Адрес: {page.url}")
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
            break  # последняя страница или дальше одни уже просмотренные
    print(f"Поиск «{label}»: найдено {found}, новых подходящих по названию {len(fresh)}, "
          f"пропущено по названию {skipped} (senior, lead, manager, стажировки и т.п."
          + (", французские названия" if os.getenv("JOB_LANGUAGE", "any") == "en"
             and os.getenv("SKIP_FRENCH_TITLES", "no").strip().lower() in gen.YES else "") + ")")
    return fresh


def english_keyword_rounds(page, seen):
    """Доп. поиск для режима en: слова из SEARCH_KEYWORDS_EN + английские фразы в описании."""
    en_keywords = [k.strip() for k in os.getenv("SEARCH_KEYWORDS_EN", "").split(";") if k.strip()]
    if os.getenv("JOB_LANGUAGE", "any") == "en" and en_keywords:
        # Английский режим: обычный поиск LinkedIn по словам + английские фразы, которые бывают
        # только в описаниях на английском. Так LinkedIn сам отдаёт в основном английские вакансии.
        marker = os.getenv("ENGLISH_MARKER",
                           '("you will" OR "we are looking for" OR "your role" OR "you have" OR "about you")')
        period = DATE_FILTERS.get(os.getenv("DATE_POSTED", "week").strip().lower(), "r604800")
        locations = [("Париж", os.getenv("SEARCH_LOCATION", "Paris, Île-de-France, France").strip())]
        if os.getenv("FRANCE_FALLBACK", "yes").strip().lower() in gen.YES:
            locations.append(("вся Франция", "France"))
        for name, location in locations:
            print(f"\n=== Поиск (английские вакансии): {name} ===")
            taken = []
            for keyword in en_keywords:
                query = f'"{keyword}" AND {marker}' if marker else keyword
                url = "https://www.linkedin.com/jobs/search/?" + urlencode(
                    {"keywords": query, "location": location, "f_TPR": period, "sortBy": "DD"})
                taken += collect(page, f"{keyword} (EN)", url, seen, taken)
            yield name, taken


def search_rounds(page, seen):
    """Круг 1 — твои фразы (Париж). Круг 2 — те же фразы по всей Франции. Отдаёт новые ссылки по кругам."""
    queries = search_queries()
    if queries:
        rounds = [("Париж", queries)]
        france = [to_france(q) for q in queries if to_france(q) != q]
        if france and os.getenv("FRANCE_FALLBACK", "yes").strip().lower() in gen.YES:
            rounds.append(("вся Франция", france))
        for name, round_queries in rounds:
            print(f"\n=== Поиск: {name} ===")
            taken = []
            for q in round_queries:
                url = "https://www.linkedin.com/jobs/search/?" + urlencode({"keywords": q})
                taken += collect(page, q, url, seen, taken, typed_query=q)
            yield name, taken
    else:  # старый способ: SEARCH_KEYWORDS + фильтры
        taken = []
        for keyword, url in search_urls():
            taken += collect(page, keyword, url, seen, taken)
        yield "поиск", taken
    # В режиме en после твоих фраз — ещё поиск по словам с английскими фразами в описании
    yield from english_keyword_rounds(page, seen)


APPLICANTS_RE = re.compile(
    r"(over|plus de)?\s*(\d[\d\s,.\u202f]*)\s*(people clicked apply|applicants|candidats|personnes ont cliqué)",
    re.IGNORECASE)


def applicants_count(page):
    """Сколько людей уже откликнулось (из строки «29 people clicked apply» / «Over 100 applicants»)."""
    try:
        text = page.locator("main").first.inner_text(timeout=5000)[:6000]
    except PlaywrightError:
        return None
    m = APPLICANTS_RE.search(text)
    if not m:
        return None
    number = int(re.sub(r"\D", "", m.group(2)) or 0)
    return number + 1 if m.group(1) else number  # «Over 100» считаем как 101


def excluded_company(company):
    names = [c.strip().lower() for c in os.getenv("EXCLUDE_COMPANIES", "").split(";") if c.strip()]
    lower = company.lower()
    return next((n for n in names if n in lower), "")


def skip_row(link, title, company, reason):
    print(f"  Пропускаю без проверки ИИ: {reason}.")
    gen.add_to_tracker({"Дата": datetime.date.today().isoformat(), "Ссылка": link, "Компания": company,
                        "Должность": title, "Подходит": "нет", "Почему": reason})


def process_job(page, link, cv_text, settings):
    """Открывает вакансию, проверяет и делает CV. True, если CV сделано."""
    goto(page, link)
    pause()
    description = read_description(page)
    if not description:
        print("  Не смог прочитать описание вакансии, пропускаю.")
        return False
    parts = [p.strip() for p in page.title().split("|")]
    title = parts[0]
    company = parts[1] if len(parts) > 2 else ""
    bad = excluded_company(company + " " + description[:300])
    if bad:
        skip_row(link, title, company, f"компания в чёрном списке ({bad})")
        return False
    limit = int(os.getenv("MAX_APPLICANTS", "0") or 0)
    count = applicants_count(page) if limit else None
    if count is not None and count >= limit:
        skip_row(link, title, company, f"уже {count} откликов (лимит {limit})")
        return False
    vacancy_text = f"{link}\n{title}\n\n{description}"
    pdf_path, row = gen.generate_for_text(vacancy_text, link, cv_text, settings,
                                          extra={"Как откликаться": apply_type(page)})
    if not row.get("Папка с CV"):
        return False  # вакансия не подошла: уже записана в таблицу
    how, company_url = apply_info(page)
    row.update({"Как откликаться": how, "Сайт компании": company_url})
    print(f"  Откликаться: {how} {company_url}")
    gen.add_to_tracker(row)
    return bool(pdf_path)


LANGUAGE_ARGS = {"en": "en", "eng": "en", "english": "en", "англ": "en",
                 "fr": "fr", "french": "fr", "франц": "fr", "any": "any", "all": "any", "все": "any"}


def language_from_command():
    """python find_jobs.py en / fr / any — язык вакансий на этот запуск (иначе берётся JOB_LANGUAGE из .env)."""
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
    settings["track"] = False  # строку в таблицу добавим сами, вместе со ссылкой на сайт компании
    browser_channel = os.getenv("BROWSER", "chrome").strip().lower()  # chrome или msedge
    max_jobs = int(os.getenv("MAX_JOBS", "5"))       # сколько новых CV сделать за запуск
    max_checks = int(os.getenv("MAX_CHECKS", "30"))  # сколько вакансий максимум открыть за запуск

    print("Читаю резюме из папки cv/ ...")
    cv_text = gen.read_cvs()
    seen = already_processed()
    # Вакансии, пропущенные раньше из-за языка: не открываем, только если язык до сих пор не тот
    for link, lang in gen.language_skips().items():
        if language in ("en", "fr") and lang != language:
            seen.add(link)
    names = {"en": "только на английском", "fr": "только на французском", "any": "на любом языке"}
    print(f"Ищу вакансии {names.get(language, 'на любом языке')}.")

    made = checked = errors_in_row = 0
    with sync_playwright() as p:
        browser = p.chromium.launch_persistent_context(
            str(PROFILE_DIR), channel=browser_channel, headless=False,
            viewport={"width": 1300, "height": 900})
        page = browser.pages[0] if browser.pages else browser.new_page()
        wait_for_login(page)

        for name, links in search_rounds(page, seen):
            print(f"\n{name}: новых вакансий {len(links)}. Смотрю, пока не сделаю {max_jobs} CV "
                  f"(открыто {checked} из {max_checks}).")
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
                    print(f"  Не получилось открыть вакансию ({str(e).splitlines()[0][:100]}). Пропускаю.")
                    if errors_in_row >= 3:
                        print("\nLinkedIn 3 раза подряд не отвечает — похоже, он просит паузу. "
                              "Останавливаюсь, запусти бота снова через час.")
                        max_checks = checked  # выходим из всех циклов
                        break
                    time.sleep(30)
            if made >= max_jobs or checked >= max_checks:
                break
            print(f"\n{name}: новые вакансии закончились.")

        if made >= max_jobs:
            print(f"\nСделал {made} CV — это MAX_JOBS, на сегодня хватит.")
        elif checked >= max_checks:
            print(f"\nОткрыл {checked} вакансий — это MAX_CHECKS, на сегодня хватит.")
        browser.close()

    gen.build_tracker_page()
    print(f"\nГотово. Новых CV: {made}.")
    print(f"Все вакансии со ссылками: {gen.TRACKER_PAGE} — открываю в браузере.")
    if gen.TRACKER_PAGE.exists():
        webbrowser.open(gen.TRACKER_PAGE.resolve().as_uri())


if __name__ == "__main__":
    main()
