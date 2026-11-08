"""
Генератор резюме под конкретную вакансию.

Как пользоваться:
  1. Положи свои резюме (PDF или TXT) в папку cv/
     (и, по желанию, фото cv/photo.jpg).
  2. Одна вакансия: вставь её текст в файл vacancy.txt
     Много вакансий: положи каждую в отдельный .txt в папку vacancies/
     (первой строкой можно вставить ссылку на вакансию).
  3. Запусти:  python generate_cv.py
  4. Готовые резюме (и письма) появятся в папке output/,
     а список всех откликов в output/applications.csv (открывается в Excel).

Скрипт НИКОГДА не придумывает опыт: он только выбирает, переставляет
и переформулирует то, что уже есть в твоих резюме.
"""

import base64
import csv
import datetime
import html
import json
import os
import re
import shutil
import subprocess
import sys
import urllib.request
from pathlib import Path

try:
    import anthropic
    from dotenv import load_dotenv
    from pypdf import PdfReader
except ImportError:
    print("Не хватает библиотек. Выполни команду:")
    print("    python -m pip install -r requirements.txt")
    sys.exit(1)

BASE_DIR = Path(__file__).resolve().parent
CV_DIR = BASE_DIR / "cv"
OUTPUT_DIR = BASE_DIR / "output"
FONTS_DIR = BASE_DIR / "fonts"
DEFAULT_VACANCY = BASE_DIR / "vacancy.txt"
VACANCIES_DIR = BASE_DIR / "vacancies"
TRACKER = OUTPUT_DIR / "applications.csv"

# Цена за 1 млн токенов (вход, выход) в долларах, чтобы показать стоимость запуска.
PRICES = {
    "claude-opus-5-5": (4.00, 20.00),
    "claude-sonnet-5-5": (2.00, 10.00),
    "claude-haiku-5-5": (0.10, 0.50),
}


# ---------- 1. Чтение исходных резюме и вакансии ----------

def read_cvs():
    if not CV_DIR.exists():
        sys.exit(f"Нет папки {CV_DIR}. Создай её и положи туда свои резюме.")
    texts = []
    for path in sorted(CV_DIR.iterdir()):
        suffix = path.suffix.lower()
        if suffix == ".pdf":
            reader = PdfReader(str(path))
            text = "\n".join(page.extract_text() or "" for page in reader.pages)
        elif suffix in (".txt", ".md"):
            text = path.read_text(encoding="utf-8", errors="ignore")
        else:
            continue
        if text.strip():
            texts.append(f"===== Файл: {path.name} =====\n{text.strip()}")
            print(f"  прочитано резюме: {path.name} ({len(text)} символов)")
    if not texts:
        sys.exit("В папке cv/ нет резюме в формате PDF или TXT.")
    return "\n\n".join(texts)


def read_vacancy(path):
    if not path.exists():
        sys.exit(f"Нет файла {path.name}. Создай его и вставь туда текст вакансии.")
    text = path.read_text(encoding="utf-8", errors="ignore").strip()
    if len(text) < 50:
        sys.exit(f"Файл {path.name} почти пустой. Вставь туда полный текст вакансии.")
    return text


# ---------- 2. Запрос к Claude ----------

# Проект, который ВСЕГДА попадает в каждое CV (раздел Projects после Experience).
# Можно править текст здесь: бот берёт отсюда факты и подстраивает формулировки под вакансию.
MANDATORY_PROJECT = """Name: AI Agent for Job Search & CV Tailoring (personal project)
Dates: 2026
Stack: Python, Claude API (Anthropic), AI agent, prompt engineering, Playwright, VS Code
Facts:
- Built an AI agent in Python on the Claude API (Anthropic): set up the Python environment in VS Code, created and connected the Claude API key securely through a .env file.
- Automated my own job search and saved hours of manual work: the agent searches LinkedIn, reads each job posting, checks if it fits my profile and writes a tailored one-page CV and cover letter in PDF.
- Prompt engineering with strict rules so the AI uses only real facts from my profile, matches the posting's keywords (ATS-friendly) and picks the CV language (French/English) automatically.
- Browser automation with Playwright and an application tracker (CSV/HTML) with links and apply method for every posting."""


SYSTEM_PROMPT = """You are an expert CV writer. You tailor a candidate's CV to one specific job posting so that a recruiter immediately sees a strong match.

Source material: the candidate's CV (PDF) and a long "master profile" (about_me.txt) with many more experiences, projects, tasks, tools and achievements than fit on one page. The source may be written in Russian, English or French; translate as needed. Treat all of it as true facts about the candidate.
If about_me.txt is present, it is the PRIMARY source: build the experience bullets, skills and order from it, and use the PDF CV only for contacts, education, languages, soft skills and hobbies (and for any fact missing from about_me.txt). Do not simply copy the PDF CV bullets.

Hard rules (truthfulness):
- Use ONLY facts from the source material. Never invent employers, job titles at employers, dates, degrees, tools, numbers, achievements or skills.
- Keep every number exactly as in the source.
- Job titles inside the experience section (role at each employer) must stay as in the source (translation is allowed).
- If the vacancy asks for something the candidate does not have, do NOT add it to the CV. List it in "missing_requirements" instead.

Tailoring (be bold here, this is the whole point):
- "headline" (the title under the name) MUST be adapted to this vacancy: start with the job title of the posting, in a clean short form (e.g. "Business Analyst", "Pricing Analyst", "Data Analyst", "Chef de projet"), then add 2-3 short keywords from the posting that the candidate genuinely has (they MUST also be in "matched_keywords" and must NOT be in "missing_requirements"; e.g. never write "Spécifications fonctionnelles" unless the source explicitly shows the candidate wrote functional specifications), separated by " | ". Example: "Business Analyst | Recueil des besoins | Spécifications fonctionnelles". Do NOT reuse the original headline unless it already matches the posting. If the source lists job titles that fit the candidate, the headline title must be the one from that list closest to the posting's title (never a title the candidate marked as not a fit).
- Follow any "note for the bot" written in the source material (e.g. always include a given experience).
- From the whole source material, pick the experiences, projects and bullets most relevant to the posting; drop or shorten the irrelevant ones. Experiences you keep must stay in exactly the same order as in about_me.txt (otherwise reverse-chronological: most recent first, e.g. a job started in April 2025 goes above one that ended in December 2024). Never reorder experiences for relevance. Look through EVERY bullet of about_me.txt for matches with the posting (e.g. Agile/sprints, Notion, SQL, stakeholder meetings, recruiting/onboarding).
- Keywords: first list the key hard skills, tools, methods and tasks the posting asks for (e.g. "recueil des besoins", "spécifications fonctionnelles", "Agile", "SQL"). For EVERY keyword the candidate genuinely has according to the source material (CVs or about_me.txt), put that exact keyword (in the posting's wording) into the CV: in the skills section if it is a skill/tool/method, and in an experience bullet if the source shows a task where the candidate did it. Return them in "matched_keywords". Keywords the candidate does not have go to "missing_requirements" and never into the CV.
- Rewrite bullets using the posting's own vocabulary (as long as the meaning stays true).
- Bullet ORDER inside each experience must follow the order of the source material (about_me.txt first, otherwise the CV), because it tells a logical story (e.g. diagnostic -> presentation -> recommendations). Choose which bullets to keep or drop for relevance, but never reorder the ones you keep.
- Order and rename skill groups so that the skills the posting asks for come first, using the posting's wording.
- {language_rule}
- The CV MUST fit on ONE A4 page (it is a junior CV). Use the space well: the right column should be filled almost to the bottom of the page, without empty space. The right column also holds the mandatory project, so aim for 12-13 experience bullets in total (plus 2-3 project bullets): 4-5 for the two main experiences, 2 for the others; each bullet at most ~200 characters; education without extra lines. (If the result is slightly too long, the script shrinks the font automatically.)
- The LEFT sidebar (skills, soft skills, hobbies, languages) is narrow and is usually what pushes the CV to a 2nd page, so keep it short: at most 5 skill groups, each with a short title and at most ~70 characters of items (only the skills most relevant to this posting, not everything the candidate knows); at most 3 soft skills of max ~40 characters each; at most 2 hobbies of max ~40 characters each; languages: the same format for every language, "C1 – <proficiency in CV language>" (e.g. "C1 – Professional proficiency" / "C1 – Maîtrise professionnelle") and the native one as "Native" / "Langue maternelle".
- Contacts (email, phone, city, LinkedIn) must be copied exactly from the source CV.
- "notes_for_ivan" and "missing_requirements" are for the candidate and must be written in Russian. In "notes_for_ivan", in at most 3 short sentences, say which headline you chose and what you emphasized. Keep each "missing_requirements" item short (max ~10 words).
- "projects": ALWAYS return exactly one project, the mandatory project below, even if the posting is not technical. It is shown in its own section after the experience, never inside "experience". Name and dates as given; "stack" = short line, max ~70 characters, that MUST contain "Python" and "Claude API". 2-3 bullets (max ~180 characters each), rephrased from its facts. The bullets MUST always say that he built an AI agent in Python with the Claude API (and connected the API key), and that he automated his own work and saved time with this AI agent; a third bullet may add the detail most relevant to the posting. Never mention costs or prices. Same truthfulness rules: only these facts. Write it in the CV language. Mandatory project:
{project}
- {cover_rule}"""

STR = {"type": "string"}


def obj(props):
    return {
        "type": "object",
        "properties": props,
        "required": list(props),
        "additionalProperties": False,
    }


def arr(items):
    return {"type": "array", "items": items}


CV_SCHEMA = obj({
    "company_name": STR,
    "job_title": STR,
    "labels": obj({
        "experience": STR, "projects": STR, "education": STR, "skills": STR,
        "soft_skills": STR, "hobbies": STR, "languages": STR,
    }),
    "name": STR,
    "headline": STR,
    "contacts": obj({"email": STR, "phone": STR, "city": STR, "linkedin": STR}),
    "skill_groups": arr(obj({"title": STR, "items": STR})),
    "soft_skills": arr(STR),
    "hobbies": arr(STR),
    "languages": arr(obj({"language": STR, "level": STR})),
    "experience": arr(obj({
        "company": STR, "location": STR, "role": STR,
        "dates": STR, "bullets": arr(STR),
    })),
    "projects": arr(obj({
        "name": STR, "stack": STR, "dates": STR, "bullets": arr(STR),
    })),
    "education": arr(obj({
        "school": STR, "degree_type": STR, "program": STR, "dates": STR,
    })),
    "matched_keywords": arr(STR),
    "cover_letter": STR,
    "notes_for_ivan": STR,
    "missing_requirements": arr(STR),
})


def language_rule(cv_language):
    if cv_language.lower() == "auto":
        return ("Write the CV in the language of the job posting: French if the posting is in French, "
                "otherwise English.")
    return f"Write the CV in {cv_language}, whatever the language of the job posting."


def cover_rule(enabled):
    if not enabled:
        return 'Return "cover_letter" as an empty string.'
    words = os.getenv("COVER_WORDS", "350-450").strip()
    return ('"cover_letter": a full cover letter / lettre de motivation of ' + words + ' words, in the same language as the CV, '
            'ready to paste into LinkedIn or an email. Structure (4-6 paragraphs, separated by blank lines): '
            '1) greeting ("Madame, Monsieur," in French, "Dear Hiring Team," in English); '
            '2) who he is and which role he applies for, and why THIS company and role interest him, using concrete details from the job posting; '
            '3-4) one paragraph per most relevant experience from the source (company name, what he actually did, tools, results with numbers), '
            'each clearly linked to a requirement of the job; '
            '5) what he brings to the team and that he is motivated to learn what he does not know yet (you may name ONE item from missing_requirements as something he wants to learn, never claim it); '
            '6) polite closing, availability for an interview, and the candidate name. '
            'Language: simple and natural, like a real young professional writes - short clear sentences, common everyday words, '
            'no fancy or rare words, no clichés or buzzwords (avoid "passionné", "dynamique", "synergy", "leverage", "spearheaded" and similar). '
            'Same truthfulness rules as the CV: use only facts from the source, never invent. No address header, no date, no placeholders.')


def call_claude(model, effort, system, user_message, schema):
    """Один запрос к Claude, ответ строго в формате schema. Возвращает (данные, стоимость $)."""
    client = anthropic.Anthropic()  # ключ берётся из ANTHROPIC_API_KEY в файле .env
    params = dict(
        model=model,
        max_tokens=16000,
        system=system,
        messages=[{"role": "user", "content": user_message}],
        output_config={
            "effort": effort,
            "format": {"type": "json_schema", "schema": schema},
        },
    )
    if "haiku" not in model:
        # Если модель вдруг откажется отвечать, API само попробует другую модель.
        params["betas"] = ["server-side-fallback-2026-07-01"]
        params["fallbacks"] = "default"
    try:
        response = client.beta.messages.create(**params)
    except anthropic.AuthenticationError:
        sys.exit("Неверный API-ключ. Проверь строку ANTHROPIC_API_KEY в файле .env")
    except anthropic.PermissionDeniedError:
        sys.exit("У ключа нет доступа. Проверь, что на балансе в console.anthropic.com есть деньги.")
    except anthropic.NotFoundError:
        sys.exit(f"Модель {model} не найдена. Проверь строку MODEL в файле .env")
    except anthropic.RateLimitError:
        sys.exit("Слишком много запросов подряд. Подожди минуту и попробуй снова.")
    except anthropic.BadRequestError as e:
        if "credit" in str(e.message).lower():
            sys.exit("На балансе закончились деньги. Пополни его в console.anthropic.com")
        sys.exit(f"Ошибка запроса: {e.message}")
    except anthropic.APIStatusError as e:
        sys.exit(f"Ошибка на стороне API ({e.status_code}). Попробуй позже.")
    except anthropic.APIConnectionError:
        sys.exit("Нет соединения с интернетом или с API. Проверь интернет.")

    if response.stop_reason == "refusal":
        sys.exit("Claude отказался отвечать на этот запрос. Проверь текст вакансии.")
    if response.stop_reason == "max_tokens":
        sys.exit("Ответ получился слишком длинным и оборвался. Попробуй ещё раз.")

    text = "".join(b.text for b in response.content if b.type == "text")
    usage = response.usage
    price_in, price_out = PRICES.get(model, (0, 0))
    cost = usage.input_tokens / 1e6 * price_in + usage.output_tokens / 1e6 * price_out
    return json.loads(text), cost


def ask_claude(cv_text, vacancy_text, model, cv_language, effort, cover_letter):
    user_message = (
        f"<source_cvs>\n{cv_text}\n</source_cvs>\n\n"
        f"<job_posting>\n{vacancy_text}\n</job_posting>\n\n"
        "Create the tailored CV for this job posting."
    )
    system = SYSTEM_PROMPT.format(language_rule=language_rule(cv_language),
                                  cover_rule=cover_rule(cover_letter),
                                  project=MANDATORY_PROJECT)
    print(f"  Делаю резюме ({model})... обычно 20-60 секунд.")
    data, cost = call_claude(model, effort, system, user_message, CV_SCHEMA)
    print(f"  Готово. Примерная стоимость: ${cost:.3f}")
    return data, cost


# ---------- 2б. Быстрая проверка: подходит ли вакансия ----------

FIT_PROMPT = """You screen job postings for a candidate before he applies.
The candidate's profile (with the job titles that fit him, titles that are NOT a fit, and seniority rules) is in <profile>.
Decide if he should apply to the posting in <job_posting>, strictly following his own rules
(e.g. skip roles marked "not a fit", skip senior/lead/manager roles, skip roles asking for more than 5 years,
3-5 years only if the job closely matches his experience).
Internships (stage, stagiaire), alternance / apprenticeship and VIE are NOT a fit unless his profile explicitly says he wants them: set "apply" to false for them.
"score" is 0-100: how likely it is worth applying. Score like a friendly recruiter for a JUNIOR candidate, not like a strict checklist:
- Start from the job title and level. If the title is in his "fit" list (Business Analyst, Data Analyst, Pricing Analyst, etc.)
  and the level is junior / confirmed / 0-5 years (and no hard blocker below), the score is at least 55, usually 60-80.
- Missing industry knowledge (banking, monétique, capital markets, insurance, supply chain, retail...) or specific tools
  (Summit, Murex, SAS, MicroStrategy, Snowflake, Jira...) are NORMAL gaps for a junior: lower the score by at most 5-10 points in total, never reject for them alone.
- Check his education before saying it is missing: a Master / Bac+5 / MSc in his profile satisfies "Bac+5".
- Hard blockers (score below 40, apply false): senior / lead / manager / expert level, more than 5 years required,
  a title from his "not a fit" list (Data Engineer, Data Scientist, ML, developer, quant), internship / alternance, a job that is not analysis at all
  (sales, technician, C-level), a required nationality he does not have.
"reasons" MUST be written in Russian (even if the job posting is in French or English), 1-2 short sentences."""

FIT_SCHEMA = obj({
    "company_name": STR,
    "job_title": STR,
    "apply": {"type": "boolean"},
    "score": {"type": "integer"},
    "reasons": STR,
})


def check_fit(cv_text, vacancy_text, model):
    user_message = (f"<profile>\n{cv_text}\n</profile>\n\n"
                    f"<job_posting>\n{vacancy_text}\n</job_posting>")
    data, cost = call_claude(model, "low", FIT_PROMPT, user_message, FIT_SCHEMA)
    return data, cost


# ---------- 3. Сборка резюме (HTML -> PDF) ----------

CSS = """
@page { size: A4; margin: 0; }
* { box-sizing: border-box; }
body { margin: 0; font-family: 'Open Sans', Arial, sans-serif;
       font-size: calc(8.6pt * var(--s)); line-height: 1.38; color: #222; }
.sidebar-bg { position: fixed; top: 0; left: 0; bottom: 0; width: 36%; background: #e7eded; }
.page { position: relative; display: flex; min-height: 297mm; }
.side { width: 36%; }
.side-top { background: #87a6a3; color: #fff; padding: 9mm 7mm 6mm; }
.photo { width: 34mm; height: calc(44mm * var(--s)); margin-bottom: 4mm;
         background-repeat: no-repeat;
         background-size: var(--zoom);
         background-position: 50% var(--top); }
.name { font-size: 18pt; font-weight: 600; line-height: 1.1; margin: 0 0 1.5mm; }
.headline { font-size: 9pt; opacity: .95; margin-bottom: 3mm; }
.contact { font-size: calc(8.3pt * var(--s)); margin: .6mm 0; word-break: break-all; }
.contact a { color: #fff; }
.side-body { padding: 4mm 7mm 8mm; }
.side h2 { color: #87a6a3; font-size: 11pt; font-weight: 600; margin: calc(4mm * var(--s)) 0 1.5mm;
           padding-bottom: 1mm; border-bottom: 1px solid #c9d6d5; }
.group-title { font-weight: 700; margin-top: calc(1.8mm * var(--s)); }
.soft { margin: 1.5mm 0; }
.lang-name { font-weight: 700; margin-top: 1.5mm; }
.main { width: 64%; padding: 8mm 9mm 8mm 8mm; }
.main h2 { color: #87a6a3; font-size: 11pt; font-weight: 600; letter-spacing: .5px;
           margin: 0 0 2mm; padding-bottom: 1mm; border-bottom: 1px solid #c9d6d5;
           text-transform: uppercase; }
.job { margin-bottom: calc(3.5mm * var(--s)); break-inside: avoid; }
.row { display: flex; justify-content: space-between; gap: 4mm; }
.company { font-weight: 700; }
ul { margin: 1mm 0 0; padding-left: 4mm; }
li { margin-bottom: calc(1mm * var(--s)); }
.edu { margin-bottom: 2.5mm; }
.edu-section { margin-top: 5mm; }
"""


def esc(value):
    return html.escape(str(value or ""))


# Статичные файлы шрифта Open Sans (как в твоём исходном CV).
# Скачиваются один раз в папку fonts/ и встраиваются прямо в резюме,
# поэтому в PDF нет странных расстояний между буквами.
FONT_URLS = {
    400: "https://fonts.gstatic.com/s/opensans/v44/memSYaGs126MiZpBA-UvWbX2vVnXBbObj2OVZyOOSr4dVJWUgsjZ0B4uaVc.ttf",
    600: "https://fonts.gstatic.com/s/opensans/v44/memSYaGs126MiZpBA-UvWbX2vVnXBbObj2OVZyOOSr4dVJWUgsgH1x4uaVc.ttf",
    700: "https://fonts.gstatic.com/s/opensans/v44/memSYaGs126MiZpBA-UvWbX2vVnXBbObj2OVZyOOSr4dVJWUgsg-1x4uaVc.ttf",
}


def font_css():
    FONTS_DIR.mkdir(exist_ok=True)
    rules = []
    for weight, url in FONT_URLS.items():
        path = FONTS_DIR / f"OpenSans-{weight}.ttf"
        if not path.exists():
            try:
                urllib.request.urlretrieve(url, path)
            except OSError:
                print("  не удалось скачать шрифт Open Sans, будет использован Arial")
                return ""
        data = base64.b64encode(path.read_bytes()).decode()
        rules.append(
            "@font-face { font-family: 'Open Sans'; font-style: normal; "
            f"font-weight: {weight}; src: url(data:font/ttf;base64,{data}) format('truetype'); }}")
    return "\n".join(rules)


def photo_tag():
    for name in ("photo.jpg", "photo.jpeg", "photo.png"):
        path = CV_DIR / name
        if path.exists():
            mime = "image/png" if path.suffix.lower() == ".png" else "image/jpeg"
            data = base64.b64encode(path.read_bytes()).decode()
            zoom = os.getenv("PHOTO_ZOOM", "165").strip()
            top = os.getenv("PHOTO_TOP", "18").strip()
            return (f'<div class="photo" style="--zoom:{zoom}%; --top:{top}%; '
                    f'background-image:url(data:{mime};base64,{data})"></div>')
    return ""


def build_html(cv, scale=1.0, move_side=False):
    lb = cv["labels"]
    c = cv["contacts"]
    linkedin = c["linkedin"]
    linkedin_url = linkedin if linkedin.startswith("http") else "https://" + linkedin

    contact_lines = []
    if c["email"]:
        contact_lines.append(f'✉ <a href="mailto:{esc(c["email"])}">{esc(c["email"])}</a>')
    if c["phone"]:
        contact_lines.append(f'☎ {esc(c["phone"])}')
    if c["city"]:
        contact_lines.append(f'⌂ {esc(c["city"])}')
    if linkedin:
        contact_lines.append(f'in <a href="{esc(linkedin_url)}">{esc(linkedin)}</a>')
    contacts = "".join(f'<div class="contact">{line}</div>' for line in contact_lines)

    skills = "".join(
        f'<div class="group-title">{esc(g["title"])}</div><div>{esc(g["items"])}</div>'
        for g in cv["skill_groups"])
    soft = "".join(f'<div class="soft">{esc(s)}</div>' for s in cv["soft_skills"])
    hobbies = "".join(f'<div class="group-title">{esc(h)}</div>' for h in cv["hobbies"])
    langs = "".join(
        f'<div class="lang-name">{esc(l["language"])}</div><div>{esc(l["level"])}</div>'
        for l in cv["languages"])

    jobs = ""
    for job in cv["experience"]:
        bullets = "".join(f"<li>{esc(b)}</li>" for b in job["bullets"])
        jobs += f"""
        <div class="job">
          <div class="row"><span class="company">{esc(job["company"])}</span><span>{esc(job["location"])}</span></div>
          <div class="row"><span>{esc(job["role"])}</span><span>{esc(job["dates"])}</span></div>
          <ul>{bullets}</ul>
        </div>"""

    projects = ""
    for pr in cv.get("projects", []):
        bullets = "".join(f"<li>{esc(b)}</li>" for b in pr["bullets"])
        projects += f"""
        <div class="job">
          <div class="row"><span class="company">{esc(pr["name"])}</span><span>{esc(pr["dates"])}</span></div>
          <div class="row"><span>{esc(pr["stack"])}</span></div>
          <ul>{bullets}</ul>
        </div>"""
    projects_block = (f'<div class="edu-section"><h2>{esc(cv["labels"].get("projects") or "Projects")}</h2>'
                      f'{projects}</div>') if projects else ""

    edu = "".join(f"""
        <div class="edu">
          <div class="row"><span class="company">{esc(e["school"])}</span><span>{esc(e["degree_type"])}</span></div>
          <div class="row"><span>{esc(e["program"])}</span><span>{esc(e["dates"])}</span></div>
        </div>""" for e in cv["education"])

    def section(title, body):
        return f"<h2>{esc(title)}</h2>{body}" if body else ""

    # Если левая колонка слишком длинная, переносим Soft skills, Hobbies и Languages
    # вниз правой колонки, где обычно остаётся свободное место.
    extra = [(lb["soft_skills"], soft), (lb["hobbies"], hobbies), (lb["languages"], langs)]
    side_extra = "" if move_side else "".join(section(t, b) for t, b in extra)
    main_extra = "".join(
        f'<div class="edu-section">{section(t, b)}</div>' for t, b in extra) if move_side else ""

    return f"""<!doctype html>
<html><head><meta charset="utf-8">
<title>{esc(cv["name"])} - CV</title>
<style>{font_css()}\n:root {{ --s: {scale}; }}\n{CSS}</style></head>
<body>
<div class="sidebar-bg"></div>
<div class="page">
  <div class="side">
    <div class="side-top">
      {photo_tag()}
      <div class="name">{esc(cv["name"])}</div>
      <div class="headline">{esc(cv["headline"])}</div>
      {contacts}
    </div>
    <div class="side-body">
      {section(lb["skills"], skills)}
      {side_extra}
    </div>
  </div>
  <div class="main">
    <h2>{esc(lb["experience"])}</h2>
    {jobs}
    {projects_block}
    <div class="edu-section"><h2>{esc(lb["education"])}</h2>{edu}</div>
    {main_extra}
  </div>
</div>
</body></html>"""


def find_browser():
    candidates = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    for name in ("msedge", "google-chrome", "chromium", "chromium-browser", "chrome"):
        found = shutil.which(name)
        if found:
            return found
    return None


def html_to_pdf(html_path, pdf_path):
    browser = find_browser()
    if not browser:
        return False
    subprocess.run([
        browser, "--headless", "--disable-gpu", "--no-pdf-header-footer",
        "--print-to-pdf-no-header", f"--print-to-pdf={pdf_path}",
        html_path.as_uri(),
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=120)
    return pdf_path.exists()


def page_count(pdf_path):
    try:
        return len(PdfReader(str(pdf_path)).pages)
    except Exception:
        return 1


def render_one_page(cv, html_path, pdf_path):
    """Делает PDF; если он вышел на 2 страницы, переносит часть левой колонки вправо
    и/или чуть уменьшает шрифт, пока резюме не влезет на 1 страницу."""
    # Сначала уменьшаем шрифт, не трогая левую колонку (правая колонка с проектом длиннее,
    # и перенос туда Soft skills/Hobbies/Languages только мешает), потом пробуем перенос.
    attempts = [(1.0, False), (0.95, False), (0.9, False), (1.0, True), (0.95, True),
                (0.9, True), (0.86, False), (0.86, True), (0.82, False), (0.82, True)]
    for scale, move_side in attempts:
        html_path.write_text(build_html(cv, scale, move_side), encoding="utf-8")
        if pdf_path.exists():
            pdf_path.unlink()
        if not html_to_pdf(html_path, pdf_path):
            return False
        pages = page_count(pdf_path)
        if pages == 1:
            return True
    # Всё ещё 2 страницы: оставляем в проекте 2 пункта и пробуем ещё раз.
    for pr in cv.get("projects", []):
        pr["bullets"] = pr["bullets"][:2]
    for scale, move_side in [(0.9, False), (0.86, False), (0.86, True), (0.82, False), (0.82, True)]:
        html_path.write_text(build_html(cv, scale, move_side), encoding="utf-8")
        if pdf_path.exists():
            pdf_path.unlink()
        if not html_to_pdf(html_path, pdf_path):
            return False
        pages = page_count(pdf_path)
        if pages == 1:
            return True
    print(f"\nВнимание: резюме не влезло на 1 страницу даже с мелким шрифтом ({pages} стр.). "
          "Запусти ещё раз: бот каждый раз пишет немного по-другому.")
    return True


def safe_name(text):
    text = re.sub(r"[^\w\- ]", "", text, flags=re.UNICODE).strip().replace(" ", "_")
    return text[:40] or "vacancy"


# ---------- 4. Таблица откликов ----------

TRACKER_COLUMNS = ["Дата", "Компания", "Должность", "Подходит", "Оценка", "Почему",
                   "Ссылка", "Как откликаться", "Сайт компании", "Папка с CV", "Чего нет",
                   "Статус отклика"]


def add_to_tracker(row):
    OUTPUT_DIR.mkdir(exist_ok=True)
    if TRACKER.exists():
        with open(TRACKER, encoding="utf-8-sig") as f:
            header = f.readline().strip().split(";")
        if header != TRACKER_COLUMNS:  # таблица от старой версии скрипта: откладываем её
            TRACKER.replace(TRACKER.with_name("applications_old.csv"))
    new_file = not TRACKER.exists()
    # utf-8-sig и ";" - чтобы файл нормально открывался в Excel с русским и французским текстом
    with open(TRACKER, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        if new_file:
            writer.writerow(TRACKER_COLUMNS)
        writer.writerow([row.get(c, "") for c in TRACKER_COLUMNS])
    build_tracker_page()


TRACKER_PAGE = OUTPUT_DIR / "vacancies.html"


def build_tracker_page():
    """Делает output/vacancies.html: все вакансии с кликабельными ссылками (открывается двойным щелчком)."""
    if not TRACKER.exists():
        return
    with open(TRACKER, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f, delimiter=";"))
    rows.reverse()  # новые сверху

    def link(url, text):
        return f'<a href="{html.escape(url)}" target="_blank">{text}</a>' if url.startswith("http") else ""

    def cv_link(folder):
        if not folder:
            return ""
        folder = Path(folder)
        pdfs = sorted((OUTPUT_DIR / folder.name).glob("*.pdf"))
        if not pdfs:
            return ""
        return (f'<a href="{html.escape(folder.name)}/{html.escape(pdfs[0].name)}" target="_blank">CV</a> · '
                f'<a href="{html.escape(folder.name)}/cover_letter.txt" target="_blank">письмо</a>')

    def table(items, with_cv):
        head = (["Отправил"] if with_cv else []) + ["Дата", "Компания", "Должность", "Оценка", "LinkedIn",
                                                     "Как откликаться", "Сайт компании"]
        head += ["CV", "Чего нет"] if with_cv else ["Почему не подходит"]
        out = [f'<table{" id=good" if with_cv else ""}><tr>' + "".join(f"<th>{h}</th>" for h in head) + "</tr>"]
        for r in items:
            key = html.escape(r["Ссылка"] or r["Папка с CV"], quote=True)
            sent = ([f'<td class=sentcell><input type=checkbox class=sent data-key="{key}" title="Отметь, когда отправишь CV"></td>']
                    if with_cv else [])
            cells = [r["Дата"], html.escape(r["Компания"]), html.escape(r["Должность"]), r["Оценка"],
                     link(r["Ссылка"], "открыть"), html.escape(r["Как откликаться"]),
                     link(r["Сайт компании"], "сайт")]
            cells += [cv_link(r["Папка с CV"]), html.escape(r["Чего нет"])] if with_cv else [html.escape(r["Почему"])]
            out.append("<tr>" + "".join(sent) + "".join(f"<td>{c}</td>" for c in cells) + "</tr>")
        return "\n".join(out) + "</table>"

    good = [r for r in rows if r.get("Папка с CV")]
    bad = [r for r in rows if not r.get("Папка с CV")]
    page = f"""<!doctype html><html lang="ru"><meta charset="utf-8"><title>Мои вакансии</title>
<style>body{{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#222}}h1{{color:#5f7f7c}}
table{{border-collapse:collapse;width:100%;font-size:14px;margin-bottom:32px}}
th,td{{border:1px solid #ddd;padding:6px 8px;text-align:left;vertical-align:top}}th{{background:#87a6a3;color:#fff}}
tr:nth-child(even) td{{background:#f6f8f8}}a{{color:#1a5fb4;font-weight:600}}
.sentcell{{text-align:center}}.sentcell input{{width:20px;height:20px;cursor:pointer}}
tr.done td{{background:#e9f3ec !important;color:#777}}
.filters button{{margin-right:6px;padding:6px 12px;border:1px solid #87a6a3;background:#fff;border-radius:6px;cursor:pointer}}
.filters button.on{{background:#87a6a3;color:#fff}}</style>
<h1>Мои вакансии</h1><p>Обновлено: {datetime.datetime.now():%Y-%m-%d %H:%M}. Новые сверху.</p>
<h2>Подходят, CV готово ({len(good)})</h2>
<p class=filters><button data-f=all class=on>Все</button><button data-f=todo>Не отправлены</button><button data-f=done>Отправлены</button>
<span id=count></span></p>
{table(good, True)}
<details><summary><b>Отсеяны ({len(bad)})</b></summary>{table(bad, False)}</details>
<script>
// Отметки «Отправил» хранятся в этом браузере и не пропадают, когда бот обновляет страницу.
const store = {{
  get(k) {{ try {{ return localStorage.getItem("sent:" + k) === "1"; }} catch (e) {{ return false; }} }},
  set(k, v) {{ try {{ v ? localStorage.setItem("sent:" + k, "1") : localStorage.removeItem("sent:" + k); }} catch (e) {{}} }}
}};
let filter = "all";
function refresh() {{
  let done = 0, all = 0;
  document.querySelectorAll("input.sent").forEach(box => {{
    const row = box.closest("tr");
    all++; if (box.checked) done++;
    row.classList.toggle("done", box.checked);
    row.style.display = filter === "all" || (filter === "done") === box.checked ? "" : "none";
  }});
  document.getElementById("count").textContent = `Отправлено ${{done}} из ${{all}}`;
}}
document.querySelectorAll("input.sent").forEach(box => {{
  box.checked = store.get(box.dataset.key);
  box.addEventListener("change", () => {{ store.set(box.dataset.key, box.checked); refresh(); }});
}});
document.querySelectorAll(".filters button").forEach(b => b.addEventListener("click", () => {{
  filter = b.dataset.f;
  document.querySelectorAll(".filters button").forEach(x => x.classList.toggle("on", x === b));
  refresh();
}}));
refresh();
</script></html>"""
    TRACKER_PAGE.write_text(page, encoding="utf-8")


# ---------- 5. Обработка одной вакансии ----------

def unescape_all(value):
    """Убирает HTML-коды вроде &#39; и &amp; из всех строк ответа."""
    if isinstance(value, str):
        return html.unescape(value)
    if isinstance(value, list):
        return [unescape_all(v) for v in value]
    if isinstance(value, dict):
        return {k: unescape_all(v) for k, v in value.items()}
    return value


def clean_headline(cv):
    """Убирает из заголовка слова, которых, по словам самого бота, у Ивана нет."""
    missing_text = " ".join(cv["missing_requirements"]).lower()
    parts = [p.strip() for p in cv["headline"].split("|")]
    keep = parts[:1]
    for part in parts[1:]:
        if part and part.lower() in missing_text:
            print(f"  Убрал из заголовка «{part}»: этого опыта у тебя нет.")
        else:
            keep.append(part)
    cv["headline"] = " | ".join(keep)


def print_cv_report(cv):
    print("  Заголовок:", cv["headline"])
    if cv["matched_keywords"]:
        print("  Ключевые слова из вакансии в резюме:", ", ".join(cv["matched_keywords"]))
    print("  Что изменено:", cv["notes_for_ivan"])
    if cv["missing_requirements"]:
        print("  Чего из требований у тебя нет (НЕ добавлено):")
        for item in cv["missing_requirements"]:
            print("    -", item)


def process_vacancy(vacancy_path, cv_text, settings):
    vacancy_text = read_vacancy(vacancy_path)
    first_line = vacancy_text.splitlines()[0].strip()
    link = first_line if first_line.startswith("http") else ""
    return generate_for_text(vacancy_text, link, cv_text, settings)


# Быстрый фильтр без ИИ: стажировки и альтернанс отсеиваются сразу, бесплатно.
# Свои слова можно добавить в .env: EXCLUDE_WORDS=CDD,freelance
TITLE_WORDS = ["stage", "stagiaire", "intern", "internship", "alternance", "alternant", "alternante",
               "apprenti", "apprentie", "apprentissage", "apprenticeship"]
TEXT_PHRASES = ["votre stage", "offre de stage", "stage de fin d'études", "stage de fin d’études",
                "en alternance", "rythme d'alternance", "rythme d’alternance", "contrat d'apprentissage",
                "contrat d’apprentissage", "contrat de professionnalisation", "internship", "stage de 6 mois",
                "stage de 4 à 6 mois", "convention de stage"]


def excluded_reason(vacancy_text):
    """Возвращает причину, если вакансия точно не подходит (стажировка, альтернанс и т.п.)."""
    extra = [w.strip() for w in os.getenv("EXCLUDE_WORDS", "").split(",") if w.strip()]
    lines = vacancy_text.splitlines()
    head = lines[1] if lines and lines[0].startswith("http") and len(lines) > 1 else vacancy_text[:150]
    for word in TITLE_WORDS + extra:
        if re.search(rf"(?<!\w){re.escape(word)}(?!\w)", head, flags=re.IGNORECASE):
            return f"в названии есть «{word}»"
    if re.search(r"(?<!\w)V\.?I\.?E(?!\w)", head):
        return "в названии есть «VIE»"
    lower = vacancy_text.lower()
    for phrase in TEXT_PHRASES:
        if phrase in lower:
            return f"в описании есть «{phrase}»"
    return ""


FR_WORDS = {"le", "la", "les", "des", "du", "et", "vous", "nous", "pour", "avec", "dans", "une", "est", "sur",
            "au", "aux", "votre", "vos", "notre", "nos", "qui", "que", "en", "être", "sera", "poste", "équipe"}
EN_WORDS = {"the", "and", "you", "we", "for", "with", "in", "a", "is", "on", "to", "of", "your", "our",
            "who", "will", "be", "are", "team", "role", "experience", "this", "as", "an"}


LANGUAGE_SKIPS = OUTPUT_DIR / "language_skipped.csv"


def remember_language(link, lang):
    """Запоминает вакансии, пропущенные из-за языка (ссылка;язык), чтобы не открывать их зря."""
    if not link:
        return
    OUTPUT_DIR.mkdir(exist_ok=True)
    with open(LANGUAGE_SKIPS, "a", encoding="utf-8") as f:
        f.write(f"{link};{lang}\n")


def language_skips():
    """{ссылка: язык} для вакансий, пропущенных из-за языка."""
    if not LANGUAGE_SKIPS.exists():
        return {}
    result = {}
    for line in LANGUAGE_SKIPS.read_text(encoding="utf-8").splitlines():
        link, _, lang = line.partition(";")
        if link:
            result[link] = lang
    return result


def posting_language(vacancy_text):
    """Грубо определяет язык вакансии по частым словам: 'en', 'fr' или 'other' (немецкий, голландский и т.п.)."""
    words = re.findall(r"[a-zàâçéèêëîïôûùüÿœäößñ']+", vacancy_text.lower())
    fr = sum(w in FR_WORDS for w in words)
    en = sum(w in EN_WORDS for w in words)
    if fr > en * 1.5:  # двуязычные вакансии (часть на французском, часть на английском) считаем английскими
        return "fr"
    if not words or en / len(words) < 0.08:  # в английском тексте таких слов обычно 20-30%
        return "other"
    return "en"


def generate_for_text(vacancy_text, link, cv_text, settings, extra=None):
    """Проверяет вакансию и делает под неё CV. Возвращает (путь к PDF, строка таблицы)
    или (None, строка таблицы), если вакансия не подошла."""
    today = datetime.date.today().isoformat()
    row = {"Дата": today, "Ссылка": link, **(extra or {})}

    reason = excluded_reason(vacancy_text)
    if reason:
        title = vacancy_text.splitlines()[1] if link and len(vacancy_text.splitlines()) > 1 else ""
        print(f"  Пропускаю без проверки ИИ: похоже на стажировку/альтернанс ({reason}).")
        row.update({"Должность": title, "Подходит": "нет", "Почему": f"стажировка/альтернанс: {reason}"})
        add_to_tracker(row)
        return None, row

    wanted = os.getenv("JOB_LANGUAGE", "any").strip().lower()[:2]  # en / fr / any
    if wanted in ("en", "fr"):
        lang = posting_language(vacancy_text)
        if lang != wanted:
            names = {"en": "английском", "fr": "французском", "other": "другом"}
            print(f"  Пропускаю без проверки ИИ: вакансия на {names[lang]} языке (сейчас ищу: {wanted}).")
            remember_language(link, lang)  # в таблицу не пишем: при другом языке вакансия снова пригодится
            return None, row

    if settings["filter"]:
        fit, _ = check_fit(cv_text, vacancy_text, settings["filter_model"])
        row.update({"Компания": fit["company_name"], "Должность": fit["job_title"],
                    "Оценка": fit["score"], "Почему": html.unescape(fit["reasons"]),
                    "Подходит": "да" if fit["apply"] else "нет"})
        print(f"  {fit['job_title']} в {fit['company_name']}: оценка {fit['score']}/100. {html.unescape(fit['reasons'])}")
        if not fit["apply"] or fit["score"] < settings["min_score"]:
            print("  Пропускаю: вакансия тебе не подходит (резюме не делаю, деньги не трачу).")
            add_to_tracker(row)
            return None, row

    cv, _ = ask_claude(cv_text, vacancy_text, settings["model"], settings["cv_language"],
                       settings["effort"], settings["cover_letter"])
    cv = unescape_all(cv)
    clean_headline(cv)
    print_cv_report(cv)

    stamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    folder = OUTPUT_DIR / f"{stamp}_{safe_name(cv['company_name'])}"
    folder.mkdir(parents=True, exist_ok=True)
    name = safe_name(cv["name"]) or "CV"
    html_path = folder / f"CV_{name}_{safe_name(cv['company_name'])}.html"
    pdf_path = html_path.with_suffix(".pdf")
    (folder / "details.json").write_text(json.dumps(cv, ensure_ascii=False, indent=2), encoding="utf-8")
    (folder / "vacancy.txt").write_text(vacancy_text, encoding="utf-8")
    if cv["cover_letter"].strip():
        (folder / "cover_letter.txt").write_text(cv["cover_letter"].strip() + "\n", encoding="utf-8")
        print("  Сопроводительное письмо: cover_letter.txt")

    if render_one_page(cv, html_path, pdf_path):
        print(f"  PDF: {pdf_path}")
    else:
        print(f"  Не получилось сделать PDF автоматически. Открой {html_path.name} "
              "в браузере и нажми Ctrl+P -> «Сохранить как PDF».")
        pdf_path = None

    row.update({"Компания": row.get("Компания") or cv["company_name"],
                "Должность": row.get("Должность") or cv["job_title"],
                "Подходит": row.get("Подходит") or "да",
                "Папка с CV": folder.name,
                "Чего нет": "; ".join(cv["missing_requirements"])})
    if settings.get("track", True):
        add_to_tracker(row)
    return pdf_path, row


# ---------- 6. Главная функция ----------

YES = ("yes", "да", "true", "1", "on")


def load_settings():
    load_dotenv(BASE_DIR / ".env")
    if not os.getenv("ANTHROPIC_API_KEY"):
        sys.exit("Не найден API-ключ. Создай файл .env со строкой:\n"
                 "    ANTHROPIC_API_KEY=sk-ant-...")
    return {
        "model": os.getenv("MODEL", "claude-opus-5-5").strip(),
        "effort": os.getenv("EFFORT", "medium").strip().lower(),
        "cv_language": os.getenv("CV_LANGUAGE", "auto").strip(),
        "cover_letter": os.getenv("COVER_LETTER", "yes").strip().lower() in YES,
        "filter": os.getenv("FILTER", "yes").strip().lower() in YES,
        "filter_model": os.getenv("FILTER_MODEL", "claude-haiku-5-5").strip(),
        "min_score": int(os.getenv("MIN_SCORE", "50")),
    }


def main():
    if sys.argv[1:] == ["--page"]:  # python generate_cv.py --page : только обновить страницу со ссылками
        build_tracker_page()
        print(f"Страница обновлена: {TRACKER_PAGE}")
        return
    settings = load_settings()

    print("Читаю резюме из папки cv/ ...")
    cv_text = read_cvs()

    if len(sys.argv) > 1:
        vacancies = [Path(sys.argv[1])]
    else:
        VACANCIES_DIR.mkdir(exist_ok=True)
        vacancies = sorted(VACANCIES_DIR.glob("*.txt"))
        if not vacancies:
            vacancies = [DEFAULT_VACANCY]

    done_dir = VACANCIES_DIR / "done"
    for i, path in enumerate(vacancies, 1):
        print(f"\n[{i}/{len(vacancies)}] {path.name}")
        process_vacancy(path, cv_text, settings)
        if path.parent == VACANCIES_DIR:
            done_dir.mkdir(exist_ok=True)
            path.replace(done_dir / path.name)

    build_tracker_page()
    print(f"\nВсё готово. Все вакансии со ссылками: {TRACKER_PAGE} (открой двойным щелчком)")


if __name__ == "__main__":
    main()
