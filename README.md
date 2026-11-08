# AI Job-Search & CV Tailoring Agent

A personal Python project that automates the boring part of a job search on LinkedIn:
it **finds job postings, checks if they fit my profile, and writes a tailored one-page CV
and cover letter (PDF) for each one**. I review everything and apply myself.

Built with **Python**, the **Claude API (Anthropic)** and **Playwright**.

## What it does

1. **Searches LinkedIn** in a real Chrome window, using the same natural-language phrases
   I would type myself (e.g. `business analyst paris posted in the past week`), and scrolls
   through several result pages.
2. **Filters for free** (no AI cost): already-seen jobs, internships / apprenticeships,
   senior / lead / manager titles, blacklisted companies, postings with too many applicants,
   and (optionally) postings that are not in English or French.
3. **Checks the fit with AI** (a small, cheap model) against my own rules: job titles that fit,
   seniority, years of experience. Gives a 0–100 score and a short reason.
4. **Writes a tailored CV** with a stronger model:
   - uses only real facts from my profile (never invents experience),
   - adapts the headline and keywords to the posting (ATS-friendly),
   - picks the CV language (French or English) from the posting,
   - always fits on **one A4 page** (auto-shrinks the font / moves sidebar blocks if needed),
   - adds a cover letter in plain, simple language.
5. **Tracks everything** in `output/applications.csv` and a local `vacancies.html` page with
   clickable links (LinkedIn, company site, CV, letter) and a "sent" checkbox per job.

## Tech

- Python 3, Anthropic Python SDK (structured JSON output with a JSON schema)
- Playwright (persistent Chrome profile, so I log in to LinkedIn once by hand)
- HTML → PDF with headless Chrome / Edge, embedded Open Sans fonts, one-page fitting loop
- pypdf, python-dotenv

## Project structure

```
generate_cv.py        CV + cover letter generator (also works alone on a pasted job text)
find_jobs.py          LinkedIn search, filters, fit check, tracker page
requirements.txt
.env.example          all settings with comments (copy to .env)
examples/             templates: profile (about_me) and a job text
```

## Quick start

```bash
pip install -r requirements.txt
```

1. Copy `.env.example` to `.env` and put your Anthropic API key in it.
2. Create a `cv/` folder with your CV (PDF) and a filled `about_me.txt`
   (see `examples/about_me.example.txt`). Optional: `cv/photo.jpg`.
3. One job: paste the posting into `vacancy.txt` and run `python generate_cv.py`.
4. Search LinkedIn: `python find_jobs.py` (or `python find_jobs.py en` / `fr` for one language).
   The first time, log in to LinkedIn in the window that opens.

Typical cost with the default models: about $0.05 per CV and well under a cent per fit check.

## Notes

- The bot **never applies** for you: it only prepares documents and links.
- Automating LinkedIn may go against its User Agreement. Use it slowly and at your own risk
  (the bot pauses randomly between pages and stops if LinkedIn stops responding).
- `.env`, `browser_profile/` (your LinkedIn session), `cv/` and `output/` are in `.gitignore`
  on purpose: never upload them.
- Code comments and console messages are in Russian (my native language).
