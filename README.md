# 🚀 AI Daily Career Coach (Google Apps Script)

An agentic daily planning system that bridges Google Calendar availability with monthly OKR progress to deliver a personalized, high-leverage action plan via email each morning.

## 📌 Project Overview

The goal is to eliminate decision fatigue. Instead of staring at a long OKR sheet and a busy calendar, this script:

1. **Analyzes** multiple Google Calendars to find available "White Space."
2. **Syncs** OKR running counts from Airtable (interview prep, CRM) and Google Calendar (French classes).
3. **Prioritizes** tasks using Google Gemini based on effort remaining vs. calendar availability.
4. **Delivers** a curated HTML + plain-text plan to Gmail by 8 AM.

---

## 🏗 System Architecture

| Layer | Technology |
|-------|-----------|
| Runtime | Google Apps Script (V8) |
| OKR data | Google Sheets (`Mon_YY` tab per month) |
| Calendar data | Google Calendar API (free/busy across multiple calendars) |
| CRM & job data | Airtable REST API (interview prep base + CRM base) |
| AI engine | Google Gemini (`gemini-2.0-flash`, with lite fallback) |
| Email delivery | GmailApp (HTML + plain-text) |
| Secrets | GAS Script Properties (never in code) |

### Pipeline (`runDailyCoach`)

```
Config (Sheet2)
  → updateFrenchProgress   — calendar keyword scan → OKR Running Count
  → updateInterviewOKR     — Airtable counts → OKR Running Count
  → getDailyAvailability   — free/busy → totalMinutes + largestBlock
  → getActiveOKRs          — OKR sheet → task list
  → prioritizeTasksWithAI  — Gemini → ranked plan (fallback: deterministic)
  → getFeaturedProductJobs — Airtable CRM → 5 unsent job postings
  → sendCoachEmail         — Gmail HTML + plain-text
```

---

## 🚀 Installation & Setup

1. **Clone & push:** `clasp push` syncs `code.js` and `appsscript.json` to your GAS project.
2. **Enable services** in the GAS editor sidebar: Google Sheets API v4, Google Calendar API v3.
3. **Script Properties:** Add all keys listed in the table below under *Project Settings → Script Properties*.
4. **OKR sheet tab:** Create a tab named `Mon_YY` (e.g. `May_26`) in your OKR Google Sheet before the 1st of each month.
5. **Trigger:** Set a time-driven trigger on `runDailyCoach` to fire daily between 7–8 AM.

### Script Properties (Secrets)

| Key | Purpose |
|-----|---------|
| `GEMINI_API_KEY` | Google Gemini API key |
| `GEMINI_MODEL` | Primary model (default: `gemini-2.0-flash`) |
| `GEMINI_FALLBACK_MODEL` | Fallback model (default: `gemini-2.0-flash-lite`) |
| `AIRTABLE_PAT` | Airtable Personal Access Token |
| `AIRTABLE_BASE_ID` | Interview prep Airtable base |
| `AIRTABLE_BASE_ID_CRM` | CRM Airtable base (jobs + meetings) |
| `AIRTABLE_TABLE_NAME` | Interview responses table (default: `Responses`) |
| `AIRTABLE_TABLE_NAME_JOBS` | Jobs table (default: `Jobs`) |
| `AIRTABLE_TABLE_NAME_MEETINGS` | Meetings table (default: `Meetings`) |

---

## 🗂 Project Management

All feature work and bug fixes are tracked as **GitHub Issues** on this repo, managed via the [Career Coach - GAS](https://github.com/users/simmyhundal/projects/4) project board.

### Milestones = Epics

Milestones group issues by theme rather than by time box. A milestone is closed when all its issues are resolved. Current milestones:

| Milestone | Purpose |
|-----------|---------|
| OKR Sync Accuracy | Correct, reliable data flowing into OKR Running Counts |
| Relevant Tasks are listed in Daily Email | Email content quality and completeness |
| Filter Applicable Key Results to be included in Daily Task List | OKR filtering and task selection logic |

New milestones are created when work doesn't fit an existing theme.

### Story Points (Estimate field)

Effort is tracked via the **Estimate** numeric field on the project board. Use this scale:

| Points | Effort |
|--------|--------|
| 1 | ≤ 30 min |
| 2 | ~2 hours |
| 3 | ~half day |
| 5 | ~full day |
| 8 | multi-day |
| 13 | week+ |

### Labels

| Label | Meaning |
|-------|---------|
| `bug` | Something broken that was previously working |
| `enhancement` | New feature or improvement |
| `maintenance` | Keeps existing features working (refactor, rename fix, etc.) |
| `data-sync` | OKR / Airtable / Calendar data pipeline |
| `priority:high` | Blocking or high user impact |
| `priority:medium` | Important, not urgent |
| `priority:low` | Nice to have |

### Velocity

Velocity = sum of **Estimate** points closed in a given week, visible by filtering the project board by closed date. Use this to calibrate how many points to take on per week.

### Prioritization

- `priority:high` bugs jump the queue automatically.
- For enhancements: rank by **Impact ÷ Estimate** — high impact, low effort goes first.
- Groom the backlog column briefly each week.

---

## 🛠 Development Workflow

```bash
# Edit code locally, then sync to GAS
clasp push

# Test the full pipeline (sends a real email)
# → Run runDailyCoach() in the GAS editor

# Test Airtable job fetching only
# → Run previewFeaturedProductJobs() in the GAS editor

# Test French OKR sync only
# → Run updateFrenchProgress() in the GAS editor
```

### Monthly maintenance

On the 1st of each month, create a new tab in the OKR Google Sheet named `Mon_YY` (e.g. `Jun_26`). The script will silently abort if the tab is missing.

---

## ⚠️ Key Invariants

- **OKR task names are load-bearing.** Gemini is instructed to use exact names from the sheet. Renaming a key result in the sheet requires no code change, but the new name takes effect immediately — including in `updateFrenchProgress` and `updateInterviewOKR` row lookups.
- **`_CM` fields drive Airtable sync.** Current-month counts live in Airtable formula fields suffixed `_CM`. The script reads these as numbers; the formula logic lives entirely in Airtable.
- **Airtable pagination is handled.** `fetchAirtableRecords()` follows `offset` tokens automatically.
