# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a **Google Apps Script (GAS)** project — a single `code.js` file that runs entirely in the Google cloud, not locally. It sends a daily email with an AI-prioritized task plan based on Google Calendar availability and OKR progress.

## Deployment

There is no local build/test/lint toolchain. Development workflow:
- **Edit via GAS editor:** Copy `code.js` content directly into the Google Apps Script project at [script.google.com](https://script.google.com).
- **Or use clasp:** `clasp push` to sync local changes to the GAS project.
- **Test manually:** Run `runDailyCoach()` from the GAS editor to execute the full pipeline. Use `Logger.log()` output in the GAS Execution Log to debug.
- **Preview jobs only:** Call `previewFeaturedProductJobs()` to test Airtable job fetching without sending an email.
- **Trigger:** A time-driven GAS trigger calls `runDailyCoach()` daily at 7–8 AM.

## Architecture

The pipeline runs top-to-bottom in `runDailyCoach()`:

1. **Config** — Read from the active spreadsheet's `Sheet2` tab, range `B2:C10`. Keys: `OKR_SHEET_ID`, `WORK_START`, `WORK_END`, `USER_EMAIL`.
2. **Calendar IDs** — Read from the `Google CalendarIds` sheet tab, column B. Multiple calendars are merged for free/busy analysis.
3. **OKR sheet** — A separate Google Sheet (identified by `OKR_SHEET_ID`). Each month gets its own tab named `Mon_YY` (e.g., `May_26`). Columns that matter: `Key Results`, `Effort (mins)`, `Daily`, `Completed?`, `Running Count`.
4. **OKR syncing** (`updateInterviewOKR`) — Before building the plan, the script pulls counts from two Airtable bases and writes them back to the `Running Count` column:
   - Interview prep base (`AIRTABLE_BASE_ID`) — table `Responses`, field `Updated Response Modified This Month`.
   - CRM base (`AIRTABLE_BASE_ID_CRM`) — tables `Meetings` and `Jobs`, using `_CM` suffixed fields for current-month counts.
5. **French OKR syncing** (`updateFrenchProgress`) — Reads yesterday's default calendar events for keyword matches (`FLE`, `PMF`, etc.) and increments the Running Count for the French classes key result.
6. **AI prioritization** (`prioritizeTasksWithAI`) — Calls Google Gemini API. Primary model from script property `GEMINI_MODEL` (default `gemini-2.0-flash`), fallback to `GEMINI_FALLBACK_MODEL`. If Gemini is unavailable or returns malformed output, `buildFallbackPlan()` runs a deterministic local fallback. The AI output is validated via `hasValidStructuredTasks()` — responses must use `### Task Name: X mins` or `### Task Name` headings exactly matching OKR task names.
7. **Email** (`sendCoachEmail`) — Sends both plain-text and HTML versions via `GmailApp`. Markdown is converted to HTML inline; HTML is never constructed from unescaped user data.

## Script Properties (Secrets)

All secrets live in **GAS Project Settings > Script Properties** — never in code:

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

## Key Invariants

- **Monthly tab must exist:** The OKR sheet must have a tab named `Mon_YY` for the current month before the trigger fires on the 1st. No tab = silent abort.
- **OKR task names are load-bearing:** The AI prompt instructs Gemini to use exact task names from the OKR sheet. `hasValidStructuredTasks()` enforces this — any mismatch triggers the fallback plan. When renaming OKR rows in the sheet, no code changes are needed, but be aware the AI prompt uses whatever strings are in the sheet.
- **`_CM` fields drive OKR sync:** Airtable fields with `_CM` suffixes represent current-month computed counts. The script reads these as numbers and writes them to `Running Count` in the OKR sheet. The Airtable formula logic for `_CM` fields lives in Airtable itself, not here.
- **Airtable pagination:** `fetchAirtableRecords()` handles Airtable's `offset`-based pagination automatically. All filtered queries (interview responses) use a separate `filterByFormula` fetch, while full record sets are fetched unfiltered for aggregation.
- **README is partially stale:** The README references OpenAI/GPT-4o, but the code uses Google Gemini. The README's architecture section reflects an earlier version.
