# 🚀 AI Daily Career Coach (v1.0 - Google Apps Script)

An agentic daily planning system that bridges Google Calendar availability with Monthly OKR progress to deliver a personalized, high-leverage "Daily Action Plan" via email.

## 📌 Project Overview
The goal of this project is to eliminate decision fatigue in the morning. Instead of looking at a long list of OKRs and a busy calendar, this script:
1.  **Analyzes** multiple Google Calendars to find "White Space."
2.  **Fetches** active OKRs from a Google Sheet.
3.  **Prioritizes** tasks using OpenAI based on the "Running Count" (progress) vs. "Goal" (target).
4.  **Delivers** a curated plan directly to Gmail.

## 🏗 System Architecture
- **Environment:** Google Apps Script (GAS)
- **Data Sources:** - Google Sheets API (OKR tracking)
  - Google Calendar API (Free/Busy checking across 6+ calendars)
- **AI Engine:** OpenAI API (GPT-4o)
- **Configuration:** Externalized `Agent_Config` sheet for modularity.

---

## 🛠 Features & Current Logic
- **Modular Config:** All IDs (Calendars, Sheets) and Work Hours are pulled from a central config file.
- **Security:** API Keys are stored in `Script Properties` (Environment Secrets) rather than hardcoded.
- **Focus Detection:** The script identifies the "Largest Contiguous Block" of time to suggest deep-work tasks.

## 🚀 Installation & Setup
1.  **Repo Structure:** Copy the `.gs` files into your Google Apps Script editor.
2.  **Enable Services:** Enable the **Google Sheets**, **Gmail**, and **Google Calendar** APIs in the GAS sidebar.
3.  **Secrets:** Add your `LLM_API_KEY` to **Project Settings > Script Properties**.
4.  **Triggers:** Set a manual **Time-Driven Trigger** to run `runDailyCoach` every morning between 7:00 AM - 8:00 AM.

---

## 📋 Roadmap & Milestones (v1.0)
- [x] **Core Integration:** Connect Calendar + Sheets + OpenAI.
- [ ] **High Priority: Reliability Fix:** Resolve silent failures of the morning trigger.
- [ ] **Feature: Progress-Awareness:** Update logic to calculate `Target - Running Count` to prioritize lagging goals.
- [ ] **UI/UX: Human-Readable Email:** Refactor the email output from plain text to a clean HTML/CSS template.

## 🐛 Known Issues & Bugs
- **Trigger Reliability:** Current GAS triggers intermittently fail to send the daily email.
- **Date Formatting:** Sheets occasionally passes a `Date Object` for work hours instead of a `String`, requiring a type-check fix.
- **Readability:** The AI response is currently raw text and lacks visual hierarchy in the email body.

---

## ✍️ Maintenance
- **Monthly Update:** On the 1st of every month, update the `OKR_TAB_NAME` in your `Agent_Config` sheet (e.g., from `Mar_26` to `Apr_26`).
