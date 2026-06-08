/**
 * Updates the French OKR running count based on calendar events since the
 * start of the current month through yesterday (idempotent).
 */
function updateFrenchProgress(config, sheet) {
  const keywords = ["FLE", "PMF", "pmf", "Soignant d'aide", "Pâtisserie", "Preply", "preply"];

  const now = new Date();
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0);

  const yesterdayEnd = new Date(now);
  yesterdayEnd.setDate(yesterdayEnd.getDate() - 1);
  yesterdayEnd.setHours(23, 59, 59, 999);

  if (yesterdayEnd < monthStart) {
    Logger.log("French Progress: no completed days in this month yet, skipping.");
    return;
  }

  const calendar = CalendarApp.getDefaultCalendar();
  const events = calendar.getEvents(monthStart, yesterdayEnd);

  let totalHours = 0;
  events.forEach(event => {
    const title = event.getTitle();
    if (keywords.some(key => title.includes(key))) {
      totalHours += (event.getEndTime() - event.getStartTime()) / (1000 * 60 * 60);
    }
  });

  Logger.log(`French Progress: ${totalHours} hours of French classes from ${monthStart.toDateString()} to yesterday.`);

  updateRunningCountForKeyResult(sheet, "French Courses", totalHours);
}

/**
 * Standalone runner for updateFrenchJournalOKR — loads config from the
 * active spreadsheet so the function can be triggered directly from the
 * GAS editor without passing arguments.
 */
function runFrenchJournalOKR() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Sheet2");
  if (!configSheet) { Logger.log("Could not find Sheet2 for config."); return; }
  const configData = configSheet.getRange("B2:C10").getValues();
  const config = {};
  configData.forEach(row => { config[row[0]] = row[1]; });
  updateFrenchJournalOKR(config);
}

/**
 * Counts H1 headings in the French journal Google Doc tab that contain the
 * current French month name (e.g. "mai 1", "mai 27") and writes the total
 * to the "Journal 2+ sentences in French" OKR Running Count.
 *
 * Doc ID and tab ID are read from script properties (JOURNAL_DOC_ID,
 * JOURNAL_TAB_ID) with hardcoded defaults so no property setup is required.
 */
function updateFrenchJournalOKR(config, sheet) {
  const FRENCH_MONTHS = [
    "janvier", "février", "mars", "avril", "mai", "juin",
    "juillet", "août", "septembre", "octobre", "novembre", "décembre"
  ];

  if (!sheet) {
    sheet = getCurrentOKRSheet(config);
    if (!sheet) return;
  }

  const props = PropertiesService.getScriptProperties();
  const docId  = props.getProperty('JOURNAL_DOC_ID')  || "1OQMDDuSLubv91YeincZ7Jtx-xkKRzdv6gSTJJTk-txI";
  const tabId  = props.getProperty('JOURNAL_TAB_ID')  || "t.m1y7ffcdjs6m";

  const currentMonth = FRENCH_MONTHS[new Date().getMonth()];

  try {
    const doc = DocumentApp.openById(docId);

    let tab = doc.getTab(tabId);
    if (!tab) {
      const allTabs = doc.getTabs();
      Logger.log(`getTab("${tabId}") returned null. Available tabs (${allTabs.length}): `
        + allTabs.map(t => `id="${t.getId()}" title="${t.getTitle()}"`).join(" | "));
      tab = allTabs.find(t =>
        t.getId() === tabId ||
        t.getId() === "t." + tabId ||
        t.getId().replace(/^t\./, "") === tabId
      ) || null;
    }

    if (!tab) {
      Logger.log(`French Journal Sync Error: No tab matched "${tabId}". Check log above for available IDs.`);
      return;
    }

    const body = tab.asDocumentTab().getBody();

    let count = 0;
    for (let i = 0; i < body.getNumChildren(); i++) {
      const child = body.getChild(i);
      if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
      const para = child.asParagraph();
      if (para.getHeading() !== DocumentApp.ParagraphHeading.HEADING1) continue;
      const text = para.getText().toLowerCase();
      if (text.includes(currentMonth)) {
        count++;
        Logger.log(`French Journal: H1 counted — "${para.getText()}"`);
      }
    }

    Logger.log(`French Journal Sync: ${count} H1 entries for '${currentMonth}'.`);
    updateRunningCountForKeyResult(sheet, "Journal 2+ sentences in French", count);

  } catch (e) {
    Logger.log("French Journal Sync Error: " + e.message);
  }
}

function getCurrentOKRSheet(config) {
  const now = new Date();
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const currentTabName = monthNames[now.getMonth()] + "_" + now.getFullYear().toString().slice(-2);

  const ss = SpreadsheetApp.openById(config.OKR_SHEET_ID);
  const sheet = ss.getSheetByName(currentTabName);

  if (!sheet) {
    Logger.log(`Error: Tab '${currentTabName}' not found. Please create it for the new month.`);
    return null;
  }

  return sheet;
}

function updateRunningCountForKeyResult(sheet, keyResultText, runningCount) {
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idxKR = headers.indexOf("Key Results");
  const idxRun = headers.indexOf("Running Count");

  if (idxKR === -1 || idxRun === -1) {
    Logger.log("Warning: Could not find 'Key Results' or 'Running Count' columns in the OKR sheet.");
    return false;
  }

  const searchText = keyResultText.trim();
  let updatedRows = 0;

  for (let i = 1; i < rows.length; i++) {
    const cellText = rows[i][idxKR].toString().trim();
    if (cellText === searchText) {
      sheet.getRange(i + 1, idxRun + 1).setValue(runningCount);
      Logger.log(`Spreadsheet Updated: row ${i + 1} ("${cellText}") for '${searchText}' set to ${runningCount}.`);
      updatedRows++;
    }
  }

  if (updatedRows > 0) {
    Logger.log(`updateRunningCountForKeyResult: updated ${updatedRows} row(s) for '${searchText}'.`);
    return true;
  }

  const allKRValues = rows.slice(1)
    .map(r => r[idxKR].toString().trim())
    .filter(v => v.length > 0);
  Logger.log(`Warning: No matching row found for '${searchText}'. KR values in sheet: ${JSON.stringify(allKRValues)}`);
  return false;
}

function fetchAirtableRecords(baseId, tableName, pat) {
  const options = {
    "method": "get",
    "headers": { "Authorization": "Bearer " + pat },
    "muteHttpExceptions": true
  };

  let allRecords = [];
  let offset;

  do {
    const url = offset
      ? `https://api.airtable.com/v0/${baseId}/${encodeURIComponent(tableName)}?offset=${encodeURIComponent(offset)}`
      : `https://api.airtable.com/v0/${baseId}/${encodeURIComponent(tableName)}`;
    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() !== 200) {
      throw new Error(`HTTP ${response.getResponseCode()}: ${response.getContentText()}`);
    }

    const data = JSON.parse(response.getContentText());
    allRecords = allRecords.concat(data.records || []);
    offset = data.offset;
  } while (offset);

  return allRecords;
}

function normalizeJobPosting(job) {
  return {
    title: String(job.title || "").trim(),
    targetDate: String(job.targetDate || "").trim(),
    url: String(job.url || "").trim()
  };
}

function getFeaturedProductJobs() {
  const pat = PropertiesService.getScriptProperties().getProperty('AIRTABLE_PAT');
  const crmBaseId = PropertiesService.getScriptProperties().getProperty('AIRTABLE_BASE_ID_CRM');
  const jobsTable = PropertiesService.getScriptProperties().getProperty('AIRTABLE_TABLE_NAME_JOBS') || "Jobs";

  if (!pat || !crmBaseId) {
    Logger.log("Recent Airtable jobs skipped: missing AIRTABLE_PAT or AIRTABLE_BASE_ID_CRM.");
    return [];
  }

  try {
    const records = fetchAirtableRecords(crmBaseId, jobsTable, pat);
    return records
      .filter(record => !hasAppSubmissionDate(record))
      .sort((a, b) => getCreatedTimestamp(b) - getCreatedTimestamp(a))
      .slice(0, 5)
      .map(mapAirtableJobRecord)
      .filter(job => job.title || job.url);
  } catch (e) {
    Logger.log("Recent Airtable jobs lookup error: " + e.message);
    return [];
  }
}

function hasAppSubmissionDate(record) {
  const value = record.fields?.["App Submission Date"];
  if (Array.isArray(value)) {
    return value.some(item => String(item || "").trim());
  }
  return String(value || "").trim() !== "";
}

function getFirstFieldValue(fields, candidateFields) {
  for (let i = 0; i < candidateFields.length; i++) {
    const rawValue = fields[candidateFields[i]];
    if (Array.isArray(rawValue)) {
      const firstValue = String(rawValue[0] || "").trim();
      if (firstValue) {
        return firstValue;
      }
    }

    const value = String(rawValue || "").trim();
    if (value) {
      return value;
    }
  }

  return "";
}

function getCreatedTimestamp(record) {
  const createdValue = record.fields?.Created || record.createdTime || "";
  const timestamp = new Date(createdValue).getTime();
  return isNaN(timestamp) ? 0 : timestamp;
}

function mapAirtableJobRecord(record) {
  const fields = record.fields || {};

  return normalizeJobPosting({
    title: getFirstFieldValue(fields, ["Job Title", "Title", "Role", "Position", "Name"]),
    targetDate: getFirstFieldValue(fields, ["Target App Submission Date"]),
    url: getFirstFieldValue(fields, ["Link", "Job URL", "URL", "Posting URL"])
  });
}

function previewFeaturedProductJobs() {
  const jobs = getFeaturedProductJobs();
  const output = buildFeaturedJobsPlainText(jobs);
  Logger.log(output);
  return output;
}

function updateInterviewOKR(config) {
  const pat = PropertiesService.getScriptProperties().getProperty('AIRTABLE_PAT');
  const baseId = PropertiesService.getScriptProperties().getProperty('AIRTABLE_BASE_ID');
  const tableName = PropertiesService.getScriptProperties().getProperty('AIRTABLE_TABLE_NAME') || "Responses";
  const starTableName = PropertiesService.getScriptProperties().getProperty('AIRTABLE_TABLE_NAME_STAR') || "STAR+";
  const crmBaseId = PropertiesService.getScriptProperties().getProperty('AIRTABLE_BASE_ID_CRM');
  const crmMeetingsTable = PropertiesService.getScriptProperties().getProperty('AIRTABLE_TABLE_NAME_MEETINGS') || "Meetings";
  const crmJobsTable = PropertiesService.getScriptProperties().getProperty('AIRTABLE_TABLE_NAME_JOBS') || "Jobs";

  try {
    const allResponseRecords = fetchAirtableRecords(baseId, tableName, pat);

    const count = allResponseRecords.filter(r =>
      (Number(r.fields?.["Updated Response Modified This Month"]) || 0) > 0
    ).length;

    const responseCmSum = allResponseRecords.reduce((total, record) => {
      return total + (Number(record.fields?.Response_CM) || 0);
    }, 0);
    Logger.log(`Airtable Sync: Found ${count} responses updated this month (from ${allResponseRecords.length} total). Response_CM sum=${responseCmSum}.`);

    const starRecords = fetchAirtableRecords(baseId, starTableName, pat);
    const starDoneCount = starRecords.filter(r => r.fields?.["Done?"] === "yes").length;
    Logger.log(`STAR+ Sync: ${starRecords.length} total stories, ${starDoneCount} marked Done.`);

    const sheet = getCurrentOKRSheet(config);
    if (!sheet) {
      return;
    }

    updateRunningCountForKeyResult(sheet, "responses to common interview questions", count);
    updateRunningCountForKeyResult(sheet, "Establish STAR Behavioral Interview Responses", starDoneCount);
    updateRunningCountForKeyResult(sheet, "Record responses to common interview questions and receive feedback from GenAI", responseCmSum);

    if (!crmBaseId) {
      Logger.log("Airtable CRM Sync Warning: Missing AIRTABLE_BASE_ID_CRM script property.");
      return;
    }

    const meetingRecords = fetchAirtableRecords(crmBaseId, crmMeetingsTable, pat);
    const clinicianInterviewSum = meetingRecords.reduce((total, record) => {
      return total + (Number(record.fields?.Clinician_User_Interview_CM) || 0);
    }, 0);
    const practiceInterviewSum = meetingRecords.reduce((total, record) => {
      return total + (Number(record.fields?.Practice_Interview_CM) || 0);
    }, 0);
    const soloCaseInterviewSum = meetingRecords.reduce((total, record) => {
      return total + (Number(record.fields?.Solo_Case_Interview_CM) || 0);
    }, 0);

    Logger.log(`Airtable CRM Sync: Found ${meetingRecords.length} meetings. Clinician sum=${clinicianInterviewSum}, Practice sum=${practiceInterviewSum}, Solo case sum=${soloCaseInterviewSum}.`);

    updateRunningCountForKeyResult(sheet, "Establish contact with active clinicians", clinicianInterviewSum);
    updateRunningCountForKeyResult(sheet, "Practice Interviews (case + behavioral ideally)", practiceInterviewSum);
    updateRunningCountForKeyResult(sheet, "Complete case interviews solo", soloCaseInterviewSum);

    const jobRecords = fetchAirtableRecords(crmBaseId, crmJobsTable, pat);
    const referralApplicationsSum = jobRecords.reduce((total, record) => {
      return total + (Number(record.fields?.AppSubmitted_Product_ActualReferral_CM) || 0);
    }, 0);
    const productApplicationsSum = jobRecords.reduce((total, record) => {
      return total + (Number(record.fields?.AppSubmitted_Product_CM) || 0);
    }, 0);
    const jobsListedSum = jobRecords.reduce((total, record) => {
      return total + (Number(record.fields?.JDCreated_Product_CM) || 0);
    }, 0);
    const strategyReferralApplicationsSum = jobRecords.reduce((total, record) => {
      return total + (Number(record.fields?.AppSubmitted_Strategy_ActualReferral_CM) || 0);
    }, 0);
    const strategyApplicationsSum = jobRecords.reduce((total, record) => {
      return total + (Number(record.fields?.AppSubmitted_Strategy_CM) || 0);
    }, 0);

    Logger.log(`Airtable CRM Jobs Sync: Found ${jobRecords.length} jobs. Referral apps=${referralApplicationsSum}, PM apps=${productApplicationsSum}, PM jobs listed=${jobsListedSum}. Strategy referral apps=${strategyReferralApplicationsSum}, Strategy apps=${strategyApplicationsSum}.`);

    updateRunningCountForKeyResult(sheet, "Apply to PM jobs that included a referral", referralApplicationsSum);
    updateRunningCountForKeyResult(sheet, "Apply to PM jobs", productApplicationsSum);
    updateRunningCountForKeyResult(sheet, "Find and list PM Jobs", jobsListedSum);
    updateRunningCountForKeyResult(sheet, "Apply to Consulting jobs that include a referral", strategyReferralApplicationsSum);
    updateRunningCountForKeyResult(sheet, "Apply to Consulting jobs", strategyApplicationsSum);

  } catch (e) {
    Logger.log("Airtable Sync Error: " + e.message);
  }
}
