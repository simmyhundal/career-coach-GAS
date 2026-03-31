// MASTER FUNCTION: Run this to test everything
function runDailyCoach() {
  const ss_config = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss_config.getSheetByName("Sheet2"); // Double check if it's "Sheet2" or " Sheet2"
  
  if (!configSheet) {
    throw new Error("Could not find tab named 'Sheet2'. Check for leading/trailing spaces.");
  }

  // 1. Map values by searching for the Key name (more robust than hardcoding C2, C3)
  const configData = configSheet.getRange("B2:C10").getValues();
  const config = {};
  configData.forEach(row => {
    config[row[0]] = row[1];
  });

  //1.2 Dynamically determine appropriate tab
    const now = new Date();
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const currentTabName = monthNames[now.getMonth()] + "_" + now.getFullYear().toString().slice(-2);

    const ss = SpreadsheetApp.openById(config.OKR_SHEET_ID);
    const sheet = ss.getSheetByName(currentTabName);
    
    if (!sheet) {
      Logger.log(`Error: Tab '${currentTabName}' not found. Please create it for the new month.`);
      return;
    }

  // Verify we have the data
  Logger.log("Config loaded: " + JSON.stringify(config));
  
  if (!config.WORK_START || !config.WORK_END) {
    throw new Error("Could not find WORK_START or WORK_END in Column B. Check spelling!");
  }

  // --- NEW STEP ---
  // 2.1 Update French progress from yesterday's calendar and Interview Prep from Airtable
  updateFrenchProgress(config); // From Calendar
  updateInterviewOKR(config);   // From Airtable
  // ----------------

  // 2.2 Calculate White Space
  const availability = getDailyAvailability(config);
  
  // 3. Get OKRs
  const tasks = getActiveOKRs(config, sheet);
  
  // 4. Get AI Recommendation
  const aiContent = prioritizeTasksWithAI(config, availability, tasks);
  
  // 5. Send the Email
  sendCoachEmail(config.USER_EMAIL, aiContent, availability);
}

/**
 * Updates the French OKR running count based on yesterday's calendar events.
 */
function updateFrenchProgress(config) {
  // const ss = SpreadsheetApp.getActiveSpreadsheet();
  // const okrSheet = SpreadsheetApp.openById(config.OKR_SHEET_ID).getSheetByName(config.OKR_TAB_NAME);
  
  // 1. Define Yesterday's Range
  let yesterdayStart = new Date();
  yesterdayStart.setDate(yesterdayStart.getDate() - 1);
  yesterdayStart.setHours(0, 0, 0, 0);
  
  let yesterdayEnd = new Date();
  yesterdayEnd.setDate(yesterdayEnd.getDate() - 1);
  yesterdayEnd.setHours(23, 59, 59, 999);

  // 2. Fetch Events from your Main Calendar
  // Assuming 'primary' or the specific ID for your French/Life calendar
  const calendar = CalendarApp.getDefaultCalendar(); 
  const events = calendar.getEvents(yesterdayStart, yesterdayEnd);
  
  let newHours = 0;
  const keywords = ["FLE", "PMF", "pmf", "Soignant d'aide", "Pâtisserie"];

  events.forEach(event => {
    const title = event.getTitle();
    if (keywords.some(key => title.includes(key))) {
      let durationInHours = (event.getEndTime() - event.getStartTime()) / (1000 * 60 * 60);
      newHours += durationInHours;
    }
  });

  if (newHours === 0) return; // No French classes yesterday, skip update.

  // 3. Find the "French Classes" row and update Running Count
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxKR = headers.indexOf("Key Results");
  const idxRun = headers.indexOf("Running Count");

  for (let i = 1; i < data.length; i++) {
    if (data[i][idxKR].toString().includes("Attend hrs of classes / month")) {
      let currentCount = parseFloat(data[i][idxRun]) || 0;
      sheet.getRange(i + 1, idxRun + 1).setValue(currentCount + newHours);
      Logger.log(`Added ${newHours} hours to French OKR. New Total: ${currentCount + newHours}`);
      break;
    }
  }
}

function getDailyAvailability(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calendarSheet = ss.getSheetByName("Google CalendarIds");
  const calIds = calendarSheet.getRange("B2:B" + calendarSheet.getLastRow()).getValues().flat().filter(String);

// Replace the time definition lines (lines 35-41) with this:
  let todayStart = new Date();
  let startTime = config.WORK_START; // Was config.workStart
  let startH, startM;
  
  if (startTime instanceof Date) {
    startH = startTime.getHours();
    startM = startTime.getMinutes();
  } else {
    [startH, startM] = startTime.toString().split(':');
  }
  todayStart.setHours(startH, startM, 0, 0);

  let todayEnd = new Date();
  let endTime = config.WORK_END;     // Was config.workEnd
  let endH, endM;

  if (endTime instanceof Date) {
    endH = endTime.getHours();
    endM = endTime.getMinutes();
  } else {
    [endH, endM] = endTime.toString().split(':');
  }
  todayEnd.setHours(endH, endM, 0, 0);

  const response = Calendar.Freebusy.query({
    timeMin: todayStart.toISOString(),
    timeMax: todayEnd.toISOString(),
    items: calIds.map(id => ({id: id}))
  });

  // Extract all busy blocks into one sorted array
  let busySlots = [];
  for (let id in response.calendars) {
    busySlots = busySlots.concat(response.calendars[id].busy);
  }
  busySlots.sort((a, b) => new Date(a.start) - new Date(b.start));

  // Calculate gaps (White Space)
  let freeMinutes = 0;
  let largestBlock = 0;
  let lastEnd = todayStart;

  busySlots.forEach(slot => {
    let start = new Date(slot.start);
    let end = new Date(slot.end);
    if (start > lastEnd) {
      let gap = (start - lastEnd) / 1000 / 60;
      freeMinutes += gap;
      if (gap > largestBlock) largestBlock = gap;
    }
    if (end > lastEnd) lastEnd = end;
  });

  // Check final gap before end of workday
  if (todayEnd > lastEnd) {
    let finalGap = (todayEnd - lastEnd) / 1000 / 60;
    freeMinutes += finalGap;
    if (finalGap > largestBlock) largestBlock = finalGap;
  }

  return { totalMinutes: Math.round(freeMinutes), largestBlock: Math.round(largestBlock) };
}

function getActiveOKRs(config, sheet) {
  const data = sheet.getDataRange().getValues();
  
  // Based on your CSV, Key Result is Column B (Index 1) 
  // and Effort (mins) is Column E (Index 4)
  return data.slice(1)
    .map(row => ({ name: row[1], effort: row[4] }))
    .filter(row => row.name && row.effort > 0);
}

/**
 * Calls OpenAI Chat Completions API securely.
 * Uses 'LLM_API_KEY' from Script Properties.
 */
function callOpenAI(prompt) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('LLM_API_KEY');
  
  if (!apiKey) throw new Error("Missing LLM_API_KEY in Script Properties.");

  const url = "https://api.openai.com/v1/chat/completions";
  const payload = {
    model: "gpt-4o", // You can also use gpt-4-turbo or gpt-3.5-turbo
    messages: [
      { role: "system", content: "You are a daily career coach. Your goal is to select the most high-leverage tasks from a list of OKRs that fit within a user's specific calendar availability. Keep recommendations concise, encouraging, and under 100 words." },
      { role: "user", content: prompt }
    ],
    temperature: 0.7
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { "Authorization": "Bearer " + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());
  
  if (json.error) throw new Error("OpenAI Error: " + json.error.message);
  return json.choices[0].message.content;
}

function prioritizeTasksWithAI(config, availability, tasks) {
  // Format OKRs for the prompt
  const okrSummary = tasks.map(t => 
      `- ${t.name}: ${t.unitsLeft} units remaining (Total effort left: ${t.totalEffortLeft} mins)`
    ).join('\n');
    
    const prompt = `
      Context: Today I have ${availability.totalMinutes} minutes of total free time. 
      My largest contiguous focus block is ${availability.largestBlock} minutes.
      
      Based on my OKRs, pick 2-4 tasks. 
      Priority Rules:
      1. If a task's total effort left is small, try to finish it completely today.
      2. If a task is large, suggest a "session" that fits into my largest block.
      3. Do not exceed 80% of my available time.
      
      Tasks available:
      ${okrSummary}
    `;

  return callOpenAI(prompt);
}

function escapeHtml(text) {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function formatInlineMarkdown(text) {
  // Convert markdown bold to HTML bold.
  return text.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>");
}

function markdownToHtml(markdown) {
  const lines = (markdown || "").split(/\r?\n/);
  const html = [];
  let inList = false;

  lines.forEach(rawLine => {
    const line = rawLine.trim();

    if (!line) {
      if (inList) {
        html.push("</ul>");
        inList = false;
      }
      return;
    }

    const bulletMatch = line.match(/^[-*]\s+(.*)$/);
    if (bulletMatch) {
      if (!inList) {
        html.push("<ul>");
        inList = true;
      }
      html.push(`<li>${formatInlineMarkdown(escapeHtml(bulletMatch[1]))}</li>`);
      return;
    }

    if (inList) {
      html.push("</ul>");
      inList = false;
    }
    html.push(`<p>${formatInlineMarkdown(escapeHtml(line))}</p>`);
  });

  if (inList) {
    html.push("</ul>");
  }

  return html.join("\n");
}

function markdownToPlainText(markdown) {
  return (markdown || "")
    .replace(/\*\*(.+?)\*\*/g, "$1")
    .replace(/^[-*]\s+/gm, "- ");
}

function sendCoachEmail(recipient, aiContent, availability) {
  const subjectDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM d, yyyy");
  const subject = `Your Daily Career Coach: ${subjectDate}`;
  const aiPlain = markdownToPlainText(aiContent).trim();
  const aiHtml = markdownToHtml(aiContent);

  const plainBody = [
    "Good morning!",
    "",
    `Today you have ${availability.totalMinutes} minutes of White Space across your calendars.`,
    "",
    aiPlain,
    "",
    "Go get 'em!"
  ].join("\n");

  const htmlBody = [
    "<div style=\"font-family:Arial,sans-serif;line-height:1.5;color:#222;\">",
    "<p><strong>Good morning!</strong></p>",
    `<p>Today you have <strong>${availability.totalMinutes} minutes</strong> of White Space across your calendars.</p>`,
    aiHtml,
    "<p>Go get &#39;em!</p>",
    "</div>"
  ].join("\n");

  GmailApp.sendEmail(recipient, subject, plainBody, { htmlBody: htmlBody });
}

function updateInterviewOKR(config) {
  const pat = PropertiesService.getScriptProperties().getProperty('AIRTABLE_PAT');
  const baseId = PropertiesService.getScriptProperties().getProperty('AIRTABLE_BASE_ID'); 
  const tableName = "Responses";
  const fieldName = "Updated Response Modified This Month";
  
  // 1. Updated Filter Syntax based on successful debug test
  const filter = `({${fieldName}} = TRUE())`;
  const url = `https://api.airtable.com/v0/${baseId}/${encodeURIComponent(tableName)}?filterByFormula=${encodeURIComponent(filter)}`;
  
  const options = {
    "method": "get",
    "headers": { "Authorization": "Bearer " + pat },
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    // Total count of records marked 'True' this month
    const count = data.records ? data.records.length : 0;
    Logger.log(`Airtable Sync: Found ${count} updated responses.`);

    // 2. Dynamic Tab Selection (e.g., Mar_26)
    const now = new Date();
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const currentTabName = monthNames[now.getMonth()] + "_" + now.getFullYear().toString().slice(-2);
    
    const ss = SpreadsheetApp.openById(config.OKR_SHEET_ID);
    const sheet = ss.getSheetByName(currentTabName);
    
    if (!sheet) {
      Logger.log(`Error: Tab '${currentTabName}' not found. Please create it for the new month.`);
      return;
    }

    // 3. Update the Spreadsheet
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const idxKR = headers.indexOf("Key Results");
    const idxRun = headers.indexOf("Running Count");

    let matchFound = false;
    for (let i = 1; i < rows.length; i++) {
      // Searching for the specific Interview Question OKR
      if (rows[i][idxKR].toString().includes("responses to common interview questions")) {
        sheet.getRange(i + 1, idxRun + 1).setValue(count);
        Logger.log(`Spreadsheet Updated: ${currentTabName} row ${i+1} set to ${count}.`);
        matchFound = true;
        break;
      }
    }
    
    if (!matchFound) Logger.log("Warning: Could not find a matching Key Result row in the sheet.");

  } catch (e) {
    Logger.log("Airtable Sync Error: " + e.message);
  }
}