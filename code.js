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
  updateFrenchProgress(config, sheet); // From Calendar
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
function updateFrenchProgress(config, sheet) {
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
  const headers = data[0];
  const idxName = headers.indexOf("Key Results");
  const idxEffort = headers.indexOf("Effort (mins)");
  const idxDaily = headers.indexOf("Daily");
  const idxCompleted = headers.indexOf("Completed?");

  const isYes = value => String(value || "").trim().toLowerCase() === "yes";
  
  return data.slice(1)
    .map(row => ({
      name: row[idxName],
      effort: Number(row[idxEffort]) || 0,
      isDaily: idxDaily > -1 ? isYes(row[idxDaily]) : false,
      isCompleted: idxCompleted > -1 ? isYes(row[idxCompleted]) : false
    }))
    .filter(task => task.name)
    .filter(task => !task.isCompleted)
    .filter(task => task.isDaily || task.effort > 0);
}

function buildFallbackPlan(availability, tasks) {
  const maxTotalMinutes = Math.max(30, Math.floor(availability.totalMinutes * 0.8));
  const maxSessionMinutes = Math.max(30, availability.largestBlock);
  const dailyTasks = tasks.filter(task => task && task.name && task.isDaily);
  const sortedTasks = tasks
    .filter(task => task && task.name && !task.isDaily && task.effort > 0)
    .sort((a, b) => b.effort - a.effort);

  const chosenTasks = dailyTasks.map(task => ({
    name: task.name,
    isDaily: true
  }));
  let usedMinutes = 0;

  sortedTasks.forEach(task => {
    if (chosenTasks.length >= 4 + dailyTasks.length) return;

    const sessionLength = Math.min(task.effort, maxSessionMinutes);
    if (usedMinutes + sessionLength > maxTotalMinutes) return;

    chosenTasks.push({
      name: task.name,
      sessionLength: sessionLength,
      effort: task.effort,
      isDaily: false
    });
    usedMinutes += sessionLength;
  });

  if (chosenTasks.length === 0) {
    return "### Plan needs attention\n- No eligible OKR tasks were found with remaining effort.\n- Check the OKR sheet values for task names and effort minutes.";
  }

  return chosenTasks.map(task => [
    task.isDaily
      ? `### ${task.name}`
      : `### ${task.name}: ${task.sessionLength} mins`
  ].join("\n")).join("\n\n");
}

/**
 * Uses Google Gemini to prioritize OKR tasks based on available calendar time.
 */
function prioritizeTasksWithAI(config, availability, tasks) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const model = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || 'gemini-2.0-flash';
  const fallbackModel = PropertiesService.getScriptProperties().getProperty('GEMINI_FALLBACK_MODEL') || 'gemini-2.0-flash-lite';
  if (!apiKey) {
    Logger.log("Gemini API Error: Missing GEMINI_API_KEY script property.");
    return "### Error generating plan\n- Could not connect to Gemini. Please check your API key.";
  }

  if (!tasks.length) {
    Logger.log("Gemini API Error: No active tasks found for prioritization.");
    return "### Plan needs attention\n- No active OKR tasks with remaining effort were found today.";
  }

  // 1. Format the OKR data for the prompt
  const okrSummary = tasks.map(t => 
    t.isDaily
      ? `- ${t.name} [Daily task, include unless completed]`
      : `- ${t.name} (${t.effort} mins remaining)`
  ).join('\n');

  // 2. Construct the Prompt
  const prompt = `
    Context: Today is ${new Date().toLocaleDateString()}. 
    I have ${availability.totalMinutes} minutes of total free time in Paris.
    My largest contiguous focus block is ${availability.largestBlock} minutes.
    
    Based on my OKRs, pick 2-4 tasks for today.
    Priority Rules:
    1. Focus on the highest-leverage tasks with meaningful effort remaining.
    2. Ensure the suggested session fits within my ${availability.largestBlock} min block.
    3. Do not exceed 80% of my total available time (${availability.totalMinutes * 0.8} mins).
    4. Every task marked as a daily task must be included unless it is completed.
    
    Current OKR Status:
    ${okrSummary}
    
    Output Requirements:
    Return your response in markdown, not HTML.
    Do not include any intro paragraph, summary paragraph, or closing sentence.
    Output only task sections and nothing else.
    Use this exact structure for each task:
    ### Task Name: X mins
    For daily tasks, use this exact structure instead:
    ### Task Name
    Every chosen task must come from the Current OKR Status list above.
    Keep session lengths within ${availability.largestBlock} mins.
  `;

  // 3. Prepare the Payload for Gemini API
  const payload = {
    "contents": [{
      "parts": [{
        "text": prompt
      }]
    }],
    "generationConfig": {
      "temperature": 0.7,
      "topK": 40,
      "topP": 0.95,
      "maxOutputTokens": 1024,
    }
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  const modelsToTry = [model];
  if (fallbackModel && fallbackModel !== model) {
    modelsToTry.push(fallbackModel);
  }

  for (let i = 0; i < modelsToTry.length; i++) {
    const currentModel = modelsToTry[i];
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${currentModel}:generateContent?key=${apiKey}`;

    try {
      const response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() !== 200) {
        throw new Error(`HTTP ${response.getResponseCode()}: ${response.getContentText()}`);
      }

      const result = JSON.parse(response.getContentText());
      
      // Gemini response structure: result.candidates[0].content.parts[0].text
      const aiText = result.candidates?.[0]?.content?.parts
        ?.map(part => part.text || "")
        .join("\n")
        .trim();

      const hasStructuredTasks = aiText && /^###\s+.+/m.test(aiText);

      if (hasStructuredTasks) {
        return aiText;
      }

      if (aiText) {
        Logger.log(`Gemini API Warning (${currentModel}): Response was unstructured, using fallback plan. Raw response: ${aiText}`);
        return buildFallbackPlan(availability, tasks);
      }

      throw new Error("Empty response from Gemini: " + response.getContentText());
    } catch (e) {
      Logger.log(`Gemini API Error (${currentModel}): ` + e.message);

      const isLastModel = i === modelsToTry.length - 1;
      if (isLastModel) {
        return buildFallbackPlan(availability, tasks);
      }

      const shouldTryNextModel = /HTTP 503|HTTP 429|HTTP 404/.test(String(e.message || ""));
      if (!shouldTryNextModel) {
        return buildFallbackPlan(availability, tasks);
      }
    }
  }

  return buildFallbackPlan(availability, tasks);
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

    const headingMatch = line.match(/^###\s+(.*)$/);
    if (headingMatch) {
      html.push(`<h3>${formatInlineMarkdown(escapeHtml(headingMatch[1]))}</h3>`);
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
    .replace(/^###\s+/gm, "")
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
    `<p>Today you have <strong>${availability.totalMinutes} minutes</strong> of white space across your calendars. The following are your proposed tasks: </p>`,
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
