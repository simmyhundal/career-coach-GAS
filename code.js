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

  // 4.1 Pull live PM jobs from five-star employers
  const featuredJobs = getFeaturedProductJobs();
  
  // 5. Send the Email
  sendCoachEmail(config.USER_EMAIL, aiContent, availability, featuredJobs);
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

function getWorkdayRange(config) {
  let todayStart = new Date();
  let startTime = config.WORK_START;
  let startH, startM;

  if (startTime instanceof Date) {
    startH = startTime.getHours();
    startM = startTime.getMinutes();
  } else {
    [startH, startM] = startTime.toString().split(':');
  }
  todayStart.setHours(startH, startM, 0, 0);

  let todayEnd = new Date();
  let endTime = config.WORK_END;
  let endH, endM;

  if (endTime instanceof Date) {
    endH = endTime.getHours();
    endM = endTime.getMinutes();
  } else {
    [endH, endM] = endTime.toString().split(':');
  }
  todayEnd.setHours(endH, endM, 0, 0);

  return { start: todayStart, end: todayEnd };
}

function getCalendarIds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calendarSheet = ss.getSheetByName("Google CalendarIds");
  return calendarSheet.getRange("B2:B" + calendarSheet.getLastRow()).getValues().flat().filter(String);
}

function formatEventTimeRange(start, end) {
  const tz = Session.getScriptTimeZone();
  return `${Utilities.formatDate(start, tz, "h:mm a")} - ${Utilities.formatDate(end, tz, "h:mm a")}`;
}

function getUpcomingWorkdayEvents(calIds, todayStart, todayEnd) {
  const seenEvents = {};
  const events = [];

  calIds.forEach(calId => {
    const calendar = CalendarApp.getCalendarById(calId);
    if (!calendar) {
      Logger.log(`Warning: Could not find calendar '${calId}' while building email event list.`);
      return;
    }

    calendar.getEvents(todayStart, todayEnd).forEach(event => {
      const title = String(event.getTitle() || "").trim() || "Untitled event";
      const start = event.getStartTime();
      const end = event.getEndTime();
      const dedupeKey = [title, start.getTime(), end.getTime()].join("|");

      if (seenEvents[dedupeKey]) {
        return;
      }

      seenEvents[dedupeKey] = true;
      events.push({
        title: title,
        start: start,
        end: end,
        timeLabel: formatEventTimeRange(start, end)
      });
    });
  });

  return events.sort((a, b) => a.start - b.start);
}

function getDailyAvailability(config) {
  const calIds = getCalendarIds();
  const workday = getWorkdayRange(config);
  const todayStart = workday.start;
  const todayEnd = workday.end;

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

  return {
    totalMinutes: Math.round(freeMinutes),
    largestBlock: Math.round(largestBlock),
    upcomingEvents: getUpcomingWorkdayEvents(calIds, todayStart, todayEnd)
  };
}

function normalizeTaskName(value) {
  if (value instanceof Date) {
    return "";
  }

  const name = String(value || "").trim();
  if (!name) {
    return "";
  }

  // Skip accidental numeric/date-like sheet values that should not become task titles.
  if (!/\p{L}/u.test(name)) {
    return "";
  }

  return name;
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
      name: normalizeTaskName(row[idxName]),
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

function hasValidStructuredTasks(aiText, tasks) {
  if (!aiText) return false;

  const taskLookup = {};
  tasks.forEach(task => {
    if (task && task.name) {
      taskLookup[String(task.name).trim()] = task;
    }
  });

  const headings = aiText.split(/\r?\n/)
    .map(line => line.trim())
    .filter(line => line.startsWith("### "));

  if (!headings.length) return false;

  return headings.every(heading => {
    const content = heading.replace(/^###\s+/, "").trim();
    if (!content) return false;

    if (taskLookup[content]?.isDaily) {
      return true;
    }

    const timedMatch = content.match(/^(.*):\s*(\d+)\s+mins$/);
    if (!timedMatch) return false;

    const taskName = timedMatch[1].trim();
    return Boolean(taskLookup[taskName] && !taskLookup[taskName].isDaily);
  });
}

function isGeminiQuotaError(message) {
  const text = String(message || "");
  return /HTTP 429|RESOURCE_EXHAUSTED|quota|out of credits|insufficient credits|billing/i.test(text);
}

function getExaApiKey() {
  return PropertiesService.getScriptProperties().getProperty('EXA_API_KEY');
}

function fetchExaSearchResults(query, extraPayload) {
  const apiKey = getExaApiKey();
  if (!apiKey) {
    throw new Error("Missing EXA_API_KEY script property.");
  }

  const payload = Object.assign({
    query: query,
    type: "auto",
    numResults: 10,
    text: false
  }, extraPayload || {});

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch("https://api.exa.ai/search", options);
  if (response.getResponseCode() !== 200) {
    throw new Error(`HTTP ${response.getResponseCode()}: ${response.getContentText()}`);
  }

  const result = JSON.parse(response.getContentText());
  return result.results || [];
}

function fetchExaContents(urls, extraPayload) {
  const apiKey = getExaApiKey();
  if (!apiKey) {
    throw new Error("Missing EXA_API_KEY script property.");
  }

  const payload = Object.assign({
    urls: urls,
    text: true,
    livecrawl: "preferred",
    livecrawlTimeout: 12000
  }, extraPayload || {});

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch("https://api.exa.ai/contents", options);
  if (response.getResponseCode() !== 200) {
    throw new Error(`HTTP ${response.getResponseCode()}: ${response.getContentText()}`);
  }

  const result = JSON.parse(response.getContentText());
  return result.results || [];
}

function fetchExaStructuredAnswer(query, outputSchema) {
  const apiKey = getExaApiKey();
  if (!apiKey) {
    throw new Error("Missing EXA_API_KEY script property.");
  }

  const payload = {
    query: query,
    text: true,
    outputSchema: outputSchema
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch("https://api.exa.ai/answer", options);
  if (response.getResponseCode() !== 200) {
    throw new Error(`HTTP ${response.getResponseCode()}: ${response.getContentText()}`);
  }

  const result = JSON.parse(response.getContentText());
  return result.answer;
}

function buildStructuredTasksMarkdown(taskItems, tasks) {
  if (!taskItems || !taskItems.length) {
    return "";
  }

  const taskLookup = tasks.reduce((lookup, task) => {
    if (task && task.name) {
      lookup[String(task.name).trim()] = task;
    }
    return lookup;
  }, {});

  const lines = taskItems
    .map(item => {
      const taskName = String(item.name || "").trim();
      const sourceTask = taskLookup[taskName];
      if (!sourceTask) {
        return "";
      }

      if (sourceTask.isDaily) {
        return `### ${taskName}`;
      }

      const minutes = Math.round(Number(item.minutes) || 0);
      if (!minutes) {
        return "";
      }

      return `### ${taskName}: ${minutes} mins`;
    })
    .filter(Boolean);

  const markdown = lines.join("\n\n");
  return hasValidStructuredTasks(markdown, tasks) ? markdown : "";
}

function prioritizeTasksWithExa(availability, tasks) {
  const outputSchema = {
    type: "object",
    properties: {
      tasks: {
        type: "array",
        items: {
          type: "object",
          properties: {
            name: { type: "string" },
            minutes: { type: "number" }
          },
          required: ["name", "minutes"],
          additionalProperties: false
        }
      }
    },
    required: ["tasks"],
    additionalProperties: false
  };

  const okrSummary = tasks.map(t =>
    t.isDaily
      ? `- ${t.name} [Daily task, include unless completed]`
      : `- ${t.name} (${t.effort} mins remaining)`
  ).join('\n');

  const query = `
Today I have ${availability.totalMinutes} minutes of total free time in Paris and my largest contiguous focus block is ${availability.largestBlock} minutes.

From the OKRs below, choose 2-4 tasks for today.

Rules:
- Focus on the highest-leverage tasks with meaningful effort remaining.
- Ensure the suggested session fits within ${availability.largestBlock} minutes.
- Do not exceed 80% of my total available time (${availability.totalMinutes * 0.8} mins).
- Every daily task must be included unless it is completed.
- Every chosen task must come exactly from this OKR list.

OKRs:
${okrSummary}

Return a structured object with a tasks array.
For daily tasks, set minutes to 0.
For non-daily tasks, set minutes to the planned session length.
`;

  try {
    const answer = fetchExaStructuredAnswer(query, outputSchema);
    const markdown = buildStructuredTasksMarkdown(answer?.tasks, tasks);
    return markdown || buildFallbackPlan(availability, tasks);
  } catch (e) {
    Logger.log("Exa task prioritization error: " + e.message);
    return buildFallbackPlan(availability, tasks);
  }
}

/**
 * Uses Google Gemini to prioritize OKR tasks based on available calendar time.
 */
function prioritizeTasksWithAI(config, availability, tasks) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const model = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || 'gemini-2.0-flash';
  const fallbackModel = PropertiesService.getScriptProperties().getProperty('GEMINI_FALLBACK_MODEL') || 'gemini-2.0-flash-lite';
  if (!tasks.length) {
    Logger.log("Gemini API Error: No active tasks found for prioritization.");
    return "### Plan needs attention\n- No active OKR tasks with remaining effort were found today.";
  }

  if (!apiKey) {
    Logger.log("Gemini API Error: Missing GEMINI_API_KEY script property. Falling back to Exa.");
    return prioritizeTasksWithExa(availability, tasks);
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

      const hasStructuredTasks = hasValidStructuredTasks(aiText, tasks);

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

      if (isGeminiQuotaError(e.message)) {
        Logger.log(`Gemini API quota exhausted (${currentModel}). Falling back to Exa for task prioritization.`);
        return prioritizeTasksWithExa(availability, tasks);
      }

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

function buildUpcomingEventsPlainText(events) {
  if (!events || !events.length) {
    return "No calendar events scheduled during your workday.";
  }

  return events.map(event => `- ${event.timeLabel}: ${event.title}`).join("\n");
}

function buildUpcomingEventsHtml(events) {
  if (!events || !events.length) {
    return "<p>No calendar events scheduled during your workday.</p>";
  }

  return [
    "<ul>",
    events.map(event => `<li><strong>${escapeHtml(event.timeLabel)}</strong>: ${escapeHtml(event.title)}</li>`).join("\n"),
    "</ul>"
  ].join("\n");
}

function buildFeaturedJobsPlainText(jobs) {
  if (!jobs || !jobs.length) {
    return "No product management jobs were found today from your five-star employers.";
  }

  return jobs.map((job, index) => {
    const parts = [
      `${index + 1}. ${job.title} - ${job.company}`,
      job.location ? `Location: ${job.location}` : "",
      job.source ? `Source: ${job.source}` : "",
      job.url || ""
    ].filter(Boolean);

    return parts.join("\n");
  }).join("\n\n");
}

function buildFeaturedJobsHtml(jobs) {
  if (!jobs || !jobs.length) {
    return "<p>No product management jobs were found today from your five-star employers.</p>";
  }

  return [
    "<ol>",
    jobs.map(job => {
      const detailParts = [
        job.location ? escapeHtml(job.location) : "",
        job.source ? escapeHtml(job.source) : ""
      ].filter(Boolean).join(" | ");

      const title = escapeHtml(job.title || "Product role");
      const company = escapeHtml(job.company || "Unknown employer");
      const safeUrl = escapeHtml(job.url || "");
      const details = detailParts ? `<div style=\"color:#555;\">${detailParts}</div>` : "";
      const link = safeUrl ? `<div><a href=\"${safeUrl}\">${safeUrl}</a></div>` : "";

      return `<li><strong>${title}</strong> - ${company}${details}${link}</li>`;
    }).join("\n"),
    "</ol>"
  ].join("\n");
}

function sendCoachEmail(recipient, aiContent, availability, featuredJobs) {
  const subjectDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM d, yyyy");
  const subject = `Your Daily Career Coach: ${subjectDate}`;
  const aiPlain = markdownToPlainText(aiContent).trim();
  const aiHtml = markdownToHtml(aiContent);
  const upcomingEventsPlain = buildUpcomingEventsPlainText(availability.upcomingEvents);
  const upcomingEventsHtml = buildUpcomingEventsHtml(availability.upcomingEvents);
  const featuredJobsPlain = buildFeaturedJobsPlainText(featuredJobs);
  const featuredJobsHtml = buildFeaturedJobsHtml(featuredJobs);

  const plainBody = [
    "Good morning!",
    "",
    `Today you have ${availability.totalMinutes} minutes of White Space across your calendars.`,
    "",
    "Coming up on your calendar:",
    upcomingEventsPlain,
    "",
    "The following are your proposed tasks:",
    "",
    aiPlain,
    "",
    "Product management jobs from your five-star employers:",
    "",
    featuredJobsPlain,
    "",
    "Go get 'em!"
  ].join("\n");

  const htmlBody = [
    "<div style=\"font-family:Arial,sans-serif;line-height:1.5;color:#222;\">",
    "<p><strong>Good morning!</strong></p>",
    `<p>Today you have <strong>${availability.totalMinutes} minutes</strong> of white space across your calendars.</p>`,
    "<p><strong>Coming up on your calendar:</strong></p>",
    upcomingEventsHtml,
    "<p>The following are your proposed tasks:</p>",
    aiHtml,
    "<p><strong>Product management jobs from your five-star employers:</strong></p>",
    featuredJobsHtml,
    "<p>Go get &#39;em!</p>",
    "</div>"
  ].join("\n");

  GmailApp.sendEmail(recipient, subject, plainBody, { htmlBody: htmlBody });
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

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][idxKR].toString().includes(keyResultText)) {
      sheet.getRange(i + 1, idxRun + 1).setValue(runningCount);
      Logger.log(`Spreadsheet Updated: row ${i + 1} for '${keyResultText}' set to ${runningCount}.`);
      return true;
    }
  }

  Logger.log(`Warning: Could not find a matching Key Result row for '${keyResultText}'.`);
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

function extractJsonArray(text) {
  if (!text) {
    return [];
  }

  const cleaned = String(text).trim().replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/i, "");
  const start = cleaned.indexOf("[");
  const end = cleaned.lastIndexOf("]");

  if (start === -1 || end === -1 || end < start) {
    return [];
  }

  try {
    const parsed = JSON.parse(cleaned.slice(start, end + 1));
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    Logger.log("Job JSON Parse Error: " + e.message);
    return [];
  }
}

function normalizeMotivationValue(value) {
  return String(value == null ? "" : value).trim().toLowerCase();
}

function isFiveStarEmployer(record) {
  const motivation = normalizeMotivationValue(record.fields?.["Max Motivation"]);
  return motivation === "5" || motivation === "five" || motivation === "5.0";
}

function getEmployerNameFromRecord(record) {
  const fields = record.fields || {};
  const candidateFields = ["Company", "Company Name", "Employer", "Employer Name", "Name"];

  for (let i = 0; i < candidateFields.length; i++) {
    const rawValue = fields[candidateFields[i]];
    const value = Array.isArray(rawValue) ? String(rawValue[0] || "").trim() : String(rawValue || "").trim();
    if (value) {
      return value;
    }
  }

  const firstStringField = Object.keys(fields).find(key => typeof fields[key] === "string" && String(fields[key]).trim());
  return firstStringField ? String(fields[firstStringField]).trim() : "";
}

function normalizeEmployerName(name) {
  return String(name || "").trim().toLowerCase();
}

function isTruthyFieldValue(value) {
  if (value === true) {
    return true;
  }

  const normalized = String(value == null ? "" : value).trim().toLowerCase();
  return normalized === "true" || normalized === "yes" || normalized === "y" || normalized === "1";
}

function getDestinationNameFromRecord(record) {
  const fields = record.fields || {};
  const candidateFields = ["City", "Metro Area", "Destination", "Location", "Name"];

  for (let i = 0; i < candidateFields.length; i++) {
    const rawValue = fields[candidateFields[i]];
    const value = Array.isArray(rawValue) ? String(rawValue[0] || "").trim() : String(rawValue || "").trim();
    if (value) {
      return value;
    }
  }

  const firstStringField = Object.keys(fields).find(key => typeof fields[key] === "string" && String(fields[key]).trim());
  return firstStringField ? String(fields[firstStringField]).trim() : "";
}

function getAcceptableDestinations() {
  const pat = PropertiesService.getScriptProperties().getProperty('AIRTABLE_PAT');
  const crmBaseId = PropertiesService.getScriptProperties().getProperty('AIRTABLE_BASE_ID_CRM');
  const citiesTable = PropertiesService.getScriptProperties().getProperty('AIRTABLE_TABLE_NAME_CITIES') || "Cities";

  if (!pat || !crmBaseId) {
    Logger.log("Acceptable destination lookup skipped: missing AIRTABLE_PAT or AIRTABLE_BASE_ID_CRM.");
    return [];
  }

  try {
    const cityRecords = fetchAirtableRecords(crmBaseId, citiesTable, pat);
    const destinationLookup = {};

    cityRecords
      .filter(record => isTruthyFieldValue(record.fields?.["Acceptable Destination?"]))
      .map(getDestinationNameFromRecord)
      .filter(Boolean)
      .forEach(name => {
        destinationLookup[name] = true;
      });

    return Object.keys(destinationLookup).sort();
  } catch (e) {
    Logger.log("Acceptable destination lookup error: " + e.message);
    return [];
  }
}

function getFiveStarEmployers() {
  const pat = PropertiesService.getScriptProperties().getProperty('AIRTABLE_PAT');
  const crmBaseId = PropertiesService.getScriptProperties().getProperty('AIRTABLE_BASE_ID_CRM');
  const companiesTable = PropertiesService.getScriptProperties().getProperty('AIRTABLE_TABLE_NAME_COMPANIES') || "Companies";

  if (!pat || !crmBaseId) {
    Logger.log("Five-star employer lookup skipped: missing AIRTABLE_PAT or AIRTABLE_BASE_ID_CRM.");
    return [];
  }

  try {
    const companyRecords = fetchAirtableRecords(crmBaseId, companiesTable, pat);
    const employerLookup = {};

    companyRecords
      .filter(isFiveStarEmployer)
      .map(getEmployerNameFromRecord)
      .filter(Boolean)
      .forEach(name => {
        employerLookup[name] = true;
      });

    return Object.keys(employerLookup).sort();
  } catch (e) {
    Logger.log("Five-star employer lookup error: " + e.message);
    return [];
  }
}

function normalizeJobPosting(job) {
  return {
    title: String(job.title || "").trim(),
    company: String(job.company || "").trim(),
    location: String(job.location || "").trim(),
    source: String(job.source || "").trim(),
    url: String(job.url || "").trim()
  };
}

function isValidFeaturedJob(job, employerLookup) {
  if (!job.title || !job.company || !job.url) {
    return false;
  }

  return Boolean(employerLookup[normalizeEmployerName(job.company)]);
}

function chunkArray(items, size) {
  const chunks = [];
  for (let i = 0; i < items.length; i += size) {
    chunks.push(items.slice(i, i + size));
  }
  return chunks;
}

function getDomainFromUrl(url) {
  const match = String(url || "").match(/^https?:\/\/([^\/?#]+)/i);
  return match ? match[1].toLowerCase() : "";
}

function looksLikeSpecificJobUrl(url) {
  const normalizedUrl = String(url || "").toLowerCase();
  if (!normalizedUrl) {
    return false;
  }

  const blockedPatterns = [
    /levels\.fyi\/jobs\/company\//,
    /linkedin\.com\/jobs\/[^\/]+-jobs/,
    /linkedin\.com\/jobs\/search/,
    /pitchmeai\.com\//,
    /indeed\.com\/q-/,
    /glassdoor\./
  ];

  if (blockedPatterns.some(pattern => pattern.test(normalizedUrl))) {
    return false;
  }

  const genericPatterns = [
    /\/careers\/?$/,
    /\/jobs\/?$/,
    /\/job-search\/?$/,
    /\/search\/?$/,
    /\/explore-careers\/?$/,
    /\/area-of-interest\//,
    /\/locations\/?$/,
    /\/teams\/?$/,
    /\/departments\/?$/
  ];

  if (genericPatterns.some(pattern => pattern.test(normalizedUrl))) {
    return false;
  }

  const jobDetailPatterns = [
    /\/jobs\/\d+/,
    /\/job\/[a-z0-9-]+/,
    /\/jobs\/[a-z0-9-]{6,}/,
    /\/positions?\//,
    /\/openings\//,
    /\/o\/[a-z0-9-]+/,
    /gh_jid=/,
    /lever\.co\/[^\/]+\/[a-z0-9-]+/,
    /smartrecruiters\.com\/[^\/]+\/[^\/]+/,
    /workdayjobs\.com\/[^\/]+\/job\//
  ];

  return jobDetailPatterns.some(pattern => pattern.test(normalizedUrl));
}

function isRelevantMidLevelProductRole(title, text) {
  const haystack = `${String(title || "")} ${String(text || "")}`.toLowerCase();

  if (!/(product manager|senior product manager|group product manager|lead product manager|technical product manager|growth product manager|platform product manager|ai product manager)/.test(haystack)) {
    return false;
  }

  return !/(principal product manager|director of product|head of product|vp product|vice president product|chief product officer|intern|internship|apm\b|associate product manager|staff product manager)/.test(haystack);
}

function isLikelyActiveJobPage(text) {
  const haystack = String(text || "").toLowerCase();
  if (!haystack) {
    return false;
  }

  const staleMarkers = [
    "job is no longer available",
    "this job is no longer available",
    "no longer accepting applications",
    "position has been filled",
    "position is filled",
    "job has expired",
    "job expired",
    "page not found",
    "404",
    "this posting is no longer available",
    "sorry this position is no longer posted",
    "this position is no longer posted",
    "we can't find the page",
    "the job you are looking for no longer exists",
    "not accepting applications",
    "this job is no longer posted"
  ];

  return !staleMarkers.some(marker => haystack.indexOf(marker) !== -1);
}

function extractLocationFromJobText(text, title) {
  const lines = String(text || "")
    .split(/\r?\n/)
    .map(line => line.trim())
    .filter(Boolean)
    .slice(0, 60);

  const titleLine = String(title || "").trim();
  if (titleLine) {
    lines.unshift(titleLine);
  }

  const locationLine = lines.find(line => /remote|hybrid|onsite|paris|france|london|berlin|munich|amsterdam|dublin|madrid|barcelona|new york|seattle|san francisco|boston|chicago|austin|los angeles|washington|philadelphia|atlanta|maple grove|india|bangalore|bengaluru|mumbai|delhi/i.test(line));
  return locationLine ? locationLine.slice(0, 120) : "";
}

function normalizeLocationText(text) {
  return String(text || "")
    .toLowerCase()
    .replace(/[^a-z0-9\s,/-]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function isInAcceptableDestination(locationText, acceptableDestinations) {
  if (!acceptableDestinations || !acceptableDestinations.length) {
    return true;
  }

  const normalizedLocation = normalizeLocationText(locationText);
  if (!normalizedLocation) {
    return false;
  }

  return acceptableDestinations.some(destination => {
    const normalizedDestination = normalizeLocationText(destination);
    if (!normalizedDestination) {
      return false;
    }

    return normalizedLocation.indexOf(normalizedDestination) !== -1;
  });
}

function inferLocationText(url, title, text) {
  return [title, text, url]
    .map(value => String(value || "").trim())
    .filter(Boolean)
    .join("\n");
}

function buildEmployerJobSearchQuery(employer) {
  return [
    `"${employer}"`,
    `("product manager" OR "senior product manager" OR "group product manager" OR "technical product manager" OR "growth product manager" OR "platform product manager")`,
    `-intern -internship -"associate product manager" -"principal product manager" -"director of product" -"head of product"`,
    `"apply" OR "job" OR "careers"`
  ].join(" ");
}

function searchSpecificJobsForEmployerWithExa(employer) {
  try {
    return fetchExaSearchResults(buildEmployerJobSearchQuery(employer), {
      type: "auto",
      numResults: 8,
      text: false,
      userLocation: "FR"
    });
  } catch (e) {
    Logger.log(`Exa employer search error for ${employer}: ` + e.message);
    return [];
  }
}

function validateJobCandidatesWithExa(candidateResults, acceptableDestinations) {
  const validJobs = [];
  const candidateUrls = candidateResults
    .filter(result => looksLikeSpecificJobUrl(result.url))
    .map(result => result.url);

  const resultLookup = candidateResults.reduce((lookup, result) => {
    lookup[result.url] = result;
    return lookup;
  }, {});

  chunkArray(candidateUrls, 10).forEach(urlBatch => {
    try {
      fetchExaContents(urlBatch).forEach(contentResult => {
        const sourceResult = resultLookup[contentResult.url] || {};
        const text = String(contentResult.text || "");
        const title = String(sourceResult.title || contentResult.title || "").trim();
        const employer = String(sourceResult.company || "").trim();

        if (!looksLikeSpecificJobUrl(contentResult.url)) {
          return;
        }

        if (!isLikelyActiveJobPage(text)) {
          return;
        }

        if (!isRelevantMidLevelProductRole(title, text)) {
          return;
        }

        const location = extractLocationFromJobText(text, title);
        const locationContext = inferLocationText(contentResult.url, title, text);
        if (!isInAcceptableDestination(location || locationContext, acceptableDestinations)) {
          return;
        }

        validJobs.push(normalizeJobPosting({
          title: title.replace(/\s+\|\s+.*$/, "").trim(),
          company: employer,
          location: location,
          source: getDomainFromUrl(contentResult.url),
          url: contentResult.url
        }));
      });
    } catch (e) {
      Logger.log("Exa content validation error: " + e.message);
    }
  });

  return validJobs;
}

function getFeaturedProductJobsWithExa(employers, employerLookup) {
  try {
    const uniqueJobs = {};
    const candidateResults = [];
    const acceptableDestinations = getAcceptableDestinations();

    employers.forEach(employer => {
      searchSpecificJobsForEmployerWithExa(employer).forEach(result => {
        candidateResults.push({
          title: String(result.title || "").trim(),
          company: employer,
          url: String(result.url || "").trim()
        });
      });
    });

    validateJobCandidatesWithExa(candidateResults, acceptableDestinations)
      .filter(job => isValidFeaturedJob(job, employerLookup))
      .forEach(job => {
        const key = [normalizeEmployerName(job.company), job.title, job.url].join("|");
        if (!uniqueJobs[key] && Object.keys(uniqueJobs).length < 14) {
          uniqueJobs[key] = job;
        }
      });

    return Object.keys(uniqueJobs).map(key => uniqueJobs[key]);
  } catch (e) {
    Logger.log("Featured jobs lookup error (Exa): " + e.message);
    return [];
  }
}

function getFeaturedProductJobs() {
  const employers = getFiveStarEmployers();
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const model = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || 'gemini-2.0-flash';
  const fallbackModel = PropertiesService.getScriptProperties().getProperty('GEMINI_FALLBACK_MODEL') || 'gemini-2.0-flash-lite';

  if (!employers.length) {
    Logger.log("Featured jobs skipped: no five-star employers found.");
    return [];
  }

  const modelsToTry = [model];
  if (fallbackModel && fallbackModel !== model) {
    modelsToTry.push(fallbackModel);
  }

  const employerLookup = employers.reduce((lookup, employer) => {
    lookup[normalizeEmployerName(employer)] = true;
    return lookup;
  }, {});

  if (!apiKey) {
    Logger.log("Featured jobs: missing GEMINI_API_KEY. Falling back to Exa.");
    return getFeaturedProductJobsWithExa(employers, employerLookup);
  }

  const employerList = employers.map(name => `- ${name}`).join("\n");
  const prompt = `
Find up to 14 current product management job openings from this employer list only:
${employerList}

Search the live web. Prefer official company career pages. If needed, use LinkedIn or Welcome to the Jungle.
Return JSON only as an array. Do not use markdown fences.

Each item must have:
- title
- company
- location
- source
- url

Rules:
- Include only roles that are clearly product management jobs or close variants such as Product Manager, Senior Product Manager, Group Product Manager, Principal Product Manager, Head of Product, Director of Product, Product Lead, Growth Product Manager, AI Product Manager, Platform Product Manager, Technical Product Manager.
- Company must exactly match one employer from the list.
- Use only live postings that appear open now.
- Prefer unique employers, but multiple jobs per employer are allowed if needed.
- Keep the total to 14 or fewer items.
`;

  const payload = {
    contents: [{
      parts: [{ text: prompt }]
    }],
    tools: [{ google_search: {} }],
    generationConfig: {
      temperature: 0.2,
      topK: 20,
      topP: 0.9,
      maxOutputTokens: 4096
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  for (let i = 0; i < modelsToTry.length; i++) {
    const currentModel = modelsToTry[i];

    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${currentModel}:generateContent?key=${apiKey}`;
      const response = UrlFetchApp.fetch(url, options);

      if (response.getResponseCode() !== 200) {
        throw new Error(`HTTP ${response.getResponseCode()}: ${response.getContentText()}`);
      }

      const result = JSON.parse(response.getContentText());
      const aiText = result.candidates?.[0]?.content?.parts
        ?.map(part => part.text || "")
        .join("\n")
        .trim();

      const uniqueJobs = {};

      extractJsonArray(aiText)
        .map(normalizeJobPosting)
        .filter(job => isValidFeaturedJob(job, employerLookup))
        .forEach(job => {
          const key = [normalizeEmployerName(job.company), job.title, job.url].join("|");
          if (!uniqueJobs[key] && Object.keys(uniqueJobs).length < 14) {
            uniqueJobs[key] = job;
          }
        });

      if (Object.keys(uniqueJobs).length) {
        return Object.keys(uniqueJobs).map(key => uniqueJobs[key]);
      }
    } catch (e) {
      Logger.log(`Featured jobs lookup error (${currentModel}): ` + e.message);

       if (isGeminiQuotaError(e.message)) {
        Logger.log(`Gemini API quota exhausted (${currentModel}). Falling back to Exa for featured jobs.`);
        return getFeaturedProductJobsWithExa(employers, employerLookup);
      }
    }
  }

  return [];
}

function previewFeaturedProductJobs() {
  const employers = getFiveStarEmployers();
  const employerLookup = employers.reduce((lookup, employer) => {
    lookup[normalizeEmployerName(employer)] = true;
    return lookup;
  }, {});
  const jobs = getFeaturedProductJobsWithExa(employers, employerLookup);
  const output = buildFeaturedJobsPlainText(jobs);

  Logger.log(output);
  return output;
}

function previewAcceptableDestinations() {
  const destinations = getAcceptableDestinations();
  const output = destinations.join("\n");
  Logger.log(output);
  return output;
}

function updateInterviewOKR(config) {
  const pat = PropertiesService.getScriptProperties().getProperty('AIRTABLE_PAT');
  const baseId = PropertiesService.getScriptProperties().getProperty('AIRTABLE_BASE_ID'); 
  const tableName = PropertiesService.getScriptProperties().getProperty('AIRTABLE_TABLE_NAME') || "Responses";
  const fieldName = "Updated Response Modified This Month";
  const crmBaseId = PropertiesService.getScriptProperties().getProperty('AIRTABLE_BASE_ID_CRM');
  const crmMeetingsTable = PropertiesService.getScriptProperties().getProperty('AIRTABLE_TABLE_NAME_MEETINGS') || "Meetings";
  const crmJobsTable = PropertiesService.getScriptProperties().getProperty('AIRTABLE_TABLE_NAME_JOBS') || "Jobs";
  
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
    if (response.getResponseCode() !== 200) {
      throw new Error(`HTTP ${response.getResponseCode()}: ${response.getContentText()}`);
    }
    const data = JSON.parse(response.getContentText());
    const allResponseRecords = fetchAirtableRecords(baseId, tableName, pat);
    
    // Total count of records marked 'True' this month
    const count = data.records ? data.records.length : 0;
    const responseCmSum = allResponseRecords.reduce((total, record) => {
      return total + (Number(record.fields?.Response_CM) || 0);
    }, 0);
    Logger.log(`Airtable Sync: Found ${count} updated responses. Response_CM sum across all records=${responseCmSum}.`);

    const sheet = getCurrentOKRSheet(config);
    if (!sheet) {
      return;
    }

    updateRunningCountForKeyResult(sheet, "responses to common interview questions", count);
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

    Logger.log(`Airtable CRM Sync: Found ${meetingRecords.length} meetings. Clinician sum=${clinicianInterviewSum}, Practice sum=${practiceInterviewSum}.`);

    updateRunningCountForKeyResult(sheet, "Establish contact with active clinicians", clinicianInterviewSum);
    updateRunningCountForKeyResult(sheet, "Practice Interviews (case + behavioral ideally)", practiceInterviewSum);

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

    Logger.log(`Airtable CRM Jobs Sync: Found ${jobRecords.length} jobs. Referral apps=${referralApplicationsSum}, PM apps=${productApplicationsSum}, PM jobs listed=${jobsListedSum}.`);

    updateRunningCountForKeyResult(sheet, "Apply to PM jobs that included a referral", referralApplicationsSum);
    updateRunningCountForKeyResult(sheet, "Apply to PM jobs", productApplicationsSum);
    updateRunningCountForKeyResult(sheet, "Find and list PM Jobs", jobsListedSum);

  } catch (e) {
    Logger.log("Airtable Sync Error: " + e.message);
  }
}
