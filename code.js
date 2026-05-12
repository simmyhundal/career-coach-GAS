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

  // 4.1 Pull the five most recent unsent Airtable jobs
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
    Logger.log("Gemini API Error: Missing GEMINI_API_KEY script property. Using fallback plan.");
    return buildFallbackPlan(availability, tasks);
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
        Logger.log(`Gemini API quota exhausted (${currentModel}). Using fallback plan for task prioritization.`);
        return buildFallbackPlan(availability, tasks);
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

function buildTaskPlanHtml(aiContent) {
  const TIMED_CARD = [
    'border-left: 3px solid #2D7DD2;',
    'padding: 12px 16px;',
    'margin-bottom: 10px;',
    'background-color: #f5f9ff;',
    'border-radius: 0 6px 6px 0;'
  ].join(' ');

  const DAILY_CARD = [
    'border-left: 3px solid #1E8A5E;',
    'padding: 12px 16px;',
    'margin-bottom: 10px;',
    'background-color: #f3fbf7;',
    'border-radius: 0 6px 6px 0;'
  ].join(' ');

  const DAILY_BADGE = [
    'display: inline-block;',
    'font-size: 10px;',
    'font-weight: 700;',
    'text-transform: uppercase;',
    'letter-spacing: 0.5px;',
    'background-color: #1E8A5E;',
    'color: #ffffff;',
    'padding: 2px 7px;',
    'border-radius: 10px;',
    'margin-left: 8px;',
    'vertical-align: middle;'
  ].join(' ');

  const lines = (aiContent || "").split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const html = [];
  let inList = false;

  lines.forEach(line => {
    const headingMatch = line.match(/^###\s+(.*)$/);
    const bulletMatch = line.match(/^[-*]\s+(.*)$/);

    if (headingMatch) {
      if (inList) { html.push("</ul>"); inList = false; }
      const content = headingMatch[1].trim();
      const timedMatch = content.match(/^(.+?):\s*(\d+)\s+mins?$/i);

      if (timedMatch) {
        const name = escapeHtml(timedMatch[1].trim());
        const mins = timedMatch[2];
        html.push(
          `<div style="${TIMED_CARD}">` +
          `<div style="font-size: 14px; font-weight: 600; color: #1A2744;">${name}</div>` +
          `<div style="font-size: 12px; color: #8695A3; margin-top: 4px;">${mins} minutes</div>` +
          `</div>`
        );
      } else {
        const name = escapeHtml(content);
        html.push(
          `<div style="${DAILY_CARD}">` +
          `<div style="font-size: 14px; font-weight: 600; color: #1A2744;">${name}` +
          `<span style="${DAILY_BADGE}">Daily</span></div>` +
          `</div>`
        );
      }
    } else if (bulletMatch) {
      if (!inList) {
        html.push('<ul style="margin: 8px 0 8px 4px; padding-left: 18px; color: #5D6D7E; font-size: 13px;">');
        inList = true;
      }
      html.push(`<li style="margin-bottom: 4px;">${escapeHtml(bulletMatch[1])}</li>`);
    } else {
      if (inList) { html.push("</ul>"); inList = false; }
      html.push(`<p style="font-size: 13px; color: #5D6D7E; margin: 6px 0;">${escapeHtml(line)}</p>`);
    }
  });

  if (inList) html.push("</ul>");

  return html.join("\n") ||
    '<p style="font-size: 13px; color: #8695A3; font-style: italic;">No tasks available today.</p>';
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
    return '<p style="font-size: 13px; color: #8695A3; font-style: italic;">No calendar events during your workday.</p>';
  }

  return events.map((event, i) => {
    const border = i < events.length - 1 ? 'border-bottom: 1px solid #f0f3f6;' : '';
    return (
      `<div style="padding: 9px 0; ${border}">` +
      `<table style="border-collapse: collapse; width: 100%;"><tr>` +
      `<td style="padding: 0; width: 160px; vertical-align: top; white-space: nowrap;">` +
      `<span style="font-size: 13px; color: #2D7DD2; font-weight: 500;">${escapeHtml(event.timeLabel)}</span>` +
      `</td>` +
      `<td style="padding: 0; vertical-align: top;">` +
      `<span style="font-size: 13px; color: #1A2744;">${escapeHtml(event.title)}</span>` +
      `</td>` +
      `</tr></table>` +
      `</div>`
    );
  }).join("\n");
}

function buildFeaturedJobsPlainText(jobs) {
  if (!jobs || !jobs.length) {
    return "No recent Airtable jobs were found today.";
  }

  return jobs.map((job, index) => {
    const parts = [
      `${index + 1}. ${job.title || "Untitled role"} - ${job.company || "Unknown company"}`,
      job.location ? `Location: ${job.location}` : "",
      job.postDate ? `App Post Date: ${job.postDate}` : "",
      job.url ? `Link: ${job.url}` : ""
    ].filter(Boolean);

    return parts.join("\n");
  }).join("\n\n");
}

function buildFeaturedJobsHtml(jobs) {
  if (!jobs || !jobs.length) {
    return '<p style="font-size: 13px; color: #8695A3; font-style: italic;">No recent job opportunities found.</p>';
  }

  return jobs.map((job, i) => {
    const border = i < jobs.length - 1 ? 'border-bottom: 1px solid #f0f3f6;' : '';
    const title = escapeHtml(job.title || "Untitled role");
    const company = escapeHtml(job.company || "Unknown company");
    const locationPart = job.location
      ? `<span style="color: #8695A3;"> &middot; ${escapeHtml(job.location)}</span>`
      : "";
    const postDatePart = job.postDate
      ? `<span style="color: #8695A3; font-size: 12px; margin-right: 10px;">Posted: ${escapeHtml(job.postDate)}</span>`
      : "";
    const linkPart = job.url
      ? `<a href="${escapeHtml(job.url)}" style="color: #2D7DD2; font-size: 12px; text-decoration: none;">View posting &rarr;</a>`
      : "";
    const meta = postDatePart || linkPart
      ? `<div style="margin-top: 5px;">${postDatePart}${linkPart}</div>`
      : "";

    return (
      `<div style="padding: 12px 0; ${border}">` +
      `<div style="font-size: 14px; font-weight: 600; color: #1A2744; margin-bottom: 3px;">${title}</div>` +
      `<div style="font-size: 13px; color: #5D6D7E;">${company}${locationPart}</div>` +
      meta +
      `</div>`
    );
  }).join("\n");
}

function buildEmailHtml(dateLabel, availability, aiContent, featuredJobs) {
  const S = {
    wrap: 'font-family: "Helvetica Neue", Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f5f7fa;',
    header: 'background-color: #1E3A5F; padding: 28px 32px 20px;',
    eyebrow: 'color: rgba(255,255,255,0.55); font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 2px; margin-bottom: 8px;',
    title: 'color: #ffffff; font-size: 26px; font-weight: 700; line-height: 1.2; margin-bottom: 4px;',
    date: 'color: rgba(255,255,255,0.65); font-size: 14px;',
    statsBar: 'background-color: #16304F; padding: 14px 32px;',
    statsTable: 'border-collapse: collapse; width: 100%;',
    statsCell: 'color: #ffffff; padding: 0; width: 50%;',
    statsNum: 'font-size: 24px; font-weight: 700;',
    statsLabel: 'font-size: 12px; color: rgba(255,255,255,0.6); margin-left: 5px;',
    body: 'padding: 20px 24px; background-color: #f5f7fa;',
    card: 'background-color: #ffffff; border-radius: 8px; padding: 20px 24px; margin-bottom: 16px; border: 1px solid #e2e8ef;',
    sectionLabel: 'font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.5px; color: #8695A3; margin-bottom: 14px;',
    tagline: 'text-align: center; padding: 6px 0 4px; color: #8695A3; font-size: 13px; font-style: italic;',
    footer: 'background-color: #1E3A5F; padding: 14px 32px; text-align: center;',
    footerText: 'color: rgba(255,255,255,0.4); font-size: 11px;'
  };

  return [
    `<div style="${S.wrap}">`,

    `<div style="${S.header}">`,
    `<div style="${S.eyebrow}">AI Career Coach</div>`,
    `<div style="${S.title}">Good morning!</div>`,
    `<div style="${S.date}">${escapeHtml(dateLabel)}</div>`,
    `</div>`,

    `<div style="${S.statsBar}">`,
    `<table style="${S.statsTable}"><tr>`,
    `<td style="${S.statsCell}">`,
    `<span style="${S.statsNum}">${availability.totalMinutes}</span>`,
    `<span style="${S.statsLabel}">min free today</span>`,
    `</td>`,
    `<td style="${S.statsCell}">`,
    `<span style="${S.statsNum}">${availability.largestBlock}</span>`,
    `<span style="${S.statsLabel}">min focus block</span>`,
    `</td>`,
    `</tr></table>`,
    `</div>`,

    `<div style="${S.body}">`,

    `<div style="${S.card}">`,
    `<div style="${S.sectionLabel}">Today&#39;s Schedule</div>`,
    buildUpcomingEventsHtml(availability.upcomingEvents),
    `</div>`,

    `<div style="${S.card}">`,
    `<div style="${S.sectionLabel}">Priority Tasks</div>`,
    buildTaskPlanHtml(aiContent),
    `</div>`,

    `<div style="${S.card}">`,
    `<div style="${S.sectionLabel}">Featured PM Roles</div>`,
    buildFeaturedJobsHtml(featuredJobs),
    `</div>`,

    `<div style="${S.tagline}">Go get &#39;em!</div>`,

    `</div>`,

    `<div style="${S.footer}">`,
    `<div style="${S.footerText}">AI Career Coach &middot; Google Apps Script</div>`,
    `</div>`,

    `</div>`
  ].join("\n");
}

function sendCoachEmail(recipient, aiContent, availability, featuredJobs) {
  const dateLabel = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy");
  const subject = `Daily Career Coach: ${dateLabel}`;
  const aiPlain = markdownToPlainText(aiContent).trim();
  const upcomingEventsPlain = buildUpcomingEventsPlainText(availability.upcomingEvents);
  const featuredJobsPlain = buildFeaturedJobsPlainText(featuredJobs);

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
    "Most recently created jobs from Airtable:",
    "",
    featuredJobsPlain,
    "",
    "Go get 'em!"
  ].join("\n");

  GmailApp.sendEmail(recipient, subject, plainBody, {
    htmlBody: buildEmailHtml(dateLabel, availability, aiContent, featuredJobs)
  });
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

  let updatedRows = 0;

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][idxKR].toString().includes(keyResultText)) {
      sheet.getRange(i + 1, idxRun + 1).setValue(runningCount);
      Logger.log(`Spreadsheet Updated: row ${i + 1} for '${keyResultText}' set to ${runningCount}.`);
      updatedRows++;
    }
  }

  if (updatedRows > 0) {
    return true;
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

function normalizeJobPosting(job) {
  return {
    title: String(job.title || "").trim(),
    company: String(job.company || "").trim(),
    location: String(job.location || "").trim(),
    postDate: String(job.postDate || "").trim(),
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
      .filter(job => job.title || job.company || job.url);
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
    company: getFirstFieldValue(fields, ["Company", "Company Name", "Employer", "Employer Name"]),
    location: getFirstFieldValue(fields, ["Location", "City", "Metro Area"]),
    postDate: getFirstFieldValue(fields, ["App Post Date"]),
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
    updateRunningCountForKeyResult(sheet, "Establish STAR Behavioral Interview Responses", count);
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
