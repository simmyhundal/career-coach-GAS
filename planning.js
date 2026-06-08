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

  let busySlots = [];
  for (let id in response.calendars) {
    busySlots = busySlots.concat(response.calendars[id].busy);
  }
  busySlots.sort((a, b) => new Date(a.start) - new Date(b.start));

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

  const okrSummary = tasks.map(t =>
    t.isDaily
      ? `- ${t.name} [Daily task, include unless completed]`
      : `- ${t.name} (${t.effort} mins remaining)`
  ).join('\n');

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
