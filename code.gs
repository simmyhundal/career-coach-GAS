// MASTER FUNCTION: Run this to test everything
function runDailyCoach() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Sheet2"); // Double check if it's "Sheet2" or " Sheet2"
  
  if (!configSheet) {
    throw new Error("Could not find tab named 'Sheet2'. Check for leading/trailing spaces.");
  }

  // 1. Map values by searching for the Key name (more robust than hardcoding C2, C3)
  const configData = configSheet.getRange("B2:C10").getValues();
  const config = {};
  configData.forEach(row => {
    config[row[0]] = row[1];
  });

  // Verify we have the data
  Logger.log("Config loaded: " + JSON.stringify(config));
  
  if (!config.WORK_START || !config.WORK_END) {
    throw new Error("Could not find WORK_START or WORK_END in Column B. Check spelling!");
  }

  // 2. Calculate White Space
  const availability = getDailyAvailability(config);
  
  // 3. Get OKRs
  const tasks = getActiveOKRs(config);
  
  // 4. Get AI Recommendation
  const aiContent = prioritizeTasksWithAI(config, availability, tasks);
  
  // 5. Send the Email
  sendCoachEmail(config.USER_EMAIL, aiContent, availability);
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

function getActiveOKRs(config) {
  // Use uppercase to match your Config log: OKR_SHEET_ID and OKR_TAB_NAME
  const okrFile = SpreadsheetApp.openById(config.OKR_SHEET_ID);
  const sheet = okrFile.getSheetByName(config.OKR_TAB_NAME);
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
  const okrSummary = tasks.map(t => `- ${t.name} (Requires: ${t.effort} mins)`).join('\n');
  
  const prompt = `
    Today I have ${availability.totalMinutes} minutes of total free time. 
    My largest contiguous focus block is ${availability.largestBlock} minutes.
    
    Based on my OKRs for March, please pick the best 2-4 tasks to tackle today:
    ${okrSummary}
    
    Instructions:
    1. Do not exceed a total of ${availability.totalMinutes * 0.8} minutes (80% capacity).
    2. Prioritize tasks that fit into my largest block.
    3. Format the response as a bulleted "Daily Action Plan".
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
