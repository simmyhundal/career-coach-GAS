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

  // 2.1 Update OKR running counts from external sources
  updateFrenchProgress(config, sheet);    // French course hours from Calendar
  updateFrenchJournalOKR(config, sheet);  // French journal entry count from Google Doc
  updateInterviewOKR(config);             // Interview prep counts from Airtable

  // 2.2 Calculate White Space
  const availability = getDailyAvailability(config);

  // 3. Get OKRs
  const tasks = getActiveOKRs(config, sheet);

  // 4. Get AI Recommendation
  const aiContent = prioritizeTasksWithAI(availability, tasks);

  // 4.1 Pull the five most recent unsent Airtable jobs
  const featuredJobs = getFeaturedProductJobs();

  // 5. Send the Email
  const okrUrl = `https://docs.google.com/spreadsheets/d/${config.OKR_SHEET_ID}`;
  sendCoachEmail(config.USER_EMAIL, aiContent, availability, featuredJobs, okrUrl);
}
