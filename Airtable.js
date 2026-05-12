/**
 * TEST FUNCTION: Run this to see what Airtable is actually sending.
 */
function debugAirtableFetch() {
  const pat = PropertiesService.getScriptProperties().getProperty('AIRTABLE_PAT');
  const baseId = PropertiesService.getScriptProperties().getProperty('AIRTABLE_BASE_ID'); 
  const tableName = "Responses";
  const fieldName = "Updated Response Modified This Month";
  
  // TEST VERSION 1: Use TRUE instead of 1 (Often required for Checkboxes/Booleans)
  const filter = `({${fieldName}} = TRUE())`; 
  
  const url = `https://api.airtable.com/v0/${baseId}/${encodeURIComponent(tableName)}?filterByFormula=${encodeURIComponent(filter)}`;
  
  const options = {
    "method": "get",
    "headers": { "Authorization": "Bearer " + pat },
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  
  Logger.log("--- DEBUG START ---");
  Logger.log("Status Code: " + response.getResponseCode());
  
  if (data.records && data.records.length > 0) {
    Logger.log("SUCCESS: Found " + data.records.length + " records.");
    Logger.log("Sample Record Field Value: " + data.records[0].fields[fieldName]);
    Logger.log("Full Record 0 Data: " + JSON.stringify(data.records[0]));
  } else {
    Logger.log("FAILURE: Found 0 records. The filter failed or no records match.");
    Logger.log("Raw API Response: " + response.getContentText());
  }
  Logger.log("--- DEBUG END ---");
}