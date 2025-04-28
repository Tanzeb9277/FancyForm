function doGet() {
  return HtmlService.createHtmlOutputFromFile("Form")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle("Submit and Search")
}

function submitQuery(submissionData) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Query Data")
  if (!sheet) {
    throw new Error("Sheet 'Query Data' not found.")
  }

  const email = String(Session.getActiveUser().getEmail() || "anonymous").trim()
  const query = String(submissionData.query).trim()

  // Get all data in the sheet
  const data = sheet.getDataRange().getValues()

  // Check if the email and query combination already exists
  for (let i = 0; i < data.length; i++) {
    const rowEmail = String(data[i][2]).trim() // Assuming email is in Column C (index 2)
    const rowQuery = String(data[i][3]).trim() // Assuming query is in Column D (index 3)

    if (rowEmail === email && rowQuery === query) {
      // Update the entry time for the existing row
      const now = new Date()
      const date = Utilities.formatDate(now, "America/New_York", "yyyy-MM-dd")
      const time = Utilities.formatDate(now, "America/New_York", "HH:mm:ss")
      
      // Update the date and time in the existing row
      sheet.getRange(i + 1, 1).setValue(date) // Update date in Column A
      sheet.getRange(i + 1, 2).setValue(time) // Update time in Column B
      
      return {
        submitted: true,
        message: "Existing query updated with new timestamp.",
        query: query,
      }
    }
  }

  // If we get here, no duplicate was found, so append the new data
  const now = new Date()
  const date = Utilities.formatDate(now, "America/New_York", "yyyy-MM-dd")
  const time = Utilities.formatDate(now, "America/New_York", "HH:mm:ss")
  sheet.appendRow([date, time, email, query])
  return {
    submitted: true,
    message: "Query submitted successfully!",
    query: query,
  }
}

function findMatchingNamesFrontend(searchValue = "Bobs Burger") {
  // Get the spreadsheet and the specific sheet by name
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "query Data"; // Specify the sheet name
  const sheet = ss.getSheetByName(sheetName);

  // Get the current user's email
  const currentUserEmail = Session.getActiveUser().getEmail();

  // Check if the sheet exists
  if (!sheet) {
    return JSON.stringify({ error: `Sheet "${sheetName}" not found.` });
  }

  // Define column indices
  const SEARCH_COLUMN = 4;
  const NAMES_COLUMN = 6;
  const TIME_COLUMN = 2;
  const EMAIL_COLUMN = 3; // Column C for email

  const data = sheet.getDataRange().getValues();
  const matchingEntries = []; // Changed from matchingNames to matchingEntries
  const trimmedSearchValue = String(searchValue).trim().toLowerCase();
  const now = new Date();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    // Skip if the email matches the current user's email
    if (row[EMAIL_COLUMN - 1] === currentUserEmail) {
      continue;
    }
    
    if (
      row.length > SEARCH_COLUMN - 1 &&
      row[SEARCH_COLUMN - 1] !== undefined &&
      String(row[SEARCH_COLUMN - 1]).trim().toLowerCase() === trimmedSearchValue
    ) {
      let timeValue = null;
      if (typeof row[TIME_COLUMN - 1] === 'string') {
        timeValue = row[TIME_COLUMN - 1];
      } else if (row[TIME_COLUMN - 1] instanceof Date) {
        // If it's interpreted as a Date, format it back to HH:mm:ss
        timeValue = Utilities.formatDate(row[TIME_COLUMN - 1], Session.getTimeZone(), "HH:mm:ss");
      }

      if (timeValue) {
        const parts = timeValue.split(':');
        if (parts.length === 3) {
          const entryHour = parseInt(parts[0], 10);
          const entryMinute = parseInt(parts[1], 10);
          const entrySecond = parseInt(parts[2], 10);

          const currentHour = now.getHours();
          const currentMinute = now.getMinutes();
          const currentSecond = now.getSeconds();

          const entryTimeInSeconds = entryHour * 3600 + entryMinute * 60 + entrySecond;
          const currentTimeInSeconds = currentHour * 3600 + currentMinute * 60 + currentSecond;

          const timeDifference = Math.abs(currentTimeInSeconds - entryTimeInSeconds);

          if (timeDifference <= 2400) {
            if (row.length > NAMES_COLUMN - 1 && row[NAMES_COLUMN - 1] !== undefined) {
              matchingEntries.push({
                name: row[NAMES_COLUMN - 1],
                email: row[EMAIL_COLUMN - 1]
              });
            }
          }
          Logger.log(`Search: ${trimmedSearchValue}, Entry Time (Parsed): ${timeValue}, Diff: ${timeDifference}`);
        } else {
          Logger.log(`Warning: Invalid time format in Column B: ${timeValue}`);
        }
      } else {
        Logger.log(`Skipping row ${i + 1}: Could not extract time from Column B.`);
        Logger.log(`Value in Column B: ${row[TIME_COLUMN - 1]}`);
      }
    }
  }

  return JSON.stringify(matchingEntries);
}
