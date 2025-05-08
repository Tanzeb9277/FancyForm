function doGet() {
  return HtmlService.createHtmlOutputFromFile("Form")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle("FloorGPT")
}

function submitQuery(submissionData) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QueryData")
  if (!sheet) {
    throw new Error("Sheet 'Query Data' not found.")
  }

  const email = String(Session.getActiveUser().getEmail() || "anonymous").trim()
  let query = String(submissionData.query).trim()
  
  // Check for different types of queries and route accordingly
  if (query.includes("Raters' Comments")) {
    query = extractTargetSentence(query)
  } else if (query.includes("Target Sentence from the above response:")) {
    query = extractTextBetweenPhrases(query)
  }
  // If neither condition is met, keep the query as is

  if (!query) {
    return {
      submitted: false,
      message: "Could not extract valid query from the input.",
      query: null,
      qaRating: null,
    }
  }

  // Remove all line breaks and extra spaces
  query = query
    .replace(/[\r\n]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()

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
      // Check QA rating for the query
      const qaRatingResult = JSON.parse(checkQARating(query))
      return {
        submitted: true,
        message: "Existing query updated with new timestamp.",
        query: query,
        qaRating: qaRatingResult,
      }
    }
  }

  // If we get here, no duplicate was found, so append the new data
  const now = new Date()
  const date = Utilities.formatDate(now, "America/New_York", "yyyy-MM-dd")
  const time = Utilities.formatDate(now, "America/New_York", "HH:mm:ss")
  sheet.appendRow([date, time, email, query])

  // Check QA rating for the new query
  const qaRatingResult = JSON.parse(checkQARating(query))
  return {
    submitted: true,
    message: "Query submitted successfully!",
    query: query,
    qaRating: qaRatingResult,
  }
}

function findMatchingNamesFrontend(searchValue = "Bobs Burger") {
  // Get the spreadsheet and the specific sheet by name
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheetName = "queryData" // Specify the sheet name
  const sheet = ss.getSheetByName(sheetName)

  // Get the current user's email
  const currentUserEmail = Session.getActiveUser().getEmail()

  // Check if the sheet exists
  if (!sheet) {
    return JSON.stringify({ error: `Sheet "${sheetName}" not found.` })
  }

  // Define column indices
  const SEARCH_COLUMN = 4
  const NAMES_COLUMN = 8
  const TIME_COLUMN = 2
  const EMAIL_COLUMN = 3 // Column C for email

  const data = sheet.getDataRange().getValues()
  const matchingEntries = []
  const trimmedSearchValue = String(searchValue).trim().toLowerCase()
  const now = new Date()

  for (let i = 0; i < data.length; i++) {
    const row = data[i]
    // Skip if the email matches the current user's email
    if (row[EMAIL_COLUMN - 1] === currentUserEmail) {
      continue
    }

    if (
      row.length > SEARCH_COLUMN - 1 &&
      row[SEARCH_COLUMN - 1] !== undefined &&
      String(row[SEARCH_COLUMN - 1])
        .trim()
        .toLowerCase() === trimmedSearchValue
    ) {
      let entryTime
      const timeValue = row[TIME_COLUMN - 1]

      if (timeValue instanceof Date) {
        entryTime = timeValue
      } else if (typeof timeValue === "string" && timeValue.trim() !== "") {
        entryTime = new Date(timeValue)
        if (isNaN(entryTime.getTime())) {
          Logger.log(
            `Warning: Could not parse date/time string in Column B: ${timeValue}`,
          )
          continue // Skip if parsing fails
        }
      } else {
        Logger.log(
          `Skipping row ${i + 1}: Invalid date/time value in Column B.`,
        )
        Logger.log(`Value in Column B: ${timeValue}`)
        continue // Skip if not a Date or a valid string
      }

      if (entryTime) {
        const timeDifferenceMillis = now.getTime() - entryTime.getTime()
        const secondsAgo = Math.round(timeDifferenceMillis / 1000)
        const minutesAgo = Math.round(secondsAgo / 60)
        const hoursAgo = Math.round(minutesAgo / 60)
        const daysAgo = Math.round(hoursAgo / 24)

        let relativeTime
        if (daysAgo >= 1) {
          relativeTime = `${daysAgo} day${daysAgo === 1 ? "" : "s"} ago`
        } else if (hoursAgo >= 1) {
          relativeTime = `${hoursAgo} hour${hoursAgo === 1 ? "" : "s"} ago`
        } else if (minutesAgo >= 1) {
          relativeTime = `${minutesAgo} minute${minutesAgo === 1 ? "" : "s"} ago`
        } else {
          relativeTime = `${secondsAgo} second${secondsAgo === 1 ? "" : "s"} ago`
        }

        if (
          row.length > NAMES_COLUMN - 1 &&
          row[NAMES_COLUMN - 1] !== undefined
        ) {
          matchingEntries.push({
            name: row[NAMES_COLUMN - 1],
            email: row[EMAIL_COLUMN - 1],
            timestamp: relativeTime,
          })
        }
        Logger.log(
          `Search: ${trimmedSearchValue}, Entry Time: ${entryTime.toLocaleString()}, Relative: ${relativeTime}`,
        )
      }
    }
  }

  return JSON.stringify(matchingEntries)
}

function flagTask(flagData) {
  // Get the spreadsheet and the specific sheet by name
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheetName = "queryData"
  const sheet = ss.getSheetByName(sheetName)

  if (!sheet) {
    return JSON.stringify({ error: `Sheet "${sheetName}" not found.` })
  }

  // Define column indices
  const SEARCH_COLUMN = 4 // Column D for search query
  const EMAIL_COLUMN = 3 // Column C for email
  const ID_COLUMN = 5 // Column E for ID
  const FLAG_COLUMN = 6 // Column F for flag

  const data = sheet.getDataRange().getValues()
  const trimmedTarget = String(flagData.targetSentence).trim().toLowerCase()
  const trimmedEmail = String(Session.getActiveUser().getEmail())
    .trim()
    .toLowerCase()

  for (let i = 0; i < data.length; i++) {
    const row = data[i]
    const rowSearch = String(row[SEARCH_COLUMN - 1])
      .trim()
      .toLowerCase()
    const rowEmail = String(row[EMAIL_COLUMN - 1])
      .trim()
      .toLowerCase()

    if (rowSearch === trimmedTarget && rowEmail === trimmedEmail) {
      // Update the ID and flag columns
      sheet.getRange(i + 1, ID_COLUMN).setValue(flagData.taskId)
      sheet.getRange(i + 1, FLAG_COLUMN).setValue(flagData.flag)
      return JSON.stringify({
        success: true,
        message: "Task flagged successfully",
        row: i + 1,
      })
    }
  }

  return JSON.stringify({
    error: "No matching row found with the given target sentence and email",
  })
}

function checkQARating(targetSentence = "Sponge Bob") {
  // Get the spreadsheet and the specific sheet by name
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheetName = "QARatings"
  const sheet = ss.getSheetByName(sheetName)

  if (!sheet) {
    return JSON.stringify({ error: `Sheet "${sheetName}" not found.` })
  }

  // Define column indices
  const TARGET_COLUMN = 1 // Column A for target sentence
  const RATING_COLUMN = 4 // Column D for rating
  const REASONING_COLUMN = 5 // Column E for QA reasoning

  const data = sheet.getDataRange().getValues()
  const trimmedTarget = String(targetSentence).trim().toLowerCase()

  for (let i = 0; i < data.length; i++) {
    const row = data[i]
    const rowTarget = String(row[TARGET_COLUMN - 1])
      .trim()
      .toLowerCase()

    if (rowTarget === trimmedTarget) {
      const rating = row[RATING_COLUMN - 1]
      const reasoning = row[REASONING_COLUMN - 1]

      if (!rating || rating === "") {
        return JSON.stringify({
          status: "awaiting",
          message: "Awaiting QA Rating",
        })
      } else {
        Logger.log(
          JSON.stringify({
            status: "rated",
            rating: rating,
            reasoning: reasoning || "",
          }),
        )
        return JSON.stringify({
          status: "rated",
          rating: rating,
          reasoning: reasoning || "",
        })
      }
    }
  }

  return JSON.stringify({
    error: "No matching target sentence found in QARatings sheet",
  })
}

function extractTextBetweenPhrases(
  text,
  startPhrase = "Question A",
  endPhrase = "A. To what extent is the Target Sentence",
) {
  // Convert to lowercase for case-insensitive search
  const lowerText = text.toLowerCase()
  const lowerStart = startPhrase.toLowerCase()
  const lowerEnd = endPhrase.toLowerCase()
  // Find the start and end positions
  const startIndex = lowerText.indexOf(lowerStart)
  const endIndex = lowerText.indexOf(lowerEnd)
  // If either phrase is not found, return null
  if (startIndex === -1 || endIndex === -1) {
    return null
  }
  // Extract the text between the phrases
  // Add the length of the start phrase to get the position after it
  const startPosition = startIndex + startPhrase.length
  const extractedText = text.substring(startPosition, endIndex).trim()
  return extractedText
}

function extractTargetSentence(text) {
  const startMarker = "Target Sentence\nRaters' Comments\nTask Questions\n\n\n"
  const startIndex = text.indexOf(startMarker)
  
  if (startIndex === -1) return null

  // Get everything after the start marker
  const afterStart = text.slice(startIndex + startMarker.length)
  
  // Find the next triple line break
  const endIndex = afterStart.indexOf("\n\n\n")
  
  if (endIndex === -1) return null
  
  // Extract the text between start marker and triple line break
  return afterStart.slice(0, endIndex).trim()
}
