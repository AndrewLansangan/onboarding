/**
 * The `timeTrackerRecap` function updates the "Notion Database (sync) - Mandates" Google Sheet with data from the
 * "All Recap - Current and Archived" (Google Sheet).
 *  Notion Database (sync) - Mandates (Google Sheet) : https://docs.google.com/spreadsheets/d/1uqCK0JDHKkuzDEfOlBcSqJCsVxHjCRouOW5F8hJZRUA/edit?gid=2131663677#gid=2131663677
 *  All Recap - Current and Archived" (Google Sheet) : https://docs.google.com/spreadsheets/d/1jnXb0JTCy6C0ORsGkepJBzQbPYIZHpBWO0gqvQVOvX4/edit?gid=0#gid=0
 *  It matches the "Greybox ID" from the target sheet with the "Recap"
 * name in the external sheet and retrieves the following columns: 'Error Detection', 'Last Update', 'Start Date',
 * and 'Hours (decimal)'. The data is then written into columns N to Q of the target sheet, ensuring accurate
 * and up-to-date information is reflected for each corresponding entry.
 */


function timeTrackerRecap() {
    logToSlack(
        "📢 Starting execution of \`TimeTrackerRecap\` script"
    );
    var externalSpreadsheetId = "1jnXb0JTCy6C0ORsGkepJBzQbPYIZHpBWO0gqvQVOvX4";
    var externalSheetName = "All Recap - Current and Archived";
    var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Ensure that the relevant columns exist in the target sheet
    var headers = ['Error Detection', 'Last Update', 'Start Date', 'Hours (decimal)'];
    for (var i = 0; i < headers.length; i++) {
        targetSheet.getRange(1, 14 + i).setValue(headers[i]); // Starting at column N (14th column)
    }

    var externalSs = SpreadsheetApp.openById(externalSpreadsheetId);
    var externalSheet = externalSs.getSheetByName(externalSheetName);
    if (!externalSheet) {
        logToSlack("External sheet not found.");
        return;
    }

    var externalData = externalSheet.getRange("A2:G" + externalSheet.getLastRow()).getValues();

    // Create a map of names to their corresponding data
    var nameToDataMap = {};
    externalData.forEach(row => {
        var name = row[0] ? row[0].trim().toLowerCase() : ""; // Recap names are in column A, normalized to lowercase and trimmed
        if (name) {
            nameToDataMap[name] = {
                errorDetection: row[2], // Error Detection in column C
                lastUpdate: formatDateTime(row[3]),     // Last Update in column D, format to "DD/MM/YYYY HH:MM:SS"
                startDate: formatDateTime(row[4]),      // Start Date in column E, format to "DD/MM/YYYY HH:MM:SS"
                hoursDecimal: row[5]    // Hours (decimal) in column F
            };
        }
    });

    // Iterate over 'Greybox ID' in the target sheet to prepare data for updating
    var targetDataRange = targetSheet.getRange("A2:A" + targetSheet.getLastRow());
    var targetNames = targetDataRange.getValues();

    var dataToWrite = targetNames.map(row => {
        var name = row[0] ? row[0].trim().toLowerCase() : ""; // Normalize and trim the name
        var data = nameToDataMap[name] || { errorDetection: "", lastUpdate: "", startDate: "", hoursDecimal: "" };
        return [data.errorDetection, data.lastUpdate, data.startDate, data.hoursDecimal];
    });

    // Update columns N to Q with the extracted data
    targetSheet.getRange("N2:Q" + (1 + dataToWrite.length)).setValues(dataToWrite);
    logToSlack(
        "📢 Execution of \`TimeTrackerRecap\` script finished."
    );
}

function formatDateTime(dateValue) {
    if (!dateValue) return "";

    // Check if the dateValue is a Date object, convert it to the desired string format
    if (Object.prototype.toString.call(dateValue) === '[object Date]') {
        return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    }

    // If dateValue is already a string, assume it's in the correct format and return it as is
    if (typeof dateValue === 'string') {
        return dateValue;
    }

    return ""; // Return an empty string if dateValue is not valid
}
