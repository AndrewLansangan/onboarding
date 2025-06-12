/**
 * Reads user data from a Google Sheet and groups users by their active team names.
 * Filters out users with invalid mandate statuses and already processed entries based on `lastRunTime`.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to read user data from.
 * @return {Map<string, string[]>} - A map of team names to arrays of user emails.
 */
function readUsers(sheet) {
    Logger.log(`📄 Reading users from sheet: ${sheet.getName()}`);

    const data = sheet.getDataRange().getValues();
    const header = data.shift();

    const emailIndex = header.indexOf('Email (Org)');
    const statusIndex = header.indexOf('Mandate (Status)');
    const teamIndex = header.indexOf('Team (Current)');
    const lastUpdateIndex = header.indexOf('Last Update');

    if ([emailIndex, statusIndex, teamIndex, lastUpdateIndex].includes(-1)) {
        Logger.log('❌ One or more required columns are missing.');
        return new Map();
    }

    const invalidStatuses = new Set(['To Verify', 'Completed', 'Archived']);
    const groupsMap = new Map();
    let processedRows = 0;
    let skippedRows = 0;

    for (const row of data) {
        const email = row[emailIndex];
        const status = row[statusIndex];
        const usergroupCell = row[teamIndex];
        const lastUpdate = row[lastUpdateIndex];
        const parsedDate = parseCustomDate(lastUpdate);

        if (!email || !usergroupCell || isNaN(parsedDate)) {
            Logger.log(`⚠️ Skipping row due to missing or invalid data: ${JSON.stringify(row)}`);
            skippedRows++;
            continue;
        }

        if (parsedDate < lastRunTime) {
            Logger.log(`⏩ Skipping ${email} - already processed.`);
            continue;
        }

        if (!invalidStatuses.has(status)) {
            const usergroups = usergroupCell.split(',').map((group) => group.trim());

            Logger.log(`✅ Processing user ${email} with ${usergroups.length} team(s).`);
            processedRows++;

            usergroups.forEach((userGroup) => {
                if (!userGroup || userGroup.toLowerCase().includes('admin')) {
                    Logger.log(`⚠️ Skipping team '${userGroup}' (admin or empty).`);
                    return;
                }

                const existing = groupsMap.get(userGroup) || [];
                existing.push(email);
                groupsMap.set(userGroup, existing);
                Logger.log(`👥 Added ${email} to group '${userGroup}'.`);
            });
        } else {
            Logger.log(`🚫 Skipping ${email} due to status: ${status}`);
            skippedRows++;
        }
    }

    Logger.log(`✅ Finished: ${processedRows} processed, ${skippedRows} skipped, ${groupsMap.size} group(s) created.`);
    return groupsMap;
}

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
    const externalSpreadsheetId = "1jnXb0JTCy6C0ORsGkepJBzQbPYIZHpBWO0gqvQVOvX4";
    const externalSheetName = "All Recap - Current and Archived";
    const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Ensure that the relevant columns exist in the target sheet
    const headers = ['Error Detection', 'Last Update', 'Start Date', 'Hours (decimal)'];
    for (let i = 0; i < headers.length; i++) {
        targetSheet.getRange(1, 14 + i).setValue(headers[i]); // Starting at column N (14th column)
    }

    const externalSs = SpreadsheetApp.openById(externalSpreadsheetId);
    const externalSheet = externalSs.getSheetByName(externalSheetName);
    if (!externalSheet) {
        logToSlack("External sheet not found.");
        return;
    }

    const externalData = externalSheet.getRange("A2:G" + externalSheet.getLastRow()).getValues();

    // Create a map of names to their corresponding data
    const nameToDataMap = {};
    externalData.forEach(row => {
        const name = row[0] ? row[0].trim().toLowerCase() : ""; // Recap names are in column A, normalized to lowercase and trimmed
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
    const targetDataRange = targetSheet.getRange("A2:A" + targetSheet.getLastRow());
    const targetNames = targetDataRange.getValues();

    const dataToWrite = targetNames.map(row => {
        const name = row[0] ? row[0].trim().toLowerCase() : ""; // Normalize and trim the name
        const data = nameToDataMap[name] || {
            errorDetection: "",
            lastUpdate: "",
            startDate: "",
            hoursDecimal: ""
        };
        return [data.errorDetection, data.lastUpdate, data.startDate, data.hoursDecimal];
    });

    // Update columns N to Q with the extracted data
    targetSheet.getRange("N2:Q" + (1 + dataToWrite.length)).setValues(dataToWrite);
    logToSlack(
        "📢 Execution of \`TimeTrackerRecap\` script finished."
    );
}

/**
 * Loads structured data from a Google Sheet, splitting header and data rows.
 *
 * @param {string} sheetName - The name of the sheet tab.
 * @returns {{ header: string[], rows: any[][] }} The header row and all subsequent data rows.
 * @throws Will throw an error if the sheet is missing or has no rows.
 */
function loadSheetData(sheetName) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error(`❌ Sheet "${sheetName}" not found.`);

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) {
        throw new Error(`❌ Sheet "${sheetName}" has no data rows (must include header and at least one row).`);
    }

    const [header, ...rows] = values;
    return { header, rows };
}
