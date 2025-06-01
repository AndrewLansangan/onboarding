/**
 * This script syncs data from the People Directory (Notion database) to the Notion Database (sync) - Mandates (Google Sheet).
 * People Directory (Notion database): https://www.notion.so/grey-box/People-da052a0ffb3a428d8e7013c540c42665
 * Notion Database (sync) - Mandates (Google Sheet): https://docs.google.com/spreadsheets/d/1uqCK0JDHKkuzDEfOlBcSqJCsVxHjCRouOW5F8hJZRUA/edit?gid=2131663677#gid=2131663677
 *
 * It retrieves entries from the Notion database using the Notion API, processes the data, and updates a Google Sheet named "Mandates".
 * The script handles pagination to retrieve all entries, supports multiple Notion property types, and fetches additional information
 * for relation properties to display related page titles instead of IDs.
 *
 * The script includes:
 * - `syncDataToSheet`: Main function that retrieves data from Notion and updates the Google Sheet.
 * - `constructPayload`: Helper function to build the request payload for pagination.
 * - `retrieveNotionData`: Function to fetch data from Notion using the Notion API.
 * - `getPropertyData`: Function that extracts data from Notion properties, including titles and relations.
 * - `fetchPageTitle`: Function that retrieves the title of related Notion pages for display in the sheet.
 * - `convertDate`: Utility function to convert dates from ISO format to a readable format.
 * - `handleNotionResponse`: Function to process the Notion API response and handle errors.
 * - `parseRowFromPage`: Function to extract and format a row of data from a Notion page.
 *
 * Key Features:
 * - Supports various Notion property types, including text, number, select, multi-select, email, date, and relation.
 * - Dynamically fetches and displays the title of related pages, ensuring user-friendly information in the "Team (Current)" column.
 * - Logs details of missing or undefined properties to assist with debugging.
 * - Utilizes a Notion API key stored securely in the script properties for authenticated API requests.
 * - Gracefully handles errors and logs any issues encountered during the sync process.
 * Notion link: https://www.notion.so/grey-box/Sync-Notion-People-Directory-to-Google-Sheet-syncNotionPeopleDirectoryToSheetsMandates-bca68b43894946e7847267aff967180e
 */

/**
 * Entrypoint: Sync mandate data from Google Sheets to Slack user profiles.
 * - Extracts data from the Google Sheet (via `extractUserDataFromSheet`)
 * - Gets Slack user ID (via `getUserIdByEmail`)
 * - Updates Slack profile fields (via `updateUserProfile`)
 */
/**
 * Synchronizes the "People Directory" Notion database with the "Mandates" sheet in Google Sheets.
 *
 * 🔁 This replaces the previous `syncDataToSheet()` function and consolidates:
 * - `retrieveNotionData()` → now uses `fetchNotionData()`
 * - `parseRowFromPage()` → now `parseNotionPageRow()` for clarity
 *
 * The script:
 * - Fetches all pages from the specified Notion database with pagination support
 * - Parses each page into a row of relevant values
 * - Clears and repopulates the "Mandates" sheet with up-to-date data
 *
 * ⚠️ Make sure the headers in the sheet match the parsed values
 * ✅ Sheet will always be cleared and reloaded from scratch
 *
 * Dependencies:
 * - `fetchNotionData(apiUrl, headers, payload)`
 * - `handleNotionApiResponse(response)`
 * - `parseNotionPageRow(page)`
 * - `logToSlack(message)`
 */

function syncNotionPeopleDirectoryToGoogleSheet() {
    logToSlack(
        "📢 Starting execution of \`syncNotionPeopleDirectoryToSheetsMandates\` script"
    );
    const databaseId = "3cf44b088a8f4d6b8abc989353abcdb1";
    const apiUrl = `https://api.notion.com/v1/databases/${databaseId}/query`;

    // Headers configuration
    const headers = {
        "Authorization": `Bearer ${NOTION_API_KEY}`,
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json"
    };

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandates");
    sheet.clearContents();
    sheet.appendRow(MANDATES_SHEET_COLUMNS);

    let allRows = [];
    let hasMore = true;
    let startCursor = null;

    while (hasMore) {
        const payload = constructPayload(startCursor);

        try {
            const response = fetchNotionData(apiUrl, headers, payload);
            const {results, has_more, next_cursor} = handleNotionApiResponse(response);

            const rows = results.map(page => parseNotionPageRow(page));
            allRows = allRows.concat(rows);

            hasMore = has_more;
            startCursor = next_cursor;
        } catch (error) {
            logToSlack(`Error fetching data: ${error}`);

            break;
        }
    }

    if (allRows.length > 0) {
        const startRow = 2;
        const startColumn = 1;
        const numRows = allRows.length;
        const numColumns = allRows[0].length;
        sheet.getRange(startRow, startColumn, numRows, numColumns).setValues(allRows);
    }

    logToSlack("Sync completed!");
    logToSlack(
        "📢 Execution of \`syncNotionPeopleDirectoryToSheetsMandates\` script finished"
    );
}

/**
 * Synchronizes data from the "Mandates" Google Sheet to Slack user profiles.
 *
 * Extracts each user's data, gets their Slack ID, and updates their profile fields.
 *
 * @function syncGoogleSheetToSlack
 * @returns {void}
 */
function syncGoogleSheetToSlack() {
    logToSlack("📢 Starting syncSheetsMandatesToSlackGreyBox...");
//reads user data sheet process them
    const users = extractUserDataFromSheet(SHEET_NAME);
    users.forEach(user => {
        const userId = getSlackUserIdByEmail(user.email);
        if (userId) {
            updateUserProfile(userId, constructProfileFields(user));
        }
    });

    logToSlack("✅ Finished syncSheetsMandatesToSlackGreyBox.");
}

/**
 * Synchronizes data from the "Mandates" Google Sheet to Notion's "People Directory" database.
 *
 * For each row in the sheet, updates "Hours (Current)" and "Hours (Last Update)" in Notion,
 * if changes are detected.
 *
 * @function syncGoogleSheetToNotion
 * @returns {void}
 */
function syncGoogleSheetToNotion() {
    logToSlack(
        "📢 Starting execution of \`syncSheetsMandatesToNotionPeopleDirectory\` script"
    );
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MANDATES_SHEET_NAME);
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    const header = data.shift(); // Remove the header row

    const notionUrlIndex = header.indexOf('Notion Page URL');
    const hoursCurrentIndex = header.indexOf('Hours (decimal)');
    const lastUpdateIndex = header.indexOf('Last Update');

    data.forEach((row, rowIndex) => {
        const notionUrl = row[notionUrlIndex];
        const hoursCurrent = row[hoursCurrentIndex];
        const lastUpdate = row[lastUpdateIndex];

        // Skip row if both relevant Google Sheet values are null/empty
        if (!hoursCurrent && !lastUpdate) {
            logInfo(`Skipping row ${rowIndex + 2} as both Hours and Last Update are empty.`);
            return;
        }

        const notionPageId = extractNotionPageId(notionUrl);
        if (!notionPageId) {
            logInfo(`Skipping row ${rowIndex + 2} as Notion Page ID is missing or invalid.`);
            return;
        }

        // Fetch current Notion page properties
        const notionProperties = getNotionPageProperties(notionPageId);

        // Prepare properties to update in Notion
        const propertiesToUpdate = {};

        if (hoursCurrent) {
            const roundedHoursCurrent = parseFloat(hoursCurrent.toFixed(1));
            if (notionProperties['Hours (Current)']?.number !== roundedHoursCurrent) {
                propertiesToUpdate['Hours (Current)'] = {number: roundedHoursCurrent};
            }
        }

        if (lastUpdate) {
            const formattedLastUpdate = new Date(lastUpdate).toISOString();
            if (notionProperties['Hours (Last Update)']?.date?.start !== formattedLastUpdate) {
                propertiesToUpdate['Hours (Last Update)'] = {date: {start: formattedLastUpdate}};
            }
        }

        // Update Notion only if there are changes
        if (Object.keys(propertiesToUpdate).length > 0) {
            const updateSuccess = updateNotionPageProperties(notionPageId, propertiesToUpdate);
            if (updateSuccess) {
                logInfo(`Successfully updated Notion page ${notionPageId}`);
            } else {
                logToSlack(`Failed to update Notion page ${notionPageId}`);
            }
        } else {
            logInfo(`No changes detected for Notion Page ID: ${notionPageId}. Skipping update.`);
        }
    });

    logToSlack(
        "📢 Execution of \`syncSheetsMandatesToNotionPeopleDirectory\` script finished"
    );
}

/**
 * Synchronizes the Notion Team Directory database to a Google Sheet.
 *
 * Fetches paginated data from the Notion Team database and writes it to the sheet.
 *
 * @function syncTeamDirectoryToSheet
 * @returns {void}
 */
function syncTeamDirectoryToSheet() {
    logInfo("📢 Starting execution of `syncTeamDirectoryToSheet` script");

    const apiUrl = NOTION_QUERY_URL(NOTION_TEAM_DB_ID);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEAM_SHEET_NAME);

    sheet.clearContents();
    sheet.appendRow(TEAM_DIRECTORY_COLUMNS);

    let allRows = [];
    let hasMore = true;
    let startCursor = null;

    while (hasMore) {
        const payload = buildNotionPayload(startCursor);

        try {
            const response = fetchNotionData(apiUrl, NOTION_HEADERS, payload);
            const { results, has_more, next_cursor } = processNotionResponse(response);

            const rows = results.map(page => transformNotionTeamPageToRow(page));
            allRows = allRows.concat(rows);

            hasMore = has_more;
            startCursor = next_cursor;
        } catch (error) {
            logError(`Error fetching data: ${error}`);
            break;
        }
    }

    if (allRows.length > 0) {
        sheet.getRange(2, 1, allRows.length, allRows[0].length).setValues(allRows);
    } else {
        logInfo("No rows found in Notion Team Directory Database to sync.");
    }

    logInfo("📢 `syncTeamDirectoryToSheet` script finished");
}
