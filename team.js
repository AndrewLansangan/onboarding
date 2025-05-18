// Claudia
/**
 * This script syncs data from the Team Directory (Notion database) to the Team Directory (Google Sheet).
 *
 * Team Directory (Notion database): https://www.notion.so/grey-box/70779d3ee3cf467b9b86171acabc3321
 * Team Directory (Google Sheet): https://docs.google.com/spreadsheets/d/YOUR_GOOGLE_SHEET_ID_HERE/edit
 *
 * It retrieves Notion database entries using the Notion API and updates a sheet named "Team Directory" with the data.
 * The script handles pagination to fetch all entries, extracts relevant properties from each page,
 * and formats the data to fit the Google Sheets structure. Additionally, it logs useful debugging information
 * and handles various Notion property types, including relations and formulas, to ensure that data is accurately
 * captured in the sheet. The script also clears the existing sheet content before inserting the new data.
 *
 * Data fields synced:
 * - Name
 * - Status
 * - People (Current)
 * - Scrum Master
 * - Activity (Epic)
 * - Date (Epic)
 *
 * Notion Page: https://www.notion.so/grey-box/Sync-Notion-Team-Directory-to-Google-Sheet-syncNotionTeamDirectoryToSheetsMandates-119815ddf30b802eb500d4c8e9251914?pvs=4
 */

function syncTeamDirectoryToSheet() {
    logToSlack(
        "📢 Starting execution of \`syncTeamDirectoryToSheet\` script"
    );
    // Securely retrieve the Notion API Key
    const notionApiKey = PropertiesService.getScriptProperties().getProperty('NOTION_API_KEY');
    const databaseId = "70779d3ee3cf467b9b86171acabc3321"; // Update with the new Team Directory database ID
    const apiUrl = `https://api.notion.com/v1/databases/${databaseId}/query`;

    // Headers configuration
    const headers = {
        "Authorization": `Bearer ${notionApiKey}`,
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json"
    };

    // Access the spreadsheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Team Directory");
    sheet.clearContents(); // Clear existing content
    sheet.appendRow(["Name", "Status", "People", "Scrum Master", "Activity (Epic)", "Date (Epic)"]);

    let allRows = [];
    let hasMore = true;
    let startCursor = null;

    while (hasMore) {
        const payload = buildPayload(startCursor);

        try {
            const response = fetchNotionData(apiUrl, headers, payload);
            const { results, has_more, next_cursor } = processNotionResponse(response);

            // Collect rows from the current batch
            const rows = results.map(page => extractTeamDirectoryRow(page));
            allRows = allRows.concat(rows);

            hasMore = has_more;
            startCursor = next_cursor;
        } catch (error) {
            logToSlack(`Error fetching data: ${error}`);
            break;
        }
    }

    // Batch update the sheet if there are rows to add
    if (allRows.length > 0) {
        sheet.getRange(2, 1, allRows.length, allRows[0].length).setValues(allRows);
    }

    logToSlack("Sync completed!");
    logToSlack(
        "📢 Execution of \`syncTeamDirectoryToSheet\` script finished"
    );
}

function extractTeamDirectoryRow(page) {
    const properties = page.properties;

    const name = extractPropertyValue(properties["Name"], "Name");
    const status = extractPropertyValue(properties["Status"], "Status");
    const people = extractPropertyValue(properties["People (Current)"], "People (Current)");
    const scrumMaster = extractPropertyValue(properties["Scrum Master"], "Scrum Master");
    const activityEpic = extractPropertyValue(properties["Activity (Epic)"], "Activity (Epic)");
    const dateEpic = extractPropertyValue(properties["Date (Epic)"], "Date (Epic)");

    return [name, status, people, scrumMaster, activityEpic, dateEpic];
}

function extractPropertyValue(property, propertyName) {
    if (!property) {
        logToSlack(`Property '${propertyName}' is undefined.`);
        return "";
    }

    const type = property.type;
    let value = "";

    switch (type) {
        case 'title':
            value = property.title.length > 0 ? property.title[0].plain_text : "";
            break;
        case 'number':
            value = property.number !== null ? property.number : "";
            break;
        case 'select':
            value = property.select ? property.select.name : "";
            break;
        case 'multi_select':
            value = property.multi_select.length > 0 ? property.multi_select.map(item => item.name).join(", ") : "";
            break;
        case 'email':
            value = property.email || "";
            break;
        case 'date':
            if (property.date && property.date.start) {
                value = formatDate(property.date.start);
            }
            break;
        case 'created_time':
            if (property.created_time) {
                value = new Date(property.created_time).toLocaleDateString('en-CA');
            }
            break;
        case 'rich_text':
            value = property.rich_text.length > 0 ? property.rich_text.map(item => item.plain_text).join("\n") : "";
            break;
        case 'status':
            value = property.status ? property.status.name : "";
            break;
        case 'relation':
            value = property.relation.length > 0 ? property.relation.map(item => item.name || item.plain_text).join(", ") : "";
            break;
        case 'formula':
            if (property.formula.string) {
                value = property.formula.string;
            } else if (property.formula.number !== null) {
                value = property.formula.number;
            } else if (property.formula.boolean !== null) {
                value = property.formula.boolean.toString();
            }
            break;
        default:
            logToSlack(`Unhandled property type: ${type}`);
            value = "";
    }

    return value;
}

function buildPayload(startCursor) {
    const payload = {
        page_size: 100
    };
    if (startCursor) {
        payload.start_cursor = startCursor;
    }
    return payload;
}

function fetchNotionData(apiUrl, headers, payload) {
    const options = {
        method: "post",
        headers: headers,
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(apiUrl, options);
    const jsonResponse = JSON.parse(response.getContentText());

    if (jsonResponse.results && jsonResponse.results.length > 0) {
        Logger.log(JSON.stringify(jsonResponse.results[0], null, 2));
    }

    return jsonResponse;
}

function processNotionResponse(response) {
    if (response.object && response.object === "error") {
        throw new Error(response.message);
    }
    return response;
}
