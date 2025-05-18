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

function syncDataToSheet() {
    logToSlack(
        "📢 Starting execution of \`syncNotionPeopleDirectoryToSheetsMandates\` script"
    );
    const notionApiKey = PropertiesService.getScriptProperties().getProperty('NOTION_API_KEY');
    const databaseId = "3cf44b088a8f4d6b8abc989353abcdb1";
    const apiUrl = `https://api.notion.com/v1/databases/${databaseId}/query`;

    // Headers configuration
    const headers = {
        "Authorization": `Bearer ${notionApiKey}`,
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json"
    };

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandates");
    sheet.clearContents();
    sheet.appendRow(["Greybox ID", "Mandate (Status)", "Name", "Position", "Team (Current)", "Team (Previous)", "Mandate (Date)", "Hours (Initial)", "Hours (Current)", "Availability (avg h/w)", "Email (Org)", "Created (Profile)", "Notion Page URL"]);

    let allRows = [];
    let hasMore = true;
    let startCursor = null;

    while (hasMore) {
        const payload = constructPayload(startCursor);

        try {
            const response = retrieveNotionData(apiUrl, headers, payload);
            const { results, has_more, next_cursor } = handleNotionResponse(response);

            const rows = results.map(page => parseRowFromPage(page));
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

function constructPayload(startCursor) {
    const payload = { page_size: 100 };
    if (startCursor) {
        payload.start_cursor = startCursor;
    }
    return payload;
}

function retrieveNotionData(apiUrl, headers, payload) {
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

function getPropertyData(property, propertyName) {
    if (!property) {
        Logger.log(`Property '${propertyName}' is undefined.`);
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
            if (property.date && property.date.start && property.date.end) {
                const startDate = convertDate(property.date.start);
                const endDate = convertDate(property.date.end);
                value = `${startDate} → ${endDate}`;
            } else if (property.date && property.date.start) {
                value = convertDate(property.date.start);
            } else {
                value = "";
            }
            break;
        case 'created_time':
            if (property.created_time) {
                const dateFormatter = new Intl.DateTimeFormat('en-CA', {
                    year: 'numeric',
                    month: '2-digit',
                    day: '2-digit',
                });
                const createdDate = new Date(property.created_time);
                value = dateFormatter.format(createdDate);
            } else {
                value = "";
            }
            break;
        case 'rich_text':
            value = property.rich_text.length > 0 ? property.rich_text.map(item => item.plain_text).join("\n") : "";
            break;
        case 'status':
            value = property.status ? property.status.name : "";
            break;
        case 'relation':
            value = property.relation.length > 0 ?
                property.relation.map(rel => fetchPageTitle(rel.id)).join(", ") : "";
            break;
        default:
            logToSlack(`Unhandled property type: ${type}`);
            value = "";
    }

    return value;
}

function fetchPageTitle(pageId) {
    const notionApiKey = PropertiesService.getScriptProperties().getProperty('NOTION_API_KEY');
    const apiUrl = `https://api.notion.com/v1/pages/${pageId}`;
    const headers = {
        "Authorization": `Bearer ${notionApiKey}`,
        "Notion-Version": "2022-06-28"
    };

    const options = {
        method: "get",
        headers: headers,
        muteHttpExceptions: true
    };

    try {
        const response = UrlFetchApp.fetch(apiUrl, options);
        const pageData = JSON.parse(response.getContentText());

        // Look for a title-type property dynamically within properties
        const titleProperty = Object.values(pageData.properties).find(prop => prop.type === 'title');
        if (titleProperty && titleProperty.title && titleProperty.title.length > 0) {
            return titleProperty.title[0].plain_text;
        } else {
            return "[No Title]";
        }
    } catch (error) {
        logToSlack(`Error fetching page title for ID ${pageId}: ${error}`);
        return "[Error Fetching Title]";
    }
}


function convertDate(isoDateString) {
    const date = new Date(isoDateString);
    const options = { year: 'numeric', month: 'long', day: 'numeric' };
    return date.toLocaleDateString('en-US', options);
}

function handleNotionResponse(response) {
    if (response.object && response.object === "error") {
        throw new Error(response.message);
    }
    return response;
}

function parseRowFromPage(page) {
    const properties = page.properties;

    Object.keys(properties).forEach(key => {
        if (!properties[key]) {
            Logger.log(`Property '${key}' is undefined.`);
        }
    });

    const mandateStatus = getPropertyData(properties["Mandate (Status)"], "Mandate (Status)");
    const name = getPropertyData(properties["Name"], "Name");
    const position = getPropertyData(properties["Position"], "Position");
    const teamCurrent = getPropertyData(properties["Team (Current)"], "Team (Current)");
    const teamPrevious = getPropertyData(properties["Team (Previous)"], "Team (Previous)");
    const mandateDate = getPropertyData(properties["Mandate (Date)"], "Mandate (Date)");
    const hoursInitial = getPropertyData(properties["Hours (Initial)"], "Hours (Initial)");
    const hoursCurrent = getPropertyData(properties["Hours (Current)"], "Hours (Current)");
    const availability = getPropertyData(properties["Availability (avg h/w)"], "Availability (avg h/w)");
    const emailOrg = getPropertyData(properties["Email (Org)"], "Email (Org)");
    const createdProfile = getPropertyData(properties["Created (Profile)"], "Created (Profile)");

    const greyboxId = emailOrg.split('@')[0].toLowerCase();
    const notionPageUrl = `https://www.notion.so/${page.id.replace(/-/g, '')}`;

    return [greyboxId, mandateStatus, name, position, teamCurrent, teamPrevious, mandateDate, hoursInitial, hoursCurrent, availability, emailOrg, createdProfile, notionPageUrl];
}

function fetchDateProperty(property, datePart) {
    if (property && property.date && property.date[datePart]) {
        return property.date[datePart];
    }
    return null;
}

function logMissingColumns(rowData) {
    const emptyColumns = [];
    rowData.forEach((value, index) => {
        if (value === "" || value == null) {
            emptyColumns.push(index + 1);
        }
    });
    if (emptyColumns.length > 0) {
        logToSlack(`Empty values found in columns: ${emptyColumns.join(", ")}`);
    }
}

function logTypes(properties) {
    Object.keys(properties).forEach(key => {
        const property = properties[key];
        if (property) {
            Logger.log(`Property: ${key}, Type: ${property.type}`);
        } else {
            Logger.log(`Property: ${key} is undefined.`);
        }
    });
}
