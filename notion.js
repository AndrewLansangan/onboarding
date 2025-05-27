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
            Logger.log(`Skipping row ${rowIndex + 2} as both Hours and Last Update are empty.`);
            return;
        }

        const notionPageId = extractNotionPageId(notionUrl);
        if (!notionPageId) {
            Logger.log(`Skipping row ${rowIndex + 2} as Notion Page ID is missing or invalid.`);
            return;
        }

        // Fetch current Notion page properties
        const notionProperties = getNotionPageProperties(notionPageId);

        // Prepare properties to update in Notion
        const propertiesToUpdate = {};

        if (hoursCurrent) {
            const roundedHoursCurrent = parseFloat(hoursCurrent.toFixed(1));
            if (notionProperties['Hours (Current)']?.number !== roundedHoursCurrent) {
                propertiesToUpdate['Hours (Current)'] = { number: roundedHoursCurrent };
            }
        }

        if (lastUpdate) {
            const formattedLastUpdate = new Date(lastUpdate).toISOString();
            if (notionProperties['Hours (Last Update)']?.date?.start !== formattedLastUpdate) {
                propertiesToUpdate['Hours (Last Update)'] = { date: { start: formattedLastUpdate } };
            }
        }

        // Update Notion only if there are changes
        if (Object.keys(propertiesToUpdate).length > 0) {
            const updateSuccess = updateNotionPageProperties(notionPageId, propertiesToUpdate);
            if (updateSuccess) {
                Logger.log(`Successfully updated Notion page ${notionPageId}`);
            } else {
                logToSlack(`Failed to update Notion page ${notionPageId}`);
            }
        } else {
            Logger.log(`No changes detected for Notion Page ID: ${notionPageId}. Skipping update.`);
        }
    });
    logToSlack(
        "📢 Execution of \`syncSheetsMandatesToNotionPeopleDirectory\` script finished"
    );
}
 function extractNotionPageId(notionUrl) {
    const match = notionUrl.match(/([0-9a-f]{32})$/);
    return match ? match[1] : null;
}

function getNotionHeaders() {
    return {
        Authorization: `Bearer ${NOTION_API_KEY}`,
        'Content-Type': 'application/json',
        'Notion-Version': '2022-06-28',
    };
}

 function getNotionPageProperties(pageId) {
    const url = `https://api.notion.com/v1/pages/${pageId}`;
    const response = UrlFetchApp.fetch(url, {headers: getNotionHeaders()});
    const data = JSON.parse(response.getContentText());
    return data.properties || {};
}

 function updateNotionPageProperties(pageId, properties) {
    const url = `https://api.notion.com/v1/pages/${pageId}`;
    const payload = JSON.stringify({properties});
    const options = {
        method: 'patch',
        contentType: 'application/json',
        headers: getNotionHeaders(),
        payload: payload,
    };

    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    return result.object !== 'error';
}
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

/**
 * Builds the payload for the Notion API query to filter by a specific property and value.
 * Assumes the property type is 'status' based on common usage for this script's purpose.
 * If your 'Status' property in Notion is actually a 'Select' type, change "status" below to "select".
 *
 * @param {string|null} startCursor - The cursor for pagination, or null to start.
 * @param {string} filterProperty - The name of the property (e.g., "Status").
 * @param {string} filterValue - The value to filter by (e.g., "Complete").
 * @return {Object} The payload object for the Notion API request.
 */
 function buildNotionFilterPayload(startCursor, filterProperty, filterValue) {
    // *** Correction: Use the property TYPE ('status') as the key, not the property NAME ('Status') ***
    // If your Notion property is actually a 'Select' type, change 'status' to 'select' below.
    const payload = {
        page_size: 100, // Max allowed by Notion API
        filter: {
            property: filterProperty, // The name of the property to filter (e.g., "Status")
            status: {
                // The TYPE of the property being filtered
                equals: filterValue, // The condition (e.g., "Complete")
            },
        },
    };

    if (startCursor) {
        payload.start_cursor = startCursor;
    }

    // Logger.log(`Notion Payload: ${JSON.stringify(payload)}`); // Keep for debugging if needed
    return payload;
}

/**
 * Fetches data from the Notion API using a POST request with retry logic.
 * (Uses fetchWithRetry internally)
 *
 * @param {string} apiUrl - The Notion API endpoint URL.
 * @param {Object} headers - The request headers including Authorization.
 * @param {Object} payload - The request payload (body).
 * @return {Object} The parsed JSON response from Notion.
 * @throws {Error} If the API call fails after retries or returns a Notion error object.
 */
 function fetchNotionDataWithRetry(apiUrl, headers, payload) {
    const options = {
        method: "post",
        headers: headers,
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
    };

    const response = fetchWithRetry(apiUrl, options); // Uses shared fetchWithRetry
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode >= 200 && responseCode < 300) {
        try {
            return JSON.parse(responseBody);
        } catch (e) {
            throw new Error(`Failed to parse Notion API response: ${e.message}`);
        }
    } else {
        let errorMessage = `Notion API request failed (${apiUrl}) with status code ${responseCode}.`;
        try {
            const errorResponse = JSON.parse(responseBody);
            errorMessage += ` Notion Error: ${errorResponse.message || responseBody}`;
        } catch (e) {
            errorMessage += ` Response Body: ${responseBody}`;
        }
        Logger.log(errorMessage); // Log detailed error
        throw new Error(errorMessage); // Throw to be caught by main loop
    }
}

/**
 * Processes the raw JSON response from the Notion API query.
 * Checks for Notion-specific error objects.
 *
 * @param {Object} responseJson - The parsed JSON response from Notion query.
 * @return {Object} The processed response, typically { results, has_more, next_cursor }.
 * @throws {Error} If the response object indicates an error or unexpected structure.
 */
 function processNotionResponse(responseJson) {
    if (responseJson.object && responseJson.object === "error") {
        Logger.log(`Notion API Error: ${JSON.stringify(responseJson)}`);
        throw new Error(
            responseJson.message ||
            `Notion API returned an error: ${responseJson.code}`
        );
    }
    if (!responseJson.results) {
        Logger.log(
            `Unexpected Notion API response structure: ${JSON.stringify(
                responseJson
            )}`
        );
        throw new Error(
            "Unexpected Notion API response structure (missing results)."
        );
    }
    return responseJson;
}

/**
 * Extracts relation IDs from a Notion relation property object.
 *
 * @param {Object} property - The Notion relation property object.
 * @param {string} propertyName - The name of the property (for logging).
 * @return {string[]} An array of relation page IDs, or an empty array if none found.
 */
 function extractRelationIds(property, propertyName) {
    if (
        !property ||
        property.type !== "relation" ||
        !property.relation ||
        property.relation.length === 0
    ) {
        // Logger.log(`Property '${propertyName}' is not a valid relation or is empty.`);
        return [];
    }
    return property.relation.map((rel) => rel.id);
}

/**
 * Fetches a specific property value from a specific Notion page ID.
 * Used here to get the SM's email from their Person page.
 *
 * @param {string} pageId - The ID of the Notion page to fetch.
 * @param {string} propertyToExtract - The exact name of the property to extract from the page.
 * @param {string} notionApiKey - The Notion API key.
 * @return {string|null} The value of the specified property, or null if not found/error.
 */
 function fetchPageProperty(pageId, propertyToExtract, notionApiKey) {
    const apiUrl = `https://api.notion.com/v1/pages/${pageId}`;
    const headers = {
        Authorization: `Bearer ${notionApiKey}`,
        "Notion-Version": "2022-06-28",
    };
    const options = {
        method: "get",
        headers: headers,
        muteHttpExceptions: true,
    };

    try {
        const response = fetchWithRetry(apiUrl, options); // Use retry mechanism
        const responseCode = response.getResponseCode();
        const responseBody = response.getContentText();

        if (responseCode >= 200 && responseCode < 300) {
            const pageData = JSON.parse(responseBody);
            if (pageData.properties && pageData.properties[propertyToExtract]) {
                // Use extractPropertyValue to handle different types correctly
                const value = extractPropertyValue(
                    pageData.properties[propertyToExtract],
                    propertyToExtract
                );
                return value; // Returns string, number, boolean, or empty string
            } else {
                Logger.log(
                    `Property '${propertyToExtract}' not found on page ID ${pageId}.`
                );
                return null;
            }
        } else {
            Logger.log(
                `Error fetching page ${pageId}: Status ${responseCode}. Response: ${responseBody}`
            );
            return null;
        }
    } catch (error) {
        logToSlack(
            `Exception fetching page property '${propertyToExtract}' for page ID ${pageId}: ${error}`
        );
        return null;
    }
}

/**
 * Extracts a property value from a Notion page property object.
 * (Adapted from your provided functions to be reusable here)
 *
 * @param {Object} property - The Notion property object.
 * @param {string} propertyName - The name of the property (for logging).
 * @return {string|number|boolean} The extracted value, or an empty string if not found or type is unhandled.
 */
 function extractPropertyValue(property, propertyName) {
    if (!property) {
        return "";
    }
    const type = property.type;
    let value = "";

    try {
        switch (type) {
            case "title":
                value =
                    property.title && property.title.length > 0
                        ? property.title[0].plain_text
                        : "";
                break;
            case "rich_text":
                value =
                    property.rich_text && property.rich_text.length > 0
                        ? property.rich_text.map((item) => item.plain_text).join("")
                        : "";
                break;
            case "number":
                value = property.number !== null ? property.number : "";
                break;
            case "select":
                value = property.select ? property.select.name : "";
                break;
            case "status":
                value = property.status ? property.status.name : "";
                break;
            case "email": // *** Crucial for getting SM email ***
                value = property.email || "";
                break;
            case "date":
                value = property.date && property.date.start ? property.date.start : "";
                break;
            case "checkbox":
                value = property.checkbox; // Returns true/false
                break;
            // Add other cases if needed (multi_select, people, files, url, phone_number, formula, relation (just IDs), created_time, etc.)
            // For this script, we primarily need title, status, email, and relation (handled by extractRelationIds)
            default:
                Logger.log(
                    `Unhandled property type: '${type}' for property '${propertyName}' in extractPropertyValue.`
                );
                value = "";
        }
    } catch (e) {
        Logger.log(
            `Error extracting value for property '${propertyName}' (Type: ${type}): ${e}`
        );
        value = "";
    }
    // Return primitive types directly, convert others to string if needed for consistency elsewhere
    return typeof value === "boolean" || typeof value === "number"
        ? value
        : String(value);
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
