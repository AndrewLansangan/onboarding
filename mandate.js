/**
 * This script syncs data from a Google Sheet to Notion by updating properties on Notion pages based on the values in the sheet.
 * Google Sheet: https://docs.google.com/spreadsheets/d/1uqCK0JDHKkuzDEfOlBcSqJCsVxHjCRouOW5F8hJZRUA/edit?gid=2131663677#gid=2131663677
 * It reads data from a specified sheet, compares the values with the corresponding Notion properties,
 * and updates the Notion page only if the values are different, minimizing unnecessary API calls.
 *
 * The script includes:
 * - `syncGoogleSheetToNotion`: Main function that processes the data from the sheet and updates Notion pages.
 * - `getNotionPageProperties`: Helper function that fetches the current properties of a Notion page.
 * - `updateNotionPageProperties`: Function that updates the Notion page properties if they differ from the sheet values.
 *
 * Key features:
 * - Compares Google Sheet values with Notion properties before updating to avoid unnecessary API calls.
 * - Handles date fields, ensuring that dates are updated only when a valid date is present in the sheet.
 * - Rounds `Hours (Current)` to one decimal place before updating Notion.
 * - Logs errors and issues encountered during the synchronization process.
 *
 * Notion Link: https://www.notion.so/grey-box/Sync-Mandate-Google-Sheet-to-Notion-People-Directory-syncSheetsMandatesToNotionPeopleDirectory-gs-2c6029e5bdb745ef84d12c11d2f27691?pvs=4
 */

const NOTION_API_KEY = PropertiesService.getScriptProperties().getProperty('NOTION_API_KEY');
const MANDATES_SHEET_NAME = 'Mandates'; // Change to your actual sheet name

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
    const response = UrlFetchApp.fetch(url, { headers: getNotionHeaders() });
    const data = JSON.parse(response.getContentText());
    return data.properties || {};
}

function updateNotionPageProperties(pageId, properties) {
    const url = `https://api.notion.com/v1/pages/${pageId}`;
    const payload = JSON.stringify({ properties });
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
