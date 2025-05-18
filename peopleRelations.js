/**
 * This script facilitates the linking of pages between the People Directory Notion databases and the internal HR-only version by establishing relations based on email addresses.
 * People Directory (Notion database) : https://www.notion.so/grey-box/People-da052a0ffb3a428d8e7013c540c42665
 * People Directory (Internal) (Notion database) : https://www.notion.so/grey-box/47fbed712f3e4558b032edb9ec081f00?v=2d969f6b09084313823bca813f39db69
 *
 * Key Components:
 * - Configuration: Retrieves Notion API credentials and initializes database IDs for the two databases to be linked.
 * - `fetchAllPages`: Fetches all pages from a specified Notion database, handling pagination to ensure complete data retrieval.
 * - `updatePageRelationWithMultiple`: Updates a specific relation property of a page in Notion, linking it to one or more pages based on shared email addresses.
 * - `linkDatabases`: Main function that orchestrates the linking process by:
 *   - Fetching pages from both databases.
 *   - Mapping emails from one database to corresponding page IDs.
 *   - Iterating over the pages in the first database and updating their relation properties to link them with the corresponding pages in the second database.
 *   - Handles cases where multiple matches are found by linking all corresponding relations to the page.
 *
 * The script is designed to run in sequence, ensuring that database pages are fetched and processed in a manner that maintains data integrity
 * and avoids redundant updates. Logging is extensively used to track the progress and identify any issues during execution.
 *
 * Notion link: https://www.notion.so/grey-box/Sync-Relation-Notion-Team-Directory-with-People-Directory-syncNotionPeopleRelations-gs-a906389d2dd440b6a65c6ffe0130787e
 */

// RUN linkDatabases first
// Configuration
function initializeConfig() {
    const notionApiKey = PropertiesService.getScriptProperties().getProperty('NOTION_API_KEY');
    const headers = {
        "Authorization": `Bearer ${notionApiKey}`,
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json"
    };
    Logger.log("Script configuration loaded.");

    // Database IDs
    const databaseId1 = "3cf44b088a8f4d6b8abc989353abcdb1"; // First database ID
    const databaseId2 = "47fbed712f3e4558b032edb9ec081f00"; // Second database ID

    // Check for undefined database IDs
    if (!databaseId1 || !databaseId2) {
        logToSlack("One or both database IDs are undefined. Please check the configuration.");
        return null;  // Return null if IDs are undefined
    }

    logToSlack(`Database IDs: ${databaseId1}, ${databaseId2}`);

    return { headers, databaseId1, databaseId2 };
}

// Fetch all pages from a given database with proper pagination handling
function fetchAllPages(databaseId, headers) {
    logToSlack(`Fetching all pages from database: ${databaseId}`);

    let hasMore = true;
    let startCursor = undefined; // Initialize as undefined for the first request
    let allPages = [];
    let pageCount = 0; // Track the number of pages fetched

    while (hasMore) {
        let apiUrl = `https://api.notion.com/v1/databases/${databaseId}/query`;
        let payload = {
            page_size: 100 // Maximum allowed by Notion API
        };

        // If there's a start cursor, add it to the payload for pagination
        if (startCursor) {
            payload.start_cursor = startCursor;
        }

        const options = {
            method: "post",
            headers: headers,
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        };

        const response = UrlFetchApp.fetch(apiUrl, options);
        const jsonResponse = JSON.parse(response.getContentText());

        if (!jsonResponse.results) {
            logToSlack("No results found in the response, possibly due to an error or malformed response.");
            logToSlack(JSON.stringify(jsonResponse));
            break;
        }

        // Append the fetched pages to the allPages array
        allPages = allPages.concat(jsonResponse.results);
        pageCount++;

        Logger.log(`Fetched ${jsonResponse.results.length} pages in this request.`);
        logToSlack(`Total pages fetched so far: ${allPages.length}`);

        hasMore = jsonResponse.has_more; // Check if there are more pages to fetch
        startCursor = jsonResponse.next_cursor; // Update start cursor for the next request

        if (!hasMore) {
            Logger.log(`Pagination complete. Fetched ${allPages.length} pages from database ${databaseId}.`);
        } else {
            Logger.log(`Fetching the next batch of pages starting from cursor: ${startCursor}`);
        }
    }

    return allPages;
}

// Update the relation property of a page to handle multiple related page IDs
function updatePageRelationWithMultiple(pageId, relationPayload, headers) {
    Logger.log(`Updating relation for page ${pageId} to link to multiple pages.`);

    const apiUrl = `https://api.notion.com/v1/pages/${pageId}`;
    const payload = {
        properties: {
            "People Directory (Sync)": {
                relation: relationPayload // Now an array of relations
            }
        }
    };

    const options = {
        method: "patch",
        headers: headers,
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseContent = JSON.parse(response.getContentText());

    if (response.getResponseCode() !== 200) {
        logToSlack(`Failed to update page relation for ${pageId}: ${JSON.stringify(responseContent)}`);
        return false;
    } else {
        Logger.log(`Successfully updated page ${pageId} with multiple relations.`);
        return true;
    }
}

// Main function to link pages across databases
function linkDatabases() {
    logToSlack(
        "📢 Starting execution of \`syncNotionPeopleRelations\` script"
    );
    const config = initializeConfig();

    if (!config) {
        logToSlack("Script aborted due to missing configuration.");
        return;  // Exit if configuration is not valid
    }

    const { headers, databaseId1, databaseId2 } = config;

    logToSlack("Starting the process to link databases.");

    const pagesDatabase1 = fetchAllPages(databaseId1, headers);
    const pagesDatabase2 = fetchAllPages(databaseId2, headers);

    if (!pagesDatabase1.length || !pagesDatabase2.length) {
        logToSlack("Failed to fetch pages from one or both databases. Please check the fetchAllPages function and the database IDs.");
        return;
    }

    Logger.log("Mapping 'Email (Org)' to page IDs for database 2");
    const emailToPageIdMap = pagesDatabase2.reduce((map, page) => {
        const email = (page.properties["Email (Org)"]?.email || "").toLowerCase();
        if (email) {
            if (!map[email]) {
                map[email] = []; // Initialize array if not already created
            }
            map[email].push(page.id); // Add the page ID to the array
        }
        return map;
    }, {});

    logToSlack("Iterating through pages in database 1 to update relations.");
    pagesDatabase1.forEach(page => {
        const email = (page.properties["Email (Org)"]?.email || "").toLowerCase();
        if (email && emailToPageIdMap[email]) {
            Logger.log(`Found matching email: ${email}`);
            if (!page.properties["People Directory (Sync)"]?.relation?.length) {
                const relatedPageIds = emailToPageIdMap[email]; // Get all matching page IDs

                // Prepare the relation payload with multiple IDs
                const relationPayload = relatedPageIds.map(pageId => ({ id: pageId }));

                Logger.log(`Updating page ${page.id} with relations: ${JSON.stringify(relationPayload)}`);

                const success = updatePageRelationWithMultiple(page.id, relationPayload, headers); // Updated function to handle multiple relations
                if (!success) {
                    logToSlack(`Failed to link page with email: ${email}`);
                }
            } else {
                Logger.log(`Page with email ${email} is already linked.`);
            }
        }
    });
    logToSlack("Linking process completed.");
    logToSlack(
        "📢 Execution of \`syncNotionPeopleRelations\` script finished"
    );
}
