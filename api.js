// --- Core Functions ---

/**
 * Checks if a user group with the specified name already exists in Slack.
 * @param slackGroupName
 * @param {string} slackToken - Slack API token with group read permissions.
 * @return {string|boolean} Group ID if found, false otherwise.
 */
function checkForExistingGroups(slackGroupName, slackToken = SLACK_USER_TOKEN) {
    if (!name || !slackToken) {
        logError("checkForExistingGroups: Name or token is missing.");
        return false;
    }

    try {
        const options = {
            method: 'get',
            headers: { Authorization: `Bearer ${slackToken}` },
            muteHttpExceptions: true,
        };

        const response = fetchWithRetry(SLACK_USERGROUP_LIST_URL, options);
        const data = JSON.parse(response.getContentText());

        if (!data.ok) {
            logError(`Slack API error ${response.getResponseCode()}: ${data.error}`);
            return false;
        }

        const match = data.usergroups.find(g => g.name === name);
        return match?.id || false;
    } catch (error) {
        logError(`checkForExistingGroups error: ${error.message || error}`);
        return false;
    }
}

/**
 * Fetches all pages from a Notion database, handling pagination.
 * @param {string} notionDatabaseId - The ID of the Notion database.
 * @param {Object} headers - HTTP headers, including Notion-Version and Authorization.
 * @return {Array<Object>} An array of Notion page objects.
 */
function fetchAllPages(notionDatabaseId, headers) {
    const pages = [];
    let hasMore = true;
    let cursor = null;

    logInfo(`🗃 Fetching pages from Notion DB: ${notionDatabaseId}`);

    while (hasMore) {
        const payload = { page_size: NOTION_PAGE_SIZE, ...(cursor && { start_cursor: cursor }) };
        const options = {
            method: "post",
            headers: {
                ...headers,
                "Content-Type": "application/json",
            },
            payload: JSON.stringify(payload),
            muteHttpExceptions: true,
        };

        try {
            const response = UrlFetchApp.fetch(NOTION_QUERY_URL(notionDatabaseId), options);
            const data = JSON.parse(response.getContentText());

            if (response.getResponseCode() !== 200) {
                logError(`❌ Notion API error for DB ${notionDatabaseId}: ${JSON.stringify(data)}`);
                break;
            }

            if (!data.results) {
                logError(`❌ No results from Notion API for DB: ${notionDatabaseId}`);
                break;
            }

            pages.push(...data.results);
            cursor = data.next_cursor;
            hasMore = data.has_more;

            logDebug(`Fetched ${data.results.length} entries. Total so far: ${pages.length}`);
        } catch (err) {
            logError(`❌ Error fetching pages for DB ${notionDatabaseId}: ${err.message || err}`);
            break;
        }
    }

    return pages;
}

/**
 * Updates a relation property on a Notion page with multiple relation values.
 * @param {string} pageId - The ID of the Notion page to update.
 * @param {Array<Object>} relationPayload - An array of relation objects (e.g., [{id: 'related_page_id'}]).
 * @param {Object} headers - HTTP headers, including Notion-Version and Authorization.
 * @return {boolean} True if the update was successful, false otherwise.
 */
function updatePageRelationWithMultiple(pageId, relationPayload, headers) {
    if (!pageId || !relationPayload || relationPayload.length === 0) {
        logError("updatePageRelationWithMultiple: Page ID or relation payload is invalid.");
        return false;
    }

    const options = {
        method: "patch",
        headers: {
            ...headers,
            "Content-Type": "application/json",
        },
        payload: JSON.stringify({
            properties: {
                [PEOPLE_RELATION_PROPERTY]: {
                    relation: relationPayload,
                }
            }
        }),
        muteHttpExceptions: true,
    };

    try {
        const response = UrlFetchApp.fetch(NOTION_PAGE_UPDATE_URL(pageId), options);
        const result = JSON.parse(response.getContentText());

        if (response.getResponseCode() !== 200) {
            logError(`Failed to update relation for page ${pageId}: ${JSON.stringify(result)}`);
            return false;
        }

        logInfo(`✅ Linked page ${pageId} to ${relationPayload.length} related pages.`);
        return true;
    } catch (error) {
        logError(`updatePageRelationWithMultiple error: ${error.message || error}`);
        return false;
    }
}

/**
 * Builds the payload for a Notion database query, including a start cursor if provided.
 * @param {string|null} startCursor - The cursor to start fetching from (for pagination).
 * @return {Object} The payload object.
 */
function buildPayload(startCursor) {
    return startCursor ? { page_size: NOTION_PAGE_SIZE, start_cursor: startCursor } : { page_size: NOTION_PAGE_SIZE };
}