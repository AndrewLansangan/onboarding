/**
 * Checks if a user group with the specified name already exists
 * @param {string} name - The name of the user group to check
 * @param {string} token - Slack API token with group read permissions
 * @return {string|boolean} Group ID if found, false otherwise
 */
function checkForExistingGroups(name, token) {
    if (!name || !token) return false;

    try {
        const options = {
            method: 'get',
            contentType: 'application/json',
            headers: { Authorization: `Bearer ${token}` },
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
        logError(`checkForExistingGroups error: ${error}`);
        return false;
    }
}

function fetchAllPages(databaseId, headers) {
    const pages = [];
    let hasMore = true;
    let cursor = null;

    logInfo(`🗃 Fetching pages from Notion DB: ${databaseId}`);

    while (hasMore) {
        const payload = { page_size: NOTION_PAGE_SIZE, ...(cursor && { start_cursor: cursor }) };
        const options = {
            method: "post",
            headers,
            payload: JSON.stringify(payload),
            muteHttpExceptions: true,
        };

        try {
            const response = UrlFetchApp.fetch(NOTION_QUERY_URL(databaseId), options);
            const data = JSON.parse(response.getContentText());

            if (!data.results) {
                logError(`❌ No results from Notion API for DB: ${databaseId}`);
                break;
            }

            pages.push(...data.results);
            cursor = data.next_cursor;
            hasMore = data.has_more;

            logDebug(`Fetched ${data.results.length} entries. Total so far: ${pages.length}`);
        } catch (err) {
            logError(`❌ Error fetching pages for DB ${databaseId}: ${err}`);
            break;
        }
    }

    return pages;
}

function updatePageRelationWithMultiple(pageId, relationPayload, headers) {
    if (!pageId || !relationPayload || relationPayload.length === 0) return false;

    const options = {
        method: "patch",
        headers,
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
        logError(`updatePageRelationWithMultiple error: ${error}`);
        return false;
    }
}

function buildPayload(startCursor) {
    return startCursor ? { page_size: NOTION_PAGE_SIZE, start_cursor: startCursor } : { page_size: NOTION_PAGE_SIZE };
}
