/* ============================================================================
 * 🔐 SECTION 1: AUTH & HEADERS
 * ========================================================================= */

/**
 * Retrieves the Notion API key from script config and returns standard headers.
 * @return {Object} Headers for Notion API requests.
 */
const notionApiKey = getScriptConfig().NOTION_API_KEY;
function getNotionHeaders() {
    return {
        Authorization: `Bearer ${notionApiKey}`,
        'Content-Type': 'application/json',
        'Notion-Version': '2022-06-28',
    };
}


/* ============================================================================
 * 📡 SECTION 2: CORE NOTION API CALLS (GET + PATCH)
 * ========================================================================= */

/**
 * Fetches the full properties of a Notion page by ID.
 * @param {string} pageId - Notion page ID.
 * @return {Object} Notion page properties.
 */
function getNotionPageProperties(pageId) {
    const url = `https://api.notion.com/v1/pages/${pageId}`;
    const response = UrlFetchApp.fetch(url, { headers: getNotionHeaders() });
    const data = JSON.parse(response.getContentText());
    return data.properties || {};
}

/**
 * Updates one or more properties on a Notion page.
 * @param {string} pageId - Notion page ID.
 * @param {Object} properties - Properties to update.
 * @return {boolean} Whether the update was successful.
 */
function updateNotionPageProperties(pageId, properties) {
    const url = `https://api.notion.com/v1/pages/${pageId}`;
    const options = {
        method: 'patch',
        contentType: 'application/json',
        headers: getNotionHeaders(),
        payload: JSON.stringify({ properties }),
    };
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    return result.object !== 'error';
}


/* ============================================================================
 * 🧱 SECTION 3: BUILDERS (Payloads, Filters, Primitives)
 * ========================================================================= */

/**
 * Builds a default Notion database query payload.
 * @param {string|null} startCursor - Optional start cursor for pagination.
 * @return {Object} Notion query payload.
 */
function constructPayload(startCursor) {
    const payload = { page_size: 100 };
    if (startCursor) payload.start_cursor = startCursor;
    return payload;
}

/**
 * Builds a filter payload for querying Notion by status (or select).
 * @param {string|null} startCursor
 * @param {string} filterProperty
 * @param {string} filterValue
 * @return {Object} Filtered query payload.
 */
function buildNotionFilterPayload(startCursor, filterProperty, filterValue) {
    const payload = {
        page_size: 100,
        filter: {
            property: filterProperty,
            status: { equals: filterValue },
        },
    };
    if (startCursor) payload.start_cursor = startCursor;
    return payload;
}


/* ============================================================================
 * 📦 SECTION 4: RESPONSE PROCESSORS
 * ========================================================================= */

/**
 * Throws if Notion response is an error object.
 * @param {Object} response
 * @return {Object} Same response if valid.
 */
function handleNotionApiResponse(response) {
    if (response.object === 'error') throw new Error(response.message);
    return response;
}

/**
 * Validates a parsed JSON response from Notion.
 * @param {Object} responseJson
 * @return {Object} Same response if valid.
 */
function processNotionResponse(responseJson) {
    if (responseJson.object === "error") throw new Error(responseJson.message);
    if (!responseJson.results) throw new Error("Missing 'results' in Notion response.");
    return responseJson;
}


/* ============================================================================
 * 🔁 SECTION 5: PAGE + PROPERTY PARSERS
 * ========================================================================= */

/**
 * Extracts a structured row of data from a Notion page.
 * @param {Object} page - A Notion page object.
 * @return {Array<string>} Array of extracted values.
 */
function parseNotionPageRow(page) {
    const props = page.properties;
    const safe = (key) => getNotionPropertyDisplayValue(props[key], key);
    const greyboxId = safe("Email (Org)").split('@')[0].toLowerCase();
    const url = `https://www.notion.so/${page.id.replace(/-/g, '')}`;

    return [
        greyboxId,
        safe("Mandate (Status)"),
        safe("Name"),
        safe("Position"),
        safe("Team (Current)"),
        safe("Team (Previous)"),
        safe("Mandate (Date)"),
        safe("Hours (Initial)"),
        safe("Hours (Current)"),
        safe("Availability (avg h/w)"),
        safe("Email (Org)"),
        safe("Created (Profile)"),
        url,
    ];
}

/**
 * Converts Notion property objects into readable values for spreadsheet.
 * @param {Object} property - A Notion property object.
 * @param {string} key - Property key (for logging).
 * @return {string} Display string.
 */
function getNotionPropertyDisplayValue(property, key) {
    if (!property) {
        Logger.log(`Property '${key}' is undefined.`);
        return '';
    }

    const type = property.type;
    switch (type) {
        case 'title':
            return property.title?.[0]?.plain_text || '';
        case 'number':
            return property.number ?? '';
        case 'select':
            return property.select?.name || '';
        case 'multi_select':
            return property.multi_select.map(i => i.name).join(', ');
        case 'email':
            return property.email || '';
        case 'date':
            const start = property.date?.start;
            const end = property.date?.end;
            return start && end ? `${convertDate(start)} → ${convertDate(end)}` :
                start ? convertDate(start) : '';
        case 'created_time':
            return property.created_time ? new Intl.DateTimeFormat('en-CA').format(new Date(property.created_time)) : '';
        case 'rich_text':
            return property.rich_text.map(i => i.plain_text).join('\n');
        case 'status':
            return property.status?.name || '';
        case 'relation':
            return property.relation.map(rel => fetchNotionRelationPageTitle(rel.id)).join(', ');
        default:
            logToSlack(`Unhandled property type: ${type}`);
            return '';
    }
}

/**
 * Extracts raw primitive values from a Notion property (used in logic).
 * @param {Object} property - A Notion property.
 * @param {string} key - Property name.
 * @return {string|boolean|number}
 */
function extractNotionPropertyPrimitive(property, key) {
    if (!property) return '';
    try {
        switch (property.type) {
            case 'title': return property.title?.[0]?.plain_text || '';
            case 'rich_text': return property.rich_text.map(i => i.plain_text).join('');
            case 'number': return property.number ?? '';
            case 'select': return property.select?.name || '';
            case 'status': return property.status?.name || '';
            case 'email': return property.email || '';
            case 'date': return property.date?.start || '';
            case 'checkbox': return property.checkbox;
            default:
                Logger.log(`Unhandled type '${property.type}' in '${key}'`);
                return '';
        }
    } catch (e) {
        Logger.log(`Error parsing '${key}' — ${e}`);
        return '';
    }
}

/**
 * Extracts an array of page IDs from a relation property.
 * @param {Object} property - A Notion relation property.
 * @param {string} key - Property name.
 * @return {string[]} Array of related page IDs.
 */
function extractRelationIds(property, key) {
    return property?.type === "relation" && property.relation?.length
        ? property.relation.map(r => r.id)
        : [];
}


/* ============================================================================
 * 🔗 SECTION 6: FETCH HELPERS (WITH RETRIES)
 * ========================================================================= */

/**
 * Basic POST to Notion API with raw UrlFetchApp (no retry).
 * @param {string} apiUrl
 * @param {Object} headers
 * @param {Object} payload
 * @return {Object} Parsed JSON response.
 */
function fetchNotionData(apiUrl, headers, payload) {
    const options = {
        method: "post",
        headers,
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(apiUrl, options);
    return JSON.parse(response.getContentText());
}

/**
 * POST to Notion API with retry logic.
 * @param {string} apiUrl
 * @param {Object} headers
 * @param {Object} payload
 * @return {Object} Parsed JSON response.
 */
function fetchNotionDataWithRetry(apiUrl, headers, payload) {
    const options = {
        method: "post",
        headers,
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };
    const response = fetchWithRetries(apiUrl, options);
    return JSON.parse(response.getContentText());
}


/* ============================================================================
 * 📄 SECTION 7: Notion Property Utilities
 * ========================================================================= */

/**
 * Returns part of a Notion date (start/end) if present.
 * @param {Object} property - Notion date property.
 * @param {string} part - "start" or "end".
 * @return {string|null}
 */
function fetchDateProperty(property, part) {
    return property?.date?.[part] || null;
}

/* ============================================================================
 * 📡 SECTION 8: API Calls: Relation Page Helpers
 * ========================================================================= */

/**
 * Fetches title of a related page via Notion API (for relation rendering).
 * @param {string} pageId - Related page ID.
 * @return {string} Page title or fallback.
 */
function fetchNotionRelationPageTitle(pageId) {
    const url = `https://api.notion.com/v1/pages/${pageId}`;
    const headers = getNotionHeaders();

    try {
        const response = UrlFetchApp.fetch(url, {
            method: "get",
            headers,
            muteHttpExceptions: true
        });
        const data = JSON.parse(response.getContentText());
        const titleProp = Object.values(data.properties).find(p => p.type === 'title');
        return titleProp?.title?.[0]?.plain_text || '[No Title]';
    } catch (e) {
        logToSlack(`Error fetching title for page ${pageId}: ${e}`);
        return '[Error Fetching Title]';
    }
}


/**
 * Fetches all pages from a Notion database, handling pagination.
 * @param {string} notionDatabaseId - The ID of the Notion database.
 * @param {Object} headers - HTTP headers, including Notion-Version and Authorization.
 * @return {Array<Object>} An array of Notion page objects.
 */
function fetchAllNotionPages(notionDatabaseId, headers) {
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
function buildNotionPayload(startCursor) {
    return startCursor ? { page_size: NOTION_PAGE_SIZE, start_cursor: startCursor } : { page_size: NOTION_PAGE_SIZE };
}

function extractNotionPageId(notionUrl) {
    const match = notionUrl.match(/([0-9a-f]{32})$/);
    return match ? match[1] : null;
}

function transformNotionTeamPageToRow(page) {
    const properties = page.properties;

    const name = extractPropertyValue(properties["Name"], "Name");
    const status = extractPropertyValue(properties["Status"], "Status");
    const people = extractPropertyValue(properties["People (Current)"], "People (Current)");
    const scrumMaster = extractPropertyValue(properties["Scrum Master"], "Scrum Master");
    const activityEpic = extractPropertyValue(properties["Activity (Epic)"], "Activity (Epic)");
    const dateEpic = extractPropertyValue(properties["Date (Epic)"], "Date (Epic)");

    return [name, status, people, scrumMaster, activityEpic, dateEpic];
}
