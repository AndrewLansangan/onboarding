/**
 * Reads the sheet and creates a Map of team names to arrays of user emails
 * Filters out users with invalid statuses (To Verify, Completed, Archived)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Google Sheet to read from
 * @return {Map<string, string[]>}  Map with team names as keys and arrays of emails as values
 */
/**
 * Reads user data from a Google Sheet and groups users by their active team names.
 * Filters out users with invalid mandate statuses and already processed entries based on `lastRunTime`.
 *
 * @return {Map<string, string[]>} - A map of team names to arrays of user emails.
 * @param dateStr
 */

function parseCustomDate(dateStr) {
    let [day, month, year] = String(dateStr).split('/').map(Number);
    return new Date(year, month - 1, day);
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

/**
 * External dependency (shared utility).
 * Defined in: `global-utils.gs` or similar.
 *
 * @function fetchWithRetries
 * @param {string} url
 * @param {Object} options
 * @param maxRetries
 * @returns {HTTPResponse}
 */
// fetchWithRetry() is assumed to be globally available

function fetchWithRetries(url, options, maxRetries = 3) {
    for (let attempt = 0; attempt <= maxRetries; attempt++) {
        const response = UrlFetchApp.fetch(url, options);
        const data = JSON.parse(response.getContentText());

        if (data.ok) return response; // ✅ Success
        else if (shouldRetry(data.error)) {
            Utilities.sleep(2 ** attempt * 1000); // wait 1s, 2s, 4s...
        } else {
            break; // non-retriable error
        }
    }
    return lastFailedResponse;
}

function formatDateTimeForSheet(dateValue) {
    if (!dateValue) return "";

    // Check if the dateValue is a Date object, convert it to the desired string format
    if (Object.prototype.toString.call(dateValue) === '[object Date]') {
        return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    }

    // If dateValue is already a string, assume it's in the correct format and return it as is
    if (typeof dateValue === 'string') {
        return dateValue;
    }

    return ""; // Return an empty string if dateValue is not valid
}

function convertDate(isoDateString) {
    const date = new Date(isoDateString);
    const options = { year: 'numeric', month: 'long', day: 'numeric' };
    return date.toLocaleDateString('en-US', options);
}

// =============================================================================
// 🔧 UNIVERSAL LOGGER — Console + Slack + Type Inspection
// =============================================================================

/* -------------------------------------------------------------------------
 * SECTION 1 — 🔁 Log Level Configuration
 * ----------------------------------------------------------------------- */

/**
 * Controls which log levels are enabled.
 * Toggle these flags to enable/disable specific log types.
 */
const LOG_LEVELS = {
    DEBUG: true,
    INFO: true,
    ERROR: true,
    ALWAYS: true,
};

/* -------------------------------------------------------------------------
 * SECTION 2 — 🪵 Console Logging Functions
 * ----------------------------------------------------------------------- */

/**
 * Logs debug-level messages to Apps Script console.
 * @param {...any} args - Arguments to log.
 */
function logDebug(...args) {
    if (LOG_LEVELS.DEBUG) Logger.log("[DEBUG] " + args.join(" "));
}

/**
 * Logs info-level messages to Apps Script console.
 * @param {...any} args - Arguments to log.
 */
function logInfo(...args) {
    if (LOG_LEVELS.INFO) Logger.log("[INFO] " + args.join(" "));
}

/**
 * Logs error-level messages to Apps Script console.
 * @param {...any} args - Arguments to log.
 */
function logError(...args) {
    if (LOG_LEVELS.ERROR) Logger.log("[ERROR] " + args.join(" "));
}

/**
 * Logs messages regardless of level (e.g. critical logs).
 * @param {...any} args - Arguments to log.
 */
function logAlways(...args) {
    if (LOG_LEVELS.ALWAYS) Logger.log("[ALWAYS] " + args.join(" "));
}


/* -------------------------------------------------------------------------
 * SECTION 3 — 📡 Slack Logging
 * ----------------------------------------------------------------------- */

/**
 * Sends important logs to a Slack channel using bot token.
 * Requires LOGGING_CHANNEL_ID and SLACK_BOT_TOKEN in ScriptProperties.
 * @param {string} message - Message to send to Slack.
 */
function logToSlack(message) {
    const props = Config
    const channelId = props?.LOGGING_CHANNEL_ID
    const token = props.getSlackToken?.("user")

    if (!channelId || !token) return;

    const payload = {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + token },
        payload: JSON.stringify({ channel: channelId, text: message }),
        muteHttpExceptions: true,
    };

    UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', payload);
}


/* -------------------------------------------------------------------------
 * SECTION 4 — 🧪 Debug Utilities
 * ----------------------------------------------------------------------- */

/**
 * Identifies and logs empty (null or blank) columns in a row.
 * @param {Array<any>} rowData - Array of row values.
 */
function logMissingColumns(rowData) {
    const empties = rowData
        .map((val, idx) => val == null || val === '' ? idx + 1 : null)
        .filter(Boolean);
    if (empties.length) logToSlack(`Empty columns: ${empties.join(', ')}`);
}

/**
 * Logs each property key with its detected type.
 * Useful for debugging Notion or API objects.
 * @param {Object} properties - Property map with possible type metadata.
 */
function logTypes(properties) {
    Object.entries(properties).forEach(([key, prop]) =>
        Logger.log(`Property: ${key}, Type: ${prop?.type ?? 'undefined'}`));
}
