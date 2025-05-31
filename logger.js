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
    const props = getScriptConfig()
    const channelId = props.LOGGING_CHANNEL_ID
    const token = props.SLACK_BOT_TOKEN

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
