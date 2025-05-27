// ==========================
// 🔧 Universal Logger
// ==========================
const LOG_LEVELS = {
    DEBUG: true,
    INFO: true,
    ERROR: true,
    ALWAYS: true,
};

function logDebug(...args) {
    if (LOG_LEVELS.DEBUG) Logger.log("[DEBUG] " + args.join(" "));
}

function logInfo(...args) {
    if (LOG_LEVELS.INFO) Logger.log("[INFO] " + args.join(" "));
}

function logError(...args) {
    if (LOG_LEVELS.ERROR) Logger.log("[ERROR] " + args.join(" "));
}

function logAlways(...args) {
    if (LOG_LEVELS.ALWAYS) Logger.log("[ALWAYS] " + args.join(" "));
}

/**
 * Sends critical or important logs to Slack if configured.
 * @param {string} message
 */
function logToSlack(message) {
    const channelId = PropertiesService.getScriptProperties().getProperty('LOGGING_CHANNEL_ID');
    const token = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
    if (!channelId || !token) return;

    const payload = {
        method: 'post',
        contentType: 'application/json',
        headers: {
            Authorization: 'Bearer ' + token,
        },
        payload: JSON.stringify({ channel: channelId, text: message }),
        muteHttpExceptions: true,
    };

    UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', payload);
}
