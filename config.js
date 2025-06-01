/**
 * =============================================
 * 🔗 Notion People Directory Linking Script
 * =============================================
 *
 * This script links people from the public-facing People Directory
 * to the internal HR-only People Directory in Notion, using email
 * address matching. The relation is stored as a Notion relation property.
 *
 * ---------------------------------------------
 * 🗃️ Databases Involved:
 * ---------------------------------------------
 * - People Directory (Public):
 *   https://www.notion.so/grey-box/People-da052a0ffb3a428d8e7013c540c42665
 * - People Directory (Internal):
 *   https://www.notion.so/grey-box/47fbed712f3e4558b032edb9ec081f00?v=2d969f6b09084313823bca813f39db69
 *
 * ---------------------------------------------
 * ⚙️ Main Functional Flow:
 * ---------------------------------------------
 * - Loads configuration using `getScriptConfig`
 * - Fetches pages from both databases
 * - Matches by email address
 * - Updates a relation property to link the matched internal pages
 *
 * ---------------------------------------------
 * 🔐 Configuration:
 * Script Properties expected:
 * - NOTION_API_KEY
 * - NOTION_PEOPLE_DB_ID
 * - NOTION_INTERNAL_PEOPLE_DB_ID
 */

const NOTION_API_VERSION = '2022-06-28';

/**
 * Loads config values from Script Properties.
 * Grouped for Notion, Slack, Sheets, Meta.
 */
function getScriptConfig() {
    const props = PropertiesService.getScriptProperties();
    return {
        NOTION: {
            API_KEY: props.getProperty('NOTION_API_KEY'),
            DB_ID_PEOPLE: props.getProperty('NOTION_PEOPLE_DB_ID'),
            DB_ID_INTERNAL_PEOPLE: props.getProperty('NOTION_INTERNAL_PEOPLE_DB_ID'),
            DB_ID_TEAM: props.getProperty('NOTION_TEAM_DB_ID')
        },
        SLACK: {
            BOT_TOKEN: props.getProperty('SLACK_BOT_TOKEN'),
            USER_TOKEN: props.getProperty('SLACK_USER_TOKEN'),
            LOGGING_CHANNEL_ID: props.getProperty('LOGGING_CHANNEL_ID')
        },
        SHEETS: {
            SPREADSHEET_ID: props.getProperty('SPREADSHEET_ID'),
            SHEET_NAME: props.getProperty('SHEET_NAME')
        },
        META: {
            LAST_RUN_TIME: props.getProperty('LAST_RUN_TIME'),
            NOTIFIED_TEAMS_PROPERTY_KEY: 'notifiedCompletedTeamIds'
        }
    };
}

/**
 * Initializes Notion headers and database IDs.
 * Returns object containing API headers and target DBs.
 */
function initializeConfig() {
    const config = getScriptConfig().NOTION;
    const API_KEY = config.API_KEY;
    const DB_ID_PEOPLE = config.DB_ID_PEOPLE || NOTION_DB_IDS.PEOPLE;
    const DB_ID_INTERNAL_PEOPLE = config.DB_ID_INTERNAL_PEOPLE || NOTION_DB_IDS.INTERNAL_PEOPLE;

    if (!API_KEY || !DB_ID_PEOPLE || !DB_ID_INTERNAL_PEOPLE) {
        logError("❌ Missing configuration values in Script Properties or fallback.");
        return null;
    }

    logInfo("✅ Notion configuration loaded.");
    logDebug(`People DB: ${DB_ID_PEOPLE}, Internal DB: ${DB_ID_INTERNAL_PEOPLE}`);

    return {
        headers: {
            Authorization: `Bearer ${API_KEY}`,
            'Notion-Version': NOTION_API_VERSION,
            'Content-Type': 'application/json',
        },
        databaseId1: DB_ID_PEOPLE,
        databaseId2: DB_ID_INTERNAL_PEOPLE,
    };
}
