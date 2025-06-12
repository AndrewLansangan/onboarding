/**
 * =============================================
 * 🔧 Central Configuration (Multi-Sheet Support)
 * =============================================
 *
 * Provides access to Script Properties for:
 * - Notion API
 * - Slack tokens
 * - Multiple external Google Sheets by logical key (MANDATES, TEAMS, TIMETRACKER, etc.)
 * - Runtime metadata (last run time, etc.)
 */
/**
 * =============================================
 * 📌 Global Constants: Notion, Slack, Sheets
 * =============================================
 */

// ===============================
// 🔐 Notion API Configuration
// ===============================
const NOTION_API_VERSION = '2022-06-28';
const NOTION_PAGE_SIZE = 100;

const NOTION_QUERY_URL = (dbId) => `https://api.notion.com/v1/databases/${dbId}/query`;
const NOTION_PAGE_UPDATE_URL = (pageId) => `https://api.notion.com/v1/pages/${pageId}`;

// ===============================
// 🧠 Notion Property References
// ===============================
const NOTION_PROPERTIES = {
    NAME: 'Name',
    POSITION: 'Position',
    TEAM_CURRENT: 'Team (Current)',
    TEAM_PREVIOUS: 'Team (Previous)',
    HOURS_INITIAL: 'Hours (Initial)',
    HOURS_CURRENT: 'Hours (Current)',
    AVAILABILITY: 'Availability (avg h/w)',
    CREATED_PROFILE: 'Created (Profile)',
    EMAIL: 'Email (Org)',
    STATUS: 'Status',
    MANDATE_DATE: 'Mandate (Date)',
    NOTION_URL: 'Notion Page URL'
};

const PEOPLE_RELATION_PROPERTY = 'People Directory (Sync)'; // For linking public/internal directories

// ===============================
// 📊 Google Sheets Configuration
// ===============================
const SHEET_NAMES = {
    MANDATES: 'Mandates',
    TEAM_DIRECTORY: 'Team Directory'
};

const SHEET_COLUMNS = {
    EMAIL: 'Email (Org)',
    STATUS: 'Mandate (Status)',
    TEAM: 'Team (Current)',
    LAST_UPDATE: 'Last Update',
    HOURS_DECIMAL: 'Hours (decimal)',
    CREATED_PROFILE: 'Created (Profile)'
};

const SHEET_HEADERS = {
    MANDATES: [
        "Greybox ID", "Mandate (Status)", "Name", "Position", "Team (Current)", "Team (Previous)",
        "Mandate (Date)", "Hours (Initial)", "Hours (Current)", "Availability (avg h/w)",
        "Email (Org)", "Created (Profile)", "Notion Page URL"
    ],
    TEAM_DIRECTORY: [
        "Name", "Status", "People", "Scrum Master", "Activity (Epic)", "Date (Epic)"
    ]
};

// ===============================
// 💼 Slack Custom Profile Fields
// ===============================
const SLACK_PROFILE_FIELDS = {
    Position: "Xf06JZK27DRA",
    Team: "Xf03V366R202",
    NotionPageURL: "Xf06JGJMBZPZ",
    MandateStatus: "Xf0759PXS7BP",
    Availability: "Xf074Y4V1KHV",
    CreatedProfile: "Xf075CJ4STEF",
    TimeTrackerLastUpdate: "Xf07HUS9SSC",
    TimeTrackerTotal: "Xf07GZDPHHV4"
};

const SLACK_URLS = {
    USERGROUP_LIST: 'https://slack.com/api/usergroups.list',
    PROFILE_SET: 'https://slack.com/api/users.profile.set',
    USER_LOOKUP_BY_EMAIL: 'https://slack.com/api/users.lookupByEmail',
    CHAT_POST_MESSAGE: 'https://slack.com/api/chat.postMessage'
};

// ===============================
// 🧩 Script Properties & Fallbacks
// ===============================
const NOTION_DB_IDS = {
    PEOPLE: '3cf44b088a8f4d6b8abc989353abcdb1',
    INTERNAL_PEOPLE: '47fbed712f3e4558b032edb9ec081f00',
    TEAM: '70779d3ee3cf467b9b86171acabc3321'
};

const TEAM_DB_ID_PROP_NAME = 'NOTION_TEAM_DB_ID';
const NOTIFIED_TEAMS_PROPERTY_KEY = 'notifiedCompletedTeamIds';

const COMPLETED_STATUS_VALUE = 'Completed';
const SILENCE_TEAMS_WITHOUT_SCRUM = false;

// ===============================
// 🧪 Debug
// ===============================
const DEBUG_MODE = true;

let _cachedConfig = null;

function getScriptConfig() {
    if (_cachedConfig) return _cachedConfig;

    const props = PropertiesService.getScriptProperties();

    return (_cachedConfig = {
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
            REGISTRY: {
                MANDATES: {
                    id: props.getProperty('SHEET_ID_MANDATES'),
                    name: props.getProperty('SHEET_NAME_MANDATES') || 'Mandates'
                },
                TEAMS: {
                    id: props.getProperty('SHEET_ID_TEAMS'),
                    name: props.getProperty('SHEET_NAME_TEAMS') || 'Teams'
                },
                TIMETRACKER: {
                    id: props.getProperty('SHEET_ID_TIMETRACKER'),
                    name: props.getProperty('SHEET_NAME_TIMETRACKER') || 'TimeTracker'
                }
            }
        },
        META: {
            LAST_RUN_TIME: props.getProperty('LAST_RUN_TIME'),
            NOTIFIED_TEAMS_PROPERTY_KEY: 'notifiedCompletedTeamIds'
        }
    });
}

const Config = {
    // --- Notion ---
    get notionHeaders() {
        const { API_KEY } = getScriptConfig().NOTION;
        if (!API_KEY) throw new Error("❌ NOTION_API_KEY is missing.");
        return {
            Authorization: `Bearer ${API_KEY}`,
            'Notion-Version': NOTION_API_VERSION,
            'Content-Type': 'application/json'
        };
    },
    get notionDbPeople() {
        return getScriptConfig().NOTION.DB_ID_PEOPLE;
    },
    get notionDbInternalPeople() {
        return getScriptConfig().NOTION.DB_ID_INTERNAL_PEOPLE;
    },
    get notionDbTeam() {
        return getScriptConfig().NOTION.DB_ID_TEAM;
    },

    // --- Slack ---
    get slackBotToken() {
        return getScriptConfig().SLACK.BOT_TOKEN;
    },
    get slackUserToken() {
        return getScriptConfig().SLACK.USER_TOKEN;
    },
    get slackLoggingChannelId() {
        return getScriptConfig().SLACK.LOGGING_CHANNEL_ID;
    },getSlackToken(purpose = 'default') {
            const config = getScriptConfig();
            const normalized = (purpose || '').toLowerCase();

            switch (normalized) {
                case 'user':
                    return config.SLACK.USER_TOKEN;
                case 'bot':
                    return config.SLACK.BOT_TOKEN;
                case 'default':
                default:
                    return config.SLACK.BOT_TOKEN || config.SLACK.USER_TOKEN;
            }
    },

    // --- Sheets ---
    getSheetConfig(key) {
        const registry = getScriptConfig().SHEETS.REGISTRY;
        const sheet = registry[key];
        if (!sheet || !sheet.id || !sheet.name) {
            throw new Error(`❌ Sheet config missing or incomplete for key: ${key}`);
        }
        return sheet;
    },

    getSheetInstance(key) {
        const { id, name } = this.getSheetConfig(key);
        return SpreadsheetApp.openById(id).getSheetByName(name);
    },

    // --- Meta ---
    get lastRunTime() {
        const raw = getScriptConfig().META.LAST_RUN_TIME;
        const date = raw ? new Date(raw) : new Date("2000-01-01");
        return isNaN(date.getTime()) ? new Date("2000-01-01") : date;
    },
    set lastRunTime(date) {
        if (!(date instanceof Date) || isNaN(date.getTime())) {
            throw new Error("❌ Invalid date passed to Config.lastRunTime");
        }
        PropertiesService.getScriptProperties().setProperty(
            'LAST_RUN_TIME',
            date.toISOString()
        );
    },
    get notifiedTeamsKey() {
        return getScriptConfig().META.NOTIFIED_TEAMS_PROPERTY_KEY;
    }
};
function getNotionHeaders() {
    return {
        Authorization: `Bearer ${notionApiKey}`,
        'Content-Type': 'application/json',
        'Notion-Version': '2022-06-28',
    };
}