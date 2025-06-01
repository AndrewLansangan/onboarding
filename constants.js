/**
 * =============================================
 * 📌 Global Constants: Notion, Slack, Sheets
 * =============================================
 *
 * Shared constants used across the automation scripts.
 * Provides structure for:
 * - Sheet column references
 * - Notion property names
 * - Notion DB/API setup
 * - Slack field mapping
 * - General configuration metadata
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

const PEOPLE_RELATION_PROPERTY = 'People Directory (Sync)'; // for relation updates

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

const { SLACK } = getScriptConfig();
const SLACK_USER_TOKEN = SLACK.USER_TOKEN;
const SLACK_BOT_TOKEN = SLACK.BOT_TOKEN;
const SLACK_LOGGING_CHANNEL_ID = SLACK.LOGGING_CHANNEL_ID;