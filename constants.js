/**
 * =============================================
 * 🧩 Global Constants & Configuration Map
 * =============================================
 *
 * Centralized reference for shared constants used
 * across Notion, Slack, and Google Sheets integrations.
 * This improves consistency, prevents hardcoding, and
 * simplifies refactoring.
 *
 * ---------------------------------------------
 * 📊 Sheet Column Identifiers
 * ---------------------------------------------
 * - `SHEET_COLUMNS`: Used to access values from rows in Google Sheets
 * - `MANDATES_SHEET_COLUMNS`: Expected column structure for "Mandates" sheet
 * - `TEAM_DIRECTORY_COLUMNS`: Expected column structure for "Team Directory" sheet
 *
 * ---------------------------------------------
 * 🧠 Notion Properties
 * ---------------------------------------------
 * - `NOTION_PROPERTIES`: Common property names from synced Notion databases
 * - `TEAM_NAME_PROPERTY`, `SCRUM_MASTER_RELATION_PROPERTY`: Specific to Team Directory DB
 * - `SM_EMAIL_PROPERTY_IN_PEOPLE_DB`: Used for linking to People DB
 * - `PEOPLE_RELATION_PROPERTY`: Relation field used for syncing people references
 *
 * ---------------------------------------------
 * 🔐 Notion Configuration
 * ---------------------------------------------
 * - `NOTION_VERSION`: API version used in headers
 * - `NOTION_QUERY_URL`, `NOTION_PAGE_UPDATE_URL`: Endpoint builders
 * - `NOTION_HEADERS`: Standard headers for all Notion API requests
 * - `NOTION_PAGE_SIZE`, `MAX_PAGE_SIZE`: Pagination limits
 * - `NOTION_PEOPLE_DB_ID`, `NOTION_INTERNAL_PEOPLE_DB_ID`, `NOTION_TEAM_DB_ID`: Hardcoded fallbacks (usually loaded via `getScriptConfig()`)
 *
 * ---------------------------------------------
 * 📅 Sheet Names & Statuses
 * ---------------------------------------------
 * - `MANDATES_SHEET_NAME`, `TEAM_SHEET_NAME`: Sheet tab names in spreadsheet
 * - `COMPLETED_STATUS_VALUE`: Used to filter out completed mandates
 * - `SILENCE_TEAMS_WITHOUT_SCRUM`: Toggle for team notifications
 *
 * ---------------------------------------------
 * 📎 Slack Custom Profile Fields
 * ---------------------------------------------
 * - `SLACK_PROFILE_FIELDS`: Field IDs for updating Slack user profiles
 * - `SLACK_USERGROUP_LIST_URL`: URL for reading Slack user groups
 *
 * ---------------------------------------------
 * 🗝️ Script Properties Keys
 * ---------------------------------------------
 * - `TEAM_DB_ID_PROP_NAME`: Property key to fetch Team Directory DB ID
 * - `NOTIFIED_TEAMS_PROPERTY_KEY`: Stores which teams have been notified
 */

// Column names in sheets
const SHEET_COLUMNS = {
    EMAIL: 'Email (Org)',
    STATUS: 'Mandate (Status)',
    TEAM: 'Team (Current)',
    LAST_UPDATE: 'Last Update',
    HOURS_DECIMAL: 'Hours (decimal)',
    CREATED_PROFILE: 'Created (Profile)'
};

// Notion property names
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

// API settings
const NOTION_VERSION = '2022-06-28';
const MAX_PAGE_SIZE = 100;

// --- Constants ---
// Property names in the *Team Directory* Database
const TEAM_DB_ID_PROP_NAME = "NOTION_TEAM_DB_ID";

const NOTIFIED_TEAMS_PROPERTY_KEY = "notifiedCompletedTeamIds"; // Key for storing notified IDs in Script Properties

 // --- Slack Field IDs ---
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

// ============================
// 🔧 Global Configuration
// ============================

// Sheet Names
const MANDATES_SHEET_NAME = 'Mandates';
const TEAM_SHEET_NAME = 'Team Directory';

// Notion Database IDs
const NOTION_PEOPLE_DB_ID = '3cf44b088a8f4d6b8abc989353abcdb1'; // People Directory
const NOTION_INTERNAL_PEOPLE_DB_ID = '47fbed712f3e4558b032edb9ec081f00'; // Internal HR DB
const NOTION_TEAM_DB_ID = '70779d3ee3cf467b9b86171acabc3321'; // Team Directory

// API Settings
const NOTION_HEADERS = {
    'Authorization': `Bearer ${getScriptConfig().NOTION_API_KEY}`,
    'Notion-Version': NOTION_VERSION,
    'Content-Type': 'application/json',
};

const SHEET_NAME = 'Mandates'; // Change to your actual sheet name
// --- Default Values ---
const COMPLETED_STATUS_VALUE = 'Completed';
const SILENCE_TEAMS_WITHOUT_SCRUM = false;

// --- Team Directory Property Keys ---
const TEAM_NAME_PROPERTY = 'Name';
const TEAM_STATUS_PROPERTY = 'Status';
const SCRUM_MASTER_RELATION_PROPERTY = 'Scrum Master';

// --- People Directory Property Keys ---
const SM_EMAIL_PROPERTY_IN_PEOPLE_DB = 'Email (Org)';

// --- Slack API ---
const SLACK_USERGROUP_LIST_URL = 'https://slack.com/api/usergroups.list';

// --- Notion API ---
const NOTION_QUERY_URL = (dbId) => `https://api.notion.com/v1/databases/${dbId}/query`;
const NOTION_PAGE_UPDATE_URL = (pageId) => `https://api.notion.com/v1/pages/${pageId}`;

// --- Default Parameters ---
const NOTION_PAGE_SIZE = 100;
const PEOPLE_RELATION_PROPERTY = "People Directory (Sync)";

const TEAM_DIRECTORY_COLUMNS = [
    "Name",
    "Status",
    "People",
    "Scrum Master",
    "Activity (Epic)",
    "Date (Epic)"
];

const MANDATES_SHEET_COLUMNS = [
    "Greybox ID", "Mandate (Status)", "Name", "Position", "Team (Current)", "Team (Previous)",
    "Mandate (Date)", "Hours (Initial)", "Hours (Current)", "Availability (avg h/w)",
    "Email (Org)", "Created (Profile)", "Notion Page URL"
];