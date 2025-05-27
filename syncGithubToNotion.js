function syncGitHubToNotion() {
    const notionApiKey = PropertiesService.getScriptProperties().getProperty('NOTION_API_KEY');
    const githubToken = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
    const repoOwner = 'your-org';
    const repoName = 'your-repo';

    const issues = fetchGitHubIssues(githubToken, repoOwner, repoName);

    issues.forEach(issue => {
        const notionPageId = getSyncedNotionId(issue.id);
        if (notionPageId) {
            updateNotionPage(notionPageId, issue.title, issue.body, notionApiKey);
        } else {
            const pageId = createNotionPage(issue.title, issue.body, notionApiKey);
            storeMapping(pageId, issue.id);
        }
    });
}

/**
 * Google Apps Script: Two-Way Sync between GitHub Issues and Notion
 * This script requires Script Properties to store:
 * - NOTION_API_KEY
 * - GITHUB_TOKEN
 *
 * Make sure you have a Google Sheet with columns:
 * Notion Page ID | GitHub Issue ID | Last Synced | Source of Truth
 */

const GITHUB_OWNER = 'your-org';
const GITHUB_REPO = 'your-repo';
const NOTION_DATABASE_ID = 'your-notion-database-id';
const NOTION_VERSION = '2022-06-28';
const SHEET_NAME = 'Mappings';

function syncGitHubToNotion() {
    const githubToken = getToken('GITHUB_TOKEN');
    const notionToken = getToken('NOTION_API_KEY');
    const issues = fetchGitHubIssues(githubToken);
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();

    issues.forEach(issue => {
        const row = data.find(r => r[1] === String(issue.id));
        if (!row) {
            const pageId = createNotionPage(issue.title, issue.body, notionToken);
            sheet.appendRow([pageId, issue.id, new Date().toISOString(), 'github']);
        } else {
            updateNotionPage(row[0], issue.title, issue.body, notionToken);
        }
    });
}

function syncNotionToGitHub() {
    const githubToken = getToken('GITHUB_TOKEN');
    const notionToken = getToken('NOTION_API_KEY');
    const pages = fetchNotionPages(notionToken);
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();

    pages.forEach(page => {
        const title = getPlainText(page.properties.Name);
        const body = getPlainText(page.properties.Description);
        const row = data.find(r => r[0] === page.id);

        if (!row) {
            const issueId = createGithubIssue(title, body, githubToken);
            sheet.appendRow([page.id, issueId, new Date().toISOString(), 'notion']);
        } else {
            updateGithubIssue(row[1], title, body, githubToken);
        }
    });
}

function fetchGitHubIssues(token) {
    const url = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/issues`;
    const options = {
        headers: { Authorization: `Bearer ${token}` },
        muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(url, options);
    return JSON.parse(response.getContentText());
}

function createGithubIssue(title, body, token) {
    const url = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/issues`;
    const options = {
        method: 'post',
        headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json'
        },
        payload: JSON.stringify({ title, body })
    };
    const res = UrlFetchApp.fetch(url, options);
    return JSON.parse(res.getContentText()).id;
}

function updateGithubIssue(issueId, title, body, token) {
    const url = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/issues/${issueId}`;
    const options = {
        method: 'patch',
        headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json'
        },
        payload: JSON.stringify({ title, body })
    };
    UrlFetchApp.fetch(url, options);
}

function fetchNotionPages(token) {
    const url = `https://api.notion.com/v1/databases/${NOTION_DATABASE_ID}/query`;
    const options = {
        method: 'post',
        headers: {
            Authorization: `Bearer ${token}`,
            'Notion-Version': NOTION_VERSION,
            'Content-Type': 'application/json'
        },
        muteHttpExceptions: true,
        payload: JSON.stringify({ page_size: 100 })
    };
    const res = UrlFetchApp.fetch(url, options);
    return JSON.parse(res.getContentText()).results;
}

function createNotionPage(title, body, token) {
    const url = 'https://api.notion.com/v1/pages';
    const options = {
        method: 'post',
        headers: {
            Authorization: `Bearer ${token}`,
            'Notion-Version': NOTION_VERSION,
            'Content-Type': 'application/json'
        },
        payload: JSON.stringify({
            parent: { database_id: NOTION_DATABASE_ID },
            properties: {
                Name: { title: [{ text: { content: title } }] },
                Description: { rich_text: [{ text: { content: body } }] }
            }
        })
    };
    const res = UrlFetchApp.fetch(url, options);
    return JSON.parse(res.getContentText()).id;
}

function updateNotionPage(pageId, title, body, token) {
    const url = `https://api.notion.com/v1/pages/${pageId}`;
    const options = {
        method: 'patch',
        headers: {
            Authorization: `Bearer ${token}`,
            'Notion-Version': NOTION_VERSION,
            'Content-Type': 'application/json'
        },
        payload: JSON.stringify({
            properties: {
                Name: { title: [{ text: { content: title } }] },
                Description: { rich_text: [{ text: { content: body } }] }
            }
        })
    };
    UrlFetchApp.fetch(url, options);
}

function getPlainText(prop) {
    if (!prop || !prop[prop.type]) return '';
    return prop[prop.type].map(el => el.plain_text).join(' ');
}

function getToken(key) {
    return PropertiesService.getScriptProperties().getProperty(key);
}

function getSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return ss.getSheetByName(SHEET_NAME);
}