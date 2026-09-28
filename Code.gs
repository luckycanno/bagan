/**
 * Jong Canno Cup Season 2 - 2026 (70 Teams Bracket)
 * Backend Google Apps Script (Code.gs)
 * 
 * GitHub Repository Integration Script
 */

const CONFIG = {
  ADMIN_PASSWORD: "qwerty",
  TOTAL_TEAMS: 70,
  SHEET_NAME_TEAMS: "Teams",
  SHEET_NAME_MATCHES: "Matches",
  SHEET_NAME_SETTINGS: "Settings"
};

/**
 * Handles HTTP GET requests to serve the HTML Web App or JSON API responses.
 * @param {Object} e - Event object containing URL query parameters.
 * @return {HtmlOutput|TextOutput} Rendered web app or JSON response.
 */
function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action;

    // Handle API requests
    if (action === "getData") {
      return createJsonResponse({ status: "success", data: getTournamentData() });
    }
    
    if (action === "verifyAdmin") {
      const pass = e.parameter.password || "";
      const isValid = pass === CONFIG.ADMIN_PASSWORD;
      return createJsonResponse({ status: isValid ? "success" : "error", authenticated: isValid });
    }

    // Default action: Render HTML Web App UI
    const template = HtmlService.createTemplateFromFile("index");
    return template.evaluate()
      .setTitle("Jong Canno Cup Season 2 - Bagan 70 Tim")
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    return createJsonResponse({ status: "error", message: error.toString() });
  }
}

/**
 * Handles HTTP POST requests for saving team rosters, match scores, and resets.
 * @param {Object} e - Event object containing POST payload body.
 * @return {TextOutput} JSON response object.
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Payload tidak ditemukan.");
    }

    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const password = payload.password;

    // Verify Admin Password for write operations
    if (password !== CONFIG.ADMIN_PASSWORD) {
      return createJsonResponse({ status: "error", message: "Kata sandi admin tidak valid!" });
    }

    // Dispatch Actions
    if (action === "saveTeams") {
      saveTeamsToSheet(payload.teams);
      return createJsonResponse({ status: "success", message: "Daftar tim berhasil disimpan." });
    }

    if (action === "updateMatch") {
      saveMatchScoreToSheet(payload.matchKey, payload.score1, payload.score2, payload.winner);
      return createJsonResponse({ status: "success", message: "Skor pertandingan berhasil diperbarui." });
    }

    if (action === "resetTournament") {
      resetTournamentSheet();
      return createJsonResponse({ status: "success", message: "Bagan pertandingan berhasil direset." });
    }

    return createJsonResponse({ status: "error", message: "Aksi tidak dikenali." });

  } catch (error) {
    return createJsonResponse({ status: "error", message: error.toString() });
  }
}

/**
 * Auto-initializes Google Sheet tabs (Teams, Matches) if they do not exist.
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Teams Sheet Setup
  let teamsSheet = ss.getSheetByName(CONFIG.SHEET_NAME_TEAMS);
  if (!teamsSheet) {
    teamsSheet = ss.insertSheet(CONFIG.SHEET_NAME_TEAMS);
    teamsSheet.appendRow(["ID", "Nama Tim"]);
    
    // Fill default 70 teams
    const defaultTeams = [];
    for (let i = 1; i <= CONFIG.TOTAL_TEAMS; i++) {
      defaultTeams.push([i, `Tim ${i}`]);
    }
    teamsSheet.getRange(2, 1, CONFIG.TOTAL_TEAMS, 2).setValues(defaultTeams);
    teamsSheet.getRange("A1:B1").setFontWeight("bold").setBackground("#0a1026").setFontColor("#00f0ff");
  }

  // 2. Matches Sheet Setup
  let matchesSheet = ss.getSheetByName(CONFIG.SHEET_NAME_MATCHES);
  if (!matchesSheet) {
    matchesSheet = ss.insertSheet(CONFIG.SHEET_NAME_MATCHES);
    matchesSheet.appendRow(["MatchKey", "Score1", "Score2", "Winner"]);
    matchesSheet.getRange("A1:D1").setFontWeight("bold").setBackground("#0a1026").setFontColor("#00f0ff");
  }
}

/**
 * Reads teams and match results from the active Google Sheet.
 * @return {Object} Tournament dataset containing team list and match dictionary.
 */
function getTournamentData() {
  setupSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Read Teams
  const teamsSheet = ss.getSheetByName(CONFIG.SHEET_NAME_TEAMS);
  const teamsData = teamsSheet.getRange(2, 2, CONFIG.TOTAL_TEAMS, 1).getValues();
  const teamNames = teamsData.map(row => row[0] ? row[0].toString() : "");

  // Read Matches
  const matchesSheet = ss.getSheetByName(CONFIG.SHEET_NAME_MATCHES);
  const lastRow = matchesSheet.getLastRow();
  const matchState = {};

  if (lastRow > 1) {
    const matchesData = matchesSheet.getRange(2, 1, lastRow - 1, 4).getValues();
    matchesData.forEach(row => {
      const matchKey = row[0];
      if (matchKey) {
        matchState[matchKey] = {
          score1: Number(row[1]) || 0,
          score2: Number(row[2]) || 0,
          winner: row[3] !== "" && row[3] !== null ? Number(row[3]) : null
        };
      }
    });
  }

  return {
    teamNames: teamNames,
    matchState: matchState
  };
}

/**
 * Saves or updates team list in Google Sheet.
 * @param {Array<string>} teams - Array of 70 team names.
 */
function saveTeamsToSheet(teams) {
  if (!Array.isArray(teams)) return;
  setupSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME_TEAMS);

  const formattedValues = [];
  for (let i = 0; i < CONFIG.TOTAL_TEAMS; i++) {
    formattedValues.push([i + 1, teams[i] || `Tim ${i + 1}`]);
  }

  sheet.getRange(2, 1, CONFIG.TOTAL_TEAMS, 2).setValues(formattedValues);
}

/**
 * Saves individual match score and winner status to Google Sheet.
 * @param {string} matchKey - Identifier for match (e.g., "L_1_0", "FINAL").
 * @param {number} score1 - Score of Team 1.
 * @param {number} score2 - Score of Team 2.
 * @param {number|null} winner - Winner index (1 or 2).
 */
function saveMatchScoreToSheet(matchKey, score1, score2, winner) {
  setupSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME_MATCHES);
  const lastRow = sheet.getLastRow();

  let targetRow = -1;

  if (lastRow > 1) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (keys[i][0] === matchKey) {
        targetRow = i + 2;
        break;
      }
    }
  }

  if (targetRow === -1) {
    targetRow = Math.max(2, lastRow + 1);
  }

  sheet.getRange(targetRow, 1, 1, 4).setValues([[matchKey, score1, score2, winner]]);
}

/**
 * Clears all match results from the Google Sheet.
 */
function resetTournamentSheet() {
  setupSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME_MATCHES);
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  }
}

/**
 * Helper to build JSON HTTP outputs.
 * @param {Object} data - JavaScript object to return as JSON response.
 * @return {TextOutput}
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
