/**
 * ==============================================================================
 * GOOGLE APPS SCRIPT (GAS) BACKEND ENGINE - IJSHOCKEY COMPETITIE DATABASE
 * ==============================================================================
 * Instructies:
 * 1. Open uw Google Sheet
 * 2. Ga naar Extensies > Apps Script
 * 3. Plak deze code in 'Code.gs'
 * 4. Klik op "Uitvoeren" > "setupDatabaseSheets" om automatisch alle tabbladen en koppen aan te maken.
 * 5. Klik op "Implementeren" > "Nieuwe implementatie" > Web-app.
 * 6. Toegang verlenen aan: "Iedereen" (Everyone).
 * 7. Kopieer de gegenereerde Web App URL en plak deze in het instellingenmenu van de webapplicatie!
 */

const SHEET_NAMES = {
  DIVISIONS: 'Divisions',
  TEAMS: 'Teams',
  PLAYERS: 'Players',
  MATCHES: 'Matches',
  MATCH_EVENTS: 'MatchEvents',
  ANNOUNCEMENTS: 'Announcements',
  FARM_AFFILIATIONS: 'FarmAffiliations',
  LOAN_RECORDS: 'LoanRecords',
  CUSTOM_FIELDS: 'CustomFields'
};

function setupDatabaseSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create Divisions Sheet
  getOrCreateSheet(ss, SHEET_NAMES.DIVISIONS, ['ID', 'Name', 'Code', 'Description', 'Level']);
  
  // Create Teams Sheet (includes farm team references)
  getOrCreateSheet(ss, SHEET_NAMES.TEAMS, ['ID', 'Name', 'ShortName', 'DivisionID', 'PrimaryColor', 'SecondaryColor', 'IsFarmTeam', 'ParentTeamID', 'HomeRink', 'City']);
  
  // Create Players Sheet (includes Leenspeler & Call-up tracking)
  getOrCreateSheet(ss, SHEET_NAMES.PLAYERS, ['ID', 'FirstName', 'LastName', 'JerseyNumber', 'Position', 'PrimaryTeamID', 'DivisionID', 'IsLoanPlayer', 'LoanOriginClub', 'LoanNotes', 'IsFarmPlayer', 'ParentTeamID', 'CallUpGamesPlayed', 'CallUpLimit']);
  
  // Create Matches Sheet
  getOrCreateSheet(ss, SHEET_NAMES.MATCHES, ['ID', 'DivisionID', 'HomeTeamID', 'AwayTeamID', 'DateTime', 'Venue', 'Status', 'HomeScore', 'AwayScore', 'IsOvertime', 'IsShootout', 'Referees', 'Spectators']);
  
  // Create MatchEvents Sheet
  getOrCreateSheet(ss, SHEET_NAMES.MATCH_EVENTS, ['ID', 'MatchID', 'Timestamp', 'Period', 'GameTimeSeconds', 'TeamID', 'Type', 'ScorerID', 'Assist1ID', 'Assist2ID', 'GoalType', 'PenaltyPlayerID', 'Infraction', 'PenaltyMinutes', 'GoalieID', 'LoanPlayerName', 'Notes']);
  
  // Create Announcements Sheet
  getOrCreateSheet(ss, SHEET_NAMES.ANNOUNCEMENTS, ['ID', 'Title', 'Content', 'Author', 'Date', 'Category', 'IsPinned']);

  SpreadsheetApp.getUi().alert('✅ IJshockey Database tabbladen en kolomstructuren succesvol aangemaakt!');
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1E293B").setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// HTTP GET: Direct JSON query interface
function doGet(e) {
  try {
    const action = e.parameter.action || 'GET_ALL';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (action === 'PING') {
      return responseJSON({ status: 'ok', message: 'GAS Hockey DB Active', timestamp: new Date().toISOString() });
    }
    
    const db = {
      divisions: readSheetData(ss, SHEET_NAMES.DIVISIONS),
      teams: readSheetData(ss, SHEET_NAMES.TEAMS),
      players: readSheetData(ss, SHEET_NAMES.PLAYERS),
      matches: readSheetData(ss, SHEET_NAMES.MATCHES),
      events: readSheetData(ss, SHEET_NAMES.MATCH_EVENTS),
      announcements: readSheetData(ss, SHEET_NAMES.ANNOUNCEMENTS)
    };
    
    return responseJSON({ status: 'success', data: db });
  } catch (err) {
    return responseJSON({ status: 'error', message: err.toString() });
  }
}

// HTTP POST: Batch write mutation handler
function doPost(e) {
  try {
    const contents = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (contents.action === 'BATCH_SYNC' && contents.queueItems) {
      let syncCount = 0;
      contents.queueItems.forEach(item => {
        processQueueItem(ss, item);
        syncCount++;
      });
      return responseJSON({ status: 'success', syncedCount: syncCount, message: 'Batch sync verwerkt' });
    }
    
    return responseJSON({ status: 'error', message: 'Ongeldige actie' });
  } catch (err) {
    return responseJSON({ status: 'error', message: err.toString() });
  }
}

function processQueueItem(ss, item) {
  const table = item.table;
  const sheetName = getSheetNameForTable(table);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  
  const payload = item.payload;
  if (item.action === 'INSERT') {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const row = headers.map(h => payload[h.toLowerCase()] !== undefined ? payload[h.toLowerCase()] : (payload[h] || ''));
    sheet.appendRow(row);
  }
}

function getSheetNameForTable(table) {
  switch(table.toLowerCase()) {
    case 'divisions': return SHEET_NAMES.DIVISIONS;
    case 'teams': return SHEET_NAMES.TEAMS;
    case 'players': return SHEET_NAMES.PLAYERS;
    case 'matches': return SHEET_NAMES.MATCHES;
    case 'events': return SHEET_NAMES.MATCH_EVENTS;
    case 'announcements': return SHEET_NAMES.ANNOUNCEMENTS;
    default: return table;
  }
}

function readSheetData(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const rowObj = {};
    headers.forEach((h, colIdx) => {
      rowObj[h] = data[i][colIdx];
    });
    rows.push(rowObj);
  }
  return rows;
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function myFunction() {
  
}
