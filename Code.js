// FUTURE DEVELOPMENT NOTES
// 1. create a audit of a Grid based on pairs of members and # of months apart
// 2. create an orphans list to give options of where unseated members can sit.

/** CONSTANTS - Array Column Indices */
const HOST_COLUMNS = {
  CODE: 0,
  COUNT: 1,
  SEATS: 2,
  ORDER: 3
};

const GUEST_COLUMNS = {
  CODE: 0,
  COUNT: 1,
  SEATED: 2,
  ORDER: 3
};

const GRID_COLUMNS = {
  HOUSE: 0,
  SEATS: 1,
  SEATED: 2,
  HOST: 3,
  GUEST_1: 4,
  GUEST_2: 5,
  GUEST_3: 6,
  GUEST_4: 7,
  GUEST_5: 8
};

const CONNECTION_COLUMNS = {
  PAIR: 0,
  MEMBER_1: 1,
  MEMBER_2: 2,
  YEAR: 3,
  DATE: 4,
  HOST_ROLE: 5,
  MONTHS_APART: 6
};

/** CACHE HELPER FUNCTIONS */
function buildConnectionsCache() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = spreadSheet.getSheetByName('connectionsDB');
  const arrDB = dbSheet.getDataRange().getValues().slice();
  arrDB.splice(0, 1); // Remove header

  // Build Map: key = "member1-member2", value = months apart
  const connectionsMap = {};
  // Build connection counts Map: key = "memberCode", value = count of connections
  const connectionCounts = {};

  for (let i = 0; i < arrDB.length; i++) {
    const key = arrDB[i][CONNECTION_COLUMNS.PAIR];
    const months = arrDB[i][CONNECTION_COLUMNS.MONTHS_APART];
    const member = key.split('-')[0]; // Extract first member code

    // Store only the minimum months for each pair (since array is sorted)
    if (!connectionsMap[key]) {
      connectionsMap[key] = months;
    }

    // Count connections per member
    connectionCounts[member] = (connectionCounts[member] || 0) + 1;
  }

  // Store in cache using chunks (CacheService has 100KB limit per entry)
  const cache = CacheService.getScriptCache();
  const jsonString = JSON.stringify(connectionsMap);
  const countsString = JSON.stringify(connectionCounts);
  const chunkSize = 90000; // 90KB chunks to stay under 100KB limit
  const chunks = [];

  for (let i = 0; i < jsonString.length; i += chunkSize) {
    chunks.push(jsonString.substring(i, i + chunkSize));
  }

  // Store chunk count and each chunk
  cache.put('connectionsMap_count', chunks.length.toString(), 21600);
  for (let i = 0; i < chunks.length; i++) {
    cache.put('connectionsMap_' + i, chunks[i], 21600);
  }

  // Store connection counts (usually small enough for single entry)
  cache.put('connectionCounts', countsString, 21600);

  return connectionsMap;
}

function getConnectionsMap() {
  const cache = CacheService.getScriptCache();
  const chunkCount = cache.get('connectionsMap_count');

  if (chunkCount) {
    // Reconstruct from chunks
    let jsonString = '';
    for (let i = 0; i < parseInt(chunkCount); i++) {
      const chunk = cache.get('connectionsMap_' + i);
      if (!chunk) {
        // Cache incomplete - rebuild
        return buildConnectionsCache();
      }
      jsonString += chunk;
    }
    try {
      return JSON.parse(jsonString);
    } catch (e) {
      Logger.log('Error parsing cached connections: ' + e.message);
      return buildConnectionsCache();
    }
  }

  // Cache miss - rebuild
  return buildConnectionsCache();
}

function getConnectionCounts() {
  const cache = CacheService.getScriptCache();
  const countsString = cache.get('connectionCounts');

  if (countsString) {
    try {
      return JSON.parse(countsString);
    } catch (e) {
      Logger.log('Error parsing cached connection counts: ' + e.message);
      buildConnectionsCache();
      return JSON.parse(cache.get('connectionCounts'));
    }
  }

  // Cache miss - rebuild entire cache and return counts
  buildConnectionsCache();
  return JSON.parse(cache.get('connectionCounts'));
}

function clearConnectionsCache() {
  const cache = CacheService.getScriptCache();
  const chunkCount = cache.get('connectionsMap_count');

  if (chunkCount) {
    for (let i = 0; i < parseInt(chunkCount); i++) {
      cache.remove('connectionsMap_' + i);
    }
    cache.remove('connectionsMap_count');
  }
  cache.remove('connectionCounts');
}

function buildGrid() {
  // Creates a draft Grid for the upcoming dinner using a list of Hosts and Guests
  let countGuests = 0;

  // Load ranges from sheet
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const gridSheet = spreadSheet.getSheetByName('GridBuilder');

  // Retrieve and validate the range of control values
  const controlRange = spreadSheet.getRangeByName('controlVariables');
  if (!controlRange) {
    throw new Error('Named range "controlVariables" not found');
  }
  const arrControl = controlRange.getValues().slice();

  // Validate control variables
  if (!arrControl || arrControl.length < 9) {
    throw new Error('Control variables range is missing or incomplete');
  }
  if (!arrControl[2] || isNaN(new Date(arrControl[2]).getTime())) {
    throw new Error('Invalid dinner date in control variables');
  }

  const ctrlTimeLapse = Number(arrControl[1]);
  const ctrlNextDinnerDate = new Date(arrControl[2]);
  const ctrlThrottleSingles = Number(arrControl[3]);
  const ctrlSortHosts = Number(arrControl[4]);
  const ctrlSortGuests = Number(arrControl[5]);
  const ctrlClearSeated = Number(arrControl[6]);
  const ctrlClearGrid = Number(arrControl[7]);
  const ctrlUnseatedOptions = Number(arrControl[8]);

  // Retrieve and validate named ranges from sheet
  const hostsRange = gridSheet.getRange('rangeHosts');
  const guestsRange = gridSheet.getRange('rangeGuests');
  const gridRange = gridSheet.getRange('rangeGrid');
  const neverMatchRange = gridSheet.getRange('rangeNeverMatch');

  if (!hostsRange || !guestsRange || !gridRange || !neverMatchRange) {
    throw new Error('One or more required named ranges not found');
  }

  const arrHosts = hostsRange.getValues().slice();
  const headerHosts = arrHosts[0];
  arrHosts.splice(0, 1);

  const arrGuests = guestsRange.getValues().slice();
  const headerGuests = arrGuests[0];
  arrGuests.splice(0, 1);

  const arrGrid = gridRange.getValues().slice();
  const headerGrid = arrGrid[0];
  arrGrid.splice(0, 1);

  const arrNeverMatch = neverMatchRange.getValues().slice();
  arrNeverMatch.splice(0, 1);

  // Load connections from cache
  const connectionsMap = getConnectionsMap();
  const connectionCounts = getConnectionCounts();

  // Count the number of connections for each Host and for each Guest and sort the list, descending order
  // This algorithm prioritizes seating members who have attended the most dinners
  // since they will be the most difficult to seat with fellow members

  // Hosts Count of Previous Connections
  if (ctrlSortHosts === 1) {
    for (let i = 0; i < arrHosts.length; i++) {
      if (arrHosts[i][HOST_COLUMNS.CODE].length > 0) {
        const count = connectionCounts[arrHosts[i][HOST_COLUMNS.CODE]] || 0;
        arrHosts[i].push(count);
      }
    }
    arrHosts.sort((a, b) => b[HOST_COLUMNS.ORDER] - a[HOST_COLUMNS.ORDER]);
  }

  // Compress the host array to remove blank rows
  for (let i = arrHosts.length - 1; i >= 0; i--) {
    if (arrHosts[i][HOST_COLUMNS.CODE] < 1) {
      arrHosts.splice(i, 1);
    }
  }

  // Guest Count of Previous Connections
  if (ctrlSortGuests === 1) {
    for (let i = 0; i < arrGuests.length; i++) {
      if (arrGuests[i][GUEST_COLUMNS.CODE].length > 0) {
        const count = connectionCounts[arrGuests[i][GUEST_COLUMNS.CODE]] || 0;
        arrGuests[i].push(count);
      }
    }
    arrGuests.sort((a, b) => b[GUEST_COLUMNS.ORDER] - a[GUEST_COLUMNS.ORDER]);
  }

  // Count how many guests to process to ignore blank rows
  for (let i = 0; i < arrGuests.length; i++) {
    if (arrGuests[i][GUEST_COLUMNS.CODE].length > 0) {
      countGuests++;
    }
  }
  //
  // Filter and create bidirectional Set for never-match pairs
  const setNeverMatch = new Set();
  for (let i = 0; i < arrNeverMatch.length; i++) {
    if (arrNeverMatch[i][0] && arrNeverMatch[i][0].toString().length > 0) {
      const pair = arrNeverMatch[i][0].toString();
      setNeverMatch.add(pair);
      // Add reverse pair
      const parts = pair.split('-');
      if (parts.length === 2) {
        setNeverMatch.add([parts[1], parts[0]].join('-'));
      }
    }
  }
  // Check control flags, clear the seated flag and the Grid if requested
  if (ctrlClearSeated === 1) {
    for (let i = 0; i < countGuests; i++) {
      arrGuests[i][GUEST_COLUMNS.SEATED] = "No";
    }
  }

  if (ctrlClearGrid === 1) {
    for (let i = 0; i < arrGrid.length; i++) {
      for (let j = 0; j <= GRID_COLUMNS.GUEST_5; j++) {
        arrGrid[i][j] = null;
      }
    }
  }

  // CREATE THE GRID
  for (let i = 0; i < arrHosts.length; i++) {
    let seatedCount;
    let arrHouse;

    if (arrGrid[i][GRID_COLUMNS.HOUSE] === null) {
      // Initialize the House on first loop
      arrGrid[i][GRID_COLUMNS.HOUSE] = i + 1;
      arrGrid[i][GRID_COLUMNS.SEATS] = arrHosts[i][HOST_COLUMNS.SEATS];
      arrGrid[i][GRID_COLUMNS.HOST] = arrHosts[i][HOST_COLUMNS.CODE];
      seatedCount = Number(arrHosts[i][HOST_COLUMNS.COUNT]);

      arrHouse = [];
      arrHouse.push(arrGrid[i][GRID_COLUMNS.HOST]);
    } else {
      // The Grid already has data stored. Load that data into the variables used for processing
      arrHouse = [];
      for (let n = GRID_COLUMNS.HOST; n <= GRID_COLUMNS.GUEST_5; n++) {
        if (arrGrid[i][n] !== null) {
          arrHouse.push(arrGrid[i][n]);
        }
      }
      seatedCount = arrGrid[i][GRID_COLUMNS.SEATED];
    }
    // Fill the house - with up to 5 guests or until seats are full
    let gCount = arrHouse.length;
    while (gCount < 5 && seatedCount < arrHosts[i][HOST_COLUMNS.SEATS]) {
      let guestWasSeated = false;

      for (let gRow = 0; gRow < countGuests; gRow++) {
        if (ctrlThrottleSingles === 1 && arrGuests[gRow][GUEST_COLUMNS.COUNT] === 1 && gCount < 2) {
          continue;
        }
        if (arrGuests[gRow][GUEST_COLUMNS.SEATED] === "No" &&
            seatedCount + arrGuests[gRow][GUEST_COLUMNS.COUNT] <= arrHosts[i][HOST_COLUMNS.SEATS]) {

          const compatibilityChecks = [];

          // Check this guest against all current members in the house
          for (let h = 0; h < gCount; h++) {
            compatibilityChecks.push(memberMatch(
              `${arrGuests[gRow][GUEST_COLUMNS.CODE]}-${arrHouse[h]}`,
              ctrlTimeLapse,
              connectionsMap,
              setNeverMatch
            ));
          }

          if (compatibilityChecks.every(check => check === true)) {
            arrHouse.push(arrGuests[gRow][GUEST_COLUMNS.CODE]);
            arrGuests[gRow][GUEST_COLUMNS.SEATED] = null;
            seatedCount = seatedCount + Number(arrGuests[gRow][GUEST_COLUMNS.COUNT]);
            gCount++;
            guestWasSeated = true;
            break;
          }
        }
      }

      if (!guestWasSeated) {
        break;
      }
    }

    // The house is finished. Add the values to arrGrid with bounds checking
    arrGrid[i][GRID_COLUMNS.SEATED] = seatedCount;
    for (let j = 0; j < 6; j++) {
      arrGrid[i][GRID_COLUMNS.HOST + j] = arrHouse[j] || null;
    }
  }
  // CLEAN UP AND FINISH
  // Write the completed arrGrid back to the sheet
  arrGrid.splice(0, 0, headerGrid);
  gridSheet.getRange('rangeGrid').setValues(arrGrid);

  // Remove the column added for sorting before writing back to the sheet
  if (ctrlSortGuests === 1) {
    for (let i = 0; i < arrGuests.length; i++) {
      if (arrGuests[i][GUEST_COLUMNS.CODE].length > 0) {
        arrGuests[i].pop();
      }
    }
  }

  // Write the arrGuests back to the sheet
  arrGuests.splice(0, 0, headerGuests);
  gridSheet.getRange('rangeGuests').setValues(arrGuests);

  // Report unseated members
  reportUnseatedMembers(arrGuests, GUEST_COLUMNS.CODE, GUEST_COLUMNS.SEATED);
}
/** HELPER FUNCTIONS */

function memberMatch(strToCheck, timeLapse, connectionsMap, setNeverMatch) {
  // Check never-match list (O(1) lookup)
  if (setNeverMatch.has(strToCheck)) {
    return false;
  }

  // Check time-based constraints using Map (O(1) lookup instead of O(n))
  const monthsApart = connectionsMap[strToCheck];
  if (monthsApart !== undefined && timeLapse > monthsApart) {
    return false;
  }

  return true;
}

function updateAndSortConnectionsDB(arrDB, ctrlNextDinnerDate) {
  // Update the connections database with the length of time, in months, since the members last met
  for (let i = 0; i < arrDB.length; i++) {
    const lastDate = new Date(arrDB[i][CONNECTION_COLUMNS.DATE]);
    arrDB[i][CONNECTION_COLUMNS.MONTHS_APART] = ctrlNextDinnerDate.getMonth() - lastDate.getMonth() +
      (12 * (ctrlNextDinnerDate.getFullYear() - lastDate.getFullYear()));
  }

  // Sort the array alphabetically by pair, then numerically by months apart
  arrDB.sort((a, b) => a[CONNECTION_COLUMNS.PAIR].localeCompare(b[CONNECTION_COLUMNS.PAIR]) ||
                       a[CONNECTION_COLUMNS.MONTHS_APART] - b[CONNECTION_COLUMNS.MONTHS_APART]);
}
function prepConnections() {
  // Recalculates the number of months between members meeting and returns the sorted array to the calling function
  const result = SpreadsheetApp.getUi().alert(
    "Run this once before building the Grid. Check that the dinner date is correct on the spreadsheet. Click OK to run this script, Cancel to stop",
    SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
  );
  if (result !== SpreadsheetApp.getUi().Button.OK) {
    return;
  }

  // Get the connections data from the sheet
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = spreadSheet.getSheetByName('connectionsDB');
  const arrDB = dbSheet.getDataRange().getValues().slice();
  const headerRow = arrDB[0];
  arrDB.splice(0, 1);

  // Retrieve the range of control values to get the upcoming dinner date
  const arrControl = spreadSheet.getRangeByName('controlVariables').getValues();
  const ctrlNextDinnerDate = new Date(arrControl[2]);

  // Update and sort the connections database
  updateAndSortConnectionsDB(arrDB, ctrlNextDinnerDate);

  // Write the array back to the sheet
  arrDB.splice(0, 0, headerRow);
  const rowCount = arrDB.length;
  const colCount = arrDB[0].length;
  const connectionsRange = dbSheet.getRange(1, 1, rowCount, colCount);
  connectionsRange.setValues(arrDB);

  // Rebuild cache after updating connections
  clearConnectionsCache();
  buildConnectionsCache();
}
function updateConnections() {
  // Takes the final grid for a dinner and creates new entries in the connectionsDB sheet
  // for each pair of members that sat together at a dinner
  const result = SpreadsheetApp.getUi().alert(
    "Run this after the dinner date. Update the Grid with any last minute changes and check that the dinner date is correct on the spreadsheet. Click OK to run this script, Cancel to stop",
    SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
  );
  if (result !== SpreadsheetApp.getUi().Button.OK) {
    return;
  }

  // Get the connections data from the sheet
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = spreadSheet.getSheetByName('connectionsDB');
  const arrDB = dbSheet.getDataRange().getValues().slice();
  const headerRow = arrDB[0];
  arrDB.splice(0, 1);

  // Get the Grid from the worksheet
  const gridSheet = spreadSheet.getSheetByName('GridBuilder');
  const arrGrid = gridSheet.getRange('rangeGrid').getValues().slice();
  arrGrid.splice(0, 1);

  // Retrieve the range of control values to get the dinner date
  const arrControl = spreadSheet.getRangeByName('controlVariables').getValues();
  const ctrlNextDinnerDate = new Date(arrControl[2]);

  // Count the number of houses to process - remove blank rows
  for (let i = arrGrid.length - 1; i >= 0; i--) {
    if (arrGrid[i][GRID_COLUMNS.HOST] < 1) {
      arrGrid.splice(i, 1);
    }
  }
  const houseCount = arrGrid.length;

  // Create connections for all pairs of members in each house
  for (let i = 0; i < houseCount; i++) {
    // Count the number of members to process in the house
    let guestCount = 0;
    for (let j = GRID_COLUMNS.HOST; j <= GRID_COLUMNS.GUEST_5; j++) {
      if (arrGrid[i][j] && arrGrid[i][j].toString().length > 0) {
        guestCount++;
      }
    }

    // Create connections for all pairs of members in the house
    for (let j = 1; j < guestCount; j++) {
      const isHost = (j === 1) ? "Host" : "";
      for (let k = j; k < guestCount; k++) {
        writeConnection(
          arrGrid[i][j + GRID_COLUMNS.HOST],
          arrGrid[i][k + GRID_COLUMNS.HOST],
          ctrlNextDinnerDate,
          isHost,
          arrDB
        );
      }
    }
  }

  // Update and sort the connections database
  updateAndSortConnectionsDB(arrDB, ctrlNextDinnerDate);

  // Write the array back to the sheet
  arrDB.splice(0, 0, headerRow);
  const rowCount = arrDB.length;
  const colCount = arrDB[0].length;
  const connectionsRange = dbSheet.getRange(1, 1, rowCount, colCount);
  connectionsRange.setValues(arrDB);

  // Rebuild cache after updating connections
  clearConnectionsCache();
  buildConnectionsCache();
}
/** SUPPORTING FUNCTIONS */

/**
 * Writes new bidirectional connection entries to the connections array
 * @param {string} memberOne - First member code
 * @param {string} memberTwo - Second member code
 * @param {Date} dinnerDate - Date of the dinner
 * @param {string} memberRole - Role (e.g., "Host" or "")
 * @param {Array} arrTmp - Array that will be mutated with new connections
 */
function writeConnection(memberOne, memberTwo, dinnerDate, memberRole, arrTmp) {
  const tmpYear = dinnerDate.getFullYear();
  arrTmp.push([`${memberOne}-${memberTwo}`, memberOne, memberTwo, tmpYear, dinnerDate, memberRole, 0]);
  // Create a duplicate entry with the key flipped (member1-member2 and member2-member1)
  arrTmp.push([`${memberTwo}-${memberOne}`, memberOne, memberTwo, tmpYear, dinnerDate, memberRole, 0]);
}

function reportUnseatedMembers(arrGuests, guestCodeIdx, guestSeatedIdx) {
  const unseated = [];

  // Skip header row (index 0)
  for (let i = 1; i < arrGuests.length; i++) {
    if (arrGuests[i][guestCodeIdx] &&
        arrGuests[i][guestCodeIdx].toString().length > 0 &&
        arrGuests[i][guestSeatedIdx] === "No") {
      unseated.push({
        code: arrGuests[i][guestCodeIdx],
        count: arrGuests[i][GUEST_COLUMNS.COUNT] || 1
      });
    }
  }

  if (unseated.length > 0) {
    const msg = `Unseated members (${unseated.length}):\n` +
                unseated.map(m => `  ${m.code} (party of ${m.count})`).join('\n');
    SpreadsheetApp.getUi().alert('Seating Incomplete', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('Success', 'All members seated successfully!', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function auditGrid() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const gridSheet = spreadSheet.getSheetByName('GridBuilder');
  const dbSheet = spreadSheet.getSheetByName('connectionsDB');

  const arrGrid = gridSheet.getRange('rangeGrid').getValues().slice();
  arrGrid.splice(0, 1); // Remove header

  const arrDB = dbSheet.getDataRange().getValues().slice();
  arrDB.splice(0, 1); // Remove header

  // Build Map for O(1) lookups instead of O(n) Array.find
  const connectionsMap = {};
  for (let i = 0; i < arrDB.length; i++) {
    connectionsMap[arrDB[i][CONNECTION_COLUMNS.PAIR]] = arrDB[i];
  }

  const auditData = [];
  let minMonths = Infinity;
  let maxMonths = -Infinity;
  let totalConnections = 0;

  for (let i = 0; i < arrGrid.length; i++) {
    if (!arrGrid[i][GRID_COLUMNS.HOST]) continue;

    const members = [];
    for (let j = GRID_COLUMNS.HOST; j <= GRID_COLUMNS.GUEST_5; j++) {
      if (arrGrid[i][j]) members.push(arrGrid[i][j]);
    }

    const houseNumber = i + 1;

    // Check all pairs and create structured data rows
    for (let m = 0; m < members.length; m++) {
      for (let n = m + 1; n < members.length; n++) {
        const pair = `${members[m]}-${members[n]}`;
        const connection = connectionsMap[pair];
        const monthsApart = connection ? connection[CONNECTION_COLUMNS.MONTHS_APART] : 'Never met';

        if (typeof monthsApart === 'number') {
          minMonths = Math.min(minMonths, monthsApart);
          maxMonths = Math.max(maxMonths, monthsApart);
          totalConnections++;
        }

        // Add row: House | Member1 | Member2 | Months Apart
        auditData.push([houseNumber, members[m], members[n], monthsApart]);
      }
    }
  }

  // Get or create Grid Audit sheet
  let auditSheet = spreadSheet.getSheetByName('Grid Audit');
  if (!auditSheet) {
    auditSheet = spreadSheet.insertSheet('Grid Audit');
  } else {
    auditSheet.clear();
  }

  // Write header row
  const header = [['House', 'Member 1', 'Member 2', 'Months Apart']];
  auditSheet.getRange(1, 1, 1, 4).setValues(header);

  // Format header
  auditSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#d9d9d9');

  // Write audit data
  if (auditData.length > 0) {
    auditSheet.getRange(2, 1, auditData.length, 4).setValues(auditData);

    // Auto-resize columns
    auditSheet.autoResizeColumns(1, 4);

    // Freeze header row
    auditSheet.setFrozenRows(1);
  }

  // Create summary message
  let summaryMsg = `Audit complete! ${totalConnections} connections analyzed.\n\n`;
  if (totalConnections > 0) {
    summaryMsg += `Minimum separation: ${minMonths} months\n`;
    summaryMsg += `Maximum separation: ${maxMonths} months\n\n`;
    summaryMsg += `Results written to "Grid Audit" sheet.`;
  }

  SpreadsheetApp.getUi().alert('Audit Complete', summaryMsg, SpreadsheetApp.getUi().ButtonSet.OK);
}
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Grid Builder')
    .addItem('Build Grid', 'buildGrid')
    .addSeparator()
    .addItem('Audit Current Grid', 'auditGrid')
    .addSeparator()
    .addItem('Prepare Connections Data', 'prepConnections')
    .addSeparator()
    .addItem('Post dinner -> Update Connections History', 'updateConnections')
    .addToUi();
}