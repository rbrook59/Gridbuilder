// FUTURE DEVELOPMENT NOTES
// 1. create a audit of a Grid based on pairs of members and # of months apart
// 2. create an orphans list to give options of where unseated members can sit.
// 3.
//
/** CACHE HELPER FUNCTIONS */
function buildConnectionsCache() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = spreadSheet.getSheetByName('connectionsDB');
  const arrDB = dbSheet.getDataRange().getValues().slice();
  arrDB.splice(0, 1); // Remove header

  // Build Map: key = "member1-member2", value = months apart
  const connectionsMap = {};
  for (let i = 0; i < arrDB.length; i++) {
    const key = arrDB[i][0]; // member1-member2 pair
    const months = arrDB[i][6]; // months apart

    // Store only the minimum months for each pair (since array is sorted)
    if (!connectionsMap[key]) {
      connectionsMap[key] = months;
    }
  }

  // Store in cache using chunks (CacheService has 100KB limit per entry)
  const cache = CacheService.getScriptCache();
  const jsonString = JSON.stringify(connectionsMap);
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
    return JSON.parse(jsonString);
  }

  // Cache miss - rebuild
  return buildConnectionsCache();
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
}

function buildGrid() {
  //
  // creates a draft Grid for the upcoming dinner using a list of Hosts and Guests.
  //
  /** Set up variables aligned with table columns from spreadsheet
   *  change these if the range format or column order is changed on the sheet */
  //
  const hostCode = 0;
  const hostCount = 1;
  const hostSeats = 2;
  const hostOrder = 3; // not saved back to sheet
  const guestCode = 0;
  const guestCount = 1;
  const guestSeated = 2;
  const guestOrder = 3; // not saved back to sheet
  const gridHouse = 0;
  const gridSeats = 1;
  const gridSeated = 2;
  const gridHost = 3;
  //
  let countGuests = 0;
  //
  /** Load ranges from sheet */
  //
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const gridSheet = spreadSheet.getSheetByName('GridBuilder');
  //
  /** retrieve the range of control values that are used by the program  */
  //
  const arrControl = spreadSheet.getRange('controlVariables').getValues().slice();

  // Validate control variables
  if (!arrControl || arrControl.length < 9) {
    throw new Error('Control variables range is missing or incomplete');
  }
  if (!arrControl[2] || isNaN(new Date(arrControl[2]).getTime())) {
    throw new Error('Invalid dinner date in control variables');
  }

  const ctrlTimeLapse = Number(arrControl[1]);
  const ctrlNextDinnerDate = Date(arrControl[2]);
  const ctrlThrottleSingles = Number(arrControl[3]);
  const ctrlSortHosts = Number(arrControl[4]);
  const ctrlSortGuests = Number(arrControl[5]);
  const ctrlClearSeated = Number(arrControl[6]);
  const ctrlClearGrid = Number(arrControl[7]);
  const ctrlUnseatedOptions = Number(arrControl[8]);
  //
  // retrieve ranges from sheet
  //
  const arrHosts = gridSheet.getRange('rangeHosts').getValues().slice();
  const headerHosts = arrHosts[0]; // save the header row to add back later
  arrHosts.splice(0, 1); //remove the header row for easier sorting
  //
  const arrGuests = gridSheet.getRange('rangeGuests').getValues().slice();
  const headerGuests = arrGuests[0]; // save the header row to add back later
  arrGuests.splice(0, 1); //remove the header row for easier sorting
  //
  const arrGrid = gridSheet.getRange('rangeGrid').getValues().slice();
  const headerGrid = arrGrid[0]; // save the header row to add back later
  arrGrid.splice(0, 1); //remove the header row for easier sorting
  //
  const arrNeverMatch = gridSheet.getRange('rangeNeverMatch').getValues().slice();
  arrNeverMatch.splice(0, 1); //remove the header row for easier processing
  //
  // Load connections from cache instead of sheet
  const connectionsMap = getConnectionsMap();

  //
  /**Count the number of connections for each Host and for each Guest and sort the list, descending order.
   * This algorithm prioritizes seating members who have attended the most dinners
   * since they will be the most difficult to seat with fellow members.*/
  //
  // Hosts Count of Previous Connections
  if (ctrlSortHosts === 1) {
    for (let i = 0; i < arrHosts.length; i++) {
      if (arrHosts[i][hostCode].length > 0) {
        // Count connections using Map keys
        let count = 0;
        for (let key in connectionsMap) {
          const parts = key.split('-');
          if (parts[0] === arrHosts[i][hostCode].toString()) {
            count++;
          }
        }
        arrHosts[i].push(count);
      }
    }
    // Sort the array of Hosts
    arrHosts.sort(function (a, b) {
      return b[3] - a[3]
    }); //sort list of Hosts from most connections to least connections
  }
  // Compress the host array to remove blank rows
  for (let i = arrHosts.length - 1; i >= 0; i--) {
    if (arrHosts[i][hostCode] < 1) {
      arrHosts.splice(i, 1);
    }
  }
  //
  //Guest Count of Previous Connections
  if (ctrlSortGuests === 1) {
    for (let i = 0; i < arrGuests.length; i++) {
      if (arrGuests[i][guestCode].length > 0) {
        // Count connections using Map keys
        let count = 0;
        for (let key in connectionsMap) {
          const parts = key.split('-');
          if (parts[0] === arrGuests[i][guestCode].toString()) {
            count++;
          }
        }
        arrGuests[i].push(count);
      }
      // Sort the array of Guests
      arrGuests.sort(function (a, b) {
        return b[3] - a[3]
      }); //sort list of Guests in descending order
    }
  }
  // Count how many guests to process to ignore blank rows
  for (let i = 0; i < arrGuests.length; i++) {
    if (arrGuests[i][guestCode].length > 0) {
      countGuests++
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
  /** check control flags, clear the seated flag and the Grid  */
  //
  if (ctrlClearSeated === 1) {
    for (let i = 0; i < countGuests; i++) {
      arrGuests[i][guestSeated] = "No"
    }
  }
  // Clear the Grid
  if (ctrlClearGrid === 1) {
    for (let i = 0; i < arrGrid.length; i++) {
      for (let j = 0; j < 9; j++) {
        arrGrid[i][j] = null;
      }
    }
  }
  //
  /**  CREATE THE GRID */
  //
  for (let i = 0; i < arrHosts.length; i++) { // This many houses to process
    //
    let seatedCount;
    let arrHouse;

    if (arrGrid[i][gridHouse] === null) { // initialize the House on first loop
      arrGrid[i][gridHouse] = i + 1; // array starts at 0, so add 1
      arrGrid[i][gridSeats] = arrHosts[i][hostSeats]; // How many can be seated at this house
      arrGrid[i][gridHost] = arrHosts[i][hostCode]; // Host name code
      seatedCount = Number(arrHosts[i][hostCount]); // this many seats for the host (will be 1 or 2)
      //
      arrHouse = []; // 1D empty array to track guests assigned
      arrHouse.push(arrGrid[i][gridHost]); // add the host code to the House array
      //
    } else { // the Grid already has data stored.  Load that data into the variables used for processing
      arrHouse = []; // 1D empty array to track guests assigned
      for (let n = 3; n < arrGrid[i].length; n++) {
        if (arrGrid[i][n] !== null) {
          arrHouse.push(arrGrid[i][n]);
        }
      }
      seatedCount = arrGrid[i][gridSeated];
    }
    //
    // Fill the house - with up to 5 guests or until seats are full
    //
    for (let gCount = arrHouse.length; gCount < 5; gCount++) {
      if (seatedCount >= arrHosts[i][hostSeats]) { // House is full, go to the next house
        break;
      }
      //
      //  The array "doesItWork" tests each potential member in the house.  Any "false" values and the guest is rejected.
      //
      //
      for (let gRow = 0; gRow < countGuests; gRow++) { // cycle through each member
        if (ctrlThrottleSingles === 1 && arrGuests[gRow][guestCount] === 1 && gCount < 2) { // wait to seat singles
          continue;
        }
        if (arrGuests[gRow][guestSeated] === "No" && seatedCount + arrGuests[gRow][guestCount] <= arrHosts[i][hostSeats]) {  // Does this guest need a seat and can they fit in this house
          //
          const doesItWork = []; // array that gets loaded with True or False values.   Any False value and the guest is rejected
          //
          // Check this guest against all current members in the house
          for (let h = 0; h < gCount; h++) {
            doesItWork.push(memberMatch(
              [arrGuests[gRow][guestCode], arrHouse[h]].join('-'),
              ctrlTimeLapse,
              connectionsMap,
              setNeverMatch
            ));
          }

          if (doesItWork.findIndex((tf) => tf === false) === -1) { // This guest works if no false values are found
            arrHouse.push(arrGuests[gRow][guestCode]);
            arrGuests[gRow][guestSeated] = null;
            seatedCount = seatedCount + Number(arrGuests[gRow][guestCount]);
            gCount++
          }
        } // if guest can be seated
        if (seatedCount >= arrHosts[i][hostSeats]) { // House is full.
          break;
        }
      } //for gRow < countGuests
    } // for gCount
    //
    //  The house is finished.  Add the values to arrGrid
    //
    arrGrid[i][2] = seatedCount;
    arrGrid[i][3] = arrHouse[0];
    arrGrid[i][4] = arrHouse[1];
    arrGrid[i][5] = arrHouse[2];
    arrGrid[i][6] = arrHouse[3];
    arrGrid[i][7] = arrHouse[4];
    arrGrid[i][8] = arrHouse[5];
  } // for House count
  /**
   * CLEAN UP AND FINISH
   */
  //
  // write the completed arrGrid back to the sheet
  // 
  arrGrid.splice(0, 0, headerGrid); // add the headerRow back to the top of the array
  gridSheet.getRange('rangeGrid').setValues(arrGrid);
  //
  // Remove the column added for sorting before writing back to the sheet
  if (ctrlSortGuests === 1) {
    for (let i = 0; i < arrGuests.length; i++) {
      if (arrGuests[i][guestCode].length > 0) {
        arrGuests[i].pop();
      }
    }
  }
  // write the arrGuests back to the sheet
  arrGuests.splice(0, 0, headerGuests);
  gridSheet.getRange('rangeGuests').setValues(arrGuests);

  // Report unseated members
  reportUnseatedMembers(arrGuests, guestCode, guestSeated);
}
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
function prepConnections() {
  //
  // recalculates the number of months between members meeting and returns the sorted array to the calling function
  //
  let result = SpreadsheetApp.getUi().alert("Run this once before building the Grid.Check that the dinner date is correct on the spreadsheet.  Click OK to run this script, Cancel to stop", SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  if (result !== SpreadsheetApp.getUi().Button.OK) {
    return;
  }
  /** connectionDB Sheet layout
   * col[0] = Member1Code-Member2Code - index for each searching of pairs
   * col[1] = Member1 code
   * col[2] = Member2 code
   * col[3] = Year of dinner
   * col[4] = Calendar date of dinner
   * col[5] = Host?
   * col[6] = Months since Member1-Member2 were together
   * */
  //
  // Get the connections data from the sheet
  //
  let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  let dbSheet = spreadSheet.getSheetByName('connectionsDB');
  let arrDB = dbSheet.getDataRange().getValues().slice();
  let headerRow = arrDB[0]; // save the header row to add back later
  arrDB.splice(0, 1); //remove the header row
  //
  // retrieve the range of control values that are used by the program.  Get the upcoming dinner date.
  //
  let arrControl = spreadSheet.getRangeByName('controlVariables').getValues();
  let ctrlNextDinnerDate = new Date(arrControl[2]);
  //
  // update the connections database with the length of time, in months, since the members last met.
  //
  for (let i = 0; i < arrDB.length; i++) {
    let lastDate = new Date(arrDB[i][4]);
    arrDB[i][6] = ctrlNextDinnerDate.getMonth() - lastDate.getMonth() +
      (12 * (ctrlNextDinnerDate.getFullYear() - lastDate.getFullYear()));  // write the number of months to the array[6]
  }
  //
  // sort the array alphabetically by col[0], "member1code-member2code" and numerically by col[6] and with smallest number of months first
  //
  arrDB.sort(function (a, b) {
    return a[0].localeCompare(b[0]) || a[6] - b[6]
  }); // sorts on first element alphabetically then ascending on # of months
  //
  // write the array back to the sheet
  //
  arrDB.splice(0, 0, headerRow); // add the headerRow back to the top of the array
  let rowCount = arrDB.length;
  let colCount = arrDB[0].length;
  let connectionsRange = dbSheet.getRange(1, 1, rowCount, colCount); // define where the array is to be written back to the sheet
  connectionsRange.setValues(arrDB); // write to the sheet

  // Rebuild cache after updating connections
  clearConnectionsCache();
  buildConnectionsCache();
}
function updateConnections() {
  //
  // takes the final grid for a dinner and creates new entries in the connectionsDB sheet for each pair of
  // members that sat together at a dinner.  This updates the history of the members that have been together.
  //
  let result = SpreadsheetApp.getUi().alert("Run this after the dinner date.  Update the Grid with any last minute changes and check that the dinner date is correct on the spreadsheet.  Click OK to run this script, Cancel to stop", SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  if (result !== SpreadsheetApp.getUi().Button.OK) {
    return;
  }
  /** connectionDB Sheet layout
   * col[0] = Member1Code-Member2Code - index for each searching of pairs
   * col[1] = Member1 code
   * col[2] = Member2 code
   * col[3] = Year of dinner
   * col[4] = Calendar date of dinner
   * col[5] = Host?
   * col[6] = Months since Member1-Member2 were together
   * */
  /** grid from worksheet with final seating assignments
   * col[0] = house number
   * col[1] = seating capacity
   * col[2] = seating count
   * col[3] = Host
   * col[4] = Guest 1
   * col[5] = Guest 2
   * col[6] = Guest 3
   * col[7] = Guest 4
   * col[8] = Guest 5
   */
  //
  // Get the connections data from the sheet
  //
  let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  let dbSheet = spreadSheet.getSheetByName('connectionsDB');
  const arrDB = dbSheet.getDataRange().getValues().slice();
  let headerRow = arrDB[0]; // save the header row to add back later
  arrDB.splice(0, 1); //remove the header row
  //
  // Get the Grid from the worksheet
  //
  const gridSheet = spreadSheet.getSheetByName('GridBuilder');
  const arrGrid = gridSheet.getRange('rangeGrid').getValues().slice();
  let headerGrid = arrGrid[0]; // save the header row to add back later
  arrGrid.splice(0, 1); //remove the header row for easier sorting
  //
  // retrieve the range of control values that are used by the program.  Get the dinner date.
  //
  let arrControl = spreadSheet.getRangeByName('controlVariables').getValues();
  let ctrlNextDinnerDate = new Date(arrControl[2]);  // the data is saved with the new connections
  //
  // count the number of houses to process
  //
  for (let i = arrGrid.length - 1; i >= 0; i--) {
    if (arrGrid[i][3] < 1) { // blank host code in Grid.  Remove the row.
      arrGrid.splice(i, 1);
    }
  }
  const houseCount = arrGrid.length;
  //
  for (let i = 0; i < houseCount; i++) {
    //
    // count the number of members to process in the house
    //
    let guestCount = 0;
    for (let j = 3; j < arrGrid[i].length; j++) {  // 3 is the array position for the host
      if (arrGrid[i][j].length > 0) {
        guestCount++;
      }
    }
    // Create connections for all pairs of members in the house
    for (let j = 1; j < guestCount; j++) {
      const isHost = (j === 1) ? "Host" : "";
      for (let k = j; k < guestCount; k++) {
        writeConnection(arrGrid[i][j + 3], arrGrid[i][k + 3], ctrlNextDinnerDate, isHost, arrDB);
      }
    }
  }
  for (let i = 0; i < arrDB.length; i++) {
    let lastDate = new Date(arrDB[i][4]);
    arrDB[i][6] = ctrlNextDinnerDate.getMonth() - lastDate.getMonth() +
      (12 * (ctrlNextDinnerDate.getFullYear() - lastDate.getFullYear()));  // write the number of months to the array[6]
  }
  //
  // sort the array alphabetically by col[0], "member1code-member2code" and numerically by col[6] and with 
  // smallest numberof months first
  //
  arrDB.sort(function (a, b) {
    return a[0].localeCompare(b[0]) || a[6] - b[6]
  }); // sorts on first element alphabetically then ascending on # of months
  //
  // write the array back to the sheet
  //
  arrDB.splice(0, 0, headerRow); // add the headerRow back to the top of the array
  let rowCount = arrDB.length;
  let colCount = arrDB[0].length;
  let connectionsRange = dbSheet.getRange(1, 1, rowCount, colCount); // define where the array is to be written back to the sheet
  connectionsRange.setValues(arrDB); // write to the sheet

  // Rebuild cache after updating connections
  clearConnectionsCache();
  buildConnectionsCache();
}
//
/** SUPPORTING FUNCTIONS */
//
//
// subroutine to write new connections to connectionsDB
//
function writeConnection(memberOne, memberTwo, dinnerDate, memberRole, arrTmp) {
  const tmpYear = dinnerDate.getFullYear()
  arrTmp.push([[memberOne, memberTwo].join('-'), memberOne, memberTwo, tmpYear, dinnerDate, memberRole, 0]);
  //
  // create a duplicate entry with the key flipped. member1-member2 and member2-member1
  //
  arrTmp.push([[memberTwo, memberOne].join('-'), memberOne, memberTwo, tmpYear, dinnerDate, memberRole, 0]);
  return;
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
        count: arrGuests[i][1] || 1
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

  const auditData = [];
  let minMonths = Infinity;
  let maxMonths = -Infinity;
  let totalConnections = 0;

  for (let i = 0; i < arrGrid.length; i++) {
    if (!arrGrid[i][3]) continue; // Skip empty houses

    const members = [];
    for (let j = 3; j < arrGrid[i].length; j++) {
      if (arrGrid[i][j]) members.push(arrGrid[i][j]);
    }

    const houseNumber = i + 1;

    // Check all pairs and create structured data rows
    for (let m = 0; m < members.length; m++) {
      for (let n = m + 1; n < members.length; n++) {
        const pair = `${members[m]}-${members[n]}`;
        const connection = arrDB.find(row => row[0] === pair);
        const monthsApart = connection ? connection[6] : 'Never met';

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