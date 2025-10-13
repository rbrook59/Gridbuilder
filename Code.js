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
  // Uses random restart strategy to try multiple seating arrangements

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

  const arrHostsOriginal = hostsRange.getValues().slice();
  const headerHosts = arrHostsOriginal[0];
  arrHostsOriginal.splice(0, 1);

  const arrGuestsOriginal = guestsRange.getValues().slice();
  const headerGuests = arrGuestsOriginal[0];
  arrGuestsOriginal.splice(0, 1);

  const arrGridOriginal = gridRange.getValues().slice();
  const headerGrid = arrGridOriginal[0];
  arrGridOriginal.splice(0, 1);

  const arrNeverMatch = neverMatchRange.getValues().slice();
  arrNeverMatch.splice(0, 1);

  // Load connections from cache
  const connectionsMap = getConnectionsMap();
  const connectionCounts = getConnectionCounts();

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

  // RANDOM RESTART LOOP - Try multiple times with randomized ordering
  const maxAttempts = 10;
  let bestGrid = null;
  let bestGuests = null;
  let bestUnseatedCount = Infinity;
  let attemptNumber = 0;

  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    attemptNumber++;

    // Create fresh copies for this attempt
    const arrHosts = arrHostsOriginal.map(row => [...row]);
    const arrGuests = arrGuestsOriginal.map(row => [...row]);
    const arrGrid = arrGridOriginal.map(row => [...row]);
    let countGuests = 0;

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

    // After attempt 1, add slight randomization to the order
    if (attempt > 0) {
      // Separate non-blank and blank hosts
      const nonBlankHosts = [];
      const blankHosts = [];
      for (let i = 0; i < arrHosts.length; i++) {
        if (arrHosts[i][HOST_COLUMNS.CODE] && arrHosts[i][HOST_COLUMNS.CODE].toString().length > 0) {
          nonBlankHosts.push(arrHosts[i]);
        } else {
          blankHosts.push(arrHosts[i]);
        }
      }

      // Shuffle last 30% of non-blank hosts (keep most constrained at top)
      const shuffleStartHost = Math.floor(nonBlankHosts.length * 0.7);
      const hostsToShuffle = nonBlankHosts.splice(shuffleStartHost);
      shuffleArray(hostsToShuffle);
      nonBlankHosts.push(...hostsToShuffle);

      // Reconstruct hosts array: non-blanks first, blanks at end
      arrHosts.length = 0;
      arrHosts.push(...nonBlankHosts, ...blankHosts);

      // Separate non-blank and blank guests
      const nonBlankGuests = [];
      const blankGuests = [];
      for (let i = 0; i < arrGuests.length; i++) {
        if (arrGuests[i][GUEST_COLUMNS.CODE] && arrGuests[i][GUEST_COLUMNS.CODE].toString().length > 0) {
          nonBlankGuests.push(arrGuests[i]);
        } else {
          blankGuests.push(arrGuests[i]);
        }
      }

      // Shuffle last 30% of non-blank guests (keep most constrained at top)
      const shuffleStartGuest = Math.floor(nonBlankGuests.length * 0.7);
      const guestsToShuffle = nonBlankGuests.splice(shuffleStartGuest);
      shuffleArray(guestsToShuffle);
      nonBlankGuests.push(...guestsToShuffle);

      // Reconstruct guests array: non-blanks first, blanks at end
      arrGuests.length = 0;
      arrGuests.push(...nonBlankGuests, ...blankGuests);
    }

    // Count how many guests to process to ignore blank rows
    for (let i = 0; i < arrGuests.length; i++) {
      if (arrGuests[i][GUEST_COLUMNS.CODE].length > 0) {
        countGuests++;
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

    // Count unseated guests for this attempt
    let unseatedCount = 0;
    for (let i = 0; i < arrGuests.length; i++) {
      if (arrGuests[i][GUEST_COLUMNS.CODE] &&
          arrGuests[i][GUEST_COLUMNS.CODE].toString().length > 0 &&
          arrGuests[i][GUEST_COLUMNS.SEATED] === "No") {
        unseatedCount++;
      }
    }

    // Track best result
    if (unseatedCount < bestUnseatedCount) {
      bestUnseatedCount = unseatedCount;
      bestGrid = arrGrid.map(row => [...row]);
      bestGuests = arrGuests.map(row => [...row]);

      // If we found a perfect solution, stop trying
      if (unseatedCount === 0) {
        Logger.log(`Perfect solution found on attempt ${attemptNumber}`);
        break;
      }
    }

    Logger.log(`Attempt ${attemptNumber}: ${unseatedCount} unseated`);
  } // End random restart loop

  // Use the best result found
  let finalGrid = bestGrid;
  let finalGuests = bestGuests;

  // PHASE 2: SWAP OPTIMIZATION
  // Try to seat remaining unseated guests by swapping with seated guests
  if (bestUnseatedCount > 0) {
    Logger.log(`Starting Phase 2: Swap Optimization for ${bestUnseatedCount} unseated guests`);

    const swapResult = attemptSwapOptimization(
      finalGrid,
      finalGuests,
      ctrlTimeLapse,
      connectionsMap,
      setNeverMatch
    );

    finalGrid = swapResult.grid;
    finalGuests = swapResult.guests;

    const finalUnseatedCount = swapResult.unseatedCount;
    Logger.log(`Phase 2 complete: ${finalUnseatedCount} unseated (reduced by ${bestUnseatedCount - finalUnseatedCount})`);
  }

  // CLEAN UP AND FINISH
  // Write the completed arrGrid back to the sheet
  finalGrid.splice(0, 0, headerGrid);
  gridSheet.getRange('rangeGrid').setValues(finalGrid);

  // Remove the column added for sorting before writing back to the sheet
  if (ctrlSortGuests === 1) {
    for (let i = 0; i < finalGuests.length; i++) {
      if (finalGuests[i][GUEST_COLUMNS.CODE].length > 0) {
        finalGuests[i].pop();
      }
    }
  }

  // Write the arrGuests back to the sheet
  finalGuests.splice(0, 0, headerGuests);
  gridSheet.getRange('rangeGuests').setValues(finalGuests);

  // Report unseated members with attempt info
  Logger.log(`Best result: ${bestUnseatedCount} unseated guests after ${attemptNumber} attempts`);
  reportUnseatedMembers(finalGuests, GUEST_COLUMNS.CODE, GUEST_COLUMNS.SEATED);
}
/** HELPER FUNCTIONS */

function attemptSwapOptimization(grid, guests, timeLapse, connectionsMap, setNeverMatch) {
  // Phase 2: Try to seat unseated guests by swapping with seated guests
  const workingGrid = grid.map(row => [...row]);
  const workingGuests = guests.map(row => [...row]);

  // Find all unseated guests
  const unseatedGuests = [];
  for (let i = 0; i < workingGuests.length; i++) {
    if (workingGuests[i][GUEST_COLUMNS.CODE] &&
        workingGuests[i][GUEST_COLUMNS.CODE].toString().length > 0 &&
        workingGuests[i][GUEST_COLUMNS.SEATED] === "No") {
      unseatedGuests.push({
        index: i,
        code: workingGuests[i][GUEST_COLUMNS.CODE],
        count: workingGuests[i][GUEST_COLUMNS.COUNT]
      });
    }
  }

  let swapsMade = 0;

  // Try to swap each unseated guest
  for (let u = 0; u < unseatedGuests.length; u++) {
    const unseatedGuest = unseatedGuests[u];
    let wasSeated = false;

    // Try each house
    for (let houseIdx = 0; houseIdx < workingGrid.length && !wasSeated; houseIdx++) {
      if (!workingGrid[houseIdx][GRID_COLUMNS.HOST]) continue;

      const houseSeats = workingGrid[houseIdx][GRID_COLUMNS.SEATS];
      const houseSeated = workingGrid[houseIdx][GRID_COLUMNS.SEATED];

      // Get all members currently in this house
      const houseMembers = [];
      for (let j = GRID_COLUMNS.HOST; j <= GRID_COLUMNS.GUEST_5; j++) {
        if (workingGrid[houseIdx][j]) {
          houseMembers.push(workingGrid[houseIdx][j]);
        }
      }

      // Try swapping with each seated guest in this house (except host)
      for (let m = 1; m < houseMembers.length && !wasSeated; m++) {
        const seatedGuestCode = houseMembers[m];

        // Find this seated guest's info
        let seatedGuestIdx = -1;
        let seatedGuestCount = 0;
        for (let i = 0; i < workingGuests.length; i++) {
          if (workingGuests[i][GUEST_COLUMNS.CODE] === seatedGuestCode) {
            seatedGuestIdx = i;
            seatedGuestCount = workingGuests[i][GUEST_COLUMNS.COUNT];
            break;
          }
        }

        if (seatedGuestIdx === -1) continue;

        // Check if unseated guest fits in terms of seats
        const newSeatedCount = houseSeated - seatedGuestCount + unseatedGuest.count;
        if (newSeatedCount > houseSeats) continue;

        // Check if unseated guest is compatible with all other members in house
        let unseatedFits = true;
        for (let h = 0; h < houseMembers.length; h++) {
          if (houseMembers[h] === seatedGuestCode) continue; // Skip the one we're removing

          if (!memberMatch(
            `${unseatedGuest.code}-${houseMembers[h]}`,
            timeLapse,
            connectionsMap,
            setNeverMatch
          )) {
            unseatedFits = false;
            break;
          }
        }

        if (!unseatedFits) continue;

        // Now check if the seated guest can fit somewhere else
        let seatedGuestReseated = false;

        for (let otherHouseIdx = 0; otherHouseIdx < workingGrid.length && !seatedGuestReseated; otherHouseIdx++) {
          if (otherHouseIdx === houseIdx) continue; // Can't put them in same house
          if (!workingGrid[otherHouseIdx][GRID_COLUMNS.HOST]) continue;

          const otherSeats = workingGrid[otherHouseIdx][GRID_COLUMNS.SEATS];
          const otherSeated = workingGrid[otherHouseIdx][GRID_COLUMNS.SEATED];

          // Check if seated guest fits
          if (otherSeated + seatedGuestCount > otherSeats) continue;

          // Get members in other house
          const otherMembers = [];
          for (let j = GRID_COLUMNS.HOST; j <= GRID_COLUMNS.GUEST_5; j++) {
            if (workingGrid[otherHouseIdx][j]) {
              otherMembers.push(workingGrid[otherHouseIdx][j]);
            }
          }

          // Check if seated guest is compatible with other house
          let seatedFitsOther = true;
          for (let h = 0; h < otherMembers.length; h++) {
            if (!memberMatch(
              `${seatedGuestCode}-${otherMembers[h]}`,
              timeLapse,
              connectionsMap,
              setNeverMatch
            )) {
              seatedFitsOther = false;
              break;
            }
          }

          if (seatedFitsOther) {
            // Perform the swap!
            // Remove seated guest from original house
            for (let j = GRID_COLUMNS.HOST; j <= GRID_COLUMNS.GUEST_5; j++) {
              if (workingGrid[houseIdx][j] === seatedGuestCode) {
                workingGrid[houseIdx][j] = null;
                break;
              }
            }
            workingGrid[houseIdx][GRID_COLUMNS.SEATED] = houseSeated - seatedGuestCount;

            // Add unseated guest to original house
            for (let j = GRID_COLUMNS.HOST; j <= GRID_COLUMNS.GUEST_5; j++) {
              if (workingGrid[houseIdx][j] === null) {
                workingGrid[houseIdx][j] = unseatedGuest.code;
                break;
              }
            }
            workingGrid[houseIdx][GRID_COLUMNS.SEATED] = workingGrid[houseIdx][GRID_COLUMNS.SEATED] + unseatedGuest.count;

            // Add seated guest to other house
            for (let j = GRID_COLUMNS.HOST; j <= GRID_COLUMNS.GUEST_5; j++) {
              if (workingGrid[otherHouseIdx][j] === null) {
                workingGrid[otherHouseIdx][j] = seatedGuestCode;
                break;
              }
            }
            workingGrid[otherHouseIdx][GRID_COLUMNS.SEATED] = otherSeated + seatedGuestCount;

            // Update guest statuses
            workingGuests[unseatedGuest.index][GUEST_COLUMNS.SEATED] = null;

            seatedGuestReseated = true;
            wasSeated = true;
            swapsMade++;

            Logger.log(`Swap: ${unseatedGuest.code} into house ${houseIdx + 1}, ${seatedGuestCode} moved to house ${otherHouseIdx + 1}`);
          }
        }
      }
    }
  }

  // Count remaining unseated
  let unseatedCount = 0;
  for (let i = 0; i < workingGuests.length; i++) {
    if (workingGuests[i][GUEST_COLUMNS.CODE] &&
        workingGuests[i][GUEST_COLUMNS.CODE].toString().length > 0 &&
        workingGuests[i][GUEST_COLUMNS.SEATED] === "No") {
      unseatedCount++;
    }
  }

  Logger.log(`Swaps made: ${swapsMade}`);

  return {
    grid: workingGrid,
    guests: workingGuests,
    unseatedCount: unseatedCount
  };
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

function shuffleArray(array) {
  // Fisher-Yates shuffle algorithm for randomizing array order
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
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

  // Get time lapse threshold from control variables
  const arrControl = spreadSheet.getRangeByName('controlVariables').getValues();
  const ctrlTimeLapse = Number(arrControl[1]);

  const arrGrid = gridSheet.getRange('rangeGrid').getValues().slice();
  arrGrid.splice(0, 1); // Remove header

  const arrDB = dbSheet.getDataRange().getValues().slice();
  arrDB.splice(0, 1); // Remove header

  // Load guests data to check for unseated members
  const arrGuests = gridSheet.getRange('rangeGuests').getValues().slice();
  arrGuests.splice(0, 1); // Remove header

  // Load never-match list
  const neverMatchRange = gridSheet.getRange('rangeNeverMatch');
  const arrNeverMatch = neverMatchRange.getValues().slice();
  arrNeverMatch.splice(0, 1); // Remove header

  // Build Map for O(1) lookups instead of O(n) Array.find
  const connectionsMap = {};
  for (let i = 0; i < arrDB.length; i++) {
    connectionsMap[arrDB[i][CONNECTION_COLUMNS.PAIR]] = arrDB[i];
  }

  // Create bidirectional Set for never-match pairs
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

  const auditData = [];
  const houseData = {}; // Track stats per house
  const problemPairs = [];
  const neverMatchViolations = [];
  const unseatedMembers = [];
  let minMonths = Infinity;
  let maxMonths = -Infinity;
  let totalConnections = 0;
  let sumMonths = 0;
  let neverMetCount = 0;

  for (let i = 0; i < arrGrid.length; i++) {
    if (!arrGrid[i][GRID_COLUMNS.HOST]) continue;

    const members = [];
    for (let j = GRID_COLUMNS.HOST; j <= GRID_COLUMNS.GUEST_5; j++) {
      if (arrGrid[i][j]) members.push(arrGrid[i][j]);
    }

    const houseNumber = i + 1;
    houseData[houseNumber] = {
      pairs: [],
      minMonths: Infinity,
      maxMonths: -Infinity,
      sumMonths: 0,
      count: 0
    };

    // Check all pairs and create structured data rows
    for (let m = 0; m < members.length; m++) {
      for (let n = m + 1; n < members.length; n++) {
        const pair = `${members[m]}-${members[n]}`;
        const connection = connectionsMap[pair];
        const monthsApart = connection ? connection[CONNECTION_COLUMNS.MONTHS_APART] : 'Never met';

        let warning = '';

        // Check never-match violations FIRST (critical)
        if (setNeverMatch.has(pair)) {
          warning = '🚫 NEVER MATCH';
          neverMatchViolations.push([houseNumber, members[m], members[n], monthsApart]);
        } else if (typeof monthsApart === 'number') {
          minMonths = Math.min(minMonths, monthsApart);
          maxMonths = Math.max(maxMonths, monthsApart);
          totalConnections++;
          sumMonths += monthsApart;

          // Track house stats
          houseData[houseNumber].minMonths = Math.min(houseData[houseNumber].minMonths, monthsApart);
          houseData[houseNumber].maxMonths = Math.max(houseData[houseNumber].maxMonths, monthsApart);
          houseData[houseNumber].sumMonths += monthsApart;
          houseData[houseNumber].count++;

          // Check if below threshold
          if (monthsApart < ctrlTimeLapse) {
            warning = '⚠️ TOO SOON';
            problemPairs.push([houseNumber, members[m], members[n], monthsApart]);
          }
        } else {
          neverMetCount++;
        }

        // Add row: House | Member1 | Member2 | Months Apart | Warning
        auditData.push([houseNumber, members[m], members[n], monthsApart, warning]);
      }
    }
  }

  // Collect unseated members
  for (let i = 0; i < arrGuests.length; i++) {
    if (arrGuests[i][GUEST_COLUMNS.CODE] &&
        arrGuests[i][GUEST_COLUMNS.CODE].toString().length > 0 &&
        arrGuests[i][GUEST_COLUMNS.SEATED] === "No") {
      unseatedMembers.push([
        arrGuests[i][GUEST_COLUMNS.CODE],
        arrGuests[i][GUEST_COLUMNS.COUNT]
      ]);
    }
  }

  // Get or create Grid Audit sheet
  let auditSheet = spreadSheet.getSheetByName('Grid Audit');
  if (!auditSheet) {
    auditSheet = spreadSheet.insertSheet('Grid Audit');
  } else {
    auditSheet.clear();
  }

  let currentRow = 1;

  // SUMMARY DASHBOARD
  const avgMonths = totalConnections > 0 ? Math.round(sumMonths / totalConnections) : 0;

  auditSheet.getRange(currentRow, 1).setValue('AUDIT SUMMARY').setFontSize(14).setFontWeight('bold');
  currentRow += 2;

  const summaryData = [
    ['Total Connections:', totalConnections],
    ['Never Met:', neverMetCount],
    ['Average Separation:', avgMonths + ' months'],
    ['Minimum Separation:', minMonths === Infinity ? 'N/A' : minMonths + ' months'],
    ['Maximum Separation:', maxMonths === -Infinity ? 'N/A' : maxMonths + ' months'],
    ['Threshold Setting:', ctrlTimeLapse + ' months'],
    ['Unseated Members:', unseatedMembers.length],
    ['🚫 NEVER MATCH VIOLATIONS:', neverMatchViolations.length],
    ['Problem Pairs (Below Threshold):', problemPairs.length]
  ];

  auditSheet.getRange(currentRow, 1, summaryData.length, 2).setValues(summaryData);
  auditSheet.getRange(currentRow, 1, summaryData.length, 1).setFontWeight('bold');
  auditSheet.getRange(currentRow, 1, summaryData.length, 2).setBackground('#f3f3f3');

  // Right justify column B in summary section
  auditSheet.getRange(currentRow, 2, summaryData.length, 1).setHorizontalAlignment('right');

  // Highlight unseated members count if there are any
  if (unseatedMembers.length > 0) {
    auditSheet.getRange(currentRow + 6, 2).setFontColor('orange').setFontWeight('bold');
  }

  // Highlight never-match violations count in red if there are any (CRITICAL)
  if (neverMatchViolations.length > 0) {
    auditSheet.getRange(currentRow + 7, 1, 1, 2).setFontColor('red').setFontWeight('bold').setBackground('#ffcccc');
  }

  // Highlight problem pairs count in red if there are any
  if (problemPairs.length > 0) {
    auditSheet.getRange(currentRow + 8, 2).setFontColor('red').setFontWeight('bold');
  }

  currentRow += summaryData.length + 2;

  // NEVER MATCH VIOLATIONS section (if any) - MOST CRITICAL
  if (neverMatchViolations.length > 0) {
    auditSheet.getRange(currentRow, 1).setValue('🚫 NEVER MATCH VIOLATIONS - CRITICAL').setFontSize(12).setFontWeight('bold').setFontColor('white').setBackground('#cc0000');
    currentRow += 1;

    const violationHeader = [['House', 'Member 1', 'Member 2', 'Months Apart']];
    auditSheet.getRange(currentRow, 1, 1, 4).setValues(violationHeader);
    auditSheet.getRange(currentRow, 1, 1, 4).setFontWeight('bold').setBackground('#cc0000').setFontColor('white');
    currentRow += 1;

    auditSheet.getRange(currentRow, 1, neverMatchViolations.length, 4).setValues(neverMatchViolations);
    auditSheet.getRange(currentRow, 1, neverMatchViolations.length, 4).setBackground('#ffcccc').setFontWeight('bold');
    // Add thick border around violation section
    auditSheet.getRange(currentRow - 1, 1, neverMatchViolations.length + 1, 4)
      .setBorder(true, true, true, true, false, false, '#cc0000', SpreadsheetApp.BorderStyle.SOLID_THICK);
    currentRow += neverMatchViolations.length + 2;
  }

  // Problem Pairs section (if any)
  if (problemPairs.length > 0) {
    auditSheet.getRange(currentRow, 1).setValue('PROBLEM PAIRS (Below Threshold)').setFontSize(12).setFontWeight('bold').setFontColor('red');
    currentRow += 1;

    const problemHeader = [['House', 'Member 1', 'Member 2', 'Months Apart']];
    auditSheet.getRange(currentRow, 1, 1, 4).setValues(problemHeader);
    auditSheet.getRange(currentRow, 1, 1, 4).setFontWeight('bold').setBackground('#ffcccc');
    currentRow += 1;

    auditSheet.getRange(currentRow, 1, problemPairs.length, 4).setValues(problemPairs);
    auditSheet.getRange(currentRow, 1, problemPairs.length, 4).setBackground('#ffe6e6');
    currentRow += problemPairs.length + 2;
  }

  // UNSEATED MEMBERS section (if any)
  if (unseatedMembers.length > 0) {
    auditSheet.getRange(currentRow, 1).setValue('UNSEATED MEMBERS').setFontSize(12).setFontWeight('bold').setFontColor('darkorange');
    currentRow += 1;

    const unseatedHeaderRow = currentRow;
    const unseatedHeader = [['Member Code', 'Party Size']];
    auditSheet.getRange(currentRow, 1, 1, 2).setValues(unseatedHeader);
    auditSheet.getRange(currentRow, 1, 1, 2).setFontWeight('bold').setBackground('#ffe6cc');
    currentRow += 1;

    const unseatedDataRow = currentRow;
    auditSheet.getRange(currentRow, 1, unseatedMembers.length, 2).setValues(unseatedMembers);
    auditSheet.getRange(currentRow, 1, unseatedMembers.length, 2).setBackground('#fff4e6');

    // Right justify column B (Party Size) in header and data
    auditSheet.getRange(unseatedHeaderRow, 2, 1, 1).setHorizontalAlignment('right');
    auditSheet.getRange(unseatedDataRow, 2, unseatedMembers.length, 1).setHorizontalAlignment('right');

    currentRow += unseatedMembers.length + 2;
  }

  // MAIN AUDIT DATA
  auditSheet.getRange(currentRow, 1).setValue('DETAILED AUDIT BY HOUSE').setFontSize(12).setFontWeight('bold');
  currentRow += 2;

  const dataStartRow = currentRow;

  // Write header row
  const header = [['House', 'Member 1', 'Member 2', 'Months Apart', 'Warning']];
  auditSheet.getRange(currentRow, 1, 1, 5).setValues(header);
  auditSheet.getRange(currentRow, 1, 1, 5).setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');

  // Right justify column D (Months Apart) in header
  auditSheet.getRange(currentRow, 4, 1, 1).setHorizontalAlignment('right');

  currentRow += 1;

  // Write audit data with house subtotals
  if (auditData.length > 0) {
    let lastHouse = auditData[0][0];
    let houseStartRow = currentRow;

    for (let i = 0; i < auditData.length; i++) {
      const row = auditData[i];
      const currentHouse = row[0];

      // If we're moving to a new house, add subtotal for previous house
      if (currentHouse !== lastHouse) {
        const houseStats = houseData[lastHouse];
        const avgHouse = houseStats.count > 0 ? Math.round(houseStats.sumMonths / houseStats.count) : 0;
        const subtotalData = [['', 'House ' + lastHouse + ' Summary:',
                              'Min: ' + houseStats.minMonths + ' | Max: ' + houseStats.maxMonths + ' | Avg: ' + avgHouse,
                              '', '']];
        auditSheet.getRange(currentRow, 1, 1, 5).setValues(subtotalData);
        auditSheet.getRange(currentRow, 1, 1, 5).setFontWeight('bold').setFontStyle('italic').setBackground('#d9d9d9');

        // Add border before subtotal
        auditSheet.getRange(houseStartRow, 1, currentRow - houseStartRow, 5)
          .setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

        currentRow += 1;
        houseStartRow = currentRow;
        lastHouse = currentHouse;
      }

      auditSheet.getRange(currentRow, 1, 1, 5).setValues([row]);
      currentRow += 1;
    }

    // Add final house subtotal
    const houseStats = houseData[lastHouse];
    const avgHouse = houseStats.count > 0 ? Math.round(houseStats.sumMonths / houseStats.count) : 0;
    const subtotalData = [['', 'House ' + lastHouse + ' Summary:',
                          'Min: ' + houseStats.minMonths + ' | Max: ' + houseStats.maxMonths + ' | Avg: ' + avgHouse,
                          '', '']];
    auditSheet.getRange(currentRow, 1, 1, 5).setValues(subtotalData);
    auditSheet.getRange(currentRow, 1, 1, 5).setFontWeight('bold').setFontStyle('italic').setBackground('#d9d9d9');

    // Add border for final house
    auditSheet.getRange(houseStartRow, 1, currentRow - houseStartRow, 5)
      .setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    currentRow += 1; // Move past the last subtotal
  }

  // Calculate actual number of rows written (including subtotals)
  const actualDataRows = currentRow - dataStartRow - 1;

  // Apply conditional formatting to Months Apart column (column D)
  const dataRange = auditSheet.getRange(dataStartRow + 1, 4, actualDataRows, 1);

  // Clear existing conditional formatting
  dataRange.clearFormat();

  // Apply formatting based on values
  const values = dataRange.getValues();
  const backgrounds = [];
  const fontColors = [];

  for (let i = 0; i < values.length; i++) {
    const val = values[i][0];

    // Skip subtotal rows
    if (typeof val !== 'number' && val !== 'Never met') {
      backgrounds.push(['#d9d9d9']);
      fontColors.push(['#000000']);
      continue;
    }

    if (val === 'Never met') {
      backgrounds.push(['#cfe2f3']);
      fontColors.push(['#1155cc']);
    } else if (val <= 6) {
      backgrounds.push(['#f4cccc']);
      fontColors.push(['#cc0000']);
    } else if (val <= 12) {
      backgrounds.push(['#fff2cc']);
      fontColors.push(['#bf9000']);
    } else {
      backgrounds.push(['#d9ead3']);
      fontColors.push(['#38761d']);
    }
  }

  dataRange.setBackgrounds(backgrounds);
  dataRange.setFontColors(fontColors);

  // Right justify column D (Months Apart) in data rows - after formatting applied
  dataRange.setHorizontalAlignment('right');

  // Apply alternating colors by house for better readability
  // Read the actual written data to handle subtotal rows correctly
  const writtenData = auditSheet.getRange(dataStartRow + 1, 1, actualDataRows, 5).getValues();
  let currentHouseNum = null;
  let colorIndex = 0;
  const houseColors = ['#ffffff', '#f8f9fa'];

  for (let i = 0; i < actualDataRows; i++) {
    const rowNum = dataStartRow + 1 + i;
    const houseNum = writtenData[i][0];
    const warningValue = writtenData[i][4]; // Warning column

    // Skip subtotal rows (they have empty house number)
    if (!houseNum || houseNum === '') continue;

    if (houseNum !== currentHouseNum) {
      currentHouseNum = houseNum;
      colorIndex = (colorIndex + 1) % 2;
    }

    // Apply to House, Member 1, Member 2 columns only (not Months Apart - already colored)
    auditSheet.getRange(rowNum, 1, 1, 3).setBackground(houseColors[colorIndex]);

    // Warning column - special styling for never-match violations
    if (warningValue && warningValue.includes('NEVER MATCH')) {
      auditSheet.getRange(rowNum, 5, 1, 1).setBackground('#cc0000').setFontColor('white').setFontWeight('bold');
    } else if (warningValue && warningValue.includes('TOO SOON')) {
      auditSheet.getRange(rowNum, 5, 1, 1).setBackground('#ffcccc').setFontColor('#cc0000').setFontWeight('bold');
    } else {
      auditSheet.getRange(rowNum, 5, 1, 1).setBackground(houseColors[colorIndex]);
    }
  }

  // SORTED VIEW BY MONTHS APART
  currentRow += 3;
  auditSheet.getRange(currentRow, 1).setValue('SORTED VIEW: ALL PAIRS BY SEPARATION TIME').setFontSize(12).setFontWeight('bold');
  currentRow += 2;

  // Create sorted copy of audit data (excluding subtotals)
  const sortedData = auditData.slice();
  sortedData.sort((a, b) => {
    const valA = a[3];
    const valB = b[3];

    // Handle "Never met" - put at end
    if (valA === 'Never met' && valB === 'Never met') return 0;
    if (valA === 'Never met') return 1;
    if (valB === 'Never met') return -1;

    // Numeric comparison
    return valA - valB;
  });

  // Write sorted header
  auditSheet.getRange(currentRow, 1, 1, 5).setValues(header);
  auditSheet.getRange(currentRow, 1, 1, 5).setFontWeight('bold').setBackground('#6aa84f').setFontColor('white');

  // Right justify column D (Months Apart) in sorted header
  auditSheet.getRange(currentRow, 4, 1, 1).setHorizontalAlignment('right');

  currentRow += 1;

  const sortedStartRow = currentRow;

  // Write sorted data
  if (sortedData.length > 0) {
    auditSheet.getRange(currentRow, 1, sortedData.length, 5).setValues(sortedData);

    // Right justify column D (Months Apart) in sorted data
    auditSheet.getRange(currentRow, 4, sortedData.length, 1).setHorizontalAlignment('right');

    // Apply conditional formatting to sorted view
    const sortedMonthsRange = auditSheet.getRange(currentRow, 4, sortedData.length, 1);
    const sortedValues = sortedMonthsRange.getValues();
    const sortedBackgrounds = [];
    const sortedFontColors = [];

    for (let i = 0; i < sortedValues.length; i++) {
      const val = sortedValues[i][0];

      if (val === 'Never met') {
        sortedBackgrounds.push(['#cfe2f3']);
        sortedFontColors.push(['#1155cc']);
      } else if (val <= 6) {
        sortedBackgrounds.push(['#f4cccc']);
        sortedFontColors.push(['#cc0000']);
      } else if (val <= 12) {
        sortedBackgrounds.push(['#fff2cc']);
        sortedFontColors.push(['#bf9000']);
      } else {
        sortedBackgrounds.push(['#d9ead3']);
        sortedFontColors.push(['#38761d']);
      }
    }

    sortedMonthsRange.setBackgrounds(sortedBackgrounds);
    sortedMonthsRange.setFontColors(sortedFontColors);

    // Apply styling to Warning column in sorted view
    for (let i = 0; i < sortedData.length; i++) {
      const rowNum = sortedStartRow + i;
      const warningValue = sortedData[i][4]; // Warning column

      if (warningValue && warningValue.includes('NEVER MATCH')) {
        auditSheet.getRange(rowNum, 5, 1, 1).setBackground('#cc0000').setFontColor('white').setFontWeight('bold');
      } else if (warningValue && warningValue.includes('TOO SOON')) {
        auditSheet.getRange(rowNum, 5, 1, 1).setBackground('#ffcccc').setFontColor('#cc0000').setFontWeight('bold');
      }
    }
  }

  // Auto-resize columns for better readability
  auditSheet.autoResizeColumns(1, 5);

  // Add extra padding to columns
  for (let col = 1; col <= 5; col++) {
    const currentWidth = auditSheet.getColumnWidth(col);
    auditSheet.setColumnWidth(col, currentWidth + 20);
  }

  // Enable filters on both main and sorted views
  // Remove existing filter first if it exists
  const existingFilter = auditSheet.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }
  auditSheet.getRange(dataStartRow, 1, actualDataRows + 1, 5).createFilter();

  // Create summary message
  let summaryMsg = `Audit complete! ${totalConnections} connections analyzed.\n\n`;

  // Highlight critical violations first
  if (neverMatchViolations.length > 0) {
    summaryMsg += `🚫 CRITICAL: ${neverMatchViolations.length} NEVER MATCH VIOLATIONS FOUND!\n`;
    summaryMsg += `These pairs should NEVER be seated together.\n\n`;
  }

  if (totalConnections > 0) {
    summaryMsg += `Average separation: ${avgMonths} months\n`;
    summaryMsg += `Range: ${minMonths} to ${maxMonths} months\n`;
    summaryMsg += `Problem pairs (below threshold): ${problemPairs.length}\n`;
    summaryMsg += `Unseated members: ${unseatedMembers.length}\n\n`;
    summaryMsg += `Results written to "Grid Audit" sheet with enhanced formatting.`;
  }

  // Use warning alert if there are never-match violations
  const alertTitle = neverMatchViolations.length > 0 ? '⚠️ CRITICAL VIOLATIONS FOUND' : 'Audit Complete';
  SpreadsheetApp.getUi().alert(alertTitle, summaryMsg, SpreadsheetApp.getUi().ButtonSet.OK);

  // Switch to Grid Audit sheet
  auditSheet.activate();
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