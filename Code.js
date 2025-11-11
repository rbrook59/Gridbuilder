// FUTURE DEVELOPMENT NOTES
// 1. create a list of options where the unseated guests will work best.
// 

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

/** CONFIGURATION CONSTANTS */
const SHUFFLE_THRESHOLD = 0.7; // Shuffle bottom 30% of arrays during random restart
const MAX_ATTEMPTS = 20; // Increased from 10 for better solutions
const LOOKAHEAD_CRITICAL_THRESHOLD = 2; // Look-ahead protection for guests with ≤ N compatible houses

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

/**
 * Builds an enhanced connections map filtered to only tonight's attendees
 * @param {Array} guests - Array of guest rows
 * @param {Array} hosts - Array of host rows
 * @param {Object} connectionsMap - Original connections map from cache
 * @param {number} timeLapse - Time constraint in months
 * @returns {Object} Enhanced map with structure: { "CODE1-CODE2": { monthsApart: number, isConstrained: boolean } }
 */
function buildEnhancedConnectionsMap(guests, hosts, connectionsMap, timeLapse) {
  const enhanced = {};

  // Collect all attendee codes
  const allAttendees = [];
  for (let i = 0; i < hosts.length; i++) {
    if (hosts[i][HOST_COLUMNS.CODE]) {
      allAttendees.push(hosts[i][HOST_COLUMNS.CODE]);
    }
  }
  for (let i = 0; i < guests.length; i++) {
    if (guests[i][GUEST_COLUMNS.CODE]) {
      allAttendees.push(guests[i][GUEST_COLUMNS.CODE]);
    }
  }

  // Build all pairs for tonight's attendees
  for (let i = 0; i < allAttendees.length; i++) {
    for (let j = i + 1; j < allAttendees.length; j++) {
      const key = `${allAttendees[i]}-${allAttendees[j]}`;
      const reverseKey = `${allAttendees[j]}-${allAttendees[i]}`;

      // Check both directions in original map
      const months = connectionsMap[key] || connectionsMap[reverseKey] || 999;

      const pairData = {
        monthsApart: months,
        isConstrained: months < timeLapse
      };

      // Store both directions
      enhanced[key] = pairData;
      enhanced[reverseKey] = pairData;
    }
  }

  Logger.log(`Enhanced connections map built: ${Object.keys(enhanced).length} pairs for ${allAttendees.length} attendees`);
  return enhanced;
}

/**
 * BuildGrid - Guest-Centric Approach
 * Places guests in order, scoring houses based on capacity and connection quality
 */
function buildGrid() {
  // SETUP PHASE
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

  // Load connections from cache and build enhanced map for tonight's attendees
  const originalConnectionsMap = getConnectionsMap();
  const setNeverMatch = buildNeverMatchSet(arrNeverMatch);

  // Build enhanced connections map (filtered to tonight's attendees with pre-computed constraints)
  const connectionsMap = buildEnhancedConnectionsMap(
    arrGuestsOriginal,
    arrHostsOriginal,
    originalConnectionsMap,
    ctrlTimeLapse
  );

  // GUEST-CENTRIC ALGORITHM
  Logger.log('\n========== Starting BuildGrid (Guest-Centric) ==========');

  // Create working copies
  const finalHosts = arrHostsOriginal.map(row => [...row]);
  let finalGuests = arrGuestsOriginal.map(row => [...row]);
  let finalGrid = arrGridOriginal.map(row => [...row]);

  // Clear seated flags if requested
  if (ctrlClearSeated === 1) {
    for (let i = 0; i < finalGuests.length; i++) {
      if (finalGuests[i][GUEST_COLUMNS.CODE]) {
        finalGuests[i][GUEST_COLUMNS.SEATED] = "No";
      }
    }
  }

  // Clear grid if requested
  if (ctrlClearGrid === 1) {
    for (let i = 0; i < finalGrid.length; i++) {
      for (let j = 0; j <= GRID_COLUMNS.GUEST_5; j++) {
        finalGrid[i][j] = null;
      }
    }
  }

  // Initialize grid with hosts
  for (let i = 0; i < finalHosts.length; i++) {
    if (finalHosts[i][HOST_COLUMNS.CODE]) {
      finalGrid[i][GRID_COLUMNS.HOUSE] = i + 1;
      finalGrid[i][GRID_COLUMNS.SEATS] = finalHosts[i][HOST_COLUMNS.SEATS];
      finalGrid[i][GRID_COLUMNS.HOST] = finalHosts[i][HOST_COLUMNS.CODE];
      finalGrid[i][GRID_COLUMNS.SEATED] = finalHosts[i][HOST_COLUMNS.COUNT];
    }
  }

  // PLACE GUESTS - Simple iteration over guest list
  for (let guestIdx = 0; guestIdx < finalGuests.length; guestIdx++) {
    const guestCode = finalGuests[guestIdx][GUEST_COLUMNS.CODE];
    if (!guestCode || guestCode.toString().length === 0) continue;

    // Skip if already seated
    if (finalGuests[guestIdx][GUEST_COLUMNS.SEATED] !== "No") continue;

    const guestCount = finalGuests[guestIdx][GUEST_COLUMNS.COUNT];
    let bestHouseIdx = -1;
    let bestHouseScore = -Infinity;

    // Try each house and score it for this guest
    for (let h = 0; h < finalGrid.length; h++) {
      const hostCode = finalGrid[h][GRID_COLUMNS.HOST];
      if (!hostCode) continue;

      const houseSeats = finalGrid[h][GRID_COLUMNS.SEATS];
      const houseSeated = finalGrid[h][GRID_COLUMNS.SEATED];
      const houseMembers = getHouseMembers(finalGrid[h]);

      // Check capacity
      if (houseSeated + guestCount > houseSeats) continue;

      // Check if slot available (max 6 members)
      if (houseMembers.length >= 6) continue;

      // Check compatibility with ALL current members
      let isCompatible = true;
      for (let m = 0; m < houseMembers.length; m++) {
        const pairKey = `${guestCode}-${houseMembers[m]}`;
        if (!memberMatch(pairKey, ctrlTimeLapse, connectionsMap, setNeverMatch)) {
          isCompatible = false;
          break;
        }
      }

      if (!isCompatible) continue;

      // Score this house for this guest
      let score = 0;

      // Factor 1: Prefer houses with more space remaining (save tight houses for tight guests)
      const spaceRemaining = houseSeats - houseSeated;
      score += spaceRemaining * 5;

      // Factor 2: Prefer houses with fewer current members (distribute evenly)
      score += (6 - houseMembers.length) * 10;

      // Factor 3: Connection quality bonus
      let totalMonthsApart = 0;
      for (let m = 0; m < houseMembers.length; m++) {
        const pairKey = `${guestCode}-${houseMembers[m]}`;
        const connection = connectionsMap[pairKey];
        // Enhanced map always has entries (999 for never met)
        totalMonthsApart += connection.monthsApart;
      }
      score += totalMonthsApart / 10;

      if (score > bestHouseScore) {
        bestHouseScore = score;
        bestHouseIdx = h;
      }
    }

    // Place guest in best house found
    if (bestHouseIdx !== -1) {
      const slot = findNullSlotInHouse(finalGrid[bestHouseIdx], GRID_COLUMNS.HOST, GRID_COLUMNS.GUEST_5);
      if (slot !== -1) {
        finalGrid[bestHouseIdx][slot] = guestCode;
        finalGrid[bestHouseIdx][GRID_COLUMNS.SEATED] += guestCount;
        finalGuests[guestIdx][GUEST_COLUMNS.SEATED] = null;
      }
    }
  }

  // Count unseated
  const unseatedCount = countUnseatedGuests(finalGuests);
  Logger.log(`Phase 1 complete: ${unseatedCount} unseated`);

  // PHASE 2 & 3: Apply existing swap optimization and constraint relaxation
  if (unseatedCount > 0) {
    Logger.log('\n========== UNSEATED GUEST DIAGNOSTIC ==========');
    analyzeUnseatedGuests(finalGrid, finalGuests, ctrlTimeLapse, connectionsMap, setNeverMatch);
    Logger.log('===============================================\n');

    Logger.log(`Starting Phase 2: Swap Optimization for ${unseatedCount} unseated guests`);
    const swapResult = attemptSwapOptimization(finalGrid, finalGuests, ctrlTimeLapse, connectionsMap, setNeverMatch);
    finalGrid = swapResult.grid;
    finalGuests = swapResult.guests;

    const phase2Unseated = swapResult.unseatedCount;
    Logger.log(`Phase 2 complete: ${phase2Unseated} unseated (reduced by ${unseatedCount - phase2Unseated})`);

    if (phase2Unseated > 0) {
      Logger.log('\n========== REMAINING UNSEATED AFTER PHASE 2 ==========');
      analyzeUnseatedGuests(finalGrid, finalGuests, ctrlTimeLapse, connectionsMap, setNeverMatch);
      Logger.log('======================================================\n');

      Logger.log(`\nStarting Phase 3: Selective Constraint Relaxation for ${phase2Unseated} unseated guests`);
      const relaxResult = attemptConstraintRelaxation(finalGrid, finalGuests, ctrlTimeLapse, connectionsMap, setNeverMatch);
      finalGrid = relaxResult.grid;
      finalGuests = relaxResult.guests;

      const phase3Unseated = relaxResult.unseatedCount;
      Logger.log(`Phase 3 complete: ${phase3Unseated} unseated (seated ${relaxResult.guestsSeated} with relaxed constraints)`);

      if (phase3Unseated > 0) {
        Logger.log('\n========== FINAL UNSEATED ANALYSIS ==========');
        analyzeUnseatedGuests(finalGrid, finalGuests, ctrlTimeLapse, connectionsMap, setNeverMatch);
        Logger.log('=============================================\n');
      }
    }
  }

  // CLEANUP (same as buildGrid)
  finalGrid.splice(0, 0, headerGrid);
  gridRange.setValues(finalGrid);

  finalGuests.splice(0, 0, headerGuests);
  guestsRange.setValues(finalGuests);

  reportUnseatedMembers(finalGuests, GUEST_COLUMNS.CODE, GUEST_COLUMNS.SEATED);
}

/** HELPER FUNCTIONS */

/**
 * Separates an array into non-blank and blank rows based on a code column
 * @param {Array} array - Array to separate
 * @param {number} codeColumn - Column index to check for blank values
 * @returns {Object} Object with nonBlank and blank arrays
 */
function separateBlankRows(array, codeColumn) {
  const nonBlank = [];
  const blank = [];

  for (let i = 0; i < array.length; i++) {
    if (array[i][codeColumn] && array[i][codeColumn].toString().length > 0) {
      nonBlank.push(array[i]);
    } else {
      blank.push(array[i]);
    }
  }

  return { nonBlank, blank };
}

/**
 * Collects all member codes from a grid row (host and guests)
 * @param {Array} gridRow - Single row from the grid
 * @returns {Array} Array of member codes
 */
function getHouseMembers(gridRow) {
  const members = [];
  for (let j = GRID_COLUMNS.HOST; j <= GRID_COLUMNS.GUEST_5; j++) {
    if (gridRow[j]) {
      members.push(gridRow[j]);
    }
  }
  return members;
}

/**
 * Counts unseated guests in the guests array
 * @param {Array} guestsArray - Array of guests with seated status
 * @returns {number} Count of unseated guests
 */
function countUnseatedGuests(guestsArray) {
  let count = 0;
  for (let i = 0; i < guestsArray.length; i++) {
    if (guestsArray[i][GUEST_COLUMNS.CODE] &&
        guestsArray[i][GUEST_COLUMNS.CODE].toString().length > 0 &&
        guestsArray[i][GUEST_COLUMNS.SEATED] === "No") {
      count++;
    }
  }
  return count;
}

/**
 * Look-Ahead Blocking Detection
 * Checks if placing a guest in a house would eliminate critical options for other highly constrained guests
 * @param {Object} guestToPlace - Guest being considered for placement {code, count, compatibleHouses}
 * @param {number} houseIdx - Index of house being considered
 * @param {Array} grid - Current grid state
 * @param {Array} guestConstraints - Array of all guest constraint objects
 * @param {Array} guestsArray - Full guests array to check seated status
 * @param {number} timeLapse - Time-lapse threshold for compatibility
 * @param {Object} connectionsMap - Map of connections
 * @param {Set} setNeverMatch - Set of never-match pairs
 * @returns {boolean} True if placement would block a critical guest
 */
function wouldBlockCriticalGuest(guestToPlace, houseIdx, grid, guestConstraints, guestsArray, timeLapse, connectionsMap, setNeverMatch) {
  // Only check for guests with very limited options (≤ LOOKAHEAD_CRITICAL_THRESHOLD compatible houses)
  const CRITICAL_THRESHOLD = LOOKAHEAD_CRITICAL_THRESHOLD;

  // Get the house details after hypothetical placement
  const houseSeats = grid[houseIdx][GRID_COLUMNS.SEATS];
  const currentSeated = grid[houseIdx][GRID_COLUMNS.SEATED];
  const newSeatedCount = currentSeated + guestToPlace.count;
  const houseMembersAfterPlacement = getHouseMembers(grid[houseIdx]);
  houseMembersAfterPlacement.push(guestToPlace.code);

  // Check each other unseated guest
  for (let i = 0; i < guestConstraints.length; i++) {
    const otherGuest = guestConstraints[i];

    // Skip if it's the same guest we're placing
    if (otherGuest.code === guestToPlace.code) continue;

    // Skip if already seated
    if (guestsArray[otherGuest.index][GUEST_COLUMNS.SEATED] !== "No") continue;

    // Only care about critically constrained guests
    if (otherGuest.compatibleHouses > CRITICAL_THRESHOLD) continue;

    // Check if this house is currently compatible for the other guest
    // (before we place guestToPlace)
    let wasCompatibleBefore = true;
    const currentHouseMembers = getHouseMembers(grid[houseIdx]);

    // Check capacity before placement
    if (currentSeated + otherGuest.count > houseSeats) {
      wasCompatibleBefore = false;
    }

    // Check member compatibility before placement
    if (wasCompatibleBefore) {
      for (let m = 0; m < currentHouseMembers.length; m++) {
        const pairKey = `${otherGuest.code}-${currentHouseMembers[m]}`;
        if (!memberMatch(pairKey, timeLapse, connectionsMap, setNeverMatch)) {
          wasCompatibleBefore = false;
          break;
        }
      }
    }

    // If house wasn't compatible before, it's not a blocking issue
    if (!wasCompatibleBefore) continue;

    // Now check if placing guestToPlace would make it incompatible
    let wouldBeCompatibleAfter = true;

    // Check capacity after placement
    if (newSeatedCount + otherGuest.count > houseSeats) {
      wouldBeCompatibleAfter = false;
    }

    // Check member compatibility after placement (including with guestToPlace)
    if (wouldBeCompatibleAfter) {
      for (let m = 0; m < houseMembersAfterPlacement.length; m++) {
        const pairKey = `${otherGuest.code}-${houseMembersAfterPlacement[m]}`;
        if (!memberMatch(pairKey, timeLapse, connectionsMap, setNeverMatch)) {
          wouldBeCompatibleAfter = false;
          break;
        }
      }
    }

    // Check slot availability after placement (max 6 members)
    if (wouldBeCompatibleAfter && houseMembersAfterPlacement.length >= 6) {
      wouldBeCompatibleAfter = false;
    }

    // If placement would remove a compatible option from a critical guest, block it
    if (wasCompatibleBefore && !wouldBeCompatibleAfter) {
      Logger.log(`  ⚠️ Look-ahead: Placing ${guestToPlace.code} in house ${houseIdx + 1} would block critical guest ${otherGuest.code} (${otherGuest.compatibleHouses} options)`);
      return true;
    }
  }

  return false;
}

/**
 * Builds a bidirectional Set of never-match pairs from array
 * @param {Array} neverMatchArray - Array of never-match pairs (already has header removed)
 * @returns {Set} Set containing both directions of each pair
 */
function buildNeverMatchSet(neverMatchArray) {
  const setNeverMatch = new Set();

  for (let i = 0; i < neverMatchArray.length; i++) {
    if (neverMatchArray[i][0] && neverMatchArray[i][0].toString().length > 0) {
      const pair = neverMatchArray[i][0].toString();
      setNeverMatch.add(pair);

      // Add reverse pair
      const parts = pair.split('-');
      if (parts.length === 2) {
        setNeverMatch.add(`${parts[1]}-${parts[0]}`);
      }
    }
  }

  return setNeverMatch;
}

/**
 * Returns background and font colors for months apart value
 * @param {number|string} monthsValue - Months apart or 'Never met'
 * @returns {Object} Object with background and fontColor properties
 */
function getMonthsApartColors(monthsValue) {
  if (monthsValue === 'Never met') {
    return { background: '#cfe2f3', fontColor: '#1155cc' };
  } else if (typeof monthsValue === 'number') {
    if (monthsValue <= 6) {
      return { background: '#f4cccc', fontColor: '#cc0000' };
    } else if (monthsValue <= 12) {
      return { background: '#fff2cc', fontColor: '#bf9000' };
    } else {
      return { background: '#d9ead3', fontColor: '#38761d' };
    }
  }
  // Subtotal or other non-numeric
  return { background: '#d9d9d9', fontColor: '#000000' };
}

/**
 * Finds the first null slot in a grid row between start and end columns
 * @param {Array} gridRow - Grid row to search
 * @param {number} startCol - Starting column index
 * @param {number} endCol - Ending column index
 * @returns {number} Column index of first null, or -1 if none found
 */
function findNullSlotInHouse(gridRow, startCol, endCol) {
  for (let j = startCol; j <= endCol; j++) {
    if (gridRow[j] === null) {
      return j;
    }
  }
  return -1;
}

/**
 * Scores a guest for placement in a house. Higher score = better candidate.
 * Scoring considers: constraint level, compatibility quality, party size balance
 * @param {Object} guestInfo - Object with guest details
 * @param {Array} houseMembers - Current members in the house
 * @param {Object} connectionsMap - Map of all connections
 * @param {Object} connectionCounts - Map of connection counts per member
 * @param {number} houseSpotsRemaining - Available spots in house
 * @param {number} totalUnseatableGuests - Count of guests still unseated
 * @param {Object} criticalGuestMap - Map of guest code -> compatible house count
 * @returns {number} Score for this guest (higher is better)
 */
function scoreGuestForHouse(guestInfo, houseMembers, connectionsMap, connectionCounts, houseSpotsRemaining, totalUnseatedGuests, criticalGuestMap) {
  let score = 0;

  // 0. CRITICAL GUEST PRIORITY: Guests with very few compatible houses MUST be seated first!
  const compatibleHouseCount = criticalGuestMap[guestInfo.code] || 999;
  if (compatibleHouseCount <= 1) {
    score += 500; // CRITICAL: Only 1 house option - must seat now!
  } else if (compatibleHouseCount <= 2) {
    score += 300; // Very constrained - high priority
  } else if (compatibleHouseCount <= 3) {
    score += 150; // Constrained - elevated priority
  } else if (compatibleHouseCount <= 5) {
    score += 50; // Somewhat constrained
  }

  // 1. Constraint Priority: More connections = harder to place = higher priority
  const guestConnectionCount = connectionCounts[guestInfo.code] || 0;
  score += guestConnectionCount * 10; // Weight: 10 points per connection

  // 2. Compatibility Quality: Fewer shared connections with house = better (saves flexibility)
  let sharedConnectionScore = 0;
  for (let i = 0; i < houseMembers.length; i++) {
    const pairKey = `${guestInfo.code}-${houseMembers[i]}`;
    const connection = connectionsMap[pairKey];
    if (connection) {
      const monthsApart = connection[CONNECTION_COLUMNS.MONTHS_APART];
      // Prefer guests who met longer ago (more flexibility for future)
      sharedConnectionScore += Math.min(monthsApart, 24); // Cap at 24 months
    } else {
      // Never met = excellent fit
      sharedConnectionScore += 30;
    }
  }
  score += sharedConnectionScore;

  // 3. Party Size Efficiency: Prefer guests that fit well in remaining space
  const spotsAfterSeating = houseSpotsRemaining - guestInfo.count;
  if (spotsAfterSeating >= 0) {
    // Bonus for filling house efficiently
    if (spotsAfterSeating === 0) {
      score += 50; // Perfect fit - fills house completely
    } else if (spotsAfterSeating === 1) {
      score += 30; // Good fit - leaves room for singles
    } else if (spotsAfterSeating <= 2) {
      score += 20; // Decent fit
    } else {
      score += 10; // Leaves lots of space, but okay
    }
  }

  // 4. Singles Consideration: Slight penalty for singles in early positions
  // But don't block completely - just deprioritize
  if (guestInfo.count === 1 && houseMembers.length < 2) {
    score -= 15; // Small penalty for singles being early
  }

  // 5. Urgency: Boost score if few unseated guests remain (don't be too picky)
  if (totalUnseatedGuests <= 5) {
    score += 25; // Be more flexible near the end
  }

  return score;
}

/**
 * Analyzes guests to identify "critical" guests with very few compatible house options
 * Critical guests should be seated first before their options disappear
 * @param {Array} guests - Array of guest data
 * @param {Array} hosts - Array of host data
 * @param {number} timeLapse - Time lapse threshold
 * @param {Object} connectionsMap - Map of all connections
 * @param {Set} setNeverMatch - Set of never-match pairs
 * @returns {Object} Map of guest code -> number of compatible houses
 */
function analyzeCriticalGuests(guests, hosts, timeLapse, connectionsMap, setNeverMatch) {
  const guestCompatibilityCount = {};

  // For each unseated guest, count how many houses they can fit in
  for (let gIdx = 0; gIdx < guests.length; gIdx++) {
    if (!guests[gIdx][GUEST_COLUMNS.CODE] || guests[gIdx][GUEST_COLUMNS.CODE].toString().length === 0) continue;
    if (guests[gIdx][GUEST_COLUMNS.SEATED] !== "No") continue;

    const guestCode = guests[gIdx][GUEST_COLUMNS.CODE];
    const guestCount = guests[gIdx][GUEST_COLUMNS.COUNT];
    let compatibleHouses = 0;

    // Check compatibility with each host
    for (let hIdx = 0; hIdx < hosts.length; hIdx++) {
      if (!hosts[hIdx][HOST_COLUMNS.CODE] || hosts[hIdx][HOST_COLUMNS.CODE] < 1) continue;

      const hostCode = hosts[hIdx][HOST_COLUMNS.CODE];
      const hostSeats = hosts[hIdx][HOST_COLUMNS.SEATS];
      const hostCount = hosts[hIdx][HOST_COLUMNS.COUNT];

      // Check if guest fits in house (capacity)
      if (guestCount + hostCount > hostSeats) continue;

      // Check if guest is compatible with host
      const pairKey = `${guestCode}-${hostCode}`;
      if (memberMatch(pairKey, timeLapse, connectionsMap, setNeverMatch)) {
        compatibleHouses++;
      }
    }

    guestCompatibilityCount[guestCode] = compatibleHouses;
  }

  return guestCompatibilityCount;
}

// Diagnostic function to analyze why guests are unseated
function analyzeUnseatedGuests(grid, guests, timeLapse, connectionsMap, setNeverMatch) {
  // Find all unseated guests
  const unseatedGuests = [];
  for (let i = 0; i < guests.length; i++) {
    if (guests[i][GUEST_COLUMNS.CODE] &&
        guests[i][GUEST_COLUMNS.CODE].toString().length > 0 &&
        guests[i][GUEST_COLUMNS.SEATED] === "No") {
      unseatedGuests.push({
        index: i,
        code: guests[i][GUEST_COLUMNS.CODE],
        count: guests[i][GUEST_COLUMNS.COUNT],
        name: guests[i][GUEST_COLUMNS.NAME] || guests[i][GUEST_COLUMNS.CODE]
      });
    }
  }

  Logger.log(`Total unseated guests: ${unseatedGuests.length}`);

  // Analyze each unseated guest
  for (let u = 0; u < unseatedGuests.length; u++) {
    const guest = unseatedGuests[u];
    Logger.log(`\nGuest ${u + 1}: ${guest.name} (${guest.code}), Party size: ${guest.count}`);

    let compatibleHouseCount = 0;
    const blockingReasons = {
      capacity: [],
      neverMatch: [],
      timeLapse: [],
      noSpace: []
    };

    // Check each house
    for (let hIdx = 0; hIdx < grid.length; hIdx++) {
      if (!grid[hIdx][GRID_COLUMNS.HOST]) continue;

      const houseNum = hIdx + 1;
      const houseSeats = grid[hIdx][GRID_COLUMNS.SEATS];
      const houseSeated = grid[hIdx][GRID_COLUMNS.SEATED];
      const houseMembers = getHouseMembers(grid[hIdx]);

      // Check capacity
      if (houseSeated + guest.count > houseSeats) {
        blockingReasons.capacity.push(`House ${houseNum} (${houseSeated}/${houseSeats} seated, need ${guest.count} more)`);
        continue;
      }

      // Check if there's actually a slot available (max 5 guests + 1 host)
      if (houseMembers.length >= 6) {
        blockingReasons.noSpace.push(`House ${houseNum} (already has ${houseMembers.length} members)`);
        continue;
      }

      // Check compatibility with each member
      let isCompatible = true;
      let blockingMember = null;
      let blockingReason = null;

      for (let m = 0; m < houseMembers.length; m++) {
        const memberCode = houseMembers[m];
        const pairKey = `${guest.code}-${memberCode}`;

        // Check never-match first
        if (setNeverMatch.has(pairKey)) {
          isCompatible = false;
          blockingMember = memberCode;
          blockingReason = 'never-match';
          break;
        }

        // Check time-lapse
        if (!memberMatch(pairKey, timeLapse, connectionsMap, setNeverMatch)) {
          isCompatible = false;
          blockingMember = memberCode;
          blockingReason = 'time-lapse';

          // Get the actual connection date for more detail
          const connection = connectionsMap[pairKey];
          if (connection) {
            blockingReason = `time-lapse (met ${connection.monthsApart} months ago, need ${timeLapse})`;
          }
          break;
        }
      }

      if (isCompatible) {
        compatibleHouseCount++;
      } else {
        if (blockingReason === 'never-match') {
          blockingReasons.neverMatch.push(`House ${houseNum} (never-match with ${blockingMember})`);
        } else if (blockingReason) {
          blockingReasons.timeLapse.push(`House ${houseNum} (${blockingReason} with ${blockingMember})`);
        }
      }
    }

    Logger.log(`  Compatible houses: ${compatibleHouseCount}`);

    if (compatibleHouseCount === 0) {
      Logger.log(`  ⚠️ CRITICAL: No compatible houses found!`);
    }

    // Log blocking reasons
    if (blockingReasons.capacity.length > 0) {
      Logger.log(`  Blocked by CAPACITY (${blockingReasons.capacity.length} houses):`);
      blockingReasons.capacity.slice(0, 3).forEach(reason => Logger.log(`    - ${reason}`));
      if (blockingReasons.capacity.length > 3) {
        Logger.log(`    ... and ${blockingReasons.capacity.length - 3} more`);
      }
    }

    if (blockingReasons.neverMatch.length > 0) {
      Logger.log(`  Blocked by NEVER-MATCH (${blockingReasons.neverMatch.length} houses):`);
      blockingReasons.neverMatch.forEach(reason => Logger.log(`    - ${reason}`));
    }

    if (blockingReasons.timeLapse.length > 0) {
      Logger.log(`  Blocked by TIME-LAPSE (${blockingReasons.timeLapse.length} houses):`);
      blockingReasons.timeLapse.slice(0, 3).forEach(reason => Logger.log(`    - ${reason}`));
      if (blockingReasons.timeLapse.length > 3) {
        Logger.log(`    ... and ${blockingReasons.timeLapse.length - 3} more`);
      }
    }

    if (blockingReasons.noSpace.length > 0) {
      Logger.log(`  Blocked by NO SLOT (${blockingReasons.noSpace.length} houses):`);
      blockingReasons.noSpace.slice(0, 3).forEach(reason => Logger.log(`    - ${reason}`));
    }
  }
}

function attemptSwapOptimization(grid, guests, timeLapse, connectionsMap, setNeverMatch) {
  // Phase 2: Enhanced multi-strategy swap optimization
  const workingGrid = grid.map(row => [...row]);
  const workingGuests = guests.map(row => [...row]);

  let totalSwapsMade = 0;

  // STRATEGY 1: Simple 1-for-1 swaps (original logic)
  Logger.log('Phase 2.1: Attempting simple 1-for-1 swaps');
  const result1 = attemptSimpleSwaps(workingGrid, workingGuests, timeLapse, connectionsMap, setNeverMatch);
  totalSwapsMade += result1.swapsMade;

  if (result1.unseatedCount === 0) {
    Logger.log(`All guests seated after simple swaps`);
    return result1;
  }

  // STRATEGY 2: 2-way swaps between seated guests
  Logger.log(`Phase 2.2: Attempting 2-way swaps (${result1.unseatedCount} still unseated)`);
  const result2 = attemptTwoWaySwaps(result1.grid, result1.guests, timeLapse, connectionsMap, setNeverMatch);
  totalSwapsMade += result2.swapsMade;

  if (result2.unseatedCount === 0) {
    Logger.log(`All guests seated after 2-way swaps`);
    return result2;
  }

  // STRATEGY 3: Chain swaps for remaining unseated guests
  Logger.log(`Phase 2.3: Attempting chain swaps (${result2.unseatedCount} still unseated)`);
  const result3 = attemptChainSwaps(result2.grid, result2.guests, timeLapse, connectionsMap, setNeverMatch);
  totalSwapsMade += result3.swapsMade;

  if (result3.unseatedCount === 0) {
    Logger.log(`All guests seated after chain swaps`);
    Logger.log(`Total swaps made: ${totalSwapsMade}`);
    return result3;
  }

  // STRATEGY 4: Capacity Consolidation - redistribute singles to create space for couples
  Logger.log(`Phase 2.4: Attempting capacity consolidation (${result3.unseatedCount} still unseated)`);
  const result4 = attemptCapacityConsolidation(result3.grid, result3.guests, timeLapse, connectionsMap, setNeverMatch);
  totalSwapsMade += result4.swapsMade;

  Logger.log(`Total swaps made: ${totalSwapsMade}, Final unseated: ${result4.unseatedCount}`);
  return result4;
}

// STRATEGY 1: Simple 1-for-1 swaps
function attemptSimpleSwaps(grid, guests, timeLapse, connectionsMap, setNeverMatch) {
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
      const houseMembers = getHouseMembers(workingGrid[houseIdx]);

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
          if (otherHouseIdx === houseIdx) continue;
          if (!workingGrid[otherHouseIdx][GRID_COLUMNS.HOST]) continue;

          const otherSeats = workingGrid[otherHouseIdx][GRID_COLUMNS.SEATS];
          const otherSeated = workingGrid[otherHouseIdx][GRID_COLUMNS.SEATED];

          if (otherSeated + seatedGuestCount > otherSeats) continue;

          const otherMembers = getHouseMembers(workingGrid[otherHouseIdx]);

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
            // Perform the swap
            for (let j = GRID_COLUMNS.HOST; j <= GRID_COLUMNS.GUEST_5; j++) {
              if (workingGrid[houseIdx][j] === seatedGuestCode) {
                workingGrid[houseIdx][j] = null;
                break;
              }
            }
            workingGrid[houseIdx][GRID_COLUMNS.SEATED] = houseSeated - seatedGuestCount;

            const slotForUnseated = findNullSlotInHouse(workingGrid[houseIdx], GRID_COLUMNS.HOST, GRID_COLUMNS.GUEST_5);
            if (slotForUnseated !== -1) {
              workingGrid[houseIdx][slotForUnseated] = unseatedGuest.code;
            }
            workingGrid[houseIdx][GRID_COLUMNS.SEATED] = workingGrid[houseIdx][GRID_COLUMNS.SEATED] + unseatedGuest.count;

            const slotForSeated = findNullSlotInHouse(workingGrid[otherHouseIdx], GRID_COLUMNS.HOST, GRID_COLUMNS.GUEST_5);
            if (slotForSeated !== -1) {
              workingGrid[otherHouseIdx][slotForSeated] = seatedGuestCode;
            }
            workingGrid[otherHouseIdx][GRID_COLUMNS.SEATED] = otherSeated + seatedGuestCount;

            workingGuests[unseatedGuest.index][GUEST_COLUMNS.SEATED] = null;

            seatedGuestReseated = true;
            wasSeated = true;
            swapsMade++;

            Logger.log(`Simple swap: ${unseatedGuest.code} into house ${houseIdx + 1}, ${seatedGuestCode} to house ${otherHouseIdx + 1}`);
          }
        }
      }
    }
  }

  return {
    grid: workingGrid,
    guests: workingGuests,
    unseatedCount: countUnseatedGuests(workingGuests),
    swapsMade: swapsMade
  };
}

// STRATEGY 2: 2-way swaps - swap two seated guests between houses
function attemptTwoWaySwaps(grid, guests, timeLapse, connectionsMap, setNeverMatch) {
  const workingGrid = grid.map(row => [...row]);
  const workingGuests = guests.map(row => [...row]);
  let swapsMade = 0;

  // Find unseated guests
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

  if (unseatedGuests.length === 0) {
    return {
      grid: workingGrid,
      guests: workingGuests,
      unseatedCount: 0,
      swapsMade: 0
    };
  }

  // Try swapping seated guests between two houses to create space for unseated
  for (let house1 = 0; house1 < workingGrid.length; house1++) {
    if (!workingGrid[house1][GRID_COLUMNS.HOST]) continue;

    const house1Members = getHouseMembers(workingGrid[house1]);

    for (let house2 = house1 + 1; house2 < workingGrid.length; house2++) {
      if (!workingGrid[house2][GRID_COLUMNS.HOST]) continue;

      const house2Members = getHouseMembers(workingGrid[house2]);

      // Try swapping each guest from house1 with each guest from house2
      for (let g1 = 1; g1 < house1Members.length; g1++) { // Skip host (index 0)
        const guest1Code = house1Members[g1];
        const guest1Info = findGuestInfo(workingGuests, guest1Code);
        if (!guest1Info) continue;

        for (let g2 = 1; g2 < house2Members.length; g2++) { // Skip host
          const guest2Code = house2Members[g2];
          const guest2Info = findGuestInfo(workingGuests, guest2Code);
          if (!guest2Info) continue;

          // Check if swapping would maintain constraints
          // Guest1 goes to house2, Guest2 goes to house1

          // Check seats
          const house1Seats = workingGrid[house1][GRID_COLUMNS.SEATS];
          const house1Seated = workingGrid[house1][GRID_COLUMNS.SEATED];
          const house2Seats = workingGrid[house2][GRID_COLUMNS.SEATS];
          const house2Seated = workingGrid[house2][GRID_COLUMNS.SEATED];

          const newHouse1Seated = house1Seated - guest1Info.count + guest2Info.count;
          const newHouse2Seated = house2Seated - guest2Info.count + guest1Info.count;

          if (newHouse1Seated > house1Seats || newHouse2Seated > house2Seats) continue;

          // Check compatibility: guest1 with house2 members (excluding guest2)
          let guest1FitsHouse2 = true;
          for (let m = 0; m < house2Members.length; m++) {
            if (house2Members[m] === guest2Code) continue;
            if (!memberMatch(`${guest1Code}-${house2Members[m]}`, timeLapse, connectionsMap, setNeverMatch)) {
              guest1FitsHouse2 = false;
              break;
            }
          }

          if (!guest1FitsHouse2) continue;

          // Check compatibility: guest2 with house1 members (excluding guest1)
          let guest2FitsHouse1 = true;
          for (let m = 0; m < house1Members.length; m++) {
            if (house1Members[m] === guest1Code) continue;
            if (!memberMatch(`${guest2Code}-${house1Members[m]}`, timeLapse, connectionsMap, setNeverMatch)) {
              guest2FitsHouse1 = false;
              break;
            }
          }

          if (!guest2FitsHouse1) continue;

          // Check if this swap creates space for an unseated guest
          let createsOpportunity = false;
          for (let u = 0; u < unseatedGuests.length; u++) {
            const unseatedGuest = unseatedGuests[u];

            // Check if unseated can fit in house1 after swap
            if (newHouse1Seated + unseatedGuest.count <= house1Seats) {
              const house1AfterSwap = house1Members.filter(m => m !== guest1Code);
              house1AfterSwap.push(guest2Code);

              let unseatedFitsHouse1 = true;
              for (let m = 0; m < house1AfterSwap.length; m++) {
                if (!memberMatch(`${unseatedGuest.code}-${house1AfterSwap[m]}`, timeLapse, connectionsMap, setNeverMatch)) {
                  unseatedFitsHouse1 = false;
                  break;
                }
              }

              if (unseatedFitsHouse1) {
                createsOpportunity = true;
                break;
              }
            }

            // Check if unseated can fit in house2 after swap
            if (newHouse2Seated + unseatedGuest.count <= house2Seats) {
              const house2AfterSwap = house2Members.filter(m => m !== guest2Code);
              house2AfterSwap.push(guest1Code);

              let unseatedFitsHouse2 = true;
              for (let m = 0; m < house2AfterSwap.length; m++) {
                if (!memberMatch(`${unseatedGuest.code}-${house2AfterSwap[m]}`, timeLapse, connectionsMap, setNeverMatch)) {
                  unseatedFitsHouse2 = false;
                  break;
                }
              }

              if (unseatedFitsHouse2) {
                createsOpportunity = true;
                break;
              }
            }
          }

          if (createsOpportunity) {
            // Perform the 2-way swap
            // Remove guest1 from house1
            for (let j = GRID_COLUMNS.HOST; j <= GRID_COLUMNS.GUEST_5; j++) {
              if (workingGrid[house1][j] === guest1Code) {
                workingGrid[house1][j] = null;
                break;
              }
            }

            // Remove guest2 from house2
            for (let j = GRID_COLUMNS.HOST; j <= GRID_COLUMNS.GUEST_5; j++) {
              if (workingGrid[house2][j] === guest2Code) {
                workingGrid[house2][j] = null;
                break;
              }
            }

            // Add guest2 to house1
            const slot1 = findNullSlotInHouse(workingGrid[house1], GRID_COLUMNS.HOST, GRID_COLUMNS.GUEST_5);
            if (slot1 !== -1) workingGrid[house1][slot1] = guest2Code;

            // Add guest1 to house2
            const slot2 = findNullSlotInHouse(workingGrid[house2], GRID_COLUMNS.HOST, GRID_COLUMNS.GUEST_5);
            if (slot2 !== -1) workingGrid[house2][slot2] = guest1Code;

            // Update seated counts
            workingGrid[house1][GRID_COLUMNS.SEATED] = newHouse1Seated;
            workingGrid[house2][GRID_COLUMNS.SEATED] = newHouse2Seated;

            swapsMade++;
            Logger.log(`2-way swap: ${guest1Code} (house ${house1 + 1}→${house2 + 1}) ↔ ${guest2Code} (house ${house2 + 1}→${house1 + 1})`);

            // Now try to place an unseated guest in the newly opened spot
            for (let u = 0; u < unseatedGuests.length; u++) {
              const unseatedGuest = unseatedGuests[u];
              if (workingGuests[unseatedGuest.index][GUEST_COLUMNS.SEATED] !== "No") continue;

              // Try house1
              if (workingGrid[house1][GRID_COLUMNS.SEATED] + unseatedGuest.count <= workingGrid[house1][GRID_COLUMNS.SEATS]) {
                const h1Members = getHouseMembers(workingGrid[house1]);
                let fits = true;
                for (let m = 0; m < h1Members.length; m++) {
                  if (!memberMatch(`${unseatedGuest.code}-${h1Members[m]}`, timeLapse, connectionsMap, setNeverMatch)) {
                    fits = false;
                    break;
                  }
                }

                if (fits) {
                  const slot = findNullSlotInHouse(workingGrid[house1], GRID_COLUMNS.HOST, GRID_COLUMNS.GUEST_5);
                  if (slot !== -1) {
                    workingGrid[house1][slot] = unseatedGuest.code;
                    workingGrid[house1][GRID_COLUMNS.SEATED] += unseatedGuest.count;
                    workingGuests[unseatedGuest.index][GUEST_COLUMNS.SEATED] = null;
                    Logger.log(`  → Seated ${unseatedGuest.code} in house ${house1 + 1}`);
                    break;
                  }
                }
              }

              // Try house2
              if (workingGrid[house2][GRID_COLUMNS.SEATED] + unseatedGuest.count <= workingGrid[house2][GRID_COLUMNS.SEATS]) {
                const h2Members = getHouseMembers(workingGrid[house2]);
                let fits = true;
                for (let m = 0; m < h2Members.length; m++) {
                  if (!memberMatch(`${unseatedGuest.code}-${h2Members[m]}`, timeLapse, connectionsMap, setNeverMatch)) {
                    fits = false;
                    break;
                  }
                }

                if (fits) {
                  const slot = findNullSlotInHouse(workingGrid[house2], GRID_COLUMNS.HOST, GRID_COLUMNS.GUEST_5);
                  if (slot !== -1) {
                    workingGrid[house2][slot] = unseatedGuest.code;
                    workingGrid[house2][GRID_COLUMNS.SEATED] += unseatedGuest.count;
                    workingGuests[unseatedGuest.index][GUEST_COLUMNS.SEATED] = null;
                    Logger.log(`  → Seated ${unseatedGuest.code} in house ${house2 + 1}`);
                    break;
                  }
                }
              }
            }
          }
        }
      }
    }
  }

  return {
    grid: workingGrid,
    guests: workingGuests,
    unseatedCount: countUnseatedGuests(workingGuests),
    swapsMade: swapsMade
  };
}

// STRATEGY 3: Chain swaps - try relocating multiple guests to seat an unseated one
function attemptChainSwaps(grid, guests, timeLapse, connectionsMap, setNeverMatch) {
  const workingGrid = grid.map(row => [...row]);
  const workingGuests = guests.map(row => [...row]);
  let swapsMade = 0;

  // Find unseated guests
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

  // For each unseated guest, try chain relocations
  for (let u = 0; u < unseatedGuests.length; u++) {
    const unseatedGuest = unseatedGuests[u];
    if (workingGuests[unseatedGuest.index][GUEST_COLUMNS.SEATED] !== "No") continue;

    // Try each house as a potential destination
    for (let targetHouse = 0; targetHouse < workingGrid.length; targetHouse++) {
      if (!workingGrid[targetHouse][GRID_COLUMNS.HOST]) continue;

      const houseSeats = workingGrid[targetHouse][GRID_COLUMNS.SEATS];
      const houseSeated = workingGrid[targetHouse][GRID_COLUMNS.SEATED];
      const houseMembers = getHouseMembers(workingGrid[targetHouse]);

      // If unseated fits directly, it would have been seated already, so skip
      if (houseSeated + unseatedGuest.count <= houseSeats) {
        let directFit = true;
        for (let m = 0; m < houseMembers.length; m++) {
          if (!memberMatch(`${unseatedGuest.code}-${houseMembers[m]}`, timeLapse, connectionsMap, setNeverMatch)) {
            directFit = false;
            break;
          }
        }
        if (directFit) continue; // Should have been seated already
      }

      // Try removing each guest (except host) and see if unseated fits
      for (let m = 1; m < houseMembers.length; m++) {
        const blockingGuestCode = houseMembers[m];
        const blockingGuestInfo = findGuestInfo(workingGuests, blockingGuestCode);
        if (!blockingGuestInfo) continue;

        // Check if removing this guest makes room
        const seatsAfterRemoval = houseSeated - blockingGuestInfo.count;
        if (seatsAfterRemoval + unseatedGuest.count > houseSeats) continue;

        // Check if unseated fits with remaining members
        let unseatedFits = true;
        for (let h = 0; h < houseMembers.length; h++) {
          if (houseMembers[h] === blockingGuestCode) continue;
          if (!memberMatch(`${unseatedGuest.code}-${houseMembers[h]}`, timeLapse, connectionsMap, setNeverMatch)) {
            unseatedFits = false;
            break;
          }
        }

        if (!unseatedFits) continue;

        // Now find where to relocate the blocking guest (chain)
        for (let newHouse = 0; newHouse < workingGrid.length; newHouse++) {
          if (newHouse === targetHouse) continue;
          if (!workingGrid[newHouse][GRID_COLUMNS.HOST]) continue;

          const newHouseSeats = workingGrid[newHouse][GRID_COLUMNS.SEATS];
          const newHouseSeated = workingGrid[newHouse][GRID_COLUMNS.SEATED];

          if (newHouseSeated + blockingGuestInfo.count > newHouseSeats) continue;

          const newHouseMembers = getHouseMembers(workingGrid[newHouse]);

          let blockingFitsNew = true;
          for (let h = 0; h < newHouseMembers.length; h++) {
            if (!memberMatch(`${blockingGuestCode}-${newHouseMembers[h]}`, timeLapse, connectionsMap, setNeverMatch)) {
              blockingFitsNew = false;
              break;
            }
          }

          if (blockingFitsNew) {
            // Execute the chain swap
            // Remove blocking guest from target house
            for (let j = GRID_COLUMNS.HOST; j <= GRID_COLUMNS.GUEST_5; j++) {
              if (workingGrid[targetHouse][j] === blockingGuestCode) {
                workingGrid[targetHouse][j] = null;
                break;
              }
            }
            workingGrid[targetHouse][GRID_COLUMNS.SEATED] -= blockingGuestInfo.count;

            // Add blocking guest to new house
            const slotForBlocking = findNullSlotInHouse(workingGrid[newHouse], GRID_COLUMNS.HOST, GRID_COLUMNS.GUEST_5);
            if (slotForBlocking !== -1) {
              workingGrid[newHouse][slotForBlocking] = blockingGuestCode;
            }
            workingGrid[newHouse][GRID_COLUMNS.SEATED] += blockingGuestInfo.count;

            // Add unseated guest to target house
            const slotForUnseated = findNullSlotInHouse(workingGrid[targetHouse], GRID_COLUMNS.HOST, GRID_COLUMNS.GUEST_5);
            if (slotForUnseated !== -1) {
              workingGrid[targetHouse][slotForUnseated] = unseatedGuest.code;
            }
            workingGrid[targetHouse][GRID_COLUMNS.SEATED] += unseatedGuest.count;

            // Update guest status
            workingGuests[unseatedGuest.index][GUEST_COLUMNS.SEATED] = null;

            swapsMade++;
            Logger.log(`Chain swap: ${blockingGuestCode} (house ${targetHouse + 1}→${newHouse + 1}), ${unseatedGuest.code} → house ${targetHouse + 1}`);

            // Move to next unseated guest
            break;
          }
        }

        // If we seated this unseated guest, move to next one
        if (workingGuests[unseatedGuest.index][GUEST_COLUMNS.SEATED] !== "No") break;
      }

      // If we seated this unseated guest, move to next one
      if (workingGuests[unseatedGuest.index][GUEST_COLUMNS.SEATED] !== "No") break;
    }
  }

  return {
    grid: workingGrid,
    guests: workingGuests,
    unseatedCount: countUnseatedGuests(workingGuests),
    swapsMade: swapsMade
  };
}

// STRATEGY 4: Capacity Consolidation - move singles around to create space for couples
function attemptCapacityConsolidation(grid, guests, timeLapse, connectionsMap, setNeverMatch) {
  const workingGrid = grid.map(row => [...row]);
  const workingGuests = guests.map(row => [...row]);
  let swapsMade = 0;

  // Find unseated guests
  const unseatedGuests = [];
  for (let i = 0; i < workingGuests.length; i++) {
    if (workingGuests[i][GUEST_COLUMNS.CODE] &&
        workingGuests[i][GUEST_COLUMNS.CODE].toString().length > 0 &&
        workingGuests[i][GUEST_COLUMNS.SEATED] === "No") {
      unseatedGuests.push({
        index: i,
        code: workingGuests[i][GUEST_COLUMNS.CODE],
        count: workingGuests[i][GUEST_COLUMNS.COUNT],
        name: workingGuests[i][GUEST_COLUMNS.NAME] || workingGuests[i][GUEST_COLUMNS.CODE]
      });
    }
  }

  if (unseatedGuests.length === 0) {
    return {
      grid: workingGrid,
      guests: workingGuests,
      unseatedCount: 0,
      swapsMade: 0
    };
  }

  // Calculate total capacity available
  let totalCapacity = 0;
  let totalSeated = 0;
  for (let h = 0; h < workingGrid.length; h++) {
    if (workingGrid[h][GRID_COLUMNS.HOST]) {
      totalCapacity += workingGrid[h][GRID_COLUMNS.SEATS];
      totalSeated += workingGrid[h][GRID_COLUMNS.SEATED];
    }
  }

  const totalOpenSpots = totalCapacity - totalSeated;
  let totalUnseatedCount = 0;
  for (let u = 0; u < unseatedGuests.length; u++) {
    totalUnseatedCount += unseatedGuests[u].count;
  }

  Logger.log(`Capacity check: ${totalOpenSpots} open spots, need ${totalUnseatedCount} spots`);

  if (totalOpenSpots < totalUnseatedCount) {
    Logger.log('Insufficient total capacity - cannot seat all guests');
    return {
      grid: workingGrid,
      guests: workingGuests,
      unseatedCount: countUnseatedGuests(workingGuests),
      swapsMade: 0
    };
  }

  // Focus on unseated couples (party size 2)
  const unseatedCouples = unseatedGuests.filter(g => g.count === 2);

  for (let u = 0; u < unseatedCouples.length; u++) {
    const couple = unseatedCouples[u];
    if (workingGuests[couple.index][GUEST_COLUMNS.SEATED] !== "No") continue;

    // Find houses with exactly 1 open spot that this couple could fit in (if they had 2 spots)
    for (let targetHouse = 0; targetHouse < workingGrid.length; targetHouse++) {
      if (!workingGrid[targetHouse][GRID_COLUMNS.HOST]) continue;

      const houseSeats = workingGrid[targetHouse][GRID_COLUMNS.SEATS];
      const houseSeated = workingGrid[targetHouse][GRID_COLUMNS.SEATED];
      const spotsAvailable = houseSeats - houseSeated;

      // Need exactly 1 spot free (will create 2nd spot by moving a single)
      if (spotsAvailable !== 1) continue;

      const houseMembers = getHouseMembers(workingGrid[targetHouse]);

      // Check if couple would be compatible with this house (ignoring capacity for now)
      let coupleCompatible = true;
      for (let m = 0; m < houseMembers.length; m++) {
        const pairKey = `${couple.code}-${houseMembers[m]}`;
        if (!memberMatch(pairKey, timeLapse, connectionsMap, setNeverMatch)) {
          coupleCompatible = false;
          break;
        }
      }

      if (!coupleCompatible) continue;

      // Find a single guest in this house to relocate
      for (let m = 1; m < houseMembers.length; m++) { // Skip host (index 0)
        const singleGuestCode = houseMembers[m];
        const singleGuestInfo = findGuestInfo(workingGuests, singleGuestCode);

        // Only move singles
        if (!singleGuestInfo || singleGuestInfo.count !== 1) continue;

        // Try to find a destination for this single
        for (let destHouse = 0; destHouse < workingGrid.length; destHouse++) {
          if (destHouse === targetHouse) continue;
          if (!workingGrid[destHouse][GRID_COLUMNS.HOST]) continue;

          const destSeats = workingGrid[destHouse][GRID_COLUMNS.SEATS];
          const destSeated = workingGrid[destHouse][GRID_COLUMNS.SEATED];

          // Check if single fits
          if (destSeated + 1 > destSeats) continue;

          const destMembers = getHouseMembers(workingGrid[destHouse]);

          // Check if single is compatible with destination house
          let singleFits = true;
          for (let m = 0; m < destMembers.length; m++) {
            const pairKey = `${singleGuestCode}-${destMembers[m]}`;
            if (!memberMatch(pairKey, timeLapse, connectionsMap, setNeverMatch)) {
              singleFits = false;
              break;
            }
          }

          if (singleFits) {
            // Execute the consolidation:
            // 1. Remove single from target house
            for (let j = GRID_COLUMNS.HOST; j <= GRID_COLUMNS.GUEST_5; j++) {
              if (workingGrid[targetHouse][j] === singleGuestCode) {
                workingGrid[targetHouse][j] = null;
                break;
              }
            }
            workingGrid[targetHouse][GRID_COLUMNS.SEATED] -= 1;

            // 2. Add single to destination house
            const slotForSingle = findNullSlotInHouse(workingGrid[destHouse], GRID_COLUMNS.HOST, GRID_COLUMNS.GUEST_5);
            if (slotForSingle !== -1) {
              workingGrid[destHouse][slotForSingle] = singleGuestCode;
            }
            workingGrid[destHouse][GRID_COLUMNS.SEATED] += 1;

            // 3. Now target house has 2 spots - place the couple
            const slot1 = findNullSlotInHouse(workingGrid[targetHouse], GRID_COLUMNS.HOST, GRID_COLUMNS.GUEST_5);
            if (slot1 !== -1) {
              workingGrid[targetHouse][slot1] = couple.code;
              workingGrid[targetHouse][GRID_COLUMNS.SEATED] += couple.count;
              workingGuests[couple.index][GUEST_COLUMNS.SEATED] = null;

              swapsMade++;
              Logger.log(`Capacity consolidation: Moved ${singleGuestCode} (house ${targetHouse + 1}→${destHouse + 1}), seated ${couple.name} in house ${targetHouse + 1}`);

              // Move to next unseated couple
              break;
            }
          }
        }

        // If couple was seated, break out of single guest loop
        if (workingGuests[couple.index][GUEST_COLUMNS.SEATED] !== "No") break;
      }

      // If couple was seated, break out of target house loop
      if (workingGuests[couple.index][GUEST_COLUMNS.SEATED] !== "No") break;
    }
  }

  return {
    grid: workingGrid,
    guests: workingGuests,
    unseatedCount: countUnseatedGuests(workingGuests),
    swapsMade: swapsMade
  };
}

// Helper function to find guest info
function findGuestInfo(guests, guestCode) {
  for (let i = 0; i < guests.length; i++) {
    if (guests[i][GUEST_COLUMNS.CODE] === guestCode) {
      return {
        index: i,
        code: guestCode,
        count: guests[i][GUEST_COLUMNS.COUNT]
      };
    }
  }
  return null;
}

// PHASE 3: Selective Constraint Relaxation
// For guests with no compatible houses, progressively lower time-lapse constraint
function attemptConstraintRelaxation(grid, guests, originalTimeLapse, connectionsMap, setNeverMatch) {
  const workingGrid = grid.map(row => [...row]);
  const workingGuests = guests.map(row => [...row]);
  let guestsSeated = 0;

  // Find unseated guests
  const unseatedGuests = [];
  for (let i = 0; i < workingGuests.length; i++) {
    if (workingGuests[i][GUEST_COLUMNS.CODE] &&
        workingGuests[i][GUEST_COLUMNS.CODE].toString().length > 0 &&
        workingGuests[i][GUEST_COLUMNS.SEATED] === "No") {
      unseatedGuests.push({
        index: i,
        code: workingGuests[i][GUEST_COLUMNS.CODE],
        count: workingGuests[i][GUEST_COLUMNS.COUNT],
        name: workingGuests[i][GUEST_COLUMNS.NAME] || workingGuests[i][GUEST_COLUMNS.CODE]
      });
    }
  }

  // Progressive thresholds to try (don't go below 12 months)
  const thresholds = [24, 18, 12];

  // For each unseated guest
  for (let u = 0; u < unseatedGuests.length; u++) {
    const guest = unseatedGuests[u];

    // Check if already seated (may have been seated by previous guest's placement)
    if (workingGuests[guest.index][GUEST_COLUMNS.SEATED] !== "No") continue;

    // First, count compatible houses at original threshold
    let compatibleAtOriginal = 0;
    for (let hIdx = 0; hIdx < workingGrid.length; hIdx++) {
      if (!workingGrid[hIdx][GRID_COLUMNS.HOST]) continue;

      const houseSeats = workingGrid[hIdx][GRID_COLUMNS.SEATS];
      const houseSeated = workingGrid[hIdx][GRID_COLUMNS.SEATED];
      const houseMembers = getHouseMembers(workingGrid[hIdx]);

      if (houseSeated + guest.count > houseSeats) continue;
      if (houseMembers.length >= 6) continue;

      let isCompatible = true;
      for (let m = 0; m < houseMembers.length; m++) {
        const pairKey = `${guest.code}-${houseMembers[m]}`;
        if (!memberMatch(pairKey, originalTimeLapse, connectionsMap, setNeverMatch)) {
          isCompatible = false;
          break;
        }
      }

      if (isCompatible) compatibleAtOriginal++;
    }

    // Only relax constraints for guests with 0 compatible houses
    if (compatibleAtOriginal > 0) {
      Logger.log(`Phase 3: ${guest.name} has ${compatibleAtOriginal} compatible house(s), skipping relaxation`);
      continue;
    }

    Logger.log(`Phase 3: ${guest.name} has 0 compatible houses at ${originalTimeLapse} months, trying relaxed constraints...`);

    // Try progressively lower thresholds
    let wasSeated = false;
    for (let t = 0; t < thresholds.length && !wasSeated; t++) {
      const relaxedThreshold = thresholds[t];

      // Try each house with relaxed constraint
      for (let hIdx = 0; hIdx < workingGrid.length && !wasSeated; hIdx++) {
        if (!workingGrid[hIdx][GRID_COLUMNS.HOST]) continue;

        const houseNum = hIdx + 1;
        const houseSeats = workingGrid[hIdx][GRID_COLUMNS.SEATS];
        const houseSeated = workingGrid[hIdx][GRID_COLUMNS.SEATED];
        const houseMembers = getHouseMembers(workingGrid[hIdx]);

        // Check capacity
        if (houseSeated + guest.count > houseSeats) continue;
        if (houseMembers.length >= 6) continue;

        // Check compatibility with relaxed time-lapse
        let isCompatible = true;
        for (let m = 0; m < houseMembers.length; m++) {
          const pairKey = `${guest.code}-${houseMembers[m]}`;

          // Never-match is still enforced
          if (setNeverMatch.has(pairKey)) {
            isCompatible = false;
            break;
          }

          // Check with relaxed threshold
          if (!memberMatch(pairKey, relaxedThreshold, connectionsMap, setNeverMatch)) {
            isCompatible = false;
            break;
          }
        }

        if (isCompatible) {
          // Seat the guest with relaxed constraint
          const slot = findNullSlotInHouse(workingGrid[hIdx], GRID_COLUMNS.HOST, GRID_COLUMNS.GUEST_5);
          if (slot !== -1) {
            workingGrid[hIdx][slot] = guest.code;
            workingGrid[hIdx][GRID_COLUMNS.SEATED] = houseSeated + guest.count;
            workingGuests[guest.index][GUEST_COLUMNS.SEATED] = null;
            wasSeated = true;
            guestsSeated++;

            Logger.log(`  ✓ Seated ${guest.name} in house ${houseNum} with ${relaxedThreshold}-month constraint (relaxed from ${originalTimeLapse})`);
          }
        }
      }

      if (wasSeated) break;
    }

    if (!wasSeated) {
      Logger.log(`  ✗ Could not seat ${guest.name} even with relaxed constraints (12+ months)`);
    }
  }

  return {
    grid: workingGrid,
    guests: workingGuests,
    unseatedCount: countUnseatedGuests(workingGuests),
    guestsSeated: guestsSeated
  };
}

function memberMatch(strToCheck, timeLapse, connectionsMap, setNeverMatch) {
  // Check never-match list (O(1) lookup)
  if (setNeverMatch.has(strToCheck)) {
    return false;
  }

  // Check time-based constraints using enhanced connections map
  const connection = connectionsMap[strToCheck];
  if (connection && connection.isConstrained) {
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
  const setNeverMatch = buildNeverMatchSet(arrNeverMatch);

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

    const members = getHouseMembers(arrGrid[i]);
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
    ['Total Prior Connections:', totalConnections],
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
    const colors = getMonthsApartColors(val);
    backgrounds.push([colors.background]);
    fontColors.push([colors.fontColor]);
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
      const colors = getMonthsApartColors(val);
      sortedBackgrounds.push([colors.background]);
      sortedFontColors.push([colors.fontColor]);
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