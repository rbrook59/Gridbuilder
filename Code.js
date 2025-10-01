// FUTURE DEVELOPMENT NOTES
// 1. create a audit of a Grid based on pairs of members and # of months apart
// 2. create an orphans list to give options of where unseated members can sit.
// 3.
//
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
  var i = 0;
  var j = 0;
  var gRow = 0; // guest row in arrGuest
  var gCount = 0; // guest count in a house
  var seatedCount = 0; // number of seats taken in a house
  var countGuests = 0;
  //
  /** Load ranges from sheet */
  //
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var gridSheet = spreadSheet.getSheetByName('GridBuilder');
  //
  /** retrieve the range of control values that are used by the program  */
  //
  var arrControl = spreadSheet.getRange('controlVariables').getValues().slice();
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
  var arrHosts = gridSheet.getRange('rangeHosts').getValues().slice();
  var headerHosts = arrHosts[0]; // save the header row to add back later
  arrHosts.splice(0, 1); //remove the header row for easier sorting
  //
  var arrGuests = gridSheet.getRange('rangeGuests').getValues().slice();
  var headerGuests = arrGuests[0]; // save the header row to add back later
  arrGuests.splice(0, 1); //remove the header row for easier sorting
  //  
  var arrGrid = gridSheet.getRange('rangeGrid').getValues().slice();
  var headerGrid = arrGrid[0]; // save the header row to add back later
  arrGrid.splice(0, 1); //remove the header row for easier sorting
  //
  var arrNeverMatch = gridSheet.getRange('rangeNeverMatch').getValues().slice();
  arrNeverMatch.splice(0, 1); //remove the header row for easier processing
  //
  var dbSheet = spreadSheet.getSheetByName('connectionsDB');
  var arrDB = dbSheet.getDataRange().getValues().slice();
  var headerConnections = arrDB[0]; // save the header row to add back later
  arrDB.splice(0, 1); //remove the header row

  //
  /**Count the number of connections for each Host and for each Guest and sort the list, descending order.
   * This algorithm prioritizes seating members who have attended the most dinners
   * since they will be the most difficult to seat with fellow members.*/
  //
  // Hosts Count of Previous Connections
  if (ctrlSortHosts === 1) {
    for (i = 0; i < arrHosts.length; i++) {
      if (arrHosts[i][hostCode].length > 0) {
        var count = arrDB.filter(x => {
          return x[1].toString() === arrHosts[i][hostCode].toString();
        }).length
        arrHosts[i].push(count);
      }
    }
    // Sort the array of Hosts  
    arrHosts.sort(function (a, b) {
      return b[3] - a[3]
    }); //sort list of Hosts from most connections to least connections
  }
  // Compress the host array to remove blank rows
  for (i = arrHosts.length - 1; i >= 0; i--) {
    if (arrHosts[i][hostCode] < 1) {
      arrHosts.splice(i, 1);
    }
  }
  //
  //Guest Count of Previous Connections
  if (ctrlSortGuests === 1) {
    for (i = 0; i < arrGuests.length; i++) {
      if (arrGuests[i][guestCode].length > 0) {
        var count = arrDB.filter(x => {
          return x[1].toString() === arrGuests[i][guestCode].toString();
        }).length
        arrGuests[i].push(count);
      }
      // Sort the array of Guests    
      arrGuests.sort(function (a, b) {
        return b[3] - a[3]
      }); //sort list of Guests in descending order 
    }
  }
  // Count how many guests to process to ignore blank rows
  for (i = 0; i < arrGuests.length; i++) {
    if (arrGuests[i][guestCode].length > 0) {
      countGuests++
    }
  }
  //
  // Compress the arrNeverMatch to remove blank rows
  for (i = arrNeverMatch.length - 1; i >= 0; i--) {
    if (arrNeverMatch[i] < 1) {
      arrNeverMatch.splice(i, 1);
    }
  }
  // Duplicate each entry in the arrNeverMatch in the reverse order
  // Use existing member1-member2 entry and create additional element of member2-member1
  //
  for (i = arrNeverMatch.length - 1; i >= 0; i--) {
    let arrTemp = [];
    arrTemp = arrNeverMatch[i].toString().split(/-/);
    arrNeverMatch.push([arrTemp[1], arrTemp[0]].join('-'));
  }
  /** check control flags, clear the seated flag and the Grid  */
  //
  if (ctrlClearSeated === 1) {
    for (i = 0; i < countGuests; i++) {
      arrGuests[i][guestSeated] = "No"
    }
  }
  // Clear the Grid
  if (ctrlClearGrid === 1) {
    for (i = 0; i < arrGrid.length; i++) {
      for (j = 0; j < 9; j++) {
        arrGrid[i][j] = null;
      }
    }
  }
  //
  /**  CREATE THE GRID */
  //
  for (i = 0; i < arrHosts.length; i++) { // This many houses to process
    //
    if (arrGrid[i][gridHouse] === null) { // initialize the House on first loop
      arrGrid[i][gridHouse] = i + 1; // array starts at 0, so add 1
      arrGrid[i][gridSeats] = arrHosts[i][hostSeats]; // How many can be seated at this house
      arrGrid[i][gridHost] = arrHosts[i][hostCode]; // Host name code
      seatedCount = Number(arrHosts[i][hostCount]); // this many seats for the host (will be 1 or 2)
      //
      var arrHouse = []; // 1D empty array to track guests assigned
      arrHouse.push(arrGrid[i][gridHost]); // add the host code to the House array
      //
    } else { // the Grid already has data stored.  Load that data into the variables used for processing
      var arrHouse = []; // 1D empty array to track guests assigned
      for (let n = 3; n < arrGrid[i].length; n++) {
        if (arrGrid[i][n] != null) {
          arrHouse.push(arrGrid[i][n]);
        }
      }
      seatedCount = arrGrid[i][gridSeated];
    }
    //
    // Fill the house - with up to 5 guests or until seats are full
    //
    for (gCount = arrHouse.length; gCount < 5; gCount++) {
      if (seatedCount >= arrHosts[i][hostSeats]) { // House is full, go to the next house
        break;
      }
      //      
      //  The array "doesItWork" tests each potential member in the house.  Any "false" values and the guest is rejected.
      //
      //
      for (gRow = 0; gRow < countGuests; gRow++) { // cycle through each member
        if (ctrlThrottleSingles === 1 && gCount < 2 && arrGuests[gRow][guestCount] < 2) { // wait to seat singles
          continue;
        }
        if (arrGuests[gRow][guestSeated] === "No" && seatedCount + arrGuests[gRow][guestCount] <= arrHosts[i][hostSeats]) {  // Does this guest need a seat and can they fit in this house
          //
          var doesItWork = []; // array that gets loaded with True or False values.   Any False value and the guest is rejected
          //
          switch (gCount) {
            case 1:
              doesItWork.push(memberMatch([arrGuests[gRow][guestCode], arrHouse[0]].join('-'), ctrlTimeLapse, arrDB, arrNeverMatch));
              break;
            case 2:
              doesItWork.push(memberMatch([arrGuests[gRow][guestCode], arrHouse[0]].join('-'), ctrlTimeLapse, arrDB, arrNeverMatch));
              doesItWork.push(memberMatch([arrGuests[gRow][guestCode], arrHouse[1]].join('-'), ctrlTimeLapse, arrDB, arrNeverMatch));
              break;
            case 3:
              doesItWork.push(memberMatch([arrGuests[gRow][guestCode], arrHouse[0]].join('-'), ctrlTimeLapse, arrDB, arrNeverMatch));
              doesItWork.push(memberMatch([arrGuests[gRow][guestCode], arrHouse[1]].join('-'), ctrlTimeLapse, arrDB, arrNeverMatch));
              doesItWork.push(memberMatch([arrGuests[gRow][guestCode], arrHouse[2]].join('-'), ctrlTimeLapse, arrDB, arrNeverMatch));
              break;
            case 4:
              doesItWork.push(memberMatch([arrGuests[gRow][guestCode], arrHouse[0]].join('-'), ctrlTimeLapse, arrDB, arrNeverMatch));
              doesItWork.push(memberMatch([arrGuests[gRow][guestCode], arrHouse[1]].join('-'), ctrlTimeLapse, arrDB, arrNeverMatch));
              doesItWork.push(memberMatch([arrGuests[gRow][guestCode], arrHouse[2]].join('-'), ctrlTimeLapse, arrDB, arrNeverMatch));
              doesItWork.push(memberMatch([arrGuests[gRow][guestCode], arrHouse[3]].join('-'), ctrlTimeLapse, arrDB, arrNeverMatch));
              break;
            case 5:
              doesItWork.push(memberMatch([arrGuests[gRow][guestCode], arrHouse[0]].join('-'), ctrlTimeLapse, arrDB, arrNeverMatch));
              doesItWork.push(memberMatch([arrGuests[gRow][guestCode], arrHouse[1]].join('-'), ctrlTimeLapse, arrDB, arrNeverMatch));
              doesItWork.push(memberMatch([arrGuests[gRow][guestCode], arrHouse[2]].join('-'), ctrlTimeLapse, arrDB, arrNeverMatch));
              doesItWork.push(memberMatch([arrGuests[gRow][guestCode], arrHouse[3]].join('-'), ctrlTimeLapse, arrDB, arrNeverMatch));
              doesItWork.push(memberMatch([arrGuests[gRow][guestCode], arrHouse[4]].join('-'), ctrlTimeLapse, arrDB, arrNeverMatch));
              break;
          } // switch
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
    for (i = 0; i < arrGuests.length; i++) {
      if (arrGuests[i][guestCode].length > 0) {
        arrGuests[i].pop();
      }
    }
  }
  // write the arrGuests back to the sheet
  arrGuests.splice(0, 0, headerGuests);
  gridSheet.getRange('rangeGuests').setValues(arrGuests);
}
function memberMatch(strToCheck, timeLapse, arrToScan, arrNever) {
  canMatch = true;
  //
  // scan the connections list for a match and check the timeLapse
  //
  for (x = 0; x < arrToScan.length; x++) {
    if (strToCheck == arrToScan[x][0] && timeLapse > arrToScan[x][6]) {
      canMatch = false
      break;
    }
  }
  //
  // Scan the exceptions list for a match.
  //
  for (x = 0; x < arrNever.length; x++) {
    if (strToCheck == arrNever[x]) {
      canMatch = false;
      break
    }
  }
  return canMatch;
}
function prepConnections() {
  //
  // recalculates the number of months between members meeting and returns the sorted array to the calling function
  //
  let result = SpreadsheetApp.getUi().alert("Run this once before building the Grid.Check that the dinner date is correct on the spreadsheet.  Click OK to run this script, Cancel to stop", SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  if (result != SpreadsheetApp.getUi().Button.OK) {
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
  for (i = 0; i < arrDB.length; i++) {
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
}
function updateConnections() {
  //
  // takes the final grid for a dinner and creates new entries in the connectionsDB sheet for each pair of
  // members that sat together at a dinner.  This updates the history of the members that have been together.
  //
  let result = SpreadsheetApp.getUi().alert("Run this after the dinner date.  Update the Grid with any last minute changes and check that the dinner date is correct on the spreadsheet.  Click OK to run this script, Cancel to stop", SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  if (result != SpreadsheetApp.getUi().Button.OK) {
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
  var arrDB = dbSheet.getDataRange().getValues().slice();
  let headerRow = arrDB[0]; // save the header row to add back later
  arrDB.splice(0, 1); //remove the header row
  //
  // Get the Grid from the worksheet
  //
  var gridSheet = spreadSheet.getSheetByName('GridBuilder');
  var arrGrid = gridSheet.getRange('rangeGrid').getValues().slice();
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
  for (i = arrGrid.length - 1; i >= 0; i--) {
    if (arrGrid[i][3] < 1) { // blank host code in Grid.  Remove the row.
      arrGrid.splice(i, 1);
    }
  }
  var houseCount = arrGrid.length;
  //
  for (i = 0; i < houseCount; i++) {
    //
    // count the number of members to process in the house
    //
    var guestCount = 0;
    for (j = 3; j < arrGrid[i].length; j++) {  // 3 is the array position for the host
      if (arrGrid[i][j].length > 0) {
        guestCount++;
      }
    }
    for (j = 1; j < guestCount; j++) {
      switch (j) {
        case 1:
          for (k = 1; k < guestCount; k++) {
            writeConnection(arrGrid[i][3], arrGrid[i][k + 3], ctrlNextDinnerDate, "Host", arrDB);
          }
          break;
        case 2:
          for (k = 2; k < guestCount; k++) {
            writeConnection(arrGrid[i][4], arrGrid[i][k + 3], ctrlNextDinnerDate, "", arrDB);
          }
          break;
        case 3:
          for (k = 3; k < guestCount; k++) {
            writeConnection(arrGrid[i][5], arrGrid[i][k + 3], ctrlNextDinnerDate, "", arrDB);
          }
          break;
        case 4:
          for (k = 4; k < guestCount; k++) {
            writeConnection(arrGrid[i][6], arrGrid[i][k + 3], ctrlNextDinnerDate, "", arrDB);
          }
          break;
        case 5:
          for (k = 5; k < guestCount; k++) {
            writeConnection(arrGrid[i][7], arrGrid[i][k + 3], ctrlNextDinnerDate, "", arrDB);
          }
          break;
      }
    }
  }
  for (i = 0; i < arrDB.length; i++) {
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
}
//
/** SUPPORTING FUNCTIONS */
//
//
// subroutine to write new connections to connectionsDB
//
function writeConnection(memberOne, memberTwo, dinnerDate, memberRole, arrTmp) {
  let tmpYear = dinnerDate.getFullYear()
  arrTmp.push([[memberOne, memberTwo].join('-'), memberOne, memberTwo, tmpYear, dinnerDate, memberRole, 0]);
  //
  // create a duplicate entry with the key flipped. member1-member2 and member2-member1
  //
  arrTmp.push([[memberTwo, memberOne].join('-'), memberOne, memberTwo, tmpYear, dinnerDate, memberRole, 0]);
  return;
}
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Grid Builder')
    .addItem('Build Grid', 'buildGrid')
    .addSeparator()
    .addItem('Prepare Connections Data', 'prepConnections')
    .addSeparator()
    .addItem('Post dinner -> Update Connections History', 'updateConnections')
    .addToUi();
}