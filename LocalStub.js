/**
 * LOCAL STUB SCRIPT
 * Paste this into the Apps Script editor of each GridBuilder sheet.
 * After pasting, add the Gridbuilder library under Libraries (use the script ID
 * from the central spreadsheet) and set its identifier to: Gridbuilder
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Grid Builder')
    .addItem('Build Grid', 'buildGrid')
    .addSeparator()
    .addItem('Audit Current Grid', 'auditGrid')
    .addSeparator()
    .addItem('Prepare Connections Data', 'prepConnections')
    .addSeparator()
    .addItem('Post dinner -> Update Connections History', 'updateConnections')
    .addToUi();
}

function buildGrid()        { Gridbuilder.buildGrid(); }
function auditGrid()        { Gridbuilder.auditGrid(); }
function prepConnections()  { Gridbuilder.prepConnections(); }
function updateConnections(){ Gridbuilder.updateConnections(); }
