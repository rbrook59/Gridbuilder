# Gridbuilder

Gridbuilder is a Google Apps Script library that automates dinner party seating across multiple GridBuilder sheets. The script and connections history live in a single central spreadsheet, shared by all GridBuilder sheets.

---

## Setting Up a New GridBuilder Sheet

Follow these steps each time you create or duplicate a GridBuilder sheet.

### Step 1 — Open the Apps Script Editor

1. Open the GridBuilder Google Sheet
2. Click **Extensions → Apps Script**

---

### Step 2 — Add the Gridbuilder Library

1. In the left sidebar, click the **+** next to **Libraries**
2. In the **Script ID** field, paste:
   ```
   1CMa1P0ItEp0Se97EvdiFAI8JsoDMP3HUYvs4Y8WMO4mVBRUTZS8iwY4S
   ```
3. Click **Look up**
4. Set the **Identifier** field to exactly:
   ```
   Gridbuilder
   ```
5. Set **Version** to **HEAD (development mode)**
6. Click **Add**

---

### Step 3 — Paste the Stub Script

1. In the left sidebar, click **Code.gs**
2. Delete all existing content
3. Paste the following code:

```javascript
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
```

4. Press **Ctrl+S** to save

---

### Step 4 — Test the Setup

1. Close the Apps Script editor and reload the Google Sheet
2. The **Grid Builder** menu should appear in the menu bar
3. Run **Prepare Connections Data** to verify the library is connected correctly

---

## Deploying Code Changes

Code is managed locally via [clasp](https://github.com/google/clasp) and pushed to the central spreadsheet.

```
clasp push
```

- Answer **y** if prompted about the manifest file
- Changes take effect immediately in all sheets using HEAD mode — no redeployment needed

---

## Central Spreadsheet

| Item | ID |
|---|---|
| Spreadsheet (connectionsDB) | `1Saq0rXFWj2IHfoYDwyjperNOXZuanYSbu57AGvB5dPk` |
| Script (Library) | `1CMa1P0ItEp0Se97EvdiFAI8JsoDMP3HUYvs4Y8WMO4mVBRUTZS8iwY4S` |
