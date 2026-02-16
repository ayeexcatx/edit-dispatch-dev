  const companyToFolderId = {
  "CCG": "1PX2ruBPu08zI6lamNXwSbYFtFiU9jyF2",
  "RTM": "12JzIRk5OcF48qTBLeyW-o2zJi2qGaJ_r",
  "EXCL": "1xl5OPhlWTubQfNBAT_JwW02js_TOyULs",
  "WAJA": "1oorhtvwsj93KSiZdJf0nGNdmyH6W6kN5",
  "WMBA": '1pJJx-A139FP1focE6RyGLZyZ1fvheo2t'
};

const truckToCompanyMap = {
    "DT02": "CCG",
    "RT03": "RTM",
    "RT12": "RTM",
    "EXCL1": "EXCL",
    "WAJA01": "WAJA",
    "WAJA03": "WAJA",
    'WMBA11': 'WMBA'
    };

const DEV_SPREADSHEET_ID = "1BRkmpO0PoYyDVfK5zskSN9UZjAV9cJpcUd76xXLQHss";

// === HANDLE AMEND OR CANCEL FORM SUBMISSIONS ===
function onAmendOrCancelFormSubmit(e) {
  Logger.log(JSON.stringify(e));

  const formResponse = e.response;
  if (!formResponse) {
    Logger.log("Unexpected form structure: response is undefined");
    return;
  }

  const itemResponses = formResponse.getItemResponses();
  const responseMap = {};

  itemResponses.forEach(response => {
    const title = response.getItem().getTitle();
    const answer = response.getResponse();
    responseMap[title] = answer;
  });

  Logger.log("Parsed responses: " + JSON.stringify(responseMap));

  // Extract values from the map
  const truckNumbersRaw = responseMap["Truck Number"] || "";
  const dispatchDate = responseMap["Dispatch Date"];
  const startTimeRaw = responseMap["Start Time"];
  const action = responseMap["Action"];
  const reason = responseMap["Reason"];
  const updatedStartTime = responseMap["Updated Start Time"];
  const updatedStartLocation = responseMap["Updated Start Location"];
  const updatedInstructions = responseMap["Updated Instructions"];
  const updatedStartTime2 = responseMap["Updated Start Time 02"];
  const updatedStartLocation2 = responseMap["Updated Start Location 02"];
  const updatedInstructions2 = responseMap["Updated Instructions 02"];

  const truckNumbers = Array.isArray(truckNumbersRaw)
  ? truckNumbersRaw.map(s => s.trim())
  : truckNumbersRaw.split(",").map(s => s.trim());

  Logger.log("Truck Numbers: " + JSON.stringify(truckNumbers));
  Logger.log("Action: " + action);
  Logger.log("Reason: " + reason);
  
  // Continue with the rest of your logic below...
  if (!dispatchDate || !startTimeRaw || !action || truckNumbers.length === 0) return;

  const truckNum = truckNumbers[0];  // assuming only one truck checked
  const companyName = truckToCompanyMap[truckNum];
  const folderId = companyToFolderId[companyName];

  const dispatchFolder = DriveApp.getFolderById("1Fic0PvyH2B-Dq7P0hYQLsn0jB09qOWLE"); // Main dispatch archive folder
  const devSpreadsheet = SpreadsheetApp.openById(DEV_SPREADSHEET_ID);
  const dispatchesSheet = devSpreadsheet.getSheetByName("Dispatches");
  if (!dispatchesSheet) {
    Logger.log('Dispatches sheet not found in DEV spreadsheet');
    return;
  }

  const dispatchesValues = dispatchesSheet.getDataRange().getValues();
  if (dispatchesValues.length < 2) {
    Logger.log('Dispatches sheet has no data rows');
    return;
  }

  const dispatchesHeaders = dispatchesValues[0];
  const dispatchesColumnIndex = {};
  dispatchesHeaders.forEach((header, index) => {
    dispatchesColumnIndex[String(header).trim()] = index;
  });

  const dispatchesByDocId = {};
  const docIdColumnIndex = dispatchesColumnIndex.doc_id;
  if (docIdColumnIndex === undefined) {
    Logger.log('Dispatches sheet is missing required doc_id header');
    return;
  }

  for (let row = 1; row < dispatchesValues.length; row++) {
    const docId = dispatchesValues[row][docIdColumnIndex];
    if (docId) {
      dispatchesByDocId[String(docId)] = row;
    }
  }

  let amendmentHistorySheet = devSpreadsheet.getSheetByName('AmendmentHistory');
  if (!amendmentHistorySheet) {
    amendmentHistorySheet = devSpreadsheet.insertSheet('AmendmentHistory');
  }
  if (amendmentHistorySheet.getLastRow() === 0) {
    amendmentHistorySheet.appendRow([
      'dispatch_id',
      'truck_number',
      'event_type',
      'event_at',
      'event_by',
      'reason',
      'old_doc_id',
      'new_doc_id'
    ]);
  }

  for (const truck of truckNumbers) {
  const companyName = truckToCompanyMap[truck];
  const folderId = companyToFolderId[companyName];
  if (!companyName || !folderId) continue;

  const companyFolder = DriveApp.getFolderById(folderId);
  const truckFolder = getOrCreateSubfolder(companyFolder, truck);
  const replacedFolder = getOrCreateSubfolder(truckFolder, "Replaced Dispatches");

  const startTime = startTimeRaw.padStart(4, "0");
  const filePrefix = `${dispatchDate}_${startTime}_Dispatch_${truck}_`;

  const files = truckFolder.getFiles();
  let targetFile = null;
  while (files.hasNext()) {
    const file = files.next();
    if (file.getName().startsWith(filePrefix)) {
      targetFile = file;
      break;
    }
  }

  if (!targetFile) continue;

    // Make a copy and edit it
    const parts = targetFile.getName().split("_");
    let amendedFile;
    if (parts.length >= 5) {
      // Change middle part to AMEND or CANCEL
      parts[2] = action.toUpperCase().includes("CANCEL") ? "CANCELLED" : "AMEND";

      // Extract original start time from filename (parts[1])
      let originalStartTime = parts[1]; // e.g. "0830"

      // Try to parse updatedStartTime if provided
      let newStartTimeStr = null;
      if (updatedStartTime && updatedStartTime !== "") {
      // Expect updatedStartTime in "HH:mm" or "H:mm" format from the form
      const timeParts = updatedStartTime.split(":");
      if (timeParts.length === 2) {
        const hours = timeParts[0].padStart(2, "0");
        const minutes = timeParts[1].padStart(2, "0");
        newStartTimeStr = hours + minutes;
      }
    }

      // Use updated start time if valid, otherwise keep original
      parts[1] = newStartTimeStr || originalStartTime;

      const newName = parts.join("_");
      amendedFile = targetFile.makeCopy(newName, truckFolder);
     } else {
     Logger.log("Unexpected filename structure: " + targetFile.getName());
      return;  // Stop the script if the filename is unexpected
    }

    const rowIndex = dispatchesByDocId[targetFile.getId()];
    if (rowIndex === undefined) {
      Logger.log('No Dispatches row found for doc_id: ' + targetFile.getId());
    } else {
      const updatedRow = dispatchesValues[rowIndex].slice();
      const now = new Date();
      const eventBy = Session.getActiveUser().getEmail() || '';
      const isCancelAction = action.toLowerCase().includes('cancel');

      updatedRow[dispatchesColumnIndex.status] = isCancelAction ? 'Canceled' : 'Amended';
      updatedRow[dispatchesColumnIndex.doc_id] = amendedFile.getId();
      updatedRow[dispatchesColumnIndex.doc_url] = amendedFile.getUrl();
      updatedRow[dispatchesColumnIndex.is_confirmed] = false;
      updatedRow[dispatchesColumnIndex.last_confirmed_at] = '';
      updatedRow[dispatchesColumnIndex.last_updated_at] = now;
      updatedRow[dispatchesColumnIndex.last_updated_by] = eventBy;

      if (isCancelAction) {
        updatedRow[dispatchesColumnIndex.cancel_reason] = reason;
        updatedRow[dispatchesColumnIndex.change_summary] = '';
      } else {
        updatedRow[dispatchesColumnIndex.change_summary] = reason;
        updatedRow[dispatchesColumnIndex.cancel_reason] = '';
      }

      dispatchesSheet
        .getRange(rowIndex + 1, 1, 1, updatedRow.length)
        .setValues([updatedRow]);

      amendmentHistorySheet.appendRow([
        updatedRow[dispatchesColumnIndex.dispatch_id],
        truck,
        isCancelAction ? 'CANCEL' : 'AMEND',
        now,
        eventBy,
        reason,
        targetFile.getId(),
        amendedFile.getId()
      ]);
    }

    amendedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // Anyone can view
    const doc = DocumentApp.openById(amendedFile.getId());
    const body = doc.getBody();

    // Replace CONFIRM DISPATCH header
    const isCancel = action.toLowerCase().includes("cancel");
    const header = isCancel ? "🚫 CANCELLED 🚫" : "❗AMENDMENT❗";

    body.replaceText("✅ CONFIRM DISPATCH ✅", header);

    // Find the index of the header so we can insert after it
    const paragraphs = body.getParagraphs();
    for (let i = 0; i < paragraphs.length; i++) {
      const text = paragraphs[i].getText();
      if (text.includes("CANCELLED") || text.includes("AMENDMENT")) {
        const headerParagraph = paragraphs[i];
        headerParagraph.setBold(true); // Make header bold

        const reasonParagraph = body.insertParagraph(i + 1, reason.trim());
        reasonParagraph.setItalic(true); // Make reason italic
        reasonParagraph.setForegroundColor("#5f6368");
        break;
      }
    }

    // Apply optional updates and re-bold labels for both assignments
    const labelPatterns = [
      "▪️Start Time:",
      "▪️Start Location:",
      "▪️Instructions:",
      "▪️Start Time 02:",
      "▪️Start Location 02:",
      "▪️Instructions 02:"
    ];

     paragraphs.forEach(p => {
      const text = p.editAsText();
      const line = text.getText();

      labelPatterns.forEach(label => {
        if (line.startsWith(label)) {
          text.setBold(0, label.length - 1, true);
        }
      });
    });

    // Helper function
    function updateDispatchContent(content, updates) {
  // Assignment 01
  if (updates.startTime) {
    content = content.replace(/(Start Time:\s*)(.*)/, `$1${updates.startTime}`);
  }
  if (updates.startLocation) {
    content = content.replace(/(Start Location:\s*)(.*)/, `$1${updates.startLocation}`);
  }
  if (updates.instructions) {
    content = content.replace(/(Instructions:\s*)(.*)/, `$1${updates.instructions}`);
  }

  // Assignment 02
  if (updates.startTime2) {
    content = content.replace(/(Start Time 02:\s*)(.*)/, `$1${updates.startTime2}`);
  }
  if (updates.startLocation2) {
    content = content.replace(/(Start Location 02:\s*)(.*)/, `$1${updates.startLocation2}`);
  }
  if (updates.instructions2) {
    content = content.replace(/(Instructions 02:\s*)(.*)/, `$1${updates.instructions2}`);
  }

  return content;
}

  // Apply optional updates to Start Time, Start Location, and Instructions
  // Replace values directly in-line while preserving formatting
const updates = {
  "Start Time:": updatedStartTime,
  "Start Location:": updatedStartLocation,
  "Instructions:": updatedInstructions,
  "Start Time 02:": updatedStartTime2,
  "Start Location 02:": updatedStartLocation2,
  "Instructions 02:": updatedInstructions2
};

for (const p of body.getParagraphs()) {
  const t = p.editAsText();
  const line = t.getText();

  for (const [label, newVal] of Object.entries(updates)) {
    if (!newVal) continue;

    const fullLabel = `▪️${label}`;
    const plainLabel = label;

    // Match either with bullet or without
    if (line.startsWith(fullLabel)) {
      const labelLength = fullLabel.length;
      t.deleteText(labelLength, t.getText().length - 1);
      t.insertText(labelLength, ` ${newVal}`);
    } else if (line.startsWith(plainLabel)) {
      const labelLength = plainLabel.length;

      // Replace line with bullet + label + new value
      t.setText(`▪️${plainLabel} ${newVal}`);
    }
  }
}

  
  // Re-bold the labels after content replacement
// Re-bold labels and italicize the reason (if needed) after text replacement
const updatedParagraphs = body.getParagraphs();
for (let i = 0; i < updatedParagraphs.length; i++) {
  const p = updatedParagraphs[i];
  const text = p.editAsText();
  const line = text.getText();

  // Bold all standard and secondary assignment labels
  // Bold all standard and secondary assignment labels (more reliable)
const labelPatterns = [
  "▪️Start Time:",
  "▪️Start Location:",
  "▪️Instructions:",
  "▪️Start Time 02:",
  "▪️Start Location 02:",
  "▪️Instructions 02:"
];

labelPatterns.forEach(label => {
  const startIndex = line.indexOf(label);
  if (startIndex !== -1) {
    try {
      text.setBold(startIndex, startIndex + label.length - 1, true);
    } catch (e) {
      Logger.log(`Error bolding label "${label}" in line: "${line}"`);
    }
  }
});


  // Make AMENDMENT or CANCELLED header bold
  if (line.includes("AMENDMENT") || line.includes("CANCELLED")) {
    text.setBold(true);
  }

  // Make the line right after the header italic (the reason)
  if (i > 0 && (updatedParagraphs[i - 1].getText().includes("AMENDMENT") || updatedParagraphs[i - 1].getText().includes("CANCELLED"))) {
    text.setItalic(true);
    text.setForegroundColor("#5f6368");
  }
}
  
    doc.saveAndClose();

    // Move original to Replaced Dispatches folder
    replacedFolder.addFile(targetFile);
    truckFolder.removeFile(targetFile);

    const loopCompanyName = truckToCompanyMap[truck];
    if (!loopCompanyName) {
      Logger.log("Unknown company for truck: " + truck);
      return;
    }

    const loopFolderId = companyToFolderId[loopCompanyName];
    if (!loopFolderId) {
      Logger.log("No folder ID found for company: " + loopCompanyName);
      return;
    }

    updateCompanyDispatchPage(loopCompanyName, loopFolderId);


  };
}

// === HELPER: Get or create subfolder by name ===
function getOrCreateSubfolder(parentFolder, subfolderName) {
  const folders = parentFolder.getFoldersByName(subfolderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(subfolderName);
}

function setupAmendmentTrigger() {
  var form = FormApp.openById("1CcmEMU_fHkdWVkpDUiIMmxbkOPDDcQk5NzqkrgQluxA");
  ScriptApp.newTrigger("onAmendOrCancelFormSubmit")
    .forForm(form)
    .onFormSubmit()
    .create();
}

// UPDATE HTML ------------------
function updateCompanyDispatchPage(companyName, folderId) {
  const folderMap = {
    "CCG": ["DT02"],
    "RTM": ["RT03", "RT12"],
    "EXCL": ["EXCL1"],
    "WAJA": ["WAJA01", "WAJA03"],
    "WMBA": ["WMBA11"]
  };

  console.log("Folder ID:", folderId);  // logging
  const companyFolder = DriveApp.getFolderById(companyToFolderId[companyName]);
  const truckFolders = folderMap[companyName];

  const dispatchLookup = getDispatchLookupByDocId();

  let allFiles = [];

  // Aggregate dispatch files from all truck folders within the company folder
  for (const truckNum of truckFolders) {
    const truckFolder = companyFolder.getFoldersByName(truckNum).next();
    const files = truckFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().endsWith(".html")) continue; // skip HTML pages
      allFiles.push(file);
    }
  }

  // Parse metadata and sort
  const dispatches = allFiles.map(file => {
    const name = file.getName();
    const dateMatch = name.match(/(\d{4}-\d{2}-\d{2})_(\d{4})/);
    const date = dateMatch ? new Date(`${dateMatch[1]}T${dateMatch[2].slice(0,2)}:${dateMatch[2].slice(2)}:00`) : new Date(0);
    return { file, name, date };
  }).sort((a, b) => a.date - b.date);

    function stripTime(date) {
      return new Date(date.getFullYear(), date.getMonth(), date.getDate());
    }

    const today = stripTime(new Date());

  const upcoming = [], todayList = [], past = [];

  for (const d of dispatches) {
    const url = d.file.getUrl();
    const parts = d.name.split('_');
    const [dispatchDateStr, timeStr, , truckNumber, jobNumberWithExt] = parts;
    const jobNumber = jobNumberWithExt.replace(/\..+$/, '');

    const dispatchDate = new Date(`${dispatchDateStr}T${timeStr.slice(0, 2)}:${timeStr.slice(2)}:00`);
    const friendlyDate = dispatchDate.toLocaleDateString('en-US', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit'
    });
    const friendlyTime = dispatchDate.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' });

  let labelContent = `${friendlyDate} @ ${friendlyTime} – ${truckNumber} – ${jobNumber}`;
  let statusLabel = "";
  let labelStyle = "";

  if (d.name.includes("CANCELLED")) {
    statusLabel = " <span style='color: #d93025; font-weight: bold;'>CANCELLED</span>";
    labelStyle = "color: grey; text-decoration: line-through;";
  } else if (d.name.includes("AMEND")) {
    statusLabel = " <span style='color: #FFBF00; font-weight: bold;'>AMENDMENT</span>";
  }

  const label = `<span style="${labelStyle}">${labelContent}</span>${statusLabel}`;
  const dispatchMeta = dispatchLookup[d.file.getId()] || {};
  const dispatchId = dispatchMeta.dispatch_id;
  const isConfirmed = dispatchMeta.is_confirmed;

  let confirmControl = "<span class='confirm-placeholder'>[NO DISPATCH ID]</span>";
  if (dispatchId !== undefined && dispatchId !== null && String(dispatchId).trim() !== "") {
    if (isConfirmed) {
      confirmControl = "<button class='confirm-btn confirmed' disabled>Confirmed ✓</button>";
    } else {
      confirmControl = `<button class='confirm-btn' data-dispatch-id='${String(dispatchId)}' data-truck-number='${truckNumber}' onclick='confirmReceipt(this)'>Confirm Receipt</button>`;
    }
  }

  const link = `<div class="dispatch-block"><a href="${url}">${label}</a>${confirmControl}</div>`;

    const entry = { date: d.date, html: link };

    const dispatchDay = stripTime(dispatchDate);
    
    if (dispatchDay.getTime() < today.getTime()) {
      past.push(entry);
    } else if (dispatchDay.getTime() === today.getTime()) {
      todayList.push(entry);
    } else {
      upcoming.push(entry);
    }
  }
  upcoming.sort((a, b) => b.date - a.date);
  todayList.sort((a, b) => b.date - a.date);
  past.sort((a, b) => b.date - a.date);

  const upcomingHTML = upcoming.map(e => e.html).join('\n');
  const todayHTML = todayList.map(e => e.html).join('\n');
  const pastHTML = past.map(e => e.html).join('\n');


  //Webpage layout
  const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>${companyName} Dispatches</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 40px; background: #f9f9f9; }
    h1 {
      font-size: 28px;
      text-align: center;
      padding: 10px 20px;
      display: inline-block;
      background-color: #FFD700; /* construction yellow */
      border-radius: 25px;
      box-shadow: 0 2px 5px rgba(0,0,0,0.2);
    }
    .section {
      margin-top: 40px;
      padding: 20px;
      border-radius: 10px;
      background-color: #ffffff;
      box-shadow: 0 0 8px rgba(0,0,0,0.05);
    }
    .upcoming { border: 3px solid #4CAF50; } 
    .today { border: 3px solid #2196F3; }
    .past { border: 3px solid #9E9E9E; }
    h2 {
      font-size: 20px;
      margin-top: 0;
    }
    .dispatch-block {
      padding: 10px 0;
      border-bottom: 1px solid #e0e0e0;
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
      flex-wrap: wrap;
    }
    .dispatch-block a {
      text-decoration: none;
      color: #333;
    }
    .dispatch-block a:hover {
      text-decoration: underline;
    }
    .dispatch-block a {
      color: #1a73e8; /* Google Blue */
      text-decoration: underline;
      font-weight: bold; /*Make links bold */
    }
    .confirm-btn {
      border: none;
      border-radius: 4px;
      padding: 8px 12px;
      font-size: 14px;
      font-weight: bold;
      cursor: pointer;
      background-color: #1a73e8;
      color: #fff;
    }
    .confirm-btn.confirmed,
    .confirm-btn:disabled {
      background-color: #d9ead3;
      color: #2e7d32;
      cursor: default;
    }
    .confirm-placeholder {
      color: #d93025;
      font-size: 13px;
      font-weight: bold;
    }
    .title-container {
      text-align: center;
      margin-bottom: 24px;
    }
    .dispatch-title {
      display: inline-block;
      background-color: #FFD700; /* Construction yellow */
      padding: 12px 24px;
      border-radius: 999px; /* Pill/bubble shape */
      font-size: 2em;
      font-weight: bold;
      color: #000;
      box-shadow: 0 2px 6px rgba(0,0,0,0.15);
    }
  </style>
</head>
<body>
    <div class="title-container">
      <h1 class="dispatch-title">${companyName} Dispatches</h1>
    </div>

  <div class="section upcoming">
    <h2>Upcoming</h2>
    ${upcomingHTML || '<p>No upcoming dispatches.</p>'}
  </div>

  <div class="section today">
    <h2>Today</h2>
    ${todayHTML || '<p>No dispatches for today.</p>'}
  </div>

  <div class="section past">
    <h2>Past</h2>
    ${pastHTML || '<p>No past dispatches.</p>'}
  </div>

  <script>
    function confirmReceipt(buttonEl) {
      const dispatchId = buttonEl.getAttribute('data-dispatch-id');
      const truckNumber = buttonEl.getAttribute('data-truck-number');

      if (!dispatchId) {
        alert('Missing dispatch ID.');
        return;
      }

      buttonEl.disabled = true;

      google.script.run
        .withSuccessHandler(function() {
          buttonEl.textContent = 'Confirmed ✓';
          buttonEl.classList.add('confirmed');
        })
        .withFailureHandler(function(error) {
          buttonEl.disabled = false;
          alert('Unable to confirm receipt: ' + (error && error.message ? error.message : error));
        })
        .confirmDispatchReceipt(dispatchId, truckNumber);
    }
  </script>
</body>
</html>`;

  const fileName = `${companyName.replace(/\s+/g, '_')}_dispatch_list.html`;
  const existing = companyFolder.getFilesByName(fileName);
  let htmlFile;

  if (existing.hasNext()) {
    htmlFile = existing.next();
    htmlFile.setContent(html);
  } else {
    htmlFile = companyFolder.createFile(fileName, html, MimeType.HTML);
  }

  htmlFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
}

function getDispatchLookupByDocId() {
  const devSpreadsheet = SpreadsheetApp.openById(DEV_SPREADSHEET_ID);
  const dispatchesSheet = devSpreadsheet.getSheetByName('Dispatches');
  if (!dispatchesSheet) return {};

  const values = dispatchesSheet.getDataRange().getValues();
  if (values.length < 2) return {};

  const headers = values[0].map(header => String(header).trim());
  const columnIndex = {};
  headers.forEach((header, index) => {
    columnIndex[header] = index;
  });

  const docIdIndex = columnIndex.doc_id;
  const dispatchIdIndex = columnIndex.dispatch_id;
  const isConfirmedIndex = columnIndex.is_confirmed;
  if (docIdIndex === undefined || dispatchIdIndex === undefined || isConfirmedIndex === undefined) {
    return {};
  }

  const lookup = {};
  for (let row = 1; row < values.length; row++) {
    const docId = values[row][docIdIndex];
    if (!docId) continue;

    const rawConfirmedValue = values[row][isConfirmedIndex];
    const normalizedConfirmed =
      rawConfirmedValue === true ||
      String(rawConfirmedValue).toLowerCase() === 'true';

    lookup[String(docId)] = {
      dispatch_id: values[row][dispatchIdIndex],
      is_confirmed: normalizedConfirmed
    };
  }

  return lookup;
}
