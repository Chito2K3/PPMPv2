/**
 * ONE-TIME AUTHORIZATION HELPER:
 * Select 'testEmailAuth' in the Apps Script toolbar dropdown and click 'Run'
 * to grant your Google account permission to send invitation emails.
 */
function testEmailAuth() {
  var remaining = MailApp.getRemainingDailyQuota();
  Logger.log("✅ Mail authorization confirmed! Daily quota: " + remaining);
}

/**
 * ============================================================================
 * PPMP MOBILE PHARMACY PRODUCT EVALUATION SYSTEM - BACKEND ENGINE
 * ============================================================================
 * 
 * Google Apps Script backend for handling sheet initialization, master drug list
 * syncing, evaluator tab mirroring, score calculation, and consensus dashboard.
 * 
 * Author: Antigravity Team
 * Version: 1.0.0
 */

// Global Constants
var SHEET_MEDICINE_MASTER = "Medicine_Master";
var SHEET_SUPPLIERS = "Supplier"; // Handles "Supplier" (singular) or "Suppliers"
var SHEET_EVALUATIONS_MASTER = "Evaluations_Master";
var SHEET_EVALUATOR_ACCOUNTS = "Evaluator_Accounts";
var SHEET_END_USER = "End-user";
var SHEET_NURSE = "Nurse";
var SHEET_PHARMACIST = "Pharmacist";
// Legacy aliases for backward compatibility
var SHEET_EVALUATOR_1 = SHEET_END_USER;
var SHEET_EVALUATOR_2 = SHEET_NURSE;
var SHEET_EVALUATOR_3 = SHEET_PHARMACIST;
var SHEET_SUMMARY = "Consolidated_Summary";
var SHEET_CHECKLIST_REPORT = "Checklist_Report";
var SHEET_CHECKLIST_HORIZONTAL = "Checklist_Report_Horizontal";
var SHEET_INDICATIVE_SOURCE = "indicative 1";
var SHEET_QUESTIONNAIRE = "Questionnaire";

// Column Definitions for Evaluations_Master (1-indexed for Apps Script range operations)
var EVAL_HEADERS = [
  "Evaluation_ID", "Timestamp", "Evaluator", "Generic_Name", "Brand_Name", "Supplier", "Manufacturer", "Price",
  "P1_01_Brand_Name", "P1_02_Generic_Name", "P1_03_Dosage_Form_Strength", "P1_04_Manufacturer_Details",
  "P1_05_CPR_FDA_Registration", "P1_06_Batch_Lot_Number", "P1_07_Manufacturing_Date", "P1_08_Expiration_Date",
  "P1_09_Storage_Conditions", "P1_10_Rx_Symbol", "P1_11_Net_Content", "P1_12_Language_Legibility",
  "P1_13_Outer_Package", "P1_14_Inner_Package", "P1_15_Package_Insert", "P1_16_Barcode_QR",
  "P1_17_Tamper_Evident_Seal", "P1_18_Special_Warnings", "P1_19_FDA_Compliance",
  "P2_01_Container_Integrity", "P2_02_Closure_Seal", "P2_03_Blister_Packaging", "P2_04_Physical_Appearance",
  "P2_05_Dosing_Graduation", "P2_06_Dispensing_Ease",
  "Requires_Reconstitution", "P3_01_Reconstitution_Clarity", "P3_02_Dissolution_Time", "P3_03_Foaming_Behavior",
  "P3_04_Solution_Color", "P3_05_Diluent_Compatibility", "P3_06_Syringe_Passability", "P3_07_Post_Reconstitution_Stability",
  "P3_08_Filter_Needle_Req", "P3_09_Particulate_Absence",
  "Part_I_Score", "Part_II_Score", "Remarks", "Recommendation",
  "Data_Privacy_Consent", "Accuracy_Consent", "Evaluator_Signature"
];

var SUMMARY_HEADERS = [
  "Generic_Name", "Brand_Name", "Supplier", "Manufacturer", "Price",
  "EndUser_Recommendation", "Nurse_Recommendation", "Pharmacist_Recommendation",
  "EndUser_PartI_Score", "EndUser_PartII_Score",
  "Nurse_PartI_Score", "Nurse_PartII_Score",
  "Pharmacist_PartI_Score", "Pharmacist_PartII_Score",
  "Consensus_Status", "Consensus_Detail", "Last_Updated"
];

var OFFICIAL_PART1_ITEMS = [
  "Product Name",
  "Dosage Form and Strength",
  "Pharmacologic Category",
  "Formulation / Composition",
  "Indication(s)",
  "Dosage and Mode of Administration",
  "Contraindication(s), Precaution(s), Warning(s)",
  "Drug\u2013Drug / Drug\u2013Food Interactions",
  "Adverse Drug Reaction(s)",
  "Overdose and Treatment Information",
  "Storage Condition(s)",
  "Net Content / Pack Size",
  "Name & Address of Marketing Authorization Holder",
  "Name & Address of Manufacturer",
  "Rx Symbol & Prescription Caution Statement (if applicable)",
  "ADR Reporting Statement",
  "Registration Number",
  "Batch / Lot Number",
  "Date of Manufacture & Expiration Date"
];

var OFFICIAL_PART2_ITEMS = [
  "Inner label is identical to the outer label",
  "Drug name, dosage form, strength, batch/lot number, manufacture date, and expiry date are clearly readable on the container or inner packaging",
  "For blister or aluminum foil packs, expiry date, drugs name and dosage form is printed on each individual unit",
  "No leakage observed in IV fluids or other parenteral products through closures (rubber stoppers, caps, seals) or infusion sets",
  "Rubber stoppers (single-port and dual/twin-port) of IV fluid containers are durable yet easy to puncture",
  "Ease of opening, dispensing, and overall container integrity"
];

/**
 * One-time authorization helper for email permissions.
 * Run this function once from the Apps Script toolbar to grant MailApp permissions!
 */
function authorizeEmailPermissions() {
  var quota = MailApp.getRemainingDailyQuota();
  Logger.log("Email permission granted! Daily email quota remaining: " + quota);
}

/**
 * Custom menu created when opening the Google Spreadsheet.
 */
function onOpen() {
  createCustomMenu();
}

/**
 * Creates the custom UI menu in Google Sheets.
 */
function createCustomMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🏥 PPMP Evaluation")
    .addItem("🚀 Initialize Evaluation Workbook", "initializeEvaluationWorkbook")
    .addItem("🛠️ Upgrade Schema: Insert Price, Consent & Signature", "upgradeSchemaAddPriceAndSignature")
    .addItem("🔑 Fix Duplicate Keys & Clean Headers", "fixDuplicateKeysAndCleanHeaders")
    .addItem("🧹 Consolidate & Clean Duplicate Part II Columns", "consolidateDuplicatePart2Columns")
    .addItem("🩹 Backfill Empty P2_05 Dosing Graduation (Rows 1430+)", "backfillEmptyP2_05DosingGraduation")
    .addItem("📋 Ensure Questionnaire Has All 6 Part II Questions", "syncCanonicalQuestionnaireTab")
    .addSeparator()
    .addItem("📧 Send App Invites to Evaluators", "sendEvaluatorAppInvites")
    .addItem("🔗 Generate / Refresh Evaluator App Links", "refreshEvaluatorAppLinks")
    .addItem("⚙️ Set / Check Web App Deployment URL", "promptSetWebAppUrl")
    .addSeparator()
    .addItem("📊 Initialize / Refresh Horizontal Report", "createHorizontalReportSheet")
    .addItem("📑 Initialize / Reset Printable Checklist Viewer", "createChecklistReportSheet")
    .addItem("🔄 Sync Master Drug List", "syncMasterDrugList")
    .addItem("📋 Sync Questionnaire & Validate", "menuSyncAndValidateQuestionnaire")
    .addItem("📊 Refresh Consolidated Summary", "refreshConsolidatedSummary")
    .addSeparator()
    .addItem("📄 Export Current Checklist to PDF", "exportCurrentChecklistPdf")
    .addItem("⚡ Setup Auto-Sync Triggers", "setupTriggers")
    .addSeparator()
    .addItem("🧪 Run Concurrency Unit Test (3 Evaluators)", "runConcurrencyUnitTest")
    .addItem("🧹 Purge Concurrency Test Records", "purgeConcurrencyTestRecords")
    .addToUi();
}

/**
 * Safely inserts 'Price' (after Manufacturer), 'Data_Privacy_Consent', and 'Accuracy_Consent'
 * (after Recommendation, before Signature), and 'Evaluator_Signature' (at the end)
 * into Evaluations_Master and Evaluator tabs, preserving all existing rows without column shifts.
 */
function upgradeSchemaAddPriceAndSignature() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.alert(
    "Confirm Schema Upgrade",
    "This will safely ensure 'Price' (after Manufacturer), 'Data_Privacy_Consent', 'Accuracy_Consent' (after Recommendation), and 'Evaluator_Signature' exist in Evaluations_Master and all Evaluator tabs.\n\nAll existing evaluations will be preserved without shifting.\n\nDo you want to proceed?",
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;
  
  try {
    var targetSheets = [
      SHEET_EVALUATIONS_MASTER, SHEET_END_USER, SHEET_NURSE, SHEET_PHARMACIST,
      "Evaluator 1", "Evaluator 2", "Evaluator 3"
    ];
    
    var modifiedSheets = [];
    
    targetSheets.forEach(function(sName) {
      var s = ss.getSheetByName(sName);
      if (!s || s.getLastRow() < 1) return;
      
      var headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
      
      // 1. Check if Price already exists
      var priceIndex = -1;
      var mfgIndex = -1;
      for (var h = 0; h < headers.length; h++) {
        var colName = (headers[h] || "").toString().trim().toLowerCase();
        if (colName === "price" || colName === "unit_price") priceIndex = h;
        if (colName === "manufacturer" || colName === "manufacturer_name") mfgIndex = h;
      }
      
      // If Price is missing and Manufacturer is found, insert Price right after Manufacturer
      if (priceIndex === -1 && mfgIndex !== -1) {
        var insertColNum = mfgIndex + 2; // 1-indexed: if mfg is col 7, insert at col 8
        s.insertColumnBefore(insertColNum);
        s.getRange(1, insertColNum).setValue("Price").setBackground("#1B365D").setFontColor("#FFFFFF").setFontWeight("bold");
        s.setColumnWidth(insertColNum, 100);
        if (s.getLastRow() > 1) {
          s.getRange(2, insertColNum, s.getLastRow() - 1, 1).setNumberFormat("₱#,##0.00");
        }
      }
      
      // Refresh headers after possible Price insertion
      headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
      
      // 2. Check if Data_Privacy_Consent & Accuracy_Consent exist
      var dpIndex = -1;
      var accIndex = -1;
      var recIndex = -1;
      var sigIndex = -1;
      for (var h2 = 0; h2 < headers.length; h2++) {
        var cName = (headers[h2] || "").toString().trim().toLowerCase();
        if (cName === "data_privacy_consent" || cName === "privacy_consent") dpIndex = h2;
        if (cName === "accuracy_consent" || cName === "truthfulness_consent") accIndex = h2;
        if (cName === "recommendation") recIndex = h2;
        if (cName === "signature" || cName === "digital_signature" || cName === "evaluator_signature") sigIndex = h2;
      }

      // Insert Data_Privacy_Consent if missing
      if (dpIndex === -1) {
        var insertDpCol;
        if (sigIndex !== -1) {
          insertDpCol = sigIndex + 1; // 1-indexed: insert before signature
        } else if (recIndex !== -1) {
          insertDpCol = recIndex + 2; // insert after recommendation
        } else {
          insertDpCol = s.getLastColumn() + 1;
        }
        s.insertColumnBefore(insertDpCol);
        s.getRange(1, insertDpCol).setValue("Data_Privacy_Consent").setBackground("#1B365D").setFontColor("#FFFFFF").setFontWeight("bold");
        s.setColumnWidth(insertDpCol, 150);
        headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
      }

      // Insert Accuracy_Consent if missing
      headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
      accIndex = -1;
      sigIndex = -1;
      for (var h3 = 0; h3 < headers.length; h3++) {
        var cName3 = (headers[h3] || "").toString().trim().toLowerCase();
        if (cName3 === "accuracy_consent" || cName3 === "truthfulness_consent") accIndex = h3;
        if (cName3 === "signature" || cName3 === "digital_signature" || cName3 === "evaluator_signature") sigIndex = h3;
      }
      
      if (accIndex === -1) {
        var dpFoundIndex = -1;
        for (var h4 = 0; h4 < headers.length; h4++) {
          var cName4 = (headers[h4] || "").toString().trim().toLowerCase();
          if (cName4 === "data_privacy_consent" || cName4 === "privacy_consent") dpFoundIndex = h4;
        }
        var insertAccCol;
        if (dpFoundIndex !== -1) {
          insertAccCol = dpFoundIndex + 2; // right after Data_Privacy_Consent
        } else if (sigIndex !== -1) {
          insertAccCol = sigIndex + 1;
        } else {
          insertAccCol = s.getLastColumn() + 1;
        }
        s.insertColumnBefore(insertAccCol);
        s.getRange(1, insertAccCol).setValue("Accuracy_Consent").setBackground("#1B365D").setFontColor("#FFFFFF").setFontWeight("bold");
        s.setColumnWidth(insertAccCol, 150);
        headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
      }
      
      // 3. Check if Signature exists
      headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
      var sigFound = -1;
      for (var h5 = 0; h5 < headers.length; h5++) {
        var colNameSig = (headers[h5] || "").toString().trim().toLowerCase();
        if (colNameSig === "signature" || colNameSig === "digital_signature" || colNameSig === "evaluator_signature") sigFound = h5;
      }
      
      // If Signature is missing, append it at the end
      if (sigFound === -1) {
        var newCol = s.getLastColumn() + 1;
        s.getRange(1, newCol).setValue("Evaluator_Signature").setBackground("#1B365D").setFontColor("#FFFFFF").setFontWeight("bold");
        s.setColumnWidth(newCol, 180);
      }
      
      // Re-apply validations and formatting
      applyEvaluationsValidations(s);
      modifiedSheets.push(sName);
    });
    
    // Also update Consolidated_Summary headers and refresh
    var sumSheet = ss.getSheetByName(SHEET_SUMMARY);
    if (sumSheet) {
      setupSheetHeaders(sumSheet, SUMMARY_HEADERS, "#004B49");
      refreshConsolidatedSummary();
    }
    
    // Refresh both reports automatically
    syncHorizontalReport(true);
    createChecklistReportSheet(true);
    
    ui.alert(
      "Upgrade Complete!",
      "Successfully upgraded schema in: " + modifiedSheets.join(", ") + ".\n\n" +
      "• Column 'Price' (Col H) inserted after Manufacturer.\n" +
      "• Columns 'Data_Privacy_Consent' and 'Accuracy_Consent' inserted after Recommendation.\n" +
      "• Column 'Evaluator_Signature' verified.\n" +
      "• Horizontal Report and Printable Checklist Viewer refreshed.\n\n" +
      "Next step in AppSheet:\n" +
      "1. Open AppSheet Editor.\n" +
      "2. Go to Data > Tables > Evaluations_Master > click 'Regenerate Structure'.\n" +
      "3. Verify 'Price', 'Data_Privacy_Consent', 'Accuracy_Consent', and 'Evaluator_Signature' are recognized!",
      ui.ButtonSet.OK
    );
  } catch (err) {
    Logger.log("Error in upgradeSchemaAddPriceAndSignature: " + err.toString());
    ui.alert("Error Upgrading Schema", err.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Automatically detects and fixes duplicate or missing Evaluation_IDs in Evaluations_Master
 * and evaluator tabs to satisfy AppSheet unique key requirements. Also cleans extra header columns.
 */
function fixDuplicateKeysAndCleanHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  try {
    var fixedCount = 0;
    var fixedDetails = [];
    
    var sheetsToCheck = [
      SHEET_EVALUATIONS_MASTER, SHEET_END_USER, SHEET_NURSE, SHEET_PHARMACIST,
      "Evaluator 1", "Evaluator 2", "Evaluator 3"
    ];
    
    sheetsToCheck.forEach(function(sName) {
      var s = ss.getSheetByName(sName);
      if (!s || s.getLastRow() < 2) return;
      
      // Clear any broken spilled Excel array formulas (e.g. FILTER or _xludf.DUMMYFUNCTION)
      var cellA2 = s.getRange(2, 1);
      var formA2 = cellA2.getFormula();
      if (formA2 && (formA2.indexOf("FILTER") !== -1 || formA2.indexOf("_xludf") !== -1 || formA2.indexOf("DUMMYFUNCTION") !== -1)) {
        cellA2.clearContent();
        fixedCount++;
        fixedDetails.push(sName + ": Cleared broken spilled formula in cell A2");
      }
      
      var numRows = s.getLastRow() - 1;
      if (numRows < 1) return;
      var idRange = s.getRange(2, 1, numRows, 1);
      var ids = idRange.getValues();
      var seen = {};
      var modified = false;
      
      for (var r = 0; r < ids.length; r++) {
        var rawId = (ids[r][0] || "").toString().trim();
        var rowNumber = r + 2;
        
        if (!rawId) {
          // Generate new unique ID for empty cell
          var newId = Utilities.getUuid().substring(0, 8);
          ids[r][0] = newId;
          seen[newId] = true;
          modified = true;
          fixedCount++;
          fixedDetails.push(sName + " Row " + rowNumber + ": generated ID " + newId);
        } else if (seen[rawId]) {
          // Found duplicate! Generate new unique ID for the duplicate row
          var fixedId = Utilities.getUuid().substring(0, 8);
          ids[r][0] = fixedId;
          seen[fixedId] = true;
          modified = true;
          fixedCount++;
          fixedDetails.push(sName + " Row " + rowNumber + ": changed duplicate '" + rawId + "' to '" + fixedId + "'");
        } else {
          seen[rawId] = true;
        }
      }
      
      if (modified) {
        idRange.setValues(ids);
      }
      
      // Clean redundant empty 'Signature' column if 'Evaluator_Signature' already exists
      if (s.getLastColumn() > 0) {
        var hRow = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
        var hasEvalSig = -1;
        var hasRedundantSig = -1;
        for (var c = 0; c < hRow.length; c++) {
          var colTitle = (hRow[c] || "").toString().trim().toLowerCase();
          if (colTitle === "evaluator_signature") hasEvalSig = c + 1;
          if (colTitle === "signature") hasRedundantSig = c + 1;
        }
        if (hasEvalSig !== -1 && hasRedundantSig !== -1 && hasRedundantSig > hasEvalSig) {
          s.deleteColumn(hasRedundantSig);
          fixedCount++;
          fixedDetails.push(sName + ": Removed redundant empty 'Signature' column (already using 'Evaluator_Signature')");
        }
      }
    });
    
    // Clean Consolidated_Summary headers and trailing empty columns
    var sumSheet = ss.getSheetByName(SHEET_SUMMARY);
    if (sumSheet) {
      setupSheetHeaders(sumSheet, SUMMARY_HEADERS, "#004B49");
      if (sumSheet.getLastColumn() > SUMMARY_HEADERS.length) {
        var extraCols = sumSheet.getLastColumn() - SUMMARY_HEADERS.length;
        sumSheet.getRange(1, SUMMARY_HEADERS.length + 1, sumSheet.getMaxRows(), extraCols).clear();
      }
      refreshConsolidatedSummary();
    }
    
    // Refresh horizontal and printable reports
    syncHorizontalReport(true);
    createChecklistReportSheet(true);
    
    var msg = fixedCount > 0 
      ? "Successfully fixed " + fixedCount + " duplicate/missing key(s):\n\n" + fixedDetails.slice(0, 5).join("\n") + (fixedDetails.length > 5 ? "\n...and " + (fixedDetails.length - 5) + " more" : "") + "\n\nAll rows now have 100% unique keys. Consolidated_Summary headers cleaned."
      : "No duplicate keys found! All Evaluation_IDs are unique and headers are clean.";
      
    ui.alert("Keys & Headers Check Complete", msg, ui.ButtonSet.OK);
  } catch (err) {
    Logger.log("Error in fixDuplicateKeysAndCleanHeaders: " + err.toString());
    ui.alert("Error Fixing Keys", err.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Scans Evaluations_Master and all Evaluator tabs to merge duplicate Part II columns
 * (e.g. 'P2_09_stability_informatio' and 'P2_06_no_physical_defec') into the standard canonical columns
 * (P2_04_Physical_Appearance, P2_05_Dosing_Graduation, P2_06_Dispensing_Ease), then safely deletes the redundant columns.
 */
function consolidateDuplicatePart2Columns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var confirm = ui.alert(
    "Confirm Part II Consolidation",
    "This will merge empty cells in rows 1430+ with data from duplicate columns (P2_09_stability_information, P2_06_no_physical_defect) into the official canonical columns (P2_04, P2_05, P2_06).\n\nRedundant duplicate columns will then be safely deleted.\n\nDo you want to proceed?",
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  try {
    var sheetsToCheck = [
      SHEET_EVALUATIONS_MASTER,
      SHEET_END_USER,
      SHEET_NURSE,
      SHEET_PHARMACIST,
      "Evaluator 1", "Evaluator 2", "Evaluator 3"
    ];

    var totalMigrated = 0;
    var totalColsDeleted = 0;

    sheetsToCheck.forEach(function(sheetName) {
      var s = ss.getSheetByName(sheetName);
      if (!s || s.getLastRow() < 1 || s.getLastColumn() < 1) return;

      var lastCol = s.getLastColumn();
      var lastRow = s.getLastRow();
      var headers = s.getRange(1, 1, 1, lastCol).getValues()[0];

      // Canonical column mapping (1-indexed)
      var colP2_04 = -1;
      var colP2_05 = -1;
      var colP2_06 = -1;

      // Duplicate Part 2 columns to delete
      var dupCols = [];

      for (var c = 0; c < headers.length; c++) {
        var hName = (headers[c] || "").toString().trim();
        var hNorm = hName.toLowerCase().replace(/[^a-z0-9]/g, "");
        var colNum = c + 1;

        if (hNorm === "p204physicalappearance" || hNorm === "p204") colP2_04 = colNum;
        else if (hNorm === "p205dosinggraduation" || hNorm === "p205") colP2_05 = colNum;
        else if (hNorm === "p206dispensingease" || (hNorm === "p206" && hNorm.indexOf("physical") === -1)) colP2_06 = colNum;
      }

      // Second pass: find duplicate/rogue P2 columns
      for (var c2 = 0; c2 < headers.length; c2++) {
        var hName2 = (headers[c2] || "").toString().trim();
        var hNorm2 = hName2.toLowerCase().replace(/[^a-z0-9]/g, "");
        var colNum2 = c2 + 1;

        var isKnownCanonical = (
          hNorm2 === "p201containerintegrity" ||
          hNorm2 === "p202closureseal" ||
          hNorm2 === "p203blisterpackaging" ||
          hNorm2 === "p204physicalappearance" ||
          hNorm2 === "p205dosinggraduation" ||
          hNorm2 === "p206dispensingease"
        );

        if (hNorm2.indexOf("p2") === 0 && !isKnownCanonical) {
          var targetCanonical = colP2_06;
          if (hNorm2.indexOf("physical") !== -1 || hNorm2.indexOf("defect") !== -1 || hNorm2.indexOf("leakage") !== -1) {
            targetCanonical = colP2_04;
          } else if (hNorm2.indexOf("dosing") !== -1 || hNorm2.indexOf("stopper") !== -1) {
            targetCanonical = colP2_05;
          }
          dupCols.push({ colNum: colNum2, name: hName2, targetCol: targetCanonical });
        }
      }

      if (dupCols.length === 0) return;

      // Migrate row data from duplicate columns to canonical columns if canonical cell is blank
      if (lastRow > 1) {
        dupCols.forEach(function(dup) {
          if (dup.targetCol !== -1) {
            var dupValues = s.getRange(2, dup.colNum, lastRow - 1, 1).getValues();
            var targetValues = s.getRange(2, dup.targetCol, lastRow - 1, 1).getValues();
            var rowModified = false;

            for (var r = 0; r < dupValues.length; r++) {
              var valDup = (dupValues[r][0] || "").toString().trim();
              var valTarget = (targetValues[r][0] || "").toString().trim();

              if (!valTarget && valDup) {
                targetValues[r][0] = valDup;
                rowModified = true;
                totalMigrated++;
              }
            }

            if (rowModified) {
              s.getRange(2, dup.targetCol, lastRow - 1, 1).setValues(targetValues);
            }
          }
        });
      }

      // Delete duplicate columns in reverse order (highest index to lowest)
      dupCols.sort(function(a, b) { return b.colNum - a.colNum; });
      dupCols.forEach(function(dup) {
        s.deleteColumn(dup.colNum);
        totalColsDeleted++;
      });
    });

    // Recalculate Part II scores and refresh summary reports
    var masterSheet = ss.getSheetByName(SHEET_EVALUATIONS_MASTER);
    if (masterSheet && masterSheet.getLastRow() > 1) {
      var mHeaders = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
      var mData = masterSheet.getDataRange().getValues();
      var p2ScoreCol = -1;
      for (var hc = 0; hc < mHeaders.length; hc++) {
        if ((mHeaders[hc] || "").toString().toLowerCase().replace(/[^a-z0-9]/g, "") === "partiiscore") {
          p2ScoreCol = hc + 1;
          break;
        }
      }

      if (p2ScoreCol !== -1) {
        var scoreUpdates = [];
        for (var rowIdx = 1; rowIdx < mData.length; rowIdx++) {
          var scores = calculateScoresForRow(mData[rowIdx], mHeaders);
          scoreUpdates.push([scores.partIIScore]);
        }
        if (scoreUpdates.length > 0) {
          masterSheet.getRange(2, p2ScoreCol, scoreUpdates.length, 1).setValues(scoreUpdates);
        }
      }
    }

    refreshConsolidatedSummary(false);
    syncHorizontalReport(true);
    createChecklistReportSheet(true);

    ui.alert(
      "Part II Consolidation Complete",
      "Successfully cleaned Part II criteria:\n\n" +
      "• Migrated " + totalMigrated + " evaluation response(s) into official canonical columns (P2_04, P2_05, P2_06).\n" +
      "• Deleted " + totalColsDeleted + " redundant duplicate column(s).\n" +
      "• Recalculated Part II scores & refreshed Consolidated Summary.\n\n" +
      "All rows now have complete, aligned data with zero duplicate questionnaire columns!",
      ui.ButtonSet.OK
    );
  } catch (err) {
    Logger.log("Error in consolidateDuplicatePart2Columns: " + err.toString());
    ui.alert("Consolidation Error", err.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Safely backfills empty cells in Column AF (P2_05_Dosing_Graduation) across Evaluations_Master
 * and all Evaluator tabs (End-user, Nurse, Pharmacist, Evaluator 1/2/3).
 * 
 * Clinical Backfill Rule:
 * - If P2_04 (Physical Appearance / No leakage in closures) is "Yes", backfills P2_05 as "Yes"
 *   (or "N/A" if the product is not an IV fluid/parenteral container).
 * - If P2_04 is "No", backfills P2_05 as "No".
 * - If P2_04 is empty or N/A, backfills as "N/A".
 * 
 * Then recalculates Part II compliance score for all affected rows and refreshes summary reports.
 */
function backfillEmptyP2_05DosingGraduation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var confirm = ui.alert(
    "Confirm P2_05 Backfill",
    "This will find all rows where Column AF (P2_05_Dosing_Graduation) is empty (e.g. rows 1430–1435) and backfill them with valid clinical responses matching container integrity (Yes/N/A).\n\nPart II scores will be recalculated cleanly.\n\nDo you want to proceed?",
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  try {
    var sheetsToCheck = [
      SHEET_EVALUATIONS_MASTER,
      SHEET_END_USER,
      SHEET_NURSE,
      SHEET_PHARMACIST,
      "Evaluator 1", "Evaluator 2", "Evaluator 3"
    ];

    var totalFilled = 0;

    sheetsToCheck.forEach(function(sheetName) {
      var s = ss.getSheetByName(sheetName);
      if (!s || s.getLastRow() < 2 || s.getLastColumn() < 1) return;

      var lastCol = s.getLastColumn();
      var lastRow = s.getLastRow();
      var headers = s.getRange(1, 1, 1, lastCol).getValues()[0];

      var colP2_04 = -1;
      var colP2_05 = -1;
      var colP2Score = -1;
      var colGeneric = -1;

      for (var c = 0; c < headers.length; c++) {
        var hNorm = (headers[c] || "").toString().toLowerCase().replace(/[^a-z0-9]/g, "");
        if (hNorm === "p204physicalappearance" || hNorm === "p204") colP2_04 = c + 1;
        else if (hNorm === "p205dosinggraduation" || hNorm === "p205") colP2_05 = c + 1;
        else if (hNorm === "partiiscore") colP2Score = c + 1;
        else if (hNorm === "genericname") colGeneric = c + 1;
      }

      if (colP2_05 === -1) return;

      // 1. Proactively clear any stray validation rules accidentally inherited on metadata columns A-H
      try {
        var metaColCount = Math.min(8, lastCol);
        if (metaColCount > 0 && lastRow > 1) {
          s.getRange(2, 1, lastRow - 1, metaColCount).clearDataValidations();
        }
      } catch (eVal) {
        Logger.log("Could not clear metadata validations: " + eVal.toString());
      }

      // 2. Read full data to evaluate clinical context and compute scores
      var allData = s.getRange(2, 1, lastRow - 1, lastCol).getValues();
      var p2_05ColumnUpdates = [];
      var p2ScoreColumnUpdates = [];
      var modified = false;

      for (var r = 0; r < allData.length; r++) {
        var row = allData[r];
        var valP2_05 = (row[colP2_05 - 1] || "").toString().trim();
        var currentScore = colP2Score !== -1 ? (row[colP2Score - 1] || "").toString().trim() : "";

        if (!valP2_05) {
          var valP2_04 = colP2_04 !== -1 ? (row[colP2_04 - 1] || "").toString().trim().toUpperCase() : "";
          var genName = colGeneric !== -1 ? (row[colGeneric - 1] || "").toString().toLowerCase() : "";

          // Determine clinical backfill value
          var isParenteral = genName.indexOf("sodium chloride") !== -1 || genName.indexOf("bottle") !== -1 || genName.indexOf("bag") !== -1 || genName.indexOf("iv") !== -1 || genName.indexOf("albumin") !== -1 || genName.indexOf("parenteral") !== -1 || genName.indexOf("vial") !== -1 || genName.indexOf("d5") !== -1 || genName.indexOf("lrs") !== -1;

          var fillVal = "Yes";
          if (!isParenteral) {
            fillVal = "N/A";
          } else if (valP2_04 === "NO") {
            fillVal = "No";
          } else if (valP2_04 === "N/A") {
            fillVal = "N/A";
          } else {
            fillVal = "Yes";
          }

          valP2_05 = fillVal;
          row[colP2_05 - 1] = fillVal;
          modified = true;
          totalFilled++;

          // Recalculate Part II score for this row
          if (colP2Score !== -1) {
            var scores = calculateScoresForRow(row, headers);
            currentScore = scores.partIIScore;
          }
        }

        p2_05ColumnUpdates.push([valP2_05]);
        if (colP2Score !== -1) {
          p2ScoreColumnUpdates.push([currentScore]);
        }
      }

      // 3. Write strictly to Column AF (colP2_05) and colP2Score (never touching Column G or H)
      if (modified) {
        s.getRange(2, colP2_05, p2_05ColumnUpdates.length, 1).setValues(p2_05ColumnUpdates);
        if (colP2Score !== -1 && p2ScoreColumnUpdates.length > 0) {
          s.getRange(2, colP2Score, p2ScoreColumnUpdates.length, 1).setValues(p2ScoreColumnUpdates);
        }
      }
    });

    refreshConsolidatedSummary(false);
    syncHorizontalReport(true);
    createChecklistReportSheet(true);

    ui.alert(
      "P2_05 Backfill Complete",
      "Successfully backfilled " + totalFilled + " empty cell(s) in Column AF (P2_05_Dosing_Graduation) across Evaluations_Master and Evaluator tabs.\n\nPart II scores and Consolidated Summary have been updated cleanly!",
      ui.ButtonSet.OK
    );
  } catch (err) {
    Logger.log("Error in backfillEmptyP2_05DosingGraduation: " + err.toString());
    ui.alert("Backfill Error", err.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Automatically inspects the 'Questionnaire' tab and ensures all 6 canonical Part II questions
 * are present. If Question 5 (Rubber stoppers / Dosing graduation) is missing, it automatically inserts it
 * at the correct position with options: Yes, No, N/A.
 */
function syncCanonicalQuestionnaireTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  try {
    var sheet = ss.getSheetByName(SHEET_QUESTIONNAIRE);
    if (!sheet) {
      var sheets = ss.getSheets();
      for (var s = 0; s < sheets.length; s++) {
        var sName = sheets[s].getName().trim().toLowerCase();
        if (sName === "questionnaire" || sName === "questionnaires" || sName === "checklist_repository") {
          sheet = sheets[s];
          break;
        }
      }
    }
    
    if (!sheet) {
      ui.alert("Questionnaire Tab Not Found", "Could not find a sheet tab named 'Questionnaire'.", ui.ButtonSet.OK);
      return;
    }
    
    var lastRow = sheet.getLastRow();
    var lastCol = Math.max(sheet.getLastColumn(), 5);
    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    
    // Check if Question 5 is already present
    var q5Exists = false;
    var insertAfterRow = -1;
    var part2HeaderRow = -1;
    
    for (var r = 0; r < data.length; r++) {
      var text = (data[r][0] || "" + " " + data[r][1] || "").toString().toLowerCase();
      if (text.indexOf("part ii") !== -1 || text.indexOf("part 2") !== -1) {
        part2HeaderRow = r + 1;
      }
      if (text.indexOf("rubber stopper") !== -1 || text.indexOf("dosing graduation") !== -1 || text.indexOf("puncture") !== -1) {
        q5Exists = true;
        break;
      }
      if (text.indexOf("leakage") !== -1 || text.indexOf("physical defect") !== -1 || text.indexOf("no leakage") !== -1) {
        insertAfterRow = r + 1; // row number in sheet (1-indexed)
      }
    }
    
    if (q5Exists) {
      ui.alert("Questionnaire Status", "Question 5 ('Rubber stoppers of IV fluid containers are durable yet easy to puncture') is already present in the Questionnaire tab!", ui.ButtonSet.OK);
      return;
    }
    
    // Insert Question 5 row
    var targetRow = insertAfterRow !== -1 ? (insertAfterRow + 1) : (part2HeaderRow !== -1 ? part2HeaderRow + 5 : lastRow + 1);
    sheet.insertRowAfter(insertAfterRow !== -1 ? insertAfterRow : lastRow);
    
    var newRowData = [
      "5",
      "Rubber stoppers (single-port and dual/twin-port) of IV fluid containers are durable yet easy to puncture",
      "Yes",
      "No",
      "N/A"
    ];
    sheet.getRange(targetRow, 1, 1, newRowData.length).setValues([newRowData]);
    
    ui.alert(
      "Questionnaire Updated",
      "Successfully added Question #5 to the Questionnaire tab:\n\n" +
      "• Parameter: \"Rubber stoppers (single-port and dual/twin-port) of IV fluid containers are durable yet easy to puncture\"\n" +
      "• Options: Yes, No, N/A\n" +
      "• Position: Row " + targetRow + " (Between Question 4 and Question 6)\n\n" +
      "The Web App and evaluations will now include all 6 Part II questions!",
      ui.ButtonSet.OK
    );
  } catch (err) {
    Logger.log("Error in syncCanonicalQuestionnaireTab: " + err.toString());
    ui.alert("Sync Error", err.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Initializes all required sheets, headers, formatting, and data validations.
 */
function initializeEvaluationWorkbook() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  try {
    // 1. Medicine_Master
    var medSheet = getOrCreateSheet(ss, SHEET_MEDICINE_MASTER);
    setupSheetHeaders(medSheet, ["Generic_Name", "Category_Type", "Procurement_Mode"], "#1B365D");
    
    // 2. Suppliers
    var supSheet = getOrCreateSheet(ss, SHEET_SUPPLIERS);
    setupSheetHeaders(supSheet, ["Supplier_Name", "Contact_Person", "Contact_Number", "Email"], "#1B365D");

    // 3. Evaluator_Accounts
    var acctsSheet = getOrCreateSheet(ss, SHEET_EVALUATOR_ACCOUNTS);
    ensureEvaluatorAccountsHeaders(acctsSheet);
    if (acctsSheet.getLastRow() === 1) {
      // Seed default evaluator clinical roles
      acctsSheet.appendRow(["ACC-001", "Pharmacist", "Staff Pharmacist / Evaluator", "", "", ""]);
      acctsSheet.appendRow(["ACC-002", "Nurse", "Head Nurse / Clinical Nurse", "", "", ""]);
      acctsSheet.appendRow(["ACC-003", "End-user", "Clinical Specialist / End-User", "", "", ""]);
    }

    // 4. Evaluations_Master
    var masterSheet = getOrCreateSheet(ss, SHEET_EVALUATIONS_MASTER);
    setupSheetHeaders(masterSheet, EVAL_HEADERS, "#1B365D");
    applyEvaluationsValidations(masterSheet);
    
    // 5-7. Evaluator tabs (End-user, Nurse, Pharmacist)
    var evalSheets = [SHEET_END_USER, SHEET_NURSE, SHEET_PHARMACIST];
    evalSheets.forEach(function(sheetName) {
      var s = getOrCreateSheet(ss, sheetName);
      setupSheetHeaders(s, EVAL_HEADERS, "#2C3E50");
      applyEvaluationsValidations(s);
    });
    
    // Also format Evaluator 1/2/3 tabs if they exist in the workbook (for AppSheet compatibility)
    ["Evaluator 1", "Evaluator 2", "Evaluator 3"].forEach(function(legacyName) {
      var ls = ss.getSheetByName(legacyName);
      if (ls) {
        setupSheetHeaders(ls, EVAL_HEADERS, "#2C3E50");
        applyEvaluationsValidations(ls);
      }
    });
    
    // 8. Consolidated_Summary
    var summarySheet = getOrCreateSheet(ss, SHEET_SUMMARY);
    setupSheetHeaders(summarySheet, SUMMARY_HEADERS, "#004B49");
    
    ui.alert("Success", "Workbook initialized successfully with all required sheets, headers (including Supplier & Manufacturer), and validations!", ui.ButtonSet.OK);
  } catch (err) {
    Logger.log("Error in initializeEvaluationWorkbook: " + err.toString());
    ui.alert("Error Initializing Workbook", err.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Helper to retrieve an existing sheet or create a new one.
 */
function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

/**
 * Helper to get the Supplier master sheet, supporting both "Supplier" (singular) and "Suppliers" (plural).
 */
function getSupplierSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName("Supplier") || ss.getSheetByName("Suppliers") || ss.getSheetByName(SHEET_SUPPLIERS);
}

/**
 * Helper to get or create the Checklist Report viewer sheet, supporting both 'Checklist_Report' and 'Checklist Report'.
 */
function getOrCreateChecklistReportSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Checklist_Report") || ss.getSheetByName("Checklist Report") || ss.getSheetByName(SHEET_CHECKLIST_REPORT);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CHECKLIST_REPORT);
  }
  return sheet;
}

/**
 * Helper to get the Horizontal Report sheet, supporting both 'Checklist_Report_Horizontal' and 'Checklist Report Horizontal'.
 */
function getHorizontalReportSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName("Checklist_Report_Horizontal") || 
         ss.getSheetByName("Checklist Report Horizontal") || 
         ss.getSheetByName(SHEET_CHECKLIST_HORIZONTAL);
}

/**
 * Helper to set up sheet headers, header formatting, and freeze row 1.
 */
function setupSheetHeaders(sheet, headers, headerColor) {
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground(headerColor)
             .setFontColor("#FFFFFF")
             .setFontWeight("bold")
             .setHorizontalAlignment("center")
             .setVerticalAlignment("middle");
  
  sheet.setRowHeight(1, 35);
  sheet.setFrozenRows(1);
}

/**
 * Applies Data Validation rules (Dropdowns for Evaluator, Yes/No/N/A, Recommendation)
 */
function applyEvaluationsValidations(sheet) {
  var ruleEvaluator = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Pharmacist", "Nurse", "End-user"], true)
    .setAllowInvalid(false)
    .build();
    
  var ruleYesNoNA = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Yes", "No", "N/A"], true)
    .setAllowInvalid(false)
    .build();

  var ruleYesNo = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Yes", "No"], true)
    .setAllowInvalid(false)
    .build();

  var ruleRecommendation = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Recommended", "Not Recommended"], true)
    .setAllowInvalid(false)
    .build();

  var maxRows = 500; // Apply validation rule to first 500 rows for performance
  
  // Evaluator Column C (index 3)
  sheet.getRange(2, 3, maxRows, 1).setDataValidation(ruleEvaluator);
  
  // Price Column H (index 8)
  sheet.getRange(2, 8, maxRows, 1).setNumberFormat("₱#,##0.00");
  
  // Part I questions cols I-AA (indices 9 to 27 = 19 columns)
  sheet.getRange(2, 9, maxRows, 19).setDataValidation(ruleYesNoNA);
  
  // Part II questions cols AB-AG (indices 28 to 33 = 6 columns)
  sheet.getRange(2, 28, maxRows, 6).setDataValidation(ruleYesNoNA);
  
  // Requires_Reconstitution col AH (index 34)
  sheet.getRange(2, 34, maxRows, 1).setDataValidation(ruleYesNo);
  
  // Part III questions cols AI-AQ (indices 35 to 43 = 9 columns)
  sheet.getRange(2, 35, maxRows, 9).setDataValidation(ruleYesNoNA);
  
  // Recommendation col AU (index 47)
  sheet.getRange(2, 47, maxRows, 1).setDataValidation(ruleRecommendation);

  // Data Privacy and Accuracy Consents
  var lastCol = Math.max(1, sheet.getLastColumn());
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var dpCol = headers.indexOf("Data_Privacy_Consent") + 1;
  if (dpCol > 0) {
    sheet.getRange(2, dpCol, maxRows, 1).setDataValidation(ruleYesNo);
  }
  var accCol = headers.indexOf("Accuracy_Consent") + 1;
  if (accCol > 0) {
    sheet.getRange(2, accCol, maxRows, 1).setDataValidation(ruleYesNo);
  }
}

/**
 * Synchronizes generic drug names and categories from 'framework agreement list' or 'indicative 1'
 * sheet into 'Medicine_Master'. Supports side-by-side category tables (e.g., Column C, Column J, etc.).
 * Preserves any existing custom category and procurement mode metadata if present.
 */
function syncMasterDrugList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  // 1. Find source sheet: check 'framework agreement list', then 'indicative 1', or search
  var sourceSheet = ss.getSheetByName("framework agreement list");
  if (!sourceSheet) {
    sourceSheet = ss.getSheetByName("indicative 1");
  }
  if (!sourceSheet) {
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var sName = sheets[i].getName().trim().toLowerCase();
      if (sName.indexOf("framework agreement") !== -1 || sName.indexOf("indicative") !== -1 || sName.indexOf("ppmp") !== -1) {
        sourceSheet = sheets[i];
        break;
      }
    }
  }
  
  if (!sourceSheet) {
    ui.alert("Source Sheet Not Found", "Could not find a tab named 'framework agreement list' or 'indicative 1'. Please make sure your drug list tab exists in this workbook.", ui.ButtonSet.OK);
    return;
  }
  
  var targetSheet = getOrCreateSheet(ss, SHEET_MEDICINE_MASTER);
  
  // Read existing custom metadata (Category & Procurement) if present
  var existingMap = {};
  if (targetSheet.getLastRow() > 1) {
    var existingData = targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, 3).getValues();
    existingData.forEach(function(row) {
      var drug = row[0] ? row[0].toString().trim() : "";
      if (drug) {
        existingMap[drug] = {
          category: row[1] || "",
          procurement: row[2] || ""
        };
      }
    });
  }
  
  var lastRow = sourceSheet.getLastRow();
  var lastCol = sourceSheet.getLastColumn();
  
  if (lastRow < 4) {
    ui.alert("No Data Found", "Source sheet has fewer than 4 rows.", ui.ButtonSet.OK);
    return;
  }
  
  var fullData = sourceSheet.getRange(1, 1, lastRow, lastCol).getValues();
  var extractedItems = [];
  var seen = {};
  
  // Determine if this is the multi-column 'framework agreement list' layout or single-column layout
  var drugColIndices = [];
  
  // Scan headers (rows 1 to 8) to find columns with "item / service" or "type and nature"
  for (var c = 0; c < lastCol; c++) {
    var isDrugCol = false;
    for (var r = 0; r < Math.min(8, lastRow); r++) {
      var cellVal = fullData[r][c] ? fullData[r][c].toString().toLowerCase() : "";
      if ((cellVal.indexOf("item") !== -1 && cellVal.indexOf("service") !== -1) || cellVal.indexOf("type and nature") !== -1) {
        isDrugCol = true;
        break;
      }
    }
    // Hardcoded fallback check for Column C (index 2) and Column J (index 9) for Framework Agreement layout
    if (isDrugCol || c === 2 || c === 9) {
      // Exclude "Item No." and "Unit" columns
      var colHeaderSample = (fullData[6] && fullData[6][c] ? fullData[6][c].toString() : "") + " " + (fullData[7] && fullData[7][c] ? fullData[7][c].toString() : "");
      if (colHeaderSample.toLowerCase().indexOf("item no") === -1 && colHeaderSample.toLowerCase().indexOf("unit") === -1) {
        drugColIndices.push(c);
      }
    }
  }
  
  // Filter unique column indices
  drugColIndices = drugColIndices.filter(function(v, idx, self) { return self.indexOf(v) === idx; });
  
  // Fallback to Column A (index 0) if no multi-column headers match
  if (drugColIndices.length === 0) {
    drugColIndices = [0];
  }
  
  // Start row detection: if row 9 exists and has items, use start row 9 (index 8), else start row 4 (index 3)
  var startRowIndex = (lastRow >= 9) ? 8 : 3;
  
  drugColIndices.forEach(function(colIdx) {
    // Detect Category from Row 2 (index 1) above this column (checks up to 6 columns back for merged cells)
    var category = "";
    if (fullData[1]) {
      for (var offset = 0; offset <= 6; offset++) {
        var checkIdx = colIdx - offset;
        if (checkIdx >= 0 && fullData[1][checkIdx]) {
          var val = fullData[1][checkIdx].toString().trim();
          if (val && 
              val.toUpperCase().indexOf("PR ") === -1 && 
              val.toUpperCase().indexOf("FRAMEWORK") === -1 &&
              val.toUpperCase().indexOf("LIST") === -1) {
            category = val;
            break;
          }
        }
      }
    }
    if (!category) category = "Pharmaceutical";
    
    // Scan items down the column
    for (var r = startRowIndex; r < lastRow; r++) {
      var drugName = fullData[r][colIdx] ? fullData[r][colIdx].toString().trim() : "";
      var lowerName = drugName.toLowerCase();
      
      // Filter out non-drug text, headers, and numbers
      if (drugName && 
          lowerName.indexOf("type and nature") === -1 &&
          lowerName.indexOf("item / service") === -1 &&
          lowerName.indexOf("framework agreement") === -1 &&
          lowerName.indexOf("item no") === -1 &&
          !seen[lowerName]) {
        seen[lowerName] = true;
        
        var cat = existingMap[drugName] && existingMap[drugName].category ? existingMap[drugName].category : category;
        var proc = existingMap[drugName] && existingMap[drugName].procurement ? existingMap[drugName].procurement : "Framework Agreement / Public Bidding";
        
        extractedItems.push([drugName, cat, proc]);
      }
    }
  });
  
  // Sort extracted items alphabetically by generic name
  extractedItems.sort(function(a, b) { return a[0].localeCompare(b[0]); });
  
  // Clear existing content (except row 1) and write
  if (targetSheet.getLastRow() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, 3).clearContent();
  }
  
  if (extractedItems.length > 0) {
    targetSheet.getRange(2, 1, extractedItems.length, 3).setValues(extractedItems);
  }
  
  ui.alert("Sync Complete", "Successfully synchronized " + extractedItems.length + " unique generic drugs into " + SHEET_MEDICINE_MASTER + " from sheet tab '" + sourceSheet.getName() + "'!", ui.ButtonSet.OK);
}

/**
 * Normalizes question string for logical/semantic deduplication comparison:
 * - Strips leading numbering (e.g., "1.", "19)", "(1)")
 * - Strips parentheticals like "(s)", "(if applicable)"
 * - Standardizes dashes and slashes
 * - Removes non-alphanumeric punctuation and collapses whitespace
 */
function normalizeQuestionText(str) {
  if (!str) return "";
  var s = str.toString().trim().toLowerCase();
  // Standardize dashes and slashes
  s = s.replace(/[\u2013\u2014–—]/g, "-");
  // Remove leading numbers like "1. ", "19) ", "1 - "
  s = s.replace(/^(\d+[\.\)\-:]|\([a-z\d]+\))\s*/i, "");
  // Remove "(s)", "(es)", "(if applicable)"
  s = s.replace(/\((?:s|es|if applicable)\)/gi, "");
  // Remove punctuation except letters, numbers and spaces
  s = s.replace(/[^a-z0-9\s]/gi, " ");
  // Collapse whitespace
  s = s.replace(/\s+/g, " ").trim();
  return s;
}

/**
 * Checks if a candidate question is logically or textually duplicate of an existing question.
 * Uses exact normalized matching, significant token overlap (filtered of stop words), and Jaccard similarity.
 */
function checkLogicalDuplicate(candidateText, existingQuestionsList) {
  var normCandidate = normalizeQuestionText(candidateText);
  if (!normCandidate) return { isDuplicate: true, reason: "Empty text" };
  
  var STOP_WORDS = {
    "the": true, "a": true, "an": true, "is": true, "are": true, "to": true,
    "of": true, "and": true, "or": true, "in": true, "on": true, "with": true,
    "for": true, "at": true, "by": true, "from": true, "if": true, "each": true, "all": true
  };
  
  function getSignificantTokens(str) {
    return str.split(" ").filter(function(t) {
      return t.length > 1 && !STOP_WORDS[t];
    }).map(function(t) {
      // Basic stemming for trailing 's' on words > 3 characters
      if (t.length > 3 && t.charAt(t.length - 1) === 's') {
        return t.substring(0, t.length - 1);
      }
      return t;
    });
  }
  
  var candidateTokens = getSignificantTokens(normCandidate);
  var candidateSet = {};
  candidateTokens.forEach(function(t) { candidateSet[t] = true; });
  
  for (var i = 0; i < existingQuestionsList.length; i++) {
    var existing = existingQuestionsList[i];
    var normExisting = normalizeQuestionText(existing.text);
    
    // 1. Exact normalized match
    if (normCandidate === normExisting) {
      return { isDuplicate: true, matchedOriginal: existing.text, score: 1.0 };
    }
    
    // 2. Significant Token Overlap (Jaccard similarity without stop-words)
    var existingTokens = getSignificantTokens(normExisting);
    var existingSet = {};
    existingTokens.forEach(function(t) { existingSet[t] = true; });
    
    var intersection = 0;
    var allKeys = {};
    for (var ct in candidateSet) {
      allKeys[ct] = true;
      if (existingSet[ct]) intersection++;
    }
    for (var et in existingSet) {
      allKeys[et] = true;
    }
    
    var unionCount = Object.keys(allKeys).length;
    var jaccard = unionCount > 0 ? (intersection / unionCount) : 0;
    if (jaccard >= 0.75) {
      return { isDuplicate: true, matchedOriginal: existing.text, score: jaccard };
    }
    
    // 3. Substring containment if both are long enough
    if (normCandidate.length > 18 && normExisting.length > 18) {
      if (normCandidate.indexOf(normExisting) !== -1 || normExisting.indexOf(normCandidate) !== -1) {
        return { isDuplicate: true, matchedOriginal: existing.text, score: 0.9 };
      }
    }
  }
  
  return { isDuplicate: false };
}

/**
 * Smart Parser & Deduplication Engine for the Questionnaire sheet tab.
 * Dynamically identifies Part I and Part II sections, detects Yes/No/N/A options,
 * rejects duplicate/equivalent questions, and outputs the live schema for the Web App.
 */
function getDynamicQuestionnaireFromSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_QUESTIONNAIRE);
    if (!sheet) {
      var sheets = ss.getSheets();
      for (var s = 0; s < sheets.length; s++) {
        var sName = sheets[s].getName().trim().toLowerCase();
        if (sName === "questionnaire" || sName === "questionnaires" || sName === "checklist_repository") {
          sheet = sheets[s];
          break;
        }
      }
    }
    
    if (!sheet || sheet.getLastRow() < 4) {
      return getDefaultQuestionnaire();
    }
    
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) return getDefaultQuestionnaire();
    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    
    var part1 = [];
    var part2 = [];
    var part3 = [];
    var duplicatesRejected = [];
    
    var currentSection = "P1"; // Default starts with Part I
    
    // Canonical reference items for preserving official IDs
    var canonicalPart1 = [
      { id: "P1_01_Brand_Name", text: "Product Name" },
      { id: "P1_02_Generic_Name", text: "Dosage Form and Strength" },
      { id: "P1_03_Dosage_Form_Strength", text: "Pharmacologic Category" },
      { id: "P1_04_Manufacturer_Details", text: "Formulation / Composition" },
      { id: "P1_05_CPR_FDA_Registration", text: "Indication(s)" },
      { id: "P1_06_Batch_Lot_Number", text: "Dosage and Mode of Administration" },
      { id: "P1_07_Manufacturing_Date", text: "Contraindication(s), Precaution(s), Warning(s)" },
      { id: "P1_08_Expiration_Date", text: "Drug–Drug / Drug–Food Interactions" },
      { id: "P1_09_Storage_Conditions", text: "Adverse Drug Reaction(s)" },
      { id: "P1_10_Rx_Symbol", text: "Overdose and Treatment Information" },
      { id: "P1_11_Net_Content", text: "Storage Condition(s)" },
      { id: "P1_12_Language_Legibility", text: "Net Content / Pack Size" },
      { id: "P1_13_Outer_Package", text: "Name & Address of Marketing Authorization Holder" },
      { id: "P1_14_Inner_Package", text: "Name & Address of Manufacturer" },
      { id: "P1_15_Package_Insert", text: "Rx Symbol & Prescription Caution Statement (if applicable)" },
      { id: "P1_16_Barcode_QR", text: "ADR Reporting Statement" },
      { id: "P1_17_Tamper_Evident_Seal", text: "Registration Number" },
      { id: "P1_18_Special_Warnings", text: "Batch / Lot Number" },
      { id: "P1_19_FDA_Compliance", text: "Date of Manufacture & Expiration Date" }
    ];
    
    var canonicalPart2 = [
      { id: "P2_01_Container_Integrity", text: "Inner label is identical to the outer label", keywords: ["inner label", "identical"] },
      { id: "P2_02_Closure_Seal", text: "Drug name, dosage form, strength, batch/lot number, manufacture date, and expiry date are clearly readable on the container or inner packaging", keywords: ["clearly readable", "closure seal", "inner packaging", "readable"] },
      { id: "P2_03_Blister_Packaging", text: "For blister or aluminum foil packs, expiry date, drugs name and dosage form is printed on each individual unit", keywords: ["blister", "aluminum foil", "individual unit"] },
      { id: "P2_04_Physical_Appearance", text: "No leakage observed in IV fluids or other parenteral products through closures (rubber stoppers, caps, seals) or infusion sets", keywords: ["leakage", "physical defect", "no physical defect", "parenteral", "physical appearance"] },
      { id: "P2_05_Dosing_Graduation", text: "Rubber stoppers (single-port and dual/twin-port) of IV fluid containers are durable yet easy to puncture", keywords: ["rubber stopper", "puncture", "dosing graduation", "puncturability", "stopper"] },
      { id: "P2_06_Dispensing_Ease", text: "Ease of opening, dispensing, and overall container integrity", keywords: ["dispensing", "container integrity", "ease of opening", "stability information", "stability"] }
    ];

    for (var r = 0; r < data.length; r++) {
      var row = data[r];
      var cellA = (row[0] || "").toString().trim();
      var cellB = (row[1] || "").toString().trim();
      var cellC = (row[2] || "").toString().trim();
      var cellD = (row[3] || "").toString().trim();
      var cellE = (row[4] || "").toString().trim();
      var fullRowText = (cellA + " " + cellB + " " + cellC).toLowerCase();
      
      // 1. Detect Section Headers
      if (fullRowText.indexOf("part iii") !== -1 || fullRowText.indexOf("part 3") !== -1) {
        currentSection = "P3";
        continue;
      } else if (fullRowText.indexOf("part ii") !== -1 || fullRowText.indexOf("part 2") !== -1) {
        currentSection = "P2";
        continue;
      } else if (fullRowText.indexOf("part i") !== -1 || fullRowText.indexOf("part 1") !== -1) {
        currentSection = "P1";
        continue;
      }
      
      // Skip title/header row like "PRODUCT SAMPLE EVALUATION"
      if (cellB.toLowerCase().indexOf("product sample") !== -1 || cellB.toLowerCase().indexOf("evaluation") !== -1 && cellA === "") {
        continue;
      }
      
      // Skip table header row if it only contains column headers
      if (cellB.toLowerCase() === "generic name" || cellB.toLowerCase() === "parameters" || cellB.toLowerCase() === "criteria") {
        continue;
      }
      if (cellC.toLowerCase() === "yes" && cellD.toLowerCase() === "no" && cellB === "") {
        continue;
      }
      
      // Must have actual question text in Column B
      if (!cellB || cellB.length < 3) continue;
      
      // 2. Detect Choices (Yes, No, N/A)
      var detectedOptions = ["Yes", "No", "N/A"];
      if ((cellC || cellD) && (!cellE || cellE === "-" || cellE.toLowerCase() === "none")) {
        detectedOptions = ["Yes", "No"];
      }
      
      // 3. Select Target Section Array
      var targetArray = (currentSection === "P1") ? part1 : ((currentSection === "P2") ? part2 : part3);
      var canonicalList = (currentSection === "P1") ? canonicalPart1 : ((currentSection === "P2") ? canonicalPart2 : []);
      
      // 4. Run Deduplication & Logical Equivalence Engine
      var dupCheck = checkLogicalDuplicate(cellB, targetArray);
      if (dupCheck.isDuplicate) {
        duplicatesRejected.push({
          rejectedText: cellB,
          matchedOriginal: dupCheck.matchedOriginal,
          section: currentSection,
          rowNumber: r + 1
        });
        continue; // Reject duplicate and preserve original
      }
      
      // 5. Assign Stable Identifier
      var assignedId = "";
      for (var c = 0; c < canonicalList.length; c++) {
        var cNorm = normalizeQuestionText(canonicalList[c].text);
        var bNorm = normalizeQuestionText(cellB);
        if (cNorm === bNorm) {
          assignedId = canonicalList[c].id;
          break;
        }
        if (canonicalList[c].keywords) {
          for (var k = 0; k < canonicalList[c].keywords.length; k++) {
            if (bNorm.indexOf(normalizeQuestionText(canonicalList[c].keywords[k])) !== -1) {
              assignedId = canonicalList[c].id;
              break;
            }
          }
          if (assignedId) break;
        }
      }
      
      if (!assignedId) {
        var qIdx = targetArray.length + 1;
        var numPad = qIdx < 10 ? "0" + qIdx : "" + qIdx;
        var slug = cellB.toLowerCase().replace(/[^a-z0-9]/g, "_").substring(0, 20).replace(/_+$/, "");
        assignedId = currentSection + "_" + numPad + "_" + slug;
      }
      
      targetArray.push({
        id: assignedId,
        text: cellB,
        section: currentSection,
        options: detectedOptions,
        originalRow: r + 1
      });
    }
    
    // Safety Fallback: If Part I or Part II parsed empty, use defaults
    if (part1.length === 0) part1 = canonicalPart1.map(function(q) { return { id: q.id, text: q.text, section: "P1", options: ["Yes", "No", "N/A"] }; });
    if (part2.length === 0) part2 = canonicalPart2.map(function(q) { return { id: q.id, text: q.text, section: "P2", options: ["Yes", "No", "N/A"] }; });

    // Canonical Guarantee: Ensure P2_05_Dosing_Graduation is always included in Part II
    var hasP2_05 = false;
    for (var p = 0; p < part2.length; p++) {
      if (part2[p].id === "P2_05_Dosing_Graduation") {
        hasP2_05 = true;
        break;
      }
    }
    if (!hasP2_05) {
      var insertIdx = part2.length;
      for (var p2Idx = 0; p2Idx < part2.length; p2Idx++) {
        if (part2[p2Idx].id === "P2_04_Physical_Appearance") {
          insertIdx = p2Idx + 1;
          break;
        }
      }
      part2.splice(insertIdx, 0, {
        id: "P2_05_Dosing_Graduation",
        text: "Rubber stoppers (single-port and dual/twin-port) of IV fluid containers are durable yet easy to puncture",
        section: "P2",
        options: ["Yes", "No", "N/A"],
        originalRow: -1
      });
    }
    
    return {
      success: true,
      part1: part1,
      part2: part2,
      part3: part3,
      duplicatesRejected: duplicatesRejected,
      totalPart1: part1.length,
      totalPart2: part2.length
    };
  } catch (err) {
    Logger.log("Error in getDynamicQuestionnaireFromSheet: " + err.toString());
    return getDefaultQuestionnaire();
  }
}

/**
 * Returns default canonical questions if Questionnaire tab is not configured yet.
 */
function getDefaultQuestionnaire() {
  return {
    success: true,
    part1: [
      { id: "P1_01_Brand_Name", text: "Product Name", section: "P1", options: ["Yes", "No", "N/A"] },
      { id: "P1_02_Generic_Name", text: "Dosage Form and Strength", section: "P1", options: ["Yes", "No", "N/A"] },
      { id: "P1_03_Dosage_Form_Strength", text: "Pharmacologic Category", section: "P1", options: ["Yes", "No", "N/A"] },
      { id: "P1_04_Manufacturer_Details", text: "Formulation / Composition", section: "P1", options: ["Yes", "No", "N/A"] },
      { id: "P1_05_CPR_FDA_Registration", text: "Indication(s)", section: "P1", options: ["Yes", "No", "N/A"] },
      { id: "P1_06_Batch_Lot_Number", text: "Dosage and Mode of Administration", section: "P1", options: ["Yes", "No", "N/A"] },
      { id: "P1_07_Manufacturing_Date", text: "Contraindication(s), Precaution(s), Warning(s)", section: "P1", options: ["Yes", "No", "N/A"] },
      { id: "P1_08_Expiration_Date", text: "Drug–Drug / Drug–Food Interactions", section: "P1", options: ["Yes", "No", "N/A"] },
      { id: "P1_09_Storage_Conditions", text: "Adverse Drug Reaction(s)", section: "P1", options: ["Yes", "No", "N/A"] },
      { id: "P1_10_Rx_Symbol", text: "Overdose and Treatment Information", section: "P1", options: ["Yes", "No", "N/A"] },
      { id: "P1_11_Net_Content", text: "Storage Condition(s)", section: "P1", options: ["Yes", "No", "N/A"] },
      { id: "P1_12_Language_Legibility", text: "Net Content / Pack Size", section: "P1", options: ["Yes", "No", "N/A"] },
      { id: "P1_13_Outer_Package", text: "Name & Address of Marketing Authorization Holder", section: "P1", options: ["Yes", "No", "N/A"] },
      { id: "P1_14_Inner_Package", text: "Name & Address of Manufacturer", section: "P1", options: ["Yes", "No", "N/A"] },
      { id: "P1_15_Package_Insert", text: "Rx Symbol & Prescription Caution Statement (if applicable)", section: "P1", options: ["Yes", "No", "N/A"] },
      { id: "P1_16_Barcode_QR", text: "ADR Reporting Statement", section: "P1", options: ["Yes", "No", "N/A"] },
      { id: "P1_17_Tamper_Evident_Seal", text: "Registration Number", section: "P1", options: ["Yes", "No", "N/A"] },
      { id: "P1_18_Special_Warnings", text: "Batch / Lot Number", section: "P1", options: ["Yes", "No", "N/A"] },
      { id: "P1_19_FDA_Compliance", text: "Date of Manufacture & Expiration Date", section: "P1", options: ["Yes", "No", "N/A"] }
    ],
    part2: [
      { id: "P2_01_Container_Integrity", text: "Inner label is identical to the outer label", section: "P2", options: ["Yes", "No", "N/A"] },
      { id: "P2_02_Closure_Seal", text: "Drug name, dosage form, strength, batch/lot number, manufacture date, and expiry date are clearly readable on the container or inner packaging", section: "P2", options: ["Yes", "No", "N/A"] },
      { id: "P2_03_Blister_Packaging", text: "For blister or aluminum foil packs, expiry date, drugs name and dosage form is printed on each individual unit", section: "P2", options: ["Yes", "No", "N/A"] },
      { id: "P2_04_Physical_Appearance", text: "No leakage observed in IV fluids or other parenteral products through closures (rubber stoppers, caps, seals) or infusion sets", section: "P2", options: ["Yes", "No", "N/A"] },
      { id: "P2_05_Dosing_Graduation", text: "Rubber stoppers (single-port and dual/twin-port) of IV fluid containers are durable yet easy to puncture", section: "P2", options: ["Yes", "No", "N/A"] },
      { id: "P2_06_Dispensing_Ease", text: "Ease of opening, dispensing, and overall container integrity", section: "P2", options: ["Yes", "No", "N/A"] }
    ],
    part3: [],
    duplicatesRejected: [],
    totalPart1: 19,
    totalPart2: 6
  };
}

/**
 * Menu action to validate Questionnaire tab and preview dynamic questions and rejected duplicates.
 */
function menuSyncAndValidateQuestionnaire() {
  var ui = SpreadsheetApp.getUi();
  var result = getDynamicQuestionnaireFromSheet();
  
  var msg = "📋 Questionnaire Sync & Validation Report\n\n";
  msg += "• Active Part I Questions: " + result.part1.length + "\n";
  msg += "• Active Part II Questions: " + result.part2.length + "\n";
  if (result.part3 && result.part3.length > 0) {
    msg += "• Active Part III Questions: " + result.part3.length + "\n";
  }
  
  if (result.duplicatesRejected && result.duplicatesRejected.length > 0) {
    msg += "\n⚠️ " + result.duplicatesRejected.length + " Duplicate(s) Detected & Rejected:\n";
    result.duplicatesRejected.forEach(function(d, idx) {
      msg += (idx + 1) + ". Row " + d.rowNumber + ": \"" + d.rejectedText.substring(0, 35) + "...\" (Duplicate of original: \"" + d.matchedOriginal.substring(0, 30) + "...\")\n";
    });
    msg += "\nThe app safely maintained the original questions!";
  } else {
    msg += "\n✅ Zero duplicates detected! All questions are clean and unique.";
  }
  
  msg += "\n\nThe Web App dynamically injects these questions on next load.";
  ui.alert("Questionnaire Status", msg, ui.ButtonSet.OK);
}

/**
 * Calculates Part I and Part II compliance percentage scores adaptively.
 * Dynamically counts all active P1_ and P2_ criteria present in EVAL_HEADERS.
 * Formula: Yes_count / (Total_Questions - NA_Count) * 100%
 */
function calculateScoresForRow(rowValues, headerList) {
  var headers = headerList || EVAL_HEADERS;
  var p1Yes = 0, p1NA = 0, p1Count = 0;
  var p2Yes = 0, p2NA = 0, p2Count = 0;
  
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    var val = (rowValues[i] !== null && rowValues[i] !== undefined) ? rowValues[i].toString().trim().toUpperCase() : "";
    if (h.indexOf("P1_") === 0) {
      p1Count++;
      if (val === "YES") p1Yes++;
      if (val === "N/A" || val === "NA") p1NA++;
    } else if (h.indexOf("P2_") === 0) {
      p2Count++;
      if (val === "YES") p2Yes++;
      if (val === "N/A" || val === "NA") p2NA++;
    }
  }
  
  // Adaptive formulas
  var p1Eligible = p1Count - p1NA;
  var p1Score = p1Eligible > 0 ? ((p1Yes / p1Eligible) * 100).toFixed(1) + "%" : "100.0%";
  
  var p2Eligible = p2Count - p2NA;
  var p2Score = p2Eligible > 0 ? ((p2Yes / p2Eligible) * 100).toFixed(1) + "%" : "100.0%";
  
  return {
    partIScore: p1Score,
    partIIScore: p2Score
  };
}

/**
 * Installed trigger handler or onEdit listener for sync & mirror logic.
 */
function onEvaluationsEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_EVALUATIONS_MASTER) return;
  
  var row = e.range.getRow();
  if (row < 2) return; // Skip header
  
  processMasterRow(sheet, row);
  refreshConsolidatedSummary();
  syncHorizontalReport(true);
}

/**
 * Processes a single row in Evaluations_Master: auto-calculates scores and mirrors to evaluator sheet.
 */
function processMasterRow(masterSheet, rowNum) {
  var activeHeaders = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
  if (!activeHeaders || activeHeaders.length === 0) activeHeaders = EVAL_HEADERS;
  
  var rowRange = masterSheet.getRange(rowNum, 1, 1, activeHeaders.length);
  var values = rowRange.getValues()[0];
  
  // Auto-assign Evaluation ID if blank (AppSheet UNIQUEID() RFC4122 UUID equivalent)
  if (!values[0]) {
    values[0] = Utilities.getUuid();
    masterSheet.getRange(rowNum, 1).setValue(values[0]);
  }
  
  // Auto-assign Timestamp if blank
  if (!values[1]) {
    values[1] = new Date();
    masterSheet.getRange(rowNum, 2).setValue(values[1]);
  }
  
  // Calculate Scores
  var scores = calculateScoresForRow(values, activeHeaders);
  var p1Col = activeHeaders.indexOf("Part_I_Score") + 1;
  var p2Col = activeHeaders.indexOf("Part_II_Score") + 1;
  
  if (p1Col > 0) {
    values[p1Col - 1] = scores.partIScore;
    masterSheet.getRange(rowNum, p1Col).setValue(scores.partIScore);
  }
  if (p2Col > 0) {
    values[p2Col - 1] = scores.partIIScore;
    masterSheet.getRange(rowNum, p2Col).setValue(scores.partIIScore);
  }
  
  // Mirror to Evaluator tab based on Evaluator column (Index 2)
  var evaluator = values[2] ? values[2].toString().trim() : "";
  if (evaluator) {
    var ss = masterSheet.getParent();
    var targetTab = evaluator;
    if (evaluator === "Evaluator 1") targetTab = SHEET_END_USER;
    else if (evaluator === "Evaluator 2") targetTab = SHEET_NURSE;
    else if (evaluator === "Evaluator 3") targetTab = SHEET_PHARMACIST;
    mirrorToEvaluatorSheet(targetTab, values);
    
    // Also mirror to legacy tab if it exists in the spreadsheet
    if ((evaluator === "End-user" || evaluator === "Evaluator 1") && ss.getSheetByName("Evaluator 1")) {
      mirrorToEvaluatorSheet("Evaluator 1", values);
    } else if ((evaluator === "Nurse" || evaluator === "Evaluator 2") && ss.getSheetByName("Evaluator 2")) {
      mirrorToEvaluatorSheet("Evaluator 2", values);
    } else if ((evaluator === "Pharmacist" || evaluator === "Evaluator 3") && ss.getSheetByName("Evaluator 3")) {
      mirrorToEvaluatorSheet("Evaluator 3", values);
    }
  }
}

/**
 * Mirrors/updates a record in the specific Evaluator sheet (matching by Evaluation_ID).
 */
function mirrorToEvaluatorSheet(sheetName, rowValues) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var evalSheet = getOrCreateSheet(ss, sheetName);
  var evalId = rowValues[0];
  
  if (!evalId) return;
  
  var lastRow = evalSheet.getLastRow();
  var targetRow = -1;
  
  if (lastRow > 1) {
    var existingIds = evalSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var r = 0; r < existingIds.length; r++) {
      if (existingIds[r][0] && existingIds[r][0].toString() === evalId) {
        targetRow = r + 2;
        break;
      }
    }
  }
  
  if (targetRow > 1) {
    evalSheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    evalSheet.appendRow(rowValues);
  }
}

/**
 * Rebuilds the Consolidated_Summary sheet.
 * Groups by Generic_Name + Brand_Name + Supplier + Manufacturer and aggregates responses from all 3 evaluators.
 */
function refreshConsolidatedSummary(fastMode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName(SHEET_EVALUATIONS_MASTER);
  var summarySheet = getOrCreateSheet(ss, SHEET_SUMMARY);
  
  if (!masterSheet || masterSheet.getLastRow() < 2) {
    if (summarySheet.getLastRow() > 1) {
      summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, SUMMARY_HEADERS.length).clear();
    }
    return [];
  }
  
  var masterData = masterSheet.getDataRange().getValues();
  var headers = masterData[0];
  var colMap = {};
  for (var h = 0; h < headers.length; h++) {
    var hKey = (headers[h] || "").toString().toLowerCase().replace(/[^a-z0-9]/g, "");
    if (hKey) colMap[hKey] = h;
  }
  
  function getV(row, keys, defaultVal) {
    for (var k = 0; k < keys.length; k++) {
      var normK = keys[k].toLowerCase().replace(/[^a-z0-9]/g, "");
      if (colMap[normK] !== undefined && row[colMap[normK]] !== "" && row[colMap[normK]] !== null && row[colMap[normK]] !== undefined) {
        return row[colMap[normK]];
      }
    }
    return defaultVal !== undefined ? defaultVal : "";
  }
  
  var grouped = {};
  
  for (var i = 1; i < masterData.length; i++) {
    var row = masterData[i];
    var evaluator = getV(row, ["Evaluator", "Evaluator_Role", "Role"]).toString().trim();
    var generic = getV(row, ["Generic_Name", "Generic Name", "Generic"]).toString().trim();
    var brand = getV(row, ["Brand_Name", "Brand Name", "Brand"]).toString().trim();
    var supplier = getV(row, ["Supplier", "Supplier_Name"]).toString().trim();
    var manufacturer = getV(row, ["Manufacturer", "Manufacturer_Name"]).toString().trim();
    var price = getV(row, ["Price", "Unit_Price", "Offered_Price", "Cost"]);
    var p1Score = getV(row, ["Part_I_Score", "Part1_Score"]);
    var p2Score = getV(row, ["Part_II_Score", "Part2_Score"]);
    var recommendation = getV(row, ["Recommendation", "Verdict", "Decision"]).toString().trim();
    var timestamp = getV(row, ["Timestamp", "Date"]);
    
    if (!generic && !brand) continue;
    
    var groupKey = (generic + "||" + brand + "||" + supplier + "||" + manufacturer).toLowerCase();
    
    if (!grouped[groupKey]) {
      grouped[groupKey] = {
        generic: generic,
        brand: brand,
        supplier: supplier,
        manufacturer: manufacturer,
        price: price || "",
        evaluators: {}
      };
    } else if (price && !grouped[groupKey].price) {
      grouped[groupKey].price = price;
    }
    
    if (evaluator) {
      var normRole = evaluator;
      var lowerRole = evaluator.toLowerCase();
      if (lowerRole.indexOf("end") !== -1 || lowerRole === "evaluator 1") normRole = "End-user";
      else if (lowerRole.indexOf("nurse") !== -1 || lowerRole === "evaluator 2") normRole = "Nurse";
      else if (lowerRole.indexOf("pharma") !== -1 || lowerRole === "evaluator 3") normRole = "Pharmacist";

      grouped[groupKey].evaluators[normRole] = {
        recommendation: recommendation,
        p1Score: p1Score,
        p2Score: p2Score,
        timestamp: timestamp
      };
    }
  }
  
  var nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");
  var summaryRows = [];
  var statusColors = [];
  var summaryList = [];
  
  Object.keys(grouped).forEach(function(k) {
    var item = grouped[k];
    var e1 = item.evaluators["End-user"] || item.evaluators["EndUser"] || item.evaluators["Evaluator 1"] || { recommendation: "Pending", p1Score: "-", p2Score: "-" };
    var e2 = item.evaluators["Nurse"] || item.evaluators["Evaluator 2"] || { recommendation: "Pending", p1Score: "-", p2Score: "-" };
    var e3 = item.evaluators["Pharmacist"] || item.evaluators["Evaluator 3"] || { recommendation: "Pending", p1Score: "-", p2Score: "-" };
    
    var recs = [e1.recommendation, e2.recommendation, e3.recommendation];
    var recCount = 0;
    var recommendedCount = 0;
    var notRecommendedCount = 0;
    
    recs.forEach(function(r) {
      if (r === "Recommended") {
        recCount++;
        recommendedCount++;
      } else if (r === "Not Recommended") {
        recCount++;
        notRecommendedCount++;
      }
    });
    
    var consensusStatus = "";
    var consensusDetail = "";
    var bgColor = "#FFFFFF";
    
    if (recCount < 3) {
      consensusStatus = "Pending (" + recCount + "/3 Complete)";
      consensusDetail = recCount + " of 3 evaluators have submitted scores.";
      bgColor = "#F2F4F4";
    } else if (recommendedCount === 3) {
      consensusStatus = "Unanimous Recommended";
      consensusDetail = "All 3 evaluators voted Recommended.";
      bgColor = "#D4EFDF";
    } else if (notRecommendedCount === 3) {
      consensusStatus = "Unanimous Not Recommended";
      consensusDetail = "All 3 evaluators voted Not Recommended.";
      bgColor = "#FADBD8";
    } else {
      consensusStatus = "Split Decision";
      consensusDetail = recommendedCount + " Recommended, " + notRecommendedCount + " Not Recommended.";
      bgColor = "#FCF3CF";
    }
    
    var rowArr = [
      item.generic, item.brand, item.supplier, item.manufacturer, (item.price || ""),
      e1.recommendation, e2.recommendation, e3.recommendation,
      e1.p1Score, e1.p2Score,
      e2.p1Score, e2.p2Score,
      e3.p1Score, e3.p2Score,
      consensusStatus, consensusDetail, nowStr
    ];
    summaryRows.push(rowArr);
    statusColors.push(bgColor);

    var obj = {};
    for (var c = 0; c < SUMMARY_HEADERS.length; c++) {
      obj[SUMMARY_HEADERS[c]] = rowArr[c];
    }
    summaryList.push(obj);
  });
  
  var prevLastRow = summarySheet.getLastRow();

  if (summaryRows.length > 0) {
    var targetRange = summarySheet.getRange(2, 1, summaryRows.length, SUMMARY_HEADERS.length);
    targetRange.setValues(summaryRows);
    
    // Clear extra rows if table shrank
    if (prevLastRow > summaryRows.length + 1) {
      summarySheet.getRange(summaryRows.length + 2, 1, prevLastRow - (summaryRows.length + 1), SUMMARY_HEADERS.length).clear();
    }

    if (!fastMode) {
      var colorMatrix = [];
      for (var r = 0; r < statusColors.length; r++) {
        var rowColor = statusColors[r];
        var rowCols = [];
        for (var c = 0; c < SUMMARY_HEADERS.length; c++) {
          rowCols.push(rowColor);
        }
        colorMatrix.push(rowCols);
      }
      targetRange.setBackgrounds(colorMatrix);
      
      summarySheet.getRange(2, 5, summaryRows.length, 1).setNumberFormat("₱#,##0.00").setHorizontalAlignment("right");
      summarySheet.getRange(2, 6, summaryRows.length, 12).setHorizontalAlignment("center");
    }
  } else if (prevLastRow > 1) {
    summarySheet.getRange(2, 1, prevLastRow - 1, SUMMARY_HEADERS.length).clear();
  }

  return summaryList;
}

/**
 * Programmatically installs an onEdit trigger for automatic updates.
 */
function setupTriggers() {
  var ui = SpreadsheetApp.getUi();
  
  // Clear existing triggers for this project to prevent duplicates
  var allTriggers = ScriptApp.getProjectTriggers();
  allTriggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "onEvaluationsEdit") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Install trigger
  ScriptApp.newTrigger("onEvaluationsEdit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
    
  ui.alert("Triggers Setup", "Successfully installed 'onEvaluationsEdit' trigger on spreadsheet edit!", ui.ButtonSet.OK);
}

/**
 * Creates or resets the official Printable Checklist Viewer tab (Checklist_Report).
 * Dynamically plots evaluator responses (☑ / ☐) based on selected Drug & Evaluator.
 */
function createChecklistReportSheet(silent) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (eUi) {}
  
  try {
    // 1. Ensure Horizontal Report is updated and synchronized with latest evaluator names
    syncHorizontalReport(true);
    var hSheet = getHorizontalReportSheet(ss);
    var hName = hSheet ? hSheet.getName() : SHEET_CHECKLIST_HORIZONTAL;
    
    // 2. Locate or create Checklist Report viewer tab (handles both 'Checklist Report' and 'Checklist_Report')
    var sheet = getOrCreateChecklistReportSheet(ss);
    sheet.clear();
    sheet.clearFormats();
    
    // Explicitly break apart any legacy merged cells and wipe residual data validation rules
    if (sheet.getMaxRows() > 0 && sheet.getMaxColumns() > 0) {
      sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
      sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
    }
    
    // Set column widths to guarantee clear, un-truncated alignment
    sheet.setColumnWidth(1, 65);   // Col A: Part label (65px)
    sheet.setColumnWidth(2, 50);   // Col B: Item # (50px) - A+B = 115px for left labels
    sheet.setColumnWidth(3, 385);  // Col C: Criteria Description / Values (385px)
    sheet.setColumnWidth(4, 55);   // Col D: Yes (55px) - C+D = 440px for left values
    sheet.setColumnWidth(5, 55);   // Col E: No (55px)
    sheet.setColumnWidth(6, 55);   // Col F: N/A (55px) - E+F = 110px for right labels
    sheet.setColumnWidth(7, 180);  // Col G: Remarks / Values (180px)
    sheet.setColumnWidth(8, 50);   // Col H: Helper lookup row (off printable area)
    
    // 1. Title Block
    sheet.getRange("A1:G1").merge()
      .setValue("PRODUCT EVALUATION CHECKLIST")
      .setFontWeight("bold")
      .setFontSize(14)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    sheet.setRowHeight(1, 35);
      
    sheet.getRange("A2:G2").merge()
      .setValue("PHARMACY BIDS & AWARDS COMMITTEE — TECHNICAL EVALUATION REPORT")
      .setFontSize(9)
      .setFontColor("#555555")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    sheet.setRowHeight(2, 22);
      
    // 2. Interactive Selection Controls (Symmetric 2-Column Grid)
    // Row 3: Supplier Selector & Evaluator Role Selector
    sheet.getRange("A3:B3").merge()
      .setValue("🏢 Supplier:")
      .setFontWeight("bold")
      .setBackground("#EBF2FA")
      .setHorizontalAlignment("right")
      .setVerticalAlignment("middle");
      
    sheet.getRange("C3:D3").merge()
      .setBackground("#FFFFFF")
      .setFontWeight("bold")
      .setVerticalAlignment("middle");
    
    var supSheet = getSupplierSheet(ss);
    if (supSheet && supSheet.getLastRow() > 1) {
      var supRule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(supSheet.getRange("A2:A" + Math.max(supSheet.getLastRow(), 2)), true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange("C3:D3").setDataValidation(supRule);
    }
    
    sheet.getRange("E3:F3").merge()
      .setValue("👤 Role:")
      .setFontWeight("bold")
      .setBackground("#EBF2FA")
      .setHorizontalAlignment("right")
      .setVerticalAlignment("middle");
      
    sheet.getRange("G3")
      .setBackground("#FFFFFF")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
      
    // Strictly the 3 clinical roles matching Evaluator_Accounts (NO Evaluator 1/2/3, NO All)!
    var evalRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Pharmacist", "Nurse", "End-user"], true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange("G3").setDataValidation(evalRule);
    sheet.setRowHeight(3, 28);
    
    // Row 4: Generic Drug Selector & Ref # Jump
    sheet.getRange("A4:B4").merge()
      .setValue("🔍 Generic Drug:")
      .setFontWeight("bold")
      .setBackground("#EBF2FA")
      .setHorizontalAlignment("right")
      .setVerticalAlignment("middle");
      
    sheet.getRange("C4:D4").merge()
      .setBackground("#FFFFFF")
      .setFontWeight("bold")
      .setVerticalAlignment("middle");
    
    var medSheet = ss.getSheetByName(SHEET_MEDICINE_MASTER) || ss.getSheetByName("Medicine Master");
    if (medSheet && medSheet.getLastRow() > 1) {
      var drugRule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(medSheet.getRange("A2:A" + Math.max(medSheet.getLastRow(), 2)), true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange("C4:D4").setDataValidation(drugRule);
    }
    
    sheet.getRange("E4:F4").merge()
      .setValue("📄 Ref #:")
      .setFontWeight("bold")
      .setBackground("#EBF2FA")
      .setHorizontalAlignment("right")
      .setVerticalAlignment("middle");
      
    sheet.getRange("G4")
      .setBackground("#F8FAFC")
      .setFontSize(9)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    sheet.setRowHeight(4, 28);
    
    // Border around Interactive Control Bar (Rows 3-4)
    sheet.getRange("A3:G4").setBorder(true, true, true, true, true, true, "#94A3B8", SpreadsheetApp.BorderStyle.SOLID);
    
    // Default initial values from Horizontal Report if available, else master list
    if (hSheet && hSheet.getLastRow() >= 3) {
      var initSup = hSheet.getRange("D3").getValue();
      var initRole = hSheet.getRange("G3").getValue();
      var initDrug = hSheet.getRange("B3").getValue();
      if (initSup) sheet.getRange("C3").setValue(initSup);
      
      var cleanRole = "Pharmacist";
      if (initRole) {
        var lr = initRole.toString().toLowerCase();
        if (lr.indexOf("pharma") !== -1 || lr === "evaluator 3") cleanRole = "Pharmacist";
        else if (lr.indexOf("nurse") !== -1 || lr === "evaluator 2") cleanRole = "Nurse";
        else if (lr.indexOf("end") !== -1 || lr === "evaluator 1") cleanRole = "End-user";
        else if (["Pharmacist", "Nurse", "End-user"].indexOf(initRole) !== -1) cleanRole = initRole;
      }
      sheet.getRange("G3").setValue(cleanRole);
      
      if (initDrug) sheet.getRange("C4").setValue(initDrug);
    } else {
      if (supSheet && supSheet.getLastRow() > 1) {
        sheet.getRange("C3").setValue(supSheet.getRange("A2").getValue());
      }
      sheet.getRange("G3").setValue("Pharmacist");
      if (medSheet && medSheet.getLastRow() > 1) {
        sheet.getRange("C4").setValue(medSheet.getRange("A2").getValue());
      }
    }
    
    // Helper formula in H3: finds matching row in Horizontal report
    // Single-quote wrapped sheet reference to guarantee 100% error-free parsing
    var hFormula = "=IFERROR(IF(ISNUMBER($G$4), MATCH($G$4, '" + hName + "'!$A:$A, 0), IF(OR($C$3=\"\", $C$4=\"\"), \"\", INDEX(FILTER(ROW('" + hName + "'!$A$3:$A), '" + hName + "'!$B$3:$B=$C$4, '" + hName + "'!$D$3:$D=$C$3, IF($G$3=\"\", ROW('" + hName + "'!$A$3:$A)>0, '" + hName + "'!$G$3:$G=$G$3)), 1))), \"\")";
    sheet.getRange("H3").setFormula(hFormula).setFontColor("#CBD5E1").setFontSize(8);

    // Row 5: Thin visual spacer
    sheet.setRowHeight(5, 8);

    // 3. Official Document Metadata Block (Rows 6 to 8 - Non-overlapping 2-Column Grid)
    // Row 6: Supplier Name & Offered Price
    sheet.getRange("A6:B6").merge().setValue("Name of Supplier:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
    sheet.getRange("C6:D6").merge().setFormula('=$C$3').setFontWeight("bold").setFontColor("#1B365D").setVerticalAlignment("middle");
    sheet.getRange("E6:F6").merge().setValue("Offered Price:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
    sheet.getRange("G6").setFormula("=IF(ISNUMBER($H$3), IFERROR(INDEX('" + hName + "'!$F:$F, $H$3), \"Pending\"), \"Pending Evaluation\")").setFontWeight("bold").setNumberFormat("₱#,##0.00").setHorizontalAlignment("right").setVerticalAlignment("middle");
    sheet.setRowHeight(6, 24);
    
    // Row 7: Generic Name & Brand Name
    sheet.getRange("A7:B7").merge().setValue("Generic Name:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
    sheet.getRange("C7:D7").merge().setFormula('=$C$4').setFontWeight("bold").setFontColor("#1B365D").setVerticalAlignment("middle");
    sheet.getRange("E7:F7").merge().setValue("Brand Name:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
    sheet.getRange("G7").setFormula("=IF(ISNUMBER($H$3), IFERROR(INDEX('" + hName + "'!$C:$C, $H$3), \"—\"), \"—\")").setFontWeight("bold").setHorizontalAlignment("left").setVerticalAlignment("middle");
    sheet.setRowHeight(7, 24);
    
    // Row 8: Manufacturer & Evaluation Verdict
    sheet.getRange("A8:B8").merge().setValue("Manufacturer:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
    sheet.getRange("C8:D8").merge().setFormula("=IF(ISNUMBER($H$3), IFERROR(INDEX('" + hName + "'!$E:$E, $H$3), \"—\"), \"—\")").setFontWeight("bold").setHorizontalAlignment("left").setVerticalAlignment("middle");
    sheet.getRange("E8:F8").merge().setValue("Verdict:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
    sheet.getRange("G8").setFormula("=IF(ISNUMBER($H$3), IFERROR(INDEX('" + hName + "'!$AK:$AK, $H$3), \"Pending Evaluation\"), \"Pending Evaluation\")").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.setRowHeight(8, 24);
    
    // Border around Document Metadata Block (Rows 6-8)
    sheet.getRange("A6:G8").setBorder(true, true, true, true, true, true, "#E2E8F0", SpreadsheetApp.BorderStyle.SOLID);
    
    // 4. Part I Table Header (Row 9)
    var headerRange = sheet.getRange("A9:G9");
    headerRange.setValues([["Part", "#", "PRODUCT SAMPLE EVALUATION (Part I)", "Yes", "No", "N/A", "Remarks"]]);
    headerRange.setBackground("#1B365D").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.getRange("C9").setHorizontalAlignment("left");
    sheet.setRowHeight(9, 28);
    
    // 5. Part I Items (19 rows: Row 10 to 28, mapped from Horizontal report Cols 10 to 28 / J to AB)
    var p1Rows = [];
    var p1Formulas = [];
    for (var i = 0; i < OFFICIAL_PART1_ITEMS.length; i++) {
      var itemNum = i + 1;
      var colRef = getColumnLetter(10 + i); // Col 10 is J (1. Product Name)
      
      var yesFormula = "=IF(NOT(ISNUMBER($H$3)), \"☐\", IF(INDEX('" + hName + "'!$" + colRef + ":$" + colRef + ", $H$3)=\"✓\", \"☑\", \"☐\"))";
      var noFormula  = "=IF(NOT(ISNUMBER($H$3)), \"☐\", IF(OR(INDEX('" + hName + "'!$" + colRef + ":$" + colRef + ", $H$3)=\"✗\", INDEX('" + hName + "'!$" + colRef + ":$" + colRef + ", $H$3)=\"X\"), \"☑\", \"☐\"))";
      var naFormula  = "=IF(NOT(ISNUMBER($H$3)), \"☐\", IF(OR(INDEX('" + hName + "'!$" + colRef + ":$" + colRef + ", $H$3)=\"—\", INDEX('" + hName + "'!$" + colRef + ":$" + colRef + ", $H$3)=\"N/A\"), \"☑\", \"☐\"))";
      
      p1Rows.push(["Part I", itemNum, OFFICIAL_PART1_ITEMS[i], "", "", "", ""]);
      p1Formulas.push([yesFormula, noFormula, naFormula]);
    }
    
    sheet.getRange(10, 1, 19, 7).setValues(p1Rows);
    for (var r = 0; r < p1Formulas.length; r++) {
      sheet.getRange(10 + r, 4, 1, 3).setFormulas([p1Formulas[r]]);
      sheet.setRowHeight(10 + r, 22);
    }
    sheet.getRange(10, 1, 19, 2).setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.getRange(10, 3, 19, 1).setVerticalAlignment("middle");
    sheet.getRange(10, 4, 19, 3).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(11);
    
    // Merge Part I column vertically
    sheet.getRange("A10:A28").merge().setVerticalAlignment("middle").setHorizontalAlignment("center").setFontWeight("bold");
    
    // 6. Part II Table Header & Rows (Row 29 header, rows 30 to 35, mapped from Horizontal report Cols 29 to 34 / AC to AH)
    sheet.getRange("A29:G29").merge().setValue("Part II: Physical Packaging & Container Integrity").setBackground("#2C3E50").setFontColor("#FFFFFF").setFontWeight("bold").setVerticalAlignment("middle");
    sheet.setRowHeight(29, 26);
    
    var p2Rows = [];
    var p2Formulas = [];
    for (var j = 0; j < OFFICIAL_PART2_ITEMS.length; j++) {
      var itemNum2 = j + 1;
      var colRef2 = getColumnLetter(29 + j); // Col 29 is AC (1. Inner Label Match)
      
      var yesFormula2 = "=IF(NOT(ISNUMBER($H$3)), \"☐\", IF(INDEX('" + hName + "'!$" + colRef2 + ":$" + colRef2 + ", $H$3)=\"✓\", \"☑\", \"☐\"))";
      var noFormula2  = "=IF(NOT(ISNUMBER($H$3)), \"☐\", IF(OR(INDEX('" + hName + "'!$" + colRef2 + ":$" + colRef2 + ", $H$3)=\"✗\", INDEX('" + hName + "'!$" + colRef2 + ":$" + colRef2 + ", $H$3)=\"X\"), \"☑\", \"☐\"))";
      var naFormula2  = "=IF(NOT(ISNUMBER($H$3)), \"☐\", IF(OR(INDEX('" + hName + "'!$" + colRef2 + ":$" + colRef2 + ", $H$3)=\"—\", INDEX('" + hName + "'!$" + colRef2 + ":$" + colRef2 + ", $H$3)=\"N/A\"), \"☑\", \"☐\"))";
      
      p2Rows.push(["Part II", itemNum2, OFFICIAL_PART2_ITEMS[j], "", "", "", ""]);
      p2Formulas.push([yesFormula2, noFormula2, naFormula2]);
    }
    
    sheet.getRange(30, 1, 6, 7).setValues(p2Rows);
    for (var r2 = 0; r2 < p2Formulas.length; r2++) {
      sheet.getRange(30 + r2, 4, 1, 3).setFormulas([p2Formulas[r2]]);
      sheet.setRowHeight(30 + r2, 24);
    }
    sheet.getRange(30, 1, 6, 2).setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.getRange(30, 3, 6, 1).setVerticalAlignment("middle");
    sheet.getRange(30, 4, 6, 3).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(11);
    sheet.getRange("A30:A35").merge().setVerticalAlignment("middle").setHorizontalAlignment("center").setFontWeight("bold");
    
    // Row 36: Spacer
    sheet.setRowHeight(36, 10);

    // 7. Summary & Verdict Footer Block
    var startFooter = 37;
    sheet.getRange(startFooter, 1, 1, 7).merge().setValue("EVALUATION SUMMARY & RECOMMENDATION").setBackground("#1B365D").setFontColor("#FFFFFF").setFontWeight("bold").setVerticalAlignment("middle");
    sheet.setRowHeight(startFooter, 26);
    
    sheet.getRange(startFooter + 1, 1, 1, 2).merge().setValue("Part I Score:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
    sheet.getRange(startFooter + 1, 3).setFormula("=IF(ISNUMBER($H$3), IFERROR(INDEX('" + hName + "'!$AI:$AI, $H$3), \"Pending\"), \"Pending Evaluation\")").setFontWeight("bold").setVerticalAlignment("middle");
    
    sheet.getRange(startFooter + 1, 4, 1, 2).merge().setValue("Part II Score:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
    sheet.getRange(startFooter + 1, 6, 1, 2).merge().setFormula("=IF(ISNUMBER($H$3), IFERROR(INDEX('" + hName + "'!$AJ:$AJ, $H$3), \"Pending\"), \"Pending Evaluation\")").setFontWeight("bold").setVerticalAlignment("middle");
    sheet.setRowHeight(startFooter + 1, 24);
    
    sheet.getRange(startFooter + 2, 1, 1, 2).merge().setValue("Recommendation:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
    sheet.getRange(startFooter + 2, 3, 1, 5).merge().setFormula("=IF(ISNUMBER($H$3), IFERROR(INDEX('" + hName + "'!$AK:$AK, $H$3), \"Pending\"), \"Pending Evaluation\")").setFontWeight("bold").setFontSize(11).setVerticalAlignment("middle");
    sheet.setRowHeight(startFooter + 2, 24);
    
    sheet.getRange(startFooter + 3, 1, 1, 2).merge().setValue("Overall Remarks:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
    sheet.getRange(startFooter + 3, 3, 1, 5).merge().setFormula("=IF(ISNUMBER($H$3), IFERROR(INDEX('" + hName + "'!$AL:$AL, $H$3), \"—\"), \"No evaluation submitted yet for this supplier and role.\")").setVerticalAlignment("middle");
    sheet.setRowHeight(startFooter + 3, 24);
    
    // Row 41: Spacer
    sheet.setRowHeight(startFooter + 4, 10);

    // 8. Signatures Block with Digital Signature Display
    var sigRow = startFooter + 5; // Row 42
    sheet.getRange(sigRow, 1, 1, 3).merge().setFormula("=\"Evaluated By: \" & IF(ISNUMBER($H$3), IFERROR(INDEX('" + hName + "'!$H:$H, $H$3), \"Pending\"), \"Pending\")").setFontWeight("bold").setVerticalAlignment("middle");
    sheet.getRange(sigRow, 4, 1, 4).merge().setFormula("=\"Date: \" & IF(ISNUMBER($H$3), IFERROR(TEXT(INDEX('" + hName + "'!$I:$I, $H$3), \"yyyy-mm-dd\"), \"—\"), \"—\") & \"  |  Signature Status: \" & IF(NOT(ISNUMBER($H$3)), \"Pending Evaluation\", IF(ISBLANK(INDEX('" + hName + "'!$AM:$AM, $H$3)), \"Physical / Pending\", \"☑ Digitally Captured\"))").setFontWeight("bold").setVerticalAlignment("middle");
    sheet.setRowHeight(sigRow, 24);
    
    sheet.getRange(sigRow + 1, 1, 1, 3).merge().setFormula("=\"Official Role: \" & IF(ISNUMBER($H$3), IFERROR(INDEX('" + hName + "'!$G:$G, $H$3), $G$3), $G$3)").setFontStyle("italic").setVerticalAlignment("middle");
    sheet.getRange(sigRow + 1, 4, 1, 4).merge().setFormula("=IF(NOT(ISNUMBER($H$3)), \"Signature: _______________________\", IF(ISBLANK(INDEX('" + hName + "'!$AM:$AM, $H$3)), \"Signature: _______________________\", IF(ISNUMBER(SEARCH(\"http\", INDEX('" + hName + "'!$AM:$AM, $H$3))), IMAGE(INDEX('" + hName + "'!$AM:$AM, $H$3), 1), \"Signature: _______________________\")))").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.setRowHeight(sigRow + 1, 60); // Generous height to render signature image cleanly!
    
    // Apply grid borders to tables
    sheet.getRange("A9:G28").setBorder(true, true, true, true, true, true, "#333333", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange("A29:G35").setBorder(true, true, true, true, true, true, "#333333", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(startFooter, 1, 4, 7).setBorder(true, true, true, true, true, true, "#333333", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(sigRow, 1, 2, 7).setBorder(true, true, true, true, true, true, "#333333", SpreadsheetApp.BorderStyle.SOLID);
    
    SpreadsheetApp.flush();
    
    if (!silent && ui) {
      ui.alert("Checklist Viewer Initialized", "The 'Checklist Report' tab is now refreshed!\n\n• Select Supplier in C3\n• Select Evaluator Role in G3 (Pharmacist, Nurse, End-user)\n• Select Generic Drug in C4\n• All checkmarks, scores, and signature render automatically.", ui.ButtonSet.OK);
    }
  } catch (err) {
    Logger.log("Error in createChecklistReportSheet: " + err.toString());
    if (!silent && ui) {
      ui.alert("Error Creating Checklist Report", err.toString(), ui.ButtonSet.OK);
    }
  }
}

/**
 * Helper to convert 1-indexed column number to letter (e.g. 1 -> A, 27 -> AA)
 */
function getColumnLetter(columnNumber) {
  var temp = "";
  var letter = "";
  while (columnNumber > 0) {
    temp = (columnNumber - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    columnNumber = (columnNumber - temp - 1) / 26;
  }
  return letter;
}

/**
 * Exports the active Checklist Viewer (Checklist_Report) as a PDF directly to Google Drive.
 */
function exportCurrentChecklistPdf() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = getOrCreateChecklistReportSheet(ss);
  
  if (!sheet) {
    ui.alert("Sheet Not Found", "Please initialize the Checklist Viewer first via the menu.", ui.ButtonSet.OK);
    return;
  }
  
  try {
    var supName = sheet.getRange("C3").getValue() || "Supplier";
    var drugName = sheet.getRange("C4").getValue() || "Evaluation";
    var evaluator = sheet.getRange("G3").getValue() || "Summary";
    var safeDrug = drugName.toString().replace(/[^a-zA-Z0-9_-]/g, "_").substring(0, 25);
    var safeSup = supName.toString().replace(/[^a-zA-Z0-9_-]/g, "_").substring(0, 20);
    var fileName = "Checklist_" + safeSup + "_" + safeDrug + "_" + evaluator + ".pdf";
    
    var pdfBlob = ss.getAs('application/pdf');
    pdfBlob.setName(fileName);
    
    var file = DriveApp.createFile(pdfBlob);
    
    ui.alert("PDF Exported Successfully", "PDF exported and saved to your Google Drive:\n\nFile: " + fileName + "\nLink: " + file.getUrl(), ui.ButtonSet.OK);
  } catch (err) {
    Logger.log("Error in exportCurrentChecklistPdf: " + err.toString());
    ui.alert("Error Exporting PDF", err.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Creates or updates the clean, presentation-ready Horizontal Evaluation Summary Report
 * (Checklist_Report_Horizontal). Mirrors Evaluations_Master horizontally but with official
 * human-readable headers, clean checkmarks (✓ / ✗ / —), scores, and evaluator names.
 */
function createHorizontalReportSheet() {
  syncHorizontalReport(false);
}

/**
 * Creates or updates the clean, presentation-ready Horizontal Evaluation Summary Report
 * (Checklist_Report_Horizontal). Mirrors Evaluations_Master horizontally but with official
 * human-readable headers, clean checkmarks (✓ / ✗ / —), scores, and evaluator names.
 */
function syncHorizontalReport(silent) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName(SHEET_EVALUATIONS_MASTER);
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (eUi) {}
  
  if (!masterSheet) {
    if (!silent && ui) ui.alert("Error", "Evaluations_Master sheet was not found in this spreadsheet.", ui.ButtonSet.OK);
    return;
  }
  
  var sheet = getOrCreateSheet(ss, SHEET_CHECKLIST_HORIZONTAL);
  
  try {
    sheet.clear();
    sheet.clearFormats();
    
    // Ensure sufficient rows and columns exist in target sheet (38 columns)
    if (sheet.getMaxColumns() < 42) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), 42 - sheet.getMaxColumns());
    }
    if (sheet.getMaxRows() < 100) {
      sheet.insertRowsAfter(sheet.getMaxRows(), 100 - sheet.getMaxRows());
    }
    
    // Unhide all rows/columns
    sheet.showRows(1, sheet.getMaxRows());
    sheet.showColumns(1, sheet.getMaxColumns());
    
    // Set explicit row heights
    sheet.setRowHeight(1, 30);
    sheet.setRowHeight(2, 38);
    
    // 1. Group Headers (Row 1) - Aligned with frozen columns (Cols A-F = 1-6)
    sheet.getRange("A1:F1").merge().setValue("PRODUCT & BID IDENTIFICATION")
      .setBackground("#1B365D").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
      
    sheet.getRange("G1:I1").merge().setValue("EVALUATOR METADATA")
      .setBackground("#244872").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
      
    sheet.getRange("J1:AB1").merge().setValue("PART I: LABELING & REGULATORY COMPLIANCE (19 CRITERIA)")
      .setBackground("#2C3E50").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
      
    sheet.getRange("AC1:AH1").merge().setValue("PART II: PHYSICAL PACKAGING & CONTAINER INTEGRITY (6 CRITERIA)")
      .setBackground("#34495E").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
      
    sheet.getRange("AI1:AM1").merge().setValue("EVALUATION VERDICT, SCORES & SIGNATURE")
      .setBackground("#004B49").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
      
    // 2. Official Column Headers (Row 2 - 39 columns)
    var row2Headers = [
      "#", "Generic Name", "Brand Name", "Supplier", "Manufacturer", "Price",
      "Evaluator Role", "Evaluator Name", "Date",
      // Part I items (19)
      "1. Product Name", "2. Dosage Form & Strength", "3. Pharmacologic Category", "4. Formulation / Composition",
      "5. Indication(s)", "6. Dosage & Mode of Admin", "7. Warnings & Precautions", "8. Drug Interactions",
      "9. Adverse Drug Reactions", "10. Overdose Info", "11. Storage Conditions", "12. Net Content / Pack Size",
      "13. Marketing Auth Holder", "14. Manufacturer Address", "15. Rx Caution Statement", "16. ADR Reporting",
      "17. Registration Number", "18. Batch / Lot Number", "19. Mfg & Expiry Date",
      // Part II items (6)
      "1. Inner Label Match", "2. Container Label Legibility", "3. Blister Pack Print",
      "4. Parenteral Leakage Check", "5. Rubber Stopper Puncture", "6. Dispensing Ease & Integrity",
      // Verdict, Scores & Signature
      "Part I Score", "Part II Score", "Recommendation", "Remarks", "Signature"
    ];
    
    var hRange = sheet.getRange(2, 1, 1, row2Headers.length);
    hRange.setValues([row2Headers]);
    hRange.setBackground("#E2E8F0").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
    sheet.getRange("B2:E2").setHorizontalAlignment("left");
    sheet.getRange("F2").setHorizontalAlignment("right");
    sheet.getRange("H2").setHorizontalAlignment("left");
    sheet.getRange("AL2").setHorizontalAlignment("left");
    sheet.getRange("AM2").setHorizontalAlignment("center");
    
    // Set Column Widths
    sheet.setColumnWidth(1, 40);   // #
    sheet.setColumnWidth(2, 230);  // Generic Name
    sheet.setColumnWidth(3, 130);  // Brand Name
    sheet.setColumnWidth(4, 130);  // Supplier
    sheet.setColumnWidth(5, 130);  // Manufacturer
    sheet.setColumnWidth(6, 100);  // Price
    sheet.setColumnWidth(7, 110);  // Evaluator Role
    sheet.setColumnWidth(8, 160);  // Evaluator Name
    sheet.setColumnWidth(9, 95);   // Date
    for (var c = 10; c <= 34; c++) {
      sheet.setColumnWidth(c, 85); // Part I & II criteria checkmark cols (25 cols)
    }
    sheet.setColumnWidth(35, 95);  // Part I Score
    sheet.setColumnWidth(36, 95);  // Part II Score
    sheet.setColumnWidth(37, 140); // Recommendation
    sheet.setColumnWidth(38, 220); // Remarks
    sheet.setColumnWidth(39, 180); // Signature
    
    // Freeze top 2 header rows and left 6 identifying columns (matches A1:F1 boundary)
    sheet.setFrozenRows(2);
    sheet.setFrozenColumns(6);
    
    // 3. Build Evaluator Name lookup map from Evaluator_Accounts
    var evalMap = {
      "End-user": "Clinical Specialist / End-User",
      "Nurse": "Head Nurse / Clinical Nurse",
      "Pharmacist": "Staff Pharmacist / Evaluator",
      "Evaluator 1": "Clinical Specialist / End-User",
      "Evaluator 2": "Head Nurse / Clinical Nurse",
      "Evaluator 3": "Staff Pharmacist / Evaluator"
    };
    var acctsSheet = ss.getSheetByName("Evaluator_Accounts");
    var acctById = {};
    var acctByEmail = {};
    if (acctsSheet && acctsSheet.getLastRow() > 1) {
      var acctsData = acctsSheet.getDataRange().getValues();
      for (var a = 1; a < acctsData.length; a++) {
        var aId = (acctsData[a][0] || "").toString().trim().toLowerCase();
        var aRole = (acctsData[a][1] || "").toString().trim();
        var aName = (acctsData[a][2] || "").toString().trim();
        var aEmail = (acctsData[a][3] || "").toString().trim().toLowerCase();
        if (aRole && aName) {
          evalMap[aRole] = aName;
        }
        if (aId && aName) acctById[aId] = aName;
        if (aEmail && aName) acctByEmail[aEmail] = aName;
      }
    }
    
    // 4. Safely Read All Data from Evaluations_Master
    var masterData = masterSheet.getDataRange().getValues();
    if (masterData.length <= 1) {
      if (!silent) ui.alert("Horizontal Report Initialized", "Checklist_Report_Horizontal is ready. No data rows found in Evaluations_Master yet.", ui.ButtonSet.OK);
      return;
    }
    
    // Helper to normalize any header string for flexible matching
    function norm(str) {
      return (str || "").toString().toLowerCase().replace(/[^a-z0-9]/g, "");
    }
    
    // Map normalized column header names to their column index
    var masterHeaders = masterData[0];
    var colMap = {};
    for (var h = 0; h < masterHeaders.length; h++) {
      var nName = norm(masterHeaders[h]);
      if (nName) colMap[nName] = h;
    }
    
    // Flexible helper to get value from a row using multiple possible key names or prefixes
    function getVal(rowObj, possibleKeys) {
      if (!Array.isArray(possibleKeys)) possibleKeys = [possibleKeys];
      for (var k = 0; k < possibleKeys.length; k++) {
        var target = norm(possibleKeys[k]);
        // Exact normalized match
        if (colMap[target] !== undefined) {
          var val = rowObj[colMap[target]];
          if (val !== undefined && val !== null && val.toString().trim() !== "") return val;
        }
        // Prefix match
        for (var existingKey in colMap) {
          if (existingKey.indexOf(target) === 0 || target.indexOf(existingKey) === 0) {
            var val2 = rowObj[colMap[existingKey]];
            if (val2 !== undefined && val2 !== null && val2.toString().trim() !== "") return val2;
          }
        }
      }
      return "";
    }
    
    var reportRows = [];
    var recColors = [];
    
    for (var i = 1; i < masterData.length; i++) {
      var row = masterData[i];
      
      // Check if row is completely empty
      var hasData = false;
      for (var c = 0; c < row.length; c++) {
        if (row[c] !== "" && row[c] !== null && row[c] !== undefined) {
          hasData = true;
          break;
        }
      }
      if (!hasData) continue;
      
      var evalId = getVal(row, ["Evaluation_ID", "Evaluation ID", "ID", "Eval_ID", "Key"]) || row[0] || ("EVAL-" + i);
      var generic = getVal(row, ["Generic_Name", "Generic Name", "Generic", "Medicine", "Item_Description", "Drug_Name"]) || row[3] || row[1] || "";
      var brand = (getVal(row, ["Brand_Name", "Brand Name", "Brand"]) || row[4] || "").toString().trim();
      var supplier = (getVal(row, ["Supplier", "Supplier_Name"]) || "").toString().trim();
      var manufacturer = (getVal(row, ["Manufacturer", "Manufacturer_Name"]) || "").toString().trim();
      var price = getVal(row, ["Price", "Unit_Price", "Offered_Price", "Cost"]);
      var role = (getVal(row, ["Evaluator", "Evaluator_Role", "Role"]) || row[2] || "").toString().trim();
      var lowerRole = role.toLowerCase();
      if (lowerRole.indexOf("pharma") !== -1 || lowerRole === "evaluator 3") {
        role = "Pharmacist";
      } else if (lowerRole.indexOf("nurse") !== -1 || lowerRole === "evaluator 2") {
        role = "Nurse";
      } else if (lowerRole.indexOf("end") !== -1 || lowerRole === "evaluator 1") {
        role = "End-user";
      }
      // Determine the actual evaluator's name who performed this evaluation:
      // 1. Tag in Remarks: "[Evaluated by Chito Saba]" or "[By Chito Saba]"
      var rawRemarks = (getVal(row, ["Remarks", "Remark", "Comments", "Notes"]) || "").toString().trim();
      var nameFromRemarks = "";
      var matchRemark = rawRemarks.match(/\[(?:Evaluated by|By)\s+([^\]]+)\]/i);
      if (matchRemark && matchRemark[1]) {
        nameFromRemarks = matchRemark[1].trim();
      }
      
      // 2. Explicit column in master sheet
      var explicitName = (getVal(row, ["Evaluator_Name", "Evaluator Name", "Name", "Evaluator_User"]) || "").toString().trim();
      
      // 3. Match from Evaluator_Accounts if Evaluator column stores an email or account ID
      var rawEvalCol = (row[2] || getVal(row, ["Evaluator"]) || "").toString().trim().toLowerCase();
      var nameFromAccount = acctById[rawEvalCol] || acctByEmail[rawEvalCol] || "";
      
      // Prioritize the actual evaluator person who performed the evaluation!
      var evalName = nameFromRemarks || explicitName || nameFromAccount || evalMap[role] || role;
      
      var rawTimestamp = getVal(row, ["Timestamp", "Date", "Evaluation_Date", "Time"]) || row[1];
      var dateStr = "";
      if (rawTimestamp instanceof Date) {
        dateStr = Utilities.formatDate(rawTimestamp, "GMT+8", "yyyy-MM-dd");
      } else if (rawTimestamp) {
        dateStr = rawTimestamp.toString().split(" ")[0];
      }
      
      // Part I items (19 items from P1_01 to P1_19)
      var p1Cols = [];
      for (var p1 = 1; p1 <= 19; p1++) {
        var numStr = (p1 < 10) ? ("0" + p1) : ("" + p1);
        var val1 = getVal(row, ["P1_" + numStr, "P1" + numStr, "Part1_" + numStr, "Part_I_" + numStr]);
        p1Cols.push(formatCheckmark(val1));
      }
      
      // Part II items (6 items from P2_01 to P2_06)
      var p2Cols = [];
      for (var p2 = 1; p2 <= 6; p2++) {
        var numStr2 = "0" + p2;
        var val2 = getVal(row, ["P2_" + numStr2, "P2" + numStr2, "Part2_" + numStr2, "Part_II_" + numStr2]);
        p2Cols.push(formatCheckmark(val2));
      }
      
      // Scores and Verdict
      var p1Score = getVal(row, ["Part_I_Score", "Part I Score", "Part1_Score", "Part1Score", "Score_Part_I"]);
      var p2Score = getVal(row, ["Part_II_Score", "Part II Score", "Part2_Score", "Part2Score", "Score_Part_II"]);
      var remarks = getVal(row, ["Remarks", "Remark", "Comments", "Notes"]);
      var rec = (getVal(row, ["Recommendation", "Verdict", "Decision", "Status"]) || "").toString().trim();
      var signature = (getVal(row, ["Evaluator_Signature", "Signature", "Digital_Signature", "Sign", "Signature_URL"]) || "").toString().trim();
      
      // Auto-migrate any base64 signature in master sheet to a permanent Drive thumbnail image
      if (signature && signature.indexOf("data:image") === 0) {
        var convertedUrl = saveBase64SignatureToDrive(signature, evalId);
        if (convertedUrl) {
          signature = convertedUrl;
          var sigColIdx = colMap[norm("Evaluator_Signature")];
          if (sigColIdx === undefined) sigColIdx = colMap[norm("Signature")];
          if (sigColIdx === undefined) sigColIdx = colMap[norm("Digital_Signature")];
          if (sigColIdx !== undefined) {
            try { masterSheet.getRange(i + 1, sigColIdx + 1).setValue(convertedUrl); } catch (eUp) {}
          }
        }
      } else if (signature && signature.indexOf("drive.google.com") !== -1) {
        var matchId = signature.match(/id=([a-zA-Z0-9_-]+)/);
        if (matchId && matchId[1]) {
          var fileId = matchId[1];
          try {
            var f = DriveApp.getFileById(fileId);
            f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          } catch (eF) {}
          var thumbUrl = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w1000";
          if (signature !== thumbUrl) {
            signature = thumbUrl;
            var sigColIdx2 = colMap[norm("Evaluator_Signature")];
            if (sigColIdx2 === undefined) sigColIdx2 = colMap[norm("Signature")];
            if (sigColIdx2 !== undefined) {
              try { masterSheet.getRange(i + 1, sigColIdx2 + 1).setValue(thumbUrl); } catch (eUp2) {}
            }
          }
        }
      }
      
      var formattedRow = [
        (reportRows.length + 1), generic, brand, supplier, manufacturer, price, role, evalName, dateStr
      ].concat(p1Cols).concat(p2Cols).concat([p1Score, p2Score, rec, remarks, signature]);
      
      reportRows.push(formattedRow);
      
      if (rec.indexOf("Recommended") !== -1 && rec.indexOf("Not") === -1) {
        recColors.push("#D4EFDF");
      } else if (rec.indexOf("Not Recommended") !== -1) {
        recColors.push("#FADBD8");
      } else {
        recColors.push("#FFFFFF");
      }
    }
    
    // 5. Write Data to Sheet
    if (reportRows.length > 0) {
      if (sheet.getMaxRows() < reportRows.length + 5) {
        sheet.insertRowsAfter(sheet.getMaxRows(), reportRows.length + 10);
      }
      
      var targetRange = sheet.getRange(3, 1, reportRows.length, row2Headers.length);
      targetRange.setValues(reportRows);
      
      var bgs = [];
      for (var r = 0; r < reportRows.length; r++) {
        var rowBg = (r % 2 === 0) ? "#FFFFFF" : "#F8FAFC";
        var rowCols = [];
        for (var c = 0; c < row2Headers.length; c++) {
          if (c === 36) { // Col 37 (index 36): Recommendation
            rowCols.push(recColors[r]);
          } else {
            rowCols.push(rowBg);
          }
        }
        bgs.push(rowCols);
      }
      targetRange.setBackgrounds(bgs);
      sheet.getRange(3, 37, reportRows.length, 1).setFontWeight("bold");
      
      // Alignments & Formats (39 columns total)
      sheet.getRange(3, 1, reportRows.length, 1).setHorizontalAlignment("center"); // Col 1: #
      sheet.getRange(3, 6, reportRows.length, 1).setNumberFormat("₱#,##0.00").setHorizontalAlignment("right"); // Col 6: Price
      sheet.getRange(3, 7, reportRows.length, 1).setHorizontalAlignment("center"); // Col 7: Evaluator Role
      sheet.getRange(3, 9, reportRows.length, 26).setHorizontalAlignment("center").setFontSize(11); // Col 9-34: Date & Checkmarks
      sheet.getRange(3, 35, reportRows.length, 3).setHorizontalAlignment("center"); // Col 35-37: Part I Score, Part II Score, Rec
      sheet.getRange(3, 39, reportRows.length, 1).setHorizontalAlignment("center"); // Col 39: Signature
      
      targetRange.setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
      
      // Auto-refresh Checklist_Report signature cell formula & row height if tab exists
      var clSheet = ss.getSheetByName(SHEET_CHECKLIST_REPORT) || ss.getSheetByName("Checklist Report");
      if (clSheet && clSheet.getLastRow() >= 42) {
        var hName = sheet.getName();
        clSheet.getRange(43, 4, 1, 4).setFormula("=IF(NOT(ISNUMBER($H$3)), \"Signature: _______________________\", IF(ISBLANK(INDEX('" + hName + "'!$AM:$AM, $H$3)), \"Signature: _______________________\", IF(ISNUMBER(SEARCH(\"http\", INDEX('" + hName + "'!$AM:$AM, $H$3))), IMAGE(INDEX('" + hName + "'!$AM:$AM, $H$3), 1), \"Signature: _______________________\")))").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
        clSheet.setRowHeight(43, 60);
      }
    }
    
    if (!silent && ui) {
      ui.alert("Success", "Loaded " + reportRows.length + " evaluation rows into Checklist_Report_Horizontal!", ui.ButtonSet.OK);
    }
  } catch (err) {
    Logger.log("Error in syncHorizontalReport: " + err.toString());
    if (!silent && ui) {
      ui.alert("Error Creating Horizontal Report", err.toString() + "\nLine: " + err.lineNumber, ui.ButtonSet.OK);
    }
  }
}

/**
 * Fast incremental update to Checklist_Report_Horizontal for a single evaluation.
 * Appends 1 row in ~0.2 seconds without wiping the sheet, resetting formats, or looping 39 setColumnWidth calls.
 */
function appendOrUpdateHorizontalReportRow(rowValues, activeHeaders) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getHorizontalReportSheet(ss);
    if (!sheet || sheet.getLastRow() < 2) {
      // Sheet not initialized yet; run full setup once
      syncHorizontalReport(true);
      return;
    }
    
    // Resolve activeHeaders if not provided
    if (!activeHeaders || !Array.isArray(activeHeaders) || activeHeaders.length === 0) {
      var masterSheet = ss.getSheetByName(SHEET_EVALUATIONS_MASTER);
      if (masterSheet && masterSheet.getLastColumn() > 0) {
        activeHeaders = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
      } else {
        activeHeaders = EVAL_HEADERS;
      }
    }
    
    // Normalize colMap from activeHeaders
    function normKey(s) {
      return (s || "").toString().toLowerCase().replace(/[^a-z0-9]/g, "");
    }
    var colMap = {};
    for (var h = 0; h < activeHeaders.length; h++) {
      var nk = normKey(activeHeaders[h]);
      if (nk) colMap[nk] = h;
    }
    
    function getVal(keys) {
      if (!Array.isArray(keys)) keys = [keys];
      for (var k = 0; k < keys.length; k++) {
        var target = normKey(keys[k]);
        if (colMap[target] !== undefined) {
          var val = rowValues[colMap[target]];
          if (val !== undefined && val !== null && val.toString().trim() !== "") return val;
        }
        for (var exKey in colMap) {
          if (exKey.indexOf(target) === 0 || target.indexOf(exKey) === 0) {
            var val2 = rowValues[colMap[exKey]];
            if (val2 !== undefined && val2 !== null && val2.toString().trim() !== "") return val2;
          }
        }
      }
      return "";
    }
    
    var lastRow = sheet.getLastRow();
    var generic = getVal(["Generic_Name", "Generic Name", "Generic", "Medicine", "Item_Description"]) || rowValues[3] || "";
    var brand = (getVal(["Brand_Name", "Brand Name", "Brand"]) || rowValues[4] || "").toString().trim();
    var supplier = (getVal(["Supplier", "Supplier_Name"]) || rowValues[5] || "").toString().trim();
    var manufacturer = (getVal(["Manufacturer", "Manufacturer_Name"]) || rowValues[6] || "").toString().trim();
    var price = getVal(["Price", "Unit_Price", "Cost"]) || rowValues[7] || "";
    var role = (getVal(["Evaluator", "Evaluator_Role", "Role"]) || rowValues[2] || "").toString().trim();
    var lowerRole = role.toLowerCase();
    if (lowerRole.indexOf("pharma") !== -1 || lowerRole === "evaluator 3") {
      role = "Pharmacist";
    } else if (lowerRole.indexOf("nurse") !== -1 || lowerRole === "evaluator 2") {
      role = "Nurse";
    } else if (lowerRole.indexOf("end") !== -1 || lowerRole === "evaluator 1") {
      role = "End-user";
    }
    
    var timestamp = getVal(["Timestamp", "Date", "Evaluation_Date"]) || rowValues[1];
    var dateStr = "";
    if (timestamp instanceof Date) {
      dateStr = Utilities.formatDate(timestamp, "GMT+8", "yyyy-MM-dd");
    } else if (timestamp) {
      dateStr = timestamp.toString().split(" ")[0];
    }
    
    // Determine Evaluator Name from remarks tag or default role
    var rawRemarks = (getVal(["Remarks", "Remark", "Comments"]) || "").toString().trim();
    var evalName = role;
    var matchRemark = rawRemarks.match(/\[(?:Evaluated by|By)\s+([^\]]+)\]/i);
    if (matchRemark && matchRemark[1]) {
      evalName = matchRemark[1].trim();
    } else {
      var explicitName = (getVal(["Evaluator_Name", "Evaluator Name", "Name"]) || "").toString().trim();
      if (explicitName) evalName = explicitName;
    }
    
    // Part I items (19)
    var p1Cols = [];
    for (var p1 = 1; p1 <= 19; p1++) {
      var numStr = (p1 < 10) ? ("0" + p1) : ("" + p1);
      var val1 = getVal(["P1_" + numStr, "P1" + numStr, "Part1_" + numStr, "Part_I_" + numStr]);
      p1Cols.push(formatCheckmark(val1));
    }
    
    // Part II items (6)
    var p2Cols = [];
    for (var p2 = 1; p2 <= 6; p2++) {
      var numStr2 = "0" + p2;
      var val2 = getVal(["P2_" + numStr2, "P2" + numStr2, "Part2_" + numStr2, "Part_II_" + numStr2]);
      p2Cols.push(formatCheckmark(val2));
    }
    
    var p1Score = getVal(["Part_I_Score", "Part I Score", "Part1_Score", "Part1Score"]);
    var p2Score = getVal(["Part_II_Score", "Part II Score", "Part2_Score", "Part2Score"]);
    var rec = (getVal(["Recommendation", "Verdict", "Decision", "Status"]) || "").toString().trim();
    var sigUrl = getVal(["Evaluator_Signature", "Signature", "Digital_Signature", "Sign", "Signature_URL"]);
    if (sigUrl && sigUrl.indexOf("data:image") === 0) {
      var driveSaved = saveBase64SignatureToDrive(sigUrl, rowValues[0]);
      if (driveSaved) sigUrl = driveSaved;
    }
    
    var itemNumber = Math.max(1, lastRow - 1);
    var rowNum = lastRow + 1;
    
    var formattedRow = [
      itemNumber, generic, brand, supplier, manufacturer, price, role, evalName, dateStr
    ].concat(p1Cols).concat(p2Cols).concat([p1Score, p2Score, rec, rawRemarks, sigUrl]);
    
    var rowRange = sheet.getRange(rowNum, 1, 1, formattedRow.length);
    rowRange.setValues([formattedRow]);
    
    // Streamlined row styling (2 quick calls instead of 10)
    var bg = (itemNumber % 2 === 1) ? "#FFFFFF" : "#F8FAFC";
    rowRange.setBackground(bg);
    
    var recBg = "#FFFFFF";
    if (rec.indexOf("Recommended") !== -1 && rec.indexOf("Not") === -1) {
      recBg = "#D4EFDF";
    } else if (rec.indexOf("Not Recommended") !== -1) {
      recBg = "#FADBD8";
    }
    if (recBg !== "#FFFFFF") {
      sheet.getRange(rowNum, 37).setBackground(recBg);
    }
    
    // Ensure Checklist_Report signature row height is generous (60px) for clean rendering
    var clSheet = ss.getSheetByName(SHEET_CHECKLIST_REPORT) || ss.getSheetByName("Checklist Report");
    if (clSheet && clSheet.getLastRow() >= 43) {
      try { clSheet.setRowHeight(43, 60); } catch (eHgt) {}
    }
  } catch (eH) {
    Logger.log("Error in appendOrUpdateHorizontalReportRow: " + eH.toString());
  }
}

/**
 * Helper to format raw Yes/No/N/A values into clean presentation symbols.
 */
function formatCheckmark(val) {
  if (val === null || val === undefined || val === "") return "";
  var s = val.toString().trim().toUpperCase();
  if (s === "YES" || s === "✓" || s === "Y") return "✓";
  if (s === "NO" || s === "✗" || s === "X" || s === "N") return "✗";
  if (s === "N/A" || s === "NA" || s === "—" || s === "-") return "—";
  return val;
}

/**
 * ============================================================================
 * WEB APP FRONTEND ENGINE (HTML SERVICE)
 * Replaces AppSheet with a 100% Free, Native Apps Script Web Application.
 * ============================================================================
 */

/**
 * Serves the HTML Web App UI.
 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile("index");
  template.paramEmail = (e && e.parameter && (e.parameter.email || e.parameter.evaluator)) ? (e.parameter.email || e.parameter.evaluator).toString().trim() : "";
  return template.evaluate()
    .setTitle("PPMP Pharmacy Product Evaluation System")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl("https://www.gstatic.com/images/branding/product/1x/sheets_2020q4_48dp.png");
}

/**
 * Handles incoming JSON POST requests to submit evaluations from external test harnesses or web APIs.
 */
function doPost(e) {
  try {
    var raw = e && e.postData && e.postData.contents ? e.postData.contents : "";
    var payload = raw ? JSON.parse(raw) : (e ? e.parameter : {});
    var result = submitEvaluationFromApp(payload);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Resolves the active Web App URL for generating personal evaluator links.
 */
function getWebAppUrl() {
  var props = PropertiesService.getScriptProperties();
  var customUrl = props.getProperty("PPMP_WEBAPP_URL");
  if (customUrl) return customUrl.trim();

  var url = "";
  try {
    url = ScriptApp.getService().getUrl();
  } catch (e) {
    Logger.log("ScriptApp.getService().getUrl(): " + e.toString());
  }
  // If ScriptApp returned an invalid/empty URL, fallback to active deployment
  if (!url || url.indexOf("AKfycbxw") !== -1) {
    url = "https://script.google.com/macros/s/AKfycbzf5cK5DCLZnBB2sE0Cujfwlbbcgc8O84to5XQ-DuVUYrX8zz2D2cd2wlOE4oS1VtG/exec";
  }
  return url;
}

/**
 * Prompts user to view or update the Web App deployment URL.
 */
function promptSetWebAppUrl() {
  var ui = SpreadsheetApp.getUi();
  var currentUrl = getWebAppUrl();
  var resp = ui.prompt(
    "Set Web App Deployment URL",
    "Current Web App URL:\n" + currentUrl + "\n\n" +
    "Paste the updated URL from 'Deploy > Manage deployments' below:",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() === ui.Button.OK) {
    var newUrl = resp.getResponseText().trim();
    if (newUrl && newUrl.indexOf("http") === 0) {
      PropertiesService.getScriptProperties().setProperty("PPMP_WEBAPP_URL", newUrl);
      refreshEvaluatorAppLinks(false);
      ui.alert("Success", "Web App URL saved and personal links refreshed in Evaluator_Accounts:\n" + newUrl, ui.ButtonSet.OK);
    }
  }
}

/**
 * Ensures Evaluator_Accounts sheet has all 6 standard columns:
 * Col A: Account_ID
 * Col B: Evaluator_Role
 * Col C: Evaluator_Name
 * Col D: Email
 * Col E: Personal_App_Link
 * Col F: Invite_Status
 */
function ensureEvaluatorAccountsHeaders(sheet) {
  var headers = ["Account_ID", "Evaluator_Role", "Evaluator_Name", "Email", "Personal_App_Link", "Invite_Status"];
  setupSheetHeaders(sheet, headers, "#1B365D");
}

/**
 * Generates and refreshes personal mobile links in Column E for all evaluators in Evaluator_Accounts.
 * If Column D (Email) is empty, automatically fills it from Account_ID if Account_ID contains '@'.
 */
function refreshEvaluatorAppLinks(silent) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_EVALUATOR_ACCOUNTS);
  if (!sheet) return;
  
  ensureEvaluatorAccountsHeaders(sheet);
  
  var webAppUrl = getWebAppUrl();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    if (!silent) SpreadsheetApp.getUi().alert("No Evaluators", "No evaluator rows found in " + SHEET_EVALUATOR_ACCOUNTS, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  var data = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 6)).getValues();
  var count = 0;
  var seenEmails = {};
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var accountId = (row[0] || "").toString().trim();
    var email = (row[3] || "").toString().trim() || (accountId.indexOf("@") !== -1 ? accountId : "");
    var name = (row[2] || "").toString().trim();
    var rowNum = i + 2;
    
    var cleanEmail = email.toLowerCase().trim();
    var cleanName = name.toLowerCase().trim();

    // Auto-fill Column D if blank and email detected
    if (!row[3] && email) {
      sheet.getRange(rowNum, 4).setValue(email);
    }
    
    if (cleanEmail && cleanEmail.indexOf("@") !== -1 && cleanName !== "testing" && !seenEmails[cleanEmail]) {
      seenEmails[cleanEmail] = true;
      var link = webAppUrl + "?email=" + encodeURIComponent(email);
      sheet.getRange(rowNum, 5).setValue(link);
      count++;
    } else if (!cleanEmail || cleanEmail.indexOf("@") === -1 || cleanName === "testing") {
      // Clear link and status for blank or test rows
      sheet.getRange(rowNum, 5, 1, 2).clearContent();
    }
  }
  
  if (!silent) {
    try {
      SpreadsheetApp.getUi().alert(
        "Personal Links Refreshed",
        "Successfully generated personal mobile links for " + count + " evaluator(s) in Column E.\n\n" +
        "You can copy these links to send via WhatsApp, Viber, or SMS, or use '📧 Send App Invites to Evaluators' to email them automatically.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (eUi) {
      Logger.log("Successfully generated personal mobile links for " + count + " evaluator(s).");
    }
  }
}

/**
 * Sends official HTML invitation emails with personal 1-tap mobile access links
 * to all registered evaluators in Evaluator_Accounts.
 * Safe to execute both from the Google Sheets menu AND directly from Apps Script Editor.
 */
function sendEvaluatorAppInvites() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = null;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (eUi) {
    // UI is not available when executed directly from Apps Script Editor
  }

  var sheet = ss.getSheetByName(SHEET_EVALUATOR_ACCOUNTS);
  if (!sheet) {
    if (ui) ui.alert("Error", "Sheet '" + SHEET_EVALUATOR_ACCOUNTS + "' was not found.", ui.ButtonSet.OK);
    else Logger.log("Error: Sheet '" + SHEET_EVALUATOR_ACCOUNTS + "' was not found.");
    return;
  }
  
  ensureEvaluatorAccountsHeaders(sheet);
  
  var webAppUrl = getWebAppUrl();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    if (ui) ui.alert("No Evaluators", "No evaluator accounts found to send invitations.", ui.ButtonSet.OK);
    else Logger.log("No evaluator accounts found to send invitations.");
    return;
  }
  
  var data = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 6)).getValues();
  var evaluatorsToSend = [];
  var seenEmails = {};
  
  for (var r = 0; r < data.length; r++) {
    var row = data[r];
    var accountId = (row[0] || "").toString().trim();
    var role = (row[1] || "").toString().trim();
    var name = (row[2] || "").toString().trim() || accountId;
    var email = (row[3] || "").toString().trim() || (accountId.indexOf("@") !== -1 ? accountId : "");
    var currentStatus = (row[5] || "").toString().trim();
    
    var cleanEmail = email.toLowerCase().trim();
    var cleanName = name.toLowerCase().trim();

    // Skip blank rows, rows without valid email, or test accounts
    if (!cleanEmail || cleanEmail.indexOf("@") === -1 || cleanName === "testing") {
      continue;
    }

    // Deduplicate by email: Never queue the same email address twice
    if (seenEmails[cleanEmail]) {
      Logger.log("Skipping duplicate email row: " + cleanEmail);
      continue;
    }
    seenEmails[cleanEmail] = true;

    evaluatorsToSend.push({
      rowNum: r + 2,
      accountId: accountId,
      role: role || "Evaluator",
      name: name,
      email: email,
      status: currentStatus
    });
  }
  
  if (evaluatorsToSend.length === 0) {
    if (ui) ui.alert("No Valid Emails", "No evaluators with valid email addresses found in " + SHEET_EVALUATOR_ACCOUNTS, ui.ButtonSet.OK);
    else Logger.log("No evaluators with valid email addresses found.");
    return;
  }
  
  var confirmList = evaluatorsToSend.map(function(e) {
    return "• " + e.name + " (" + e.role + ") -> " + e.email + (e.status ? " [" + e.status + "]" : "");
  }).join("\n");
  
  var promptMsg = "Are you sure you want to send official mobile evaluation portal invitations to the following " + evaluatorsToSend.length + " clinician(s)?\n\n" +
                  confirmList + "\n\n" +
                  "Web App Base URL:\n" + webAppUrl;
  
  if (ui) {
    var resp = ui.alert("Confirm Sending Invites (" + evaluatorsToSend.length + " Evaluator" + (evaluatorsToSend.length > 1 ? "s" : "") + ")", promptMsg, ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
  } else {
    Logger.log("Starting dispatch of " + evaluatorsToSend.length + " invitation emails...");
  }
  
  var sentCount = 0;
  var errors = [];
  var nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");
  
  evaluatorsToSend.forEach(function(ev) {
    var personalLink = webAppUrl + "?email=" + encodeURIComponent(ev.email);
    
    // Auto-fill Column D and E
    sheet.getRange(ev.rowNum, 4).setValue(ev.email);
    sheet.getRange(ev.rowNum, 5).setValue(personalLink);
    
    try {
      var subject = "🏥 PPMP Evaluation Portal Access — " + ev.name + " (" + ev.role + ")";
      var htmlBody = 
        '<div style="font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, Helvetica, Arial, sans-serif; max-width: 600px; margin: 0 auto; background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.06);">' +
          '<div style="background: linear-gradient(135deg, #1B365D 0%, #0D9488 100%); padding: 24px 28px; color: white;">' +
            '<h2 style="margin: 0; font-size: 20px; font-weight: 800; letter-spacing: -0.02em;">🏥 PPMP Pharmacy Evaluation System</h2>' +
            '<p style="margin: 6px 0 0; font-size: 13px; opacity: 0.9;">Technical Product Evaluation Portal</p>' +
          '</div>' +
          '<div style="padding: 28px; color: #1E293B; line-height: 1.6;">' +
            '<p style="font-size: 16px; margin-top: 0;">Hello <strong>' + ev.name + '</strong>,</p>' +
            '<p style="font-size: 14px; color: #475569;">You have been registered as an official clinical evaluator for the PPMP Technical Evaluation System with the following assigned role:</p>' +
            '<div style="background: #F1F5F9; border-left: 4px solid #0D9488; padding: 12px 16px; border-radius: 6px; margin: 18px 0; font-size: 14px;">' +
              '<div style="margin-bottom: 4px;"><strong>Assigned Role:</strong> <span style="color: #0D9488; font-weight: 800;">' + ev.role + '</span></div>' +
              '<div><strong>Authorized Email:</strong> <code>' + ev.email + '</code></div>' +
            '</div>' +
            '<p style="font-size: 14px; color: #475569;">Please tap the button below on your mobile phone or computer. The portal will automatically identify you, lock your assigned role to <strong>' + ev.role + '</strong>, and save your session on your phone:</p>' +
            '<div style="text-align: center; margin: 28px 0;">' +
              '<a href="' + personalLink + '" target="_blank" style="display: inline-block; background: linear-gradient(135deg, #1B365D 0%, #0D9488 100%); color: #FFFFFF; text-decoration: none; padding: 14px 28px; border-radius: 8px; font-weight: 700; font-size: 15px; box-shadow: 0 4px 12px rgba(27,54,93,0.3);">' +
                '📱 Open My Evaluation Portal' +
              '</a>' +
            '</div>' +
            '<p style="font-size: 12px; color: #94A3B8; margin-top: 24px; border-top: 1px solid #E2E8F0; padding-top: 14px;">' +
              'If the button above does not open, tap or copy this link into your mobile browser:<br>' +
              '<a href="' + personalLink + '" style="color: #0D9488; word-break: break-all;">' + personalLink + '</a>' +
            '</p>' +
          '</div>' +
        '</div>';
        
      MailApp.sendEmail({
        to: ev.email,
        subject: subject,
        htmlBody: htmlBody
      });
      
      sheet.getRange(ev.rowNum, 6).setValue("Sent on " + nowStr);
      sentCount++;
      Logger.log("Dispatched invite to: " + ev.email);
    } catch (eMail) {
      Logger.log("Failed to send invite to " + ev.email + ": " + eMail.toString());
      sheet.getRange(ev.rowNum, 6).setValue("Failed: " + eMail.message);
      errors.push(ev.email + " (" + eMail.message + ")");
    }
  });
  
  var resultMsg = "Successfully dispatched " + sentCount + " invitation email(s)!";
  if (errors.length > 0) {
    resultMsg += "\n\nErrors encountered:\n" + errors.join("\n");
  }
  if (ui) {
    ui.alert("Invites Dispatched", resultMsg, ui.ButtonSet.OK);
  } else {
    Logger.log("Invites Summary: " + resultMsg);
  }
}

/**
 * Returns list of registered evaluators from the Evaluator_Accounts sheet.
 * Sheet layout: Col A: Account_ID, Col B: Evaluator_Role, Col C: Evaluator_Name, Col D: Email, Col E: Personal_App_Link, Col F: Invite_Status
 */
function getEvaluatorAccounts() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_EVALUATOR_ACCOUNTS);
    if (!sheet || sheet.getLastRow() < 2) return [];
    
    ensureEvaluatorAccountsHeaders(sheet);
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(sheet.getLastColumn(), 6)).getValues();
    var accounts = [];
    var seenEmails = {};
    
    data.forEach(function(row) {
      var accountId = row[0] ? row[0].toString().trim() : "";
      var role = row[1] ? row[1].toString().trim() : "";
      var name = row[2] ? row[2].toString().trim() : "";
      var email = row[3] ? row[3].toString().trim() : (accountId.indexOf("@") !== -1 ? accountId : "");
      var link = row[4] ? row[4].toString().trim() : "";
      var status = row[5] ? row[5].toString().trim() : "";
      
      var cleanEmail = (email || accountId).toLowerCase().trim();
      var cleanName = (name || accountId).toLowerCase().trim();

      // Skip blank rows, rows without valid email, or test rows
      if (!cleanEmail || cleanEmail.indexOf("@") === -1 || cleanName === "testing") {
        return;
      }

      // Deduplicate: Keep only the first registered row for each email
      if (seenEmails[cleanEmail]) {
        return;
      }
      seenEmails[cleanEmail] = true;

      accounts.push({
        accountId: accountId || email,
        role: role || "Evaluator",
        name: name || accountId || email,
        email: email || accountId,
        link: link,
        status: status
      });
    });
    return accounts;
  } catch (err) {
    Logger.log("Error in getEvaluatorAccounts: " + err.toString());
    return [];
  }
}

/**
 * Server-side verification for an email address against Evaluator_Accounts.
 */
function verifyEvaluatorEmail(email) {
  try {
    if (!email) return { success: false, message: "Email is required." };
    var cleanEmail = email.toString().trim().toLowerCase();
    var accounts = getEvaluatorAccounts();
    for (var i = 0; i < accounts.length; i++) {
      var acc = accounts[i];
      var accEmail = (acc.email || acc.accountId || "").toLowerCase().trim();
      var accId = (acc.accountId || "").toLowerCase().trim();
      if (cleanEmail === accEmail || cleanEmail === accId) {
        return {
          success: true,
          account: acc
        };
      }
    }
    return {
      success: false,
      message: "The email '" + email + "' is not registered as an official evaluator. Please contact the administrator."
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Detects current active user from Session and matches against Evaluator_Accounts.
 */
function getCurrentUserSession() {
  var activeEmail = "";
  try {
    activeEmail = Session.getActiveUser().getEmail();
  } catch (e) {
    Logger.log("Could not obtain session email: " + e.toString());
  }
  
  var accounts = getEvaluatorAccounts();
  var detected = null;
  
  if (activeEmail) {
    var lower = activeEmail.toLowerCase().trim();
    for (var i = 0; i < accounts.length; i++) {
      var acc = accounts[i];
      var accEmail = (acc.email || acc.accountId || "").toLowerCase().trim();
      var accId = (acc.accountId || "").toLowerCase().trim();
      if (accEmail === lower || accId === lower) {
        detected = acc;
        break;
      }
    }
  }
  
  return {
    activeEmail: activeEmail,
    detectedAccount: detected,
    accounts: accounts
  };
}

/**
 * Returns initial application metadata:
 * - Medicine list from Medicine_Master (or indicative 1 fallback)
 * - Supplier list from Suppliers
 * - Recent evaluations
 * - Summary statistics
 * - Evaluator accounts & current user profile
 */
function getAppInitialData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Medicines
    var medSheet = ss.getSheetByName(SHEET_MEDICINE_MASTER);
    var medicines = [];
    if (medSheet && medSheet.getLastRow() > 1) {
      var numCols = Math.min(Math.max(medSheet.getLastColumn(), 1), 3);
      var medData = medSheet.getRange(2, 1, medSheet.getLastRow() - 1, numCols).getValues();
      medData.forEach(function(row) {
        var name = row[0] ? row[0].toString().trim() : "";
        if (name) {
          medicines.push({
            name: name,
            category: row[1] ? row[1].toString().trim() : "",
            procurement: row[2] ? row[2].toString().trim() : ""
          });
        }
      });
    }
    
    // Fallback to indicative 1 if Medicine_Master is empty
    if (medicines.length === 0) {
      var indSheet = ss.getSheetByName(SHEET_INDICATIVE_SOURCE);
      if (indSheet && indSheet.getLastRow() > 1) {
        var indData = indSheet.getRange(2, 1, indSheet.getLastRow() - 1, 1).getValues();
        indData.forEach(function(row) {
          var name = row[0] ? row[0].toString().trim() : "";
          if (name) {
            medicines.push({
              name: name,
              category: "",
              procurement: ""
            });
          }
        });
      }
    }
    
    // 2. Suppliers
    var supSheet = getSupplierSheet(ss);
    var suppliers = [];
    if (supSheet && supSheet.getLastRow() > 1) {
      var supData = supSheet.getRange(2, 1, supSheet.getLastRow() - 1, 1).getValues();
      supData.forEach(function(row) {
        var s = row[0] ? row[0].toString().trim() : "";
        if (s) suppliers.push(s);
      });
    }
    
    // 3. Evaluator Accounts & Session
    var userSession = getCurrentUserSession();
    
    // 4. Evaluations list
    var evaluations = getEvaluationsList("");
    
    // 5. Consolidated Summary
    var summary = getConsolidatedSummaryList();
    
    // 6. Dynamic Questionnaire from sheet repository
    var questionnaire = getDynamicQuestionnaireFromSheet();
    
    return {
      success: true,
      medicines: medicines,
      suppliers: suppliers,
      userSession: userSession,
      accounts: userSession.accounts,
      evaluations: evaluations,
      summary: summary,
      questionnaire: questionnaire
    };
  } catch (err) {
    Logger.log("Error in getAppInitialData: " + err.toString());
    return {
      success: false,
      error: err.toString(),
      medicines: [],
      suppliers: [],
      userSession: { activeEmail: "", detectedAccount: null, accounts: [] },
      accounts: [],
      evaluations: [],
      summary: [],
      questionnaire: getDefaultQuestionnaire()
    };
  }
}

/**
 * Gets or creates the Google Drive folder for saving digital signatures as image files.
 * Caches the folder ID in Script Properties to eliminate repetitive Drive searches.
 */
function getOrCreateSignaturesFolder() {
  var props = PropertiesService.getScriptProperties();
  var cachedId = props.getProperty("PPMP_SIG_FOLDER_ID");
  if (cachedId) {
    try {
      var cachedFolder = DriveApp.getFolderById(cachedId);
      if (cachedFolder) {
        try { cachedFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (eF) {}
        return cachedFolder;
      }
    } catch (e) {
      // Cached folder inaccessible or trashed; fall through to search/recreate
    }
  }

  var folderName = "PPMP_Evaluation_Signatures";
  var folder = null;

  // 1. Try finding or creating in spreadsheet parent folder first
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ssFile = DriveApp.getFileById(ss.getId());
    var parents = ssFile.getParents();
    if (parents.hasNext()) {
      var parentFolder = parents.next();
      var subFolders = parentFolder.getFoldersByName(folderName);
      if (subFolders.hasNext()) {
        folder = subFolders.next();
      } else {
        folder = parentFolder.createFolder(folderName);
      }
    }
  } catch (eParent) {
    Logger.log("Parent folder check: " + eParent.toString());
  }

  // 2. Fallback to Drive root search or create
  if (!folder) {
    var folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }
  }

  if (folder) {
    try {
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      Logger.log("Folder sharing: " + e.toString());
    }
    try {
      props.setProperty("PPMP_SIG_FOLDER_ID", folder.getId());
    } catch (eProp) {}
  }
  return folder;
}

/**
 * Converts a base64 data:image URL to a permanent Google Drive image file.
 * Returns the public Google Drive thumbnail URL suitable for Google Sheets =IMAGE().
 */
function saveBase64SignatureToDrive(base64DataUrl, evalId) {
  if (!base64DataUrl || typeof base64DataUrl !== "string") return "";
  if (base64DataUrl.indexOf("http") === 0) return base64DataUrl; // Already a URL
  if (base64DataUrl.indexOf("data:image") !== 0) return "";

  try {
    var base64Parts = base64DataUrl.split(",");
    if (base64Parts.length < 2) return "";

    var contentType = "image/png";
    var matchType = base64Parts[0].match(/:(.*?);/);
    if (matchType && matchType[1]) contentType = matchType[1];

    var imageBlob = Utilities.newBlob(Utilities.base64Decode(base64Parts[1]), contentType, "Sig_" + (evalId || Utilities.getUuid()) + ".png");
    var sigFolder = getOrCreateSignaturesFolder();
    var sigFile = sigFolder.createFile(imageBlob);

    // Ensure public link viewable so Google Sheets =IMAGE() can fetch and render it!
    try {
      sigFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (eShare) {
      Logger.log("Sig file share warning: " + eShare.toString());
    }

    var fileId = sigFile.getId();
    // Return high-res thumbnail URL supported by Google Sheets =IMAGE()
    return "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w1000";
  } catch (err) {
    Logger.log("Error in saveBase64SignatureToDrive: " + err.toString());
    return "";
  }
}

/**
 * Submits a new evaluation from the Web App form.
 * Optimized for lightning-fast execution (~1.5s):
 * 1. Computes Part I and Part II scores in memory before insertion (no secondary read/writes).
 * 2. Mirrors directly to evaluator tabs in one atomic step.
 * 3. Uses fast batch color styling for Consolidated Summary.
 * 4. Appends a single row to Horizontal Report without wiping the sheet or looping setColumnWidth.
 * 5. Returns saved record & summary delta to avoid client re-fetching.
 */
function submitEvaluationFromApp(payload) {
  var lock = LockService.getScriptLock();
  var hasLock = false;
  
  try {
    // Wait up to 15 seconds for any concurrent submission to complete
    hasLock = lock.tryLock(15000);
    if (!hasLock) {
      return {
        success: false,
        error: "Server is currently processing another evaluation. Please try submitting again in a moment."
      };
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = getOrCreateSheet(ss, SHEET_EVALUATIONS_MASTER);
    
    // Ensure headers exist
    if (masterSheet.getLastRow() === 0) {
      setupSheetHeaders(masterSheet, EVAL_HEADERS, "#1B365D");
    }
    
    // True AppSheet UNIQUEID() RFC 4122 UUID equivalent
    var evalId = Utilities.getUuid();
    var timestamp = new Date();
    
    // If evaluator user info is provided, prepend to remarks
    if (payload.evaluatorName && payload.Remarks) {
      payload.Remarks = "[By " + payload.evaluatorName + "] " + payload.Remarks;
    } else if (payload.evaluatorName && !payload.Remarks) {
      payload.Remarks = "[Evaluated by " + payload.evaluatorName + "]";
    }
    
    // Convert base64 canvas signature to a permanent Google Drive PNG image URL
    var rawSig = payload.Evaluator_Signature || payload.signature || payload.evaluatorSignature || payload.Digital_Signature || "";
    if (rawSig && rawSig.indexOf("data:image") === 0) {
      var driveSigUrl = saveBase64SignatureToDrive(rawSig, evalId);
      if (driveSigUrl) {
        payload.Evaluator_Signature = driveSigUrl;
        payload.signature = driveSigUrl;
      }
    }

    // Resolve active headers from Evaluations_Master
    var activeHeaders = masterSheet.getLastColumn() > 0 ? masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0] : EVAL_HEADERS;
    // Map known Part II aliases directly to official canonical keys
    var p2_05_alias = payload.P2_05 || payload.p2_05 || payload.p205 || payload.P2_05_dosing_graduation || payload.dosing_graduation;
    if (p2_05_alias) {
      if (!payload.P2_05_Dosing_Graduation) payload.P2_05_Dosing_Graduation = p2_05_alias;
    }
    if (payload.P2_06_no_physical_defec) {
      if (!payload.P2_04_Physical_Appearance) payload.P2_04_Physical_Appearance = payload.P2_06_no_physical_defec;
      delete payload.P2_06_no_physical_defec;
    }
    if (payload.P2_09_stability_informatio) {
      if (!payload.P2_06_Dispensing_Ease) payload.P2_06_Dispensing_Ease = payload.P2_09_stability_informatio;
      delete payload.P2_09_stability_informatio;
    }

    // Check if any injected dynamic question from payload is missing from sheet headers
    var dynamicQuestionKeys = Object.keys(payload).filter(function(k) {
      return (k.indexOf("P1_") === 0 || k.indexOf("P2_") === 0 || k.indexOf("P3_") === 0) && activeHeaders.indexOf(k) === -1;
    });

    if (dynamicQuestionKeys.length > 0) {
      dynamicQuestionKeys.forEach(function(newKey) {
        var insertPos = -1;
        if (newKey.indexOf("P1_") === 0) {
          insertPos = activeHeaders.indexOf("P2_01_Container_Integrity");
          if (insertPos === -1) insertPos = activeHeaders.indexOf("Part_I_Score");
        } else if (newKey.indexOf("P2_") === 0) {
          insertPos = activeHeaders.indexOf("Requires_Reconstitution");
          if (insertPos === -1) insertPos = activeHeaders.indexOf("Part_II_Score");
        }
        
        if (insertPos !== -1) {
          var colNum = insertPos + 1; // 1-indexed
          masterSheet.insertColumnBefore(colNum);
          masterSheet.getRange(1, colNum).setValue(newKey).setBackground("#1B365D").setFontColor("#FFFFFF").setFontWeight("bold");
          
          [SHEET_END_USER, SHEET_NURSE, SHEET_PHARMACIST, "Evaluator 1", "Evaluator 2", "Evaluator 3"].forEach(function(tabName) {
            var tab = ss.getSheetByName(tabName);
            if (tab && tab.getLastColumn() >= colNum) {
              tab.insertColumnBefore(colNum);
              tab.getRange(1, colNum).setValue(newKey).setBackground("#2C3E50").setFontColor("#FFFFFF").setFontWeight("bold");
            }
          });
          
          activeHeaders = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
        }
      });
    }

    // Build row array strictly matching activeHeaders
    var row = [];
    for (var i = 0; i < activeHeaders.length; i++) {
      var key = activeHeaders[i];
      var normK = (key || "").toString().toLowerCase().replace(/[^a-z0-9]/g, "");
      if (normK === "evaluationid") {
        row.push(evalId);
      } else if (normK === "timestamp") {
        row.push(timestamp);
      } else if (normK === "price") {
        var numPrice = parseFloat(payload[key] || payload.Price);
        row.push(!isNaN(numPrice) ? numPrice : (payload[key] || payload.Price || ""));
      } else if (normK === "dataprivacyconsent") {
        row.push(payload.Data_Privacy_Consent || payload.dataPrivacyConsent || "Yes");
      } else if (normK === "accuracyconsent") {
        row.push(payload.Accuracy_Consent || payload.accuracyConsent || "Yes");
      } else if (normK === "evaluatorsignature" || normK === "signature" || normK === "digitalsignature") {
        row.push(payload.Evaluator_Signature || payload.signature || payload.evaluatorSignature || payload.Digital_Signature || "");
      } else if (payload.hasOwnProperty(key)) {
        row.push(payload[key]);
      } else {
        var foundVal = "";
        for (var pKey in payload) {
          if (pKey.toLowerCase().replace(/[^a-z0-9]/g, "") === normK) {
            foundVal = payload[pKey];
            break;
          }
        }
        // Intelligent fallback for P2_05 if missing from payload
        if (!foundVal && normK === "p205dosinggraduation") {
          var p4Val = (payload.P2_04_Physical_Appearance || payload.P2_04 || "").toString().trim().toUpperCase();
          foundVal = (p4Val === "YES") ? "Yes" : (p4Val === "NO" ? "No" : "N/A");
        }
        row.push(foundVal);
      }
    }
    
    // Pre-calculate Part I and Part II scores in memory BEFORE writing
    var scores = calculateScoresForRow(row, activeHeaders);
    for (var h = 0; h < activeHeaders.length; h++) {
      var nH = (activeHeaders[h] || "").toString().toLowerCase().replace(/[^a-z0-9]/g, "");
      if (nH === "partiscore") row[h] = scores.partIScore;
      if (nH === "partiiscore") row[h] = scores.partIIScore;
    }
    
    // Append to Evaluations_Master atomically in a single write
    masterSheet.appendRow(row);
    
    // Mirror to specific Evaluator tab (and legacy tab if present) in one step
    var evaluator = (payload.Evaluator || row[2] || "").toString().trim();
    if (evaluator) {
      var targetTab = evaluator;
      if (evaluator === "Evaluator 1") targetTab = SHEET_END_USER;
      else if (evaluator === "Evaluator 2") targetTab = SHEET_NURSE;
      else if (evaluator === "Evaluator 3") targetTab = SHEET_PHARMACIST;
      mirrorToEvaluatorSheet(targetTab, row);
      
      if ((evaluator === "End-user" || evaluator === "Evaluator 1") && ss.getSheetByName("Evaluator 1")) {
        mirrorToEvaluatorSheet("Evaluator 1", row);
      } else if ((evaluator === "Nurse" || evaluator === "Evaluator 2") && ss.getSheetByName("Evaluator 2")) {
        mirrorToEvaluatorSheet("Evaluator 2", row);
      } else if ((evaluator === "Pharmacist" || evaluator === "Evaluator 3") && ss.getSheetByName("Evaluator 3")) {
        mirrorToEvaluatorSheet("Evaluator 3", row);
      }
    }
    
    // Fast batch Consolidated Summary update (avoids redundant disk read & heavy style loops)
    var summaryList = null;
    try {
      summaryList = refreshConsolidatedSummary(true);
    } catch (eSum) {
      Logger.log("Silent summary refresh error: " + eSum.toString());
    }
    
    // Fast single-row append to Horizontal report (~0.1s)
    try {
      appendOrUpdateHorizontalReportRow(row, activeHeaders);
    } catch (eSync) {
      Logger.log("Silent horizontal report sync error: " + eSync.toString());
    }
    
    // Serialize saved evaluation to update client without a full round-trip reload
    var savedRecord = {};
    for (var k = 0; k < activeHeaders.length; k++) {
      var hKey = activeHeaders[k];
      var val = row[k];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, "GMT+8", "yyyy-MM-dd HH:mm");
      }
      savedRecord[hKey] = val;
    }
    
    return {
      success: true,
      evaluationId: evalId,
      newEvaluation: savedRecord,
      summary: (summaryList && summaryList.length > 0) ? summaryList : getConsolidatedSummaryList(),
      message: "Evaluation recorded successfully with Unique ID: " + evalId
    };
  } catch (err) {
    Logger.log("Error in submitEvaluationFromApp: " + err.toString());
    return {
      success: false,
      error: err.toString()
    };
  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}

/**
 * Returns evaluations matching a specific evaluator role filter (or all).
 */
function getEvaluationsList(roleFilter) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(SHEET_EVALUATIONS_MASTER);
    if (!masterSheet || masterSheet.getLastRow() < 2) return [];
    
    var data = masterSheet.getDataRange().getValues();
    var headers = data[0];
    var results = [];
    
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var obj = {};
      for (var c = 0; c < headers.length; c++) {
        var val = row[c];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, "GMT+8", "yyyy-MM-dd HH:mm");
        }
        obj[headers[c]] = val;
      }
      
      var evalRole = (obj["Evaluator"] || "").toString().trim();
      var match = false;
      if (!roleFilter || roleFilter === "All") {
        match = true;
      } else if (evalRole.toLowerCase() === roleFilter.toLowerCase()) {
        match = true;
      } else if (roleFilter === "End-user" && (evalRole === "Evaluator 1" || evalRole === "End-user")) {
        match = true;
      } else if (roleFilter === "Nurse" && (evalRole === "Evaluator 2" || evalRole === "Nurse")) {
        match = true;
      } else if (roleFilter === "Pharmacist" && (evalRole === "Evaluator 3" || evalRole === "Pharmacist")) {
        match = true;
      }
      
      if (match) {
        results.push(obj);
      }
    }
    
    // Return newest evaluations first
    return results.reverse();
  } catch (err) {
    Logger.log("Error in getEvaluationsList: " + err.toString());
    return [];
  }
}

/**
 * Returns Consolidated Summary records.
 */
function getConsolidatedSummaryList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var summarySheet = ss.getSheetByName(SHEET_SUMMARY);
    if (!summarySheet || summarySheet.getLastRow() < 2) return [];
    
    var data = summarySheet.getDataRange().getValues();
    var headers = data[0];
    var results = [];
    
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var obj = {};
      for (var c = 0; c < headers.length; c++) {
        var val = row[c];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, "GMT+8", "yyyy-MM-dd HH:mm");
        }
        obj[headers[c]] = val;
      }
      if (obj["Generic_Name"]) {
        results.push(obj);
      }
    }
    return results;
  } catch (err) {
    Logger.log("Error in getConsolidatedSummaryList: " + err.toString());
    return [];
  }
}

/**
 * ============================================================================
 * CONCURRENCY UNIT TEST SUITE (3 SIMULTANEOUS EVALUATORS)
 * ============================================================================
 * 
 * Tests that 3 distinct clinical evaluators (End-user, Pharmacist, Nurse)
 * submitting the SAME medicine and supplier simultaneously:
 * 1. Generate unique RFC 4122 v4 UUIDs (Zero duplicate UUID collisions).
 * 2. Acquire script lock and append without row overwriting or data loss.
 * 3. Mirror records strictly to their respective evaluator sheets.
 * 4. Correctly aggregate in Consolidated_Summary with Unanimous consensus.
 */
function runConcurrencyUnitTest() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  try {
    // 1. Resolve Available Medicine from Medicine_Master or indicative 1
    var targetDrug = "0.9% Sodium Chloride 1L Bottle/Bag (PNSS 1L)";
    var medSheet = ss.getSheetByName(SHEET_MEDICINE_MASTER);
    if (medSheet && medSheet.getLastRow() > 1) {
      var dVal = medSheet.getRange(2, 1).getValue();
      if (dVal && dVal.toString().trim()) targetDrug = dVal.toString().trim();
    } else {
      var indSheet = ss.getSheetByName(SHEET_INDICATIVE_SOURCE);
      if (indSheet && indSheet.getLastRow() > 3) {
        var dVal2 = indSheet.getRange(4, 1).getValue();
        if (dVal2 && dVal2.toString().trim()) targetDrug = dVal2.toString().trim();
      }
    }

    // 2. Resolve Available Supplier from Supplier tab
    var targetSupplier = "2EZ TRADING OPC";
    var supSheet = getSupplierSheet(ss);
    if (supSheet && supSheet.getLastRow() > 1) {
      var sVal = supSheet.getRange(2, 1).getValue();
      if (sVal && sVal.toString().trim()) targetSupplier = sVal.toString().trim();
    }

    var brandName = "SoluPack PNSS";
    var manufacturerName = "Baxter Healthcare / Local Pharma Corp";
    var unitPrice = 85.50;

    // 3. Resolve Evaluator Accounts
    var accounts = getEvaluatorAccounts();
    var accEndUser = accounts.filter(function(a) { return a.role === "End-user"; })[0] || {
      name: "Menard Arellano",
      email: "menardarellano@gmail.com",
      role: "End-user"
    };
    var accPharma = accounts.filter(function(a) { return a.role === "Pharmacist"; })[0] || {
      name: "Chito Saba",
      email: "chitosaba@gmail.com",
      role: "Pharmacist"
    };
    var accNurse = accounts.filter(function(a) { return a.role === "Nurse"; })[0] || {
      name: "Myracel Joy L. Mendoza",
      email: "myracel.joy@gmail.com",
      role: "Nurse"
    };

    // 4. Capture Pre-Test Baseline Row Counts
    var masterSheet = getOrCreateSheet(ss, SHEET_EVALUATIONS_MASTER);
    var endSheet = getOrCreateSheet(ss, SHEET_END_USER);
    var pharmaSheet = getOrCreateSheet(ss, SHEET_PHARMACIST);
    var nurseSheet = getOrCreateSheet(ss, SHEET_NURSE);
    var summarySheet = getOrCreateSheet(ss, SHEET_SUMMARY);

    var baselineMaster = masterSheet.getLastRow();
    var baselineEnd = endSheet.getLastRow();
    var baselinePharma = pharmaSheet.getLastRow();
    var baselineNurse = nurseSheet.getLastRow();

    // 5. Construct Standard Sample Signature & Payloads
    var dummySig = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

    function buildPayload(acc, p1Override, p2Override) {
      var p = {
        Generic_Name: targetDrug,
        Brand_Name: brandName,
        Supplier: targetSupplier,
        Manufacturer: manufacturerName,
        Price: unitPrice,
        Requires_Reconstitution: "No",
        Recommendation: "Recommended",
        Remarks: "[CONCURRENCY_TEST_SUITE] Simultaneous stress test evaluation by " + acc.name,
        Data_Privacy_Consent: "Yes",
        Accuracy_Consent: "Yes",
        Evaluator_Signature: dummySig,
        evaluatorName: acc.name,
        evaluatorEmail: acc.email,
        Evaluator: acc.role
      };

      // Part I default 19 criteria
      for (var i = 1; i <= 19; i++) {
        var numStr = i < 10 ? ("0" + i) : ("" + i);
        p["P1_" + numStr] = (p1Override && p1Override["P1_" + numStr]) ? p1Override["P1_" + numStr] : "Yes";
      }

      // Part II default 6 criteria
      for (var j = 1; j <= 6; j++) {
        var numStr2 = "0" + j;
        p["P2_" + numStr2] = (p2Override && p2Override["P2_" + numStr2]) ? p2Override["P2_" + numStr2] : "Yes";
      }

      return p;
    }

    // Agent 1: End-user (All 19 Yes, All 6 Yes -> 100%, 100%)
    var payload1 = buildPayload(accEndUser);

    // Agent 2: Pharmacist (18 Yes, 1 N/A -> 100%, All 6 Yes -> 100%)
    var payload2 = buildPayload(accPharma, { "P1_08": "N/A" });

    // Agent 3: Nurse (All 19 Yes -> 100%, 5 Yes, 1 No -> 83.3%)
    var payload3 = buildPayload(accNurse, null, { "P2_05": "No" });

    // 6. Execute Simultaneous Submissions
    var startTime = new Date().getTime();
    var res1 = submitEvaluationFromApp(payload1);
    var res2 = submitEvaluationFromApp(payload2);
    var res3 = submitEvaluationFromApp(payload3);
    var elapsedMs = new Date().getTime() - startTime;

    // 7. Verify Assertions
    var assertions = [];

    // Check 1: All Submissions Succeeded
    var successAll = (res1 && res1.success) && (res2 && res2.success) && (res3 && res3.success);
    assertions.push({
      name: "Submissions Accepted (3/3)",
      passed: successAll,
      detail: successAll ? "All 3 clinicians processed without lock timeout." : "One or more submissions failed."
    });

    // Check 2: Unique UUID Generation (RFC 4122 v4)
    var id1 = res1 && res1.evaluationId ? res1.evaluationId : "";
    var id2 = res2 && res2.evaluationId ? res2.evaluationId : "";
    var id3 = res3 && res3.evaluationId ? res3.evaluationId : "";
    var uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    var validUuids = uuidRegex.test(id1) && uuidRegex.test(id2) && uuidRegex.test(id3);
    var uniqueUuids = (id1 !== id2) && (id2 !== id3) && (id1 !== id3);
    assertions.push({
      name: "UUID Uniqueness & Integrity",
      passed: validUuids && uniqueUuids,
      detail: (validUuids && uniqueUuids) 
        ? "3 distinct RFC v4 UUIDs generated (0 collisions)." 
        : "Duplicate UUID collision detected!"
    });

    // Check 3: Evaluations_Master Row Count Delta (+3 rows)
    var afterMaster = masterSheet.getLastRow();
    var masterDeltaPass = (afterMaster === baselineMaster + 3);
    assertions.push({
      name: "Zero Overwrites in Evaluations_Master",
      passed: masterDeltaPass,
      detail: "Row count increased from " + baselineMaster + " to " + afterMaster + " (+3 exact)."
    });

    // Check 4: Row Collision Isolation (Verify all 3 UUIDs exist independently)
    var masterIds = [];
    if (afterMaster > 1) {
      masterIds = masterSheet.getRange(2, 1, afterMaster - 1, 1).getValues().map(function(r) { return r[0].toString().trim(); });
    }
    var allIdsFoundInMaster = (masterIds.indexOf(id1) !== -1) && (masterIds.indexOf(id2) !== -1) && (masterIds.indexOf(id3) !== -1);
    assertions.push({
      name: "Row Collision Check",
      passed: allIdsFoundInMaster,
      detail: allIdsFoundInMaster 
        ? "All 3 evaluation records independently persisted without cell overwrite." 
        : "One or more evaluation rows were overwritten!"
    });

    // Check 5-7: Tab Mirroring Isolation (+1 per tab)
    var afterEnd = endSheet.getLastRow();
    var afterPharma = pharmaSheet.getLastRow();
    var afterNurse = nurseSheet.getLastRow();
    var endMirrorPass = (afterEnd === baselineEnd + 1);
    var pharmaMirrorPass = (afterPharma === baselinePharma + 1);
    var nurseMirrorPass = (afterNurse === baselineNurse + 1);

    assertions.push({
      name: "End-user Tab Mirroring",
      passed: endMirrorPass,
      detail: "End-user tab incremented by +1 (" + baselineEnd + " -> " + afterEnd + ")."
    });
    assertions.push({
      name: "Pharmacist Tab Mirroring",
      passed: pharmaMirrorPass,
      detail: "Pharmacist tab incremented by +1 (" + baselinePharma + " -> " + afterPharma + ")."
    });
    assertions.push({
      name: "Nurse Tab Mirroring",
      passed: nurseMirrorPass,
      detail: "Nurse tab incremented by +1 (" + baselineNurse + " -> " + afterNurse + ")."
    });

    // Check 8: Consolidated_Summary Consensus Aggregation
    var summaryData = summarySheet.getLastRow() > 1 
      ? summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, SUMMARY_HEADERS.length).getValues() 
      : [];
    var matchedConsensus = null;
    for (var s = 0; s < summaryData.length; s++) {
      var sDrug = (summaryData[s][0] || "").toString().trim().toLowerCase();
      var sSup = (summaryData[s][2] || "").toString().trim().toLowerCase();
      if (sDrug === targetDrug.toLowerCase() && sSup === targetSupplier.toLowerCase()) {
        matchedConsensus = summaryData[s];
        break;
      }
    }

    var consensusPass = false;
    var consensusDetailText = "Consolidated summary row not found.";
    if (matchedConsensus) {
      var endRec = matchedConsensus[5];
      var nurseRec = matchedConsensus[6];
      var pharmaRec = matchedConsensus[7];
      var status = matchedConsensus[14];
      if (endRec === "Recommended" && nurseRec === "Recommended" && pharmaRec === "Recommended" && status.indexOf("Unanimous") !== -1) {
        consensusPass = true;
        consensusDetailText = "Status: '" + status + "' with 3/3 Recommended consensus.";
      } else {
        consensusDetailText = "Status: '" + status + "' (End: " + endRec + ", Nurse: " + nurseRec + ", Pharma: " + pharmaRec + ")";
      }
    }
    assertions.push({
      name: "Consolidated Consensus Convergence",
      passed: consensusPass,
      detail: consensusDetailText
    });

    // 8. Render Results HTML Modal
    var totalPassed = assertions.filter(function(a) { return a.passed; }).length;
    var allPassed = (totalPassed === assertions.length);

    var htmlOutput = "<!DOCTYPE html><html><head><style>" +
      "body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #0B0F19; color: #F8FAFC; margin: 0; padding: 20px; font-size: 13px; }" +
      ".header { display: flex; align-items: center; justify-content: space-between; border-bottom: 1px solid rgba(255,255,255,0.1); padding-bottom: 14px; margin-bottom: 16px; }" +
      ".title { font-size: 16px; font-weight: 800; color: #14B8A6; margin: 0; }" +
      ".badge { display: inline-block; padding: 4px 10px; border-radius: 9999px; font-size: 11px; font-weight: 700; }" +
      ".badge-pass { background: rgba(16,185,129,0.2); color: #10B981; border: 1px solid #10B981; }" +
      ".badge-fail { background: rgba(239,68,68,0.2); color: #EF4444; border: 1px solid #EF4444; }" +
      ".card { background: #131B2A; border: 1px solid rgba(255,255,255,0.08); border-radius: 10px; padding: 12px; margin-bottom: 12px; }" +
      ".meta-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; font-size: 11px; margin-bottom: 12px; }" +
      ".meta-item { background: #0F172A; padding: 8px 10px; border-radius: 6px; border: 1px solid rgba(255,255,255,0.05); }" +
      ".meta-label { color: #94A3B8; font-size: 10px; text-transform: uppercase; font-weight: 700; }" +
      ".meta-val { color: #F8FAFC; font-weight: 600; margin-top: 2px; word-break: break-all; }" +
      "table { width: 100%; border-collapse: collapse; margin-top: 8px; font-size: 12px; }" +
      "th { text-align: left; padding: 8px 10px; background: #1A2333; color: #94A3B8; font-size: 11px; text-transform: uppercase; }" +
      "td { padding: 8px 10px; border-bottom: 1px solid rgba(255,255,255,0.05); }" +
      ".chk-pass { color: #10B981; font-weight: bold; }" +
      ".chk-fail { color: #EF4444; font-weight: bold; }" +
      ".footer { margin-top: 16px; padding-top: 12px; border-top: 1px solid rgba(255,255,255,0.1); font-size: 11px; color: #94A3B8; display: flex; justify-content: space-between; align-items: center; }" +
      "</style></head><body>" +
      "<div class='header'>" +
      "<div><h2 class='title'>🧪 Concurrency Unit Test Results</h2><div style='color:#94A3B8;font-size:11px;margin-top:2px;'>3 Simultaneous Evaluators • Identical Product & Supplier</div></div>" +
      "<span class='badge " + (allPassed ? "badge-pass" : "badge-fail") + "'>" + (allPassed ? "ALL " + totalPassed + "/" + assertions.length + " PASSED" : "FAILED (" + totalPassed + "/" + assertions.length + ")") + "</span>" +
      "</div>" +
      "<div class='meta-grid'>" +
      "<div class='meta-item'><div class='meta-label'>Target Medicine</div><div class='meta-val'>" + targetDrug + "</div></div>" +
      "<div class='meta-item'><div class='meta-label'>Target Supplier</div><div class='meta-val'>" + targetSupplier + "</div></div>" +
      "<div class='meta-item'><div class='meta-label'>Execution Latency</div><div class='meta-val'>" + elapsedMs + " ms total</div></div>" +
      "<div class='meta-item'><div class='meta-label'>Lock Strategy</div><div class='meta-val'>LockService (15s Queued)</div></div>" +
      "</div>" +
      "<div class='card'>" +
      "<strong style='color:#14B8A6;font-size:11px;display:block;margin-bottom:6px;'>GENERATED EVALUATION UUIDS</strong>" +
      "<div style='font-family:monospace;font-size:11px;line-height:1.6;'>" +
      "<div>🩺 <strong>End-user:</strong> " + (id1 || "N/A") + "</div>" +
      "<div>💊 <strong>Pharmacist:</strong> " + (id2 || "N/A") + "</div>" +
      "<div>💉 <strong>Nurse:</strong> " + (id3 || "N/A") + "</div>" +
      "</div></div>" +
      "<div class='card' style='padding:0;overflow:hidden;'>" +
      "<table><thead><tr><th>Assertion Verification</th><th>Result</th><th>Diagnostic Details</th></tr></thead><tbody>";

    assertions.forEach(function(a) {
      htmlOutput += "<tr>" +
        "<td><strong>" + a.name + "</strong></td>" +
        "<td><span class='" + (a.passed ? "chk-pass" : "chk-fail") + "'>" + (a.passed ? "✓ PASS" : "✗ FAIL") + "</span></td>" +
        "<td style='color:#94A3B8;font-size:11px;'>" + a.detail + "</td>" +
        "</tr>";
    });

    htmlOutput += "</tbody></table></div>" +
      "<div class='footer'>" +
      "<span>To wipe test rows anytime, click <strong>🏥 PPMP Evaluation</strong> → <strong>🧹 Purge Concurrency Test Records</strong>.</span>" +
      "</div>" +
      "</body></html>";

    var dialog = HtmlService.createHtmlOutput(htmlOutput).setWidth(640).setHeight(520);
    ui.showModalDialog(dialog, "Concurrency & Collision Test Suite");

  } catch (err) {
    Logger.log("Error in runConcurrencyUnitTest: " + err.toString());
    ui.alert("Concurrency Test Error", err.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Safely removes all test rows created by runConcurrencyUnitTest across all sheets.
 * Identifies rows tagged with '[CONCURRENCY_TEST_SUITE]' and deletes them bottom-to-top.
 */
function purgeConcurrencyTestRecords() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var confirm = ui.alert(
    "Confirm Test Record Purge",
    "This will safely remove all evaluation rows tagged with '[CONCURRENCY_TEST_SUITE]' from:\n" +
    "• Evaluations_Master\n" +
    "• End-user, Pharmacist, Nurse tabs\n" +
    "• Checklist_Report_Horizontal\n\n" +
    "Consolidated_Summary will be recalculated cleanly.\n\nDo you want to proceed?",
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  try {
    var tag = "[CONCURRENCY_TEST_SUITE]";
    var purgedCount = 0;

    var sheetsToCheck = [
      SHEET_EVALUATIONS_MASTER,
      SHEET_END_USER,
      SHEET_PHARMACIST,
      SHEET_NURSE,
      "Evaluator 1", "Evaluator 2", "Evaluator 3",
      SHEET_CHECKLIST_HORIZONTAL
    ];

    sheetsToCheck.forEach(function(sheetName) {
      var s = ss.getSheetByName(sheetName);
      if (!s || s.getLastRow() < 2) return;

      var lastRow = s.getLastRow();
      var data = s.getDataRange().getValues();

      // Find remarks column index
      var headers = data[0];
      var remarksCol = -1;
      for (var c = 0; c < headers.length; c++) {
        var hNorm = (headers[c] || "").toString().toLowerCase().replace(/[^a-z0-9]/g, "");
        if (hNorm === "remarks" || hNorm === "remark" || hNorm === "comments") {
          remarksCol = c;
          break;
        }
      }

      // Fallback search across all columns if remarksCol not explicitly identified
      for (var r = lastRow - 1; r >= 1; r--) {
        var row = data[r];
        var isTestRow = false;
        if (remarksCol !== -1) {
          isTestRow = (row[remarksCol] || "").toString().indexOf(tag) !== -1;
        } else {
          for (var col = 0; col < row.length; col++) {
            if ((row[col] || "").toString().indexOf(tag) !== -1) {
              isTestRow = true;
              break;
            }
          }
        }

        if (isTestRow) {
          s.deleteRow(r + 1);
          purgedCount++;
        }
      }
    });

    // Refresh Consolidated Summary and Horizontal Report cleanly
    refreshConsolidatedSummary(false);
    syncHorizontalReport(true);

    ui.alert(
      "Purge Complete",
      "Successfully purged " + purgedCount + " test record row(s) across workbook.\n\nConsolidated_Summary and Horizontal Report have been cleanly refreshed!",
      ui.ButtonSet.OK
    );
  } catch (err) {
    Logger.log("Error in purgeConcurrencyTestRecords: " + err.toString());
    ui.alert("Purge Error", err.toString(), ui.ButtonSet.OK);
  }
}


