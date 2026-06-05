// ====================================================================
// 🛠️ SETTINGS: APNI RAW MASTER SHEET KI ID YAHAN DAALEIN
// ====================================================================
var RAW_MASTER_SHEET_ID = "1nCZGpXlXgHjRhvFfoCQCvGrrjs7wzE7GGZJPCinLeB4"; 
// ====================================================================

// CUSTOM MENU: Sheet refresh karne par chalega
function onOpen() {
  CREATE_MENU_NOW();
}

// 🚀 FORCE MENU BUILDER FUNCTION
function CREATE_MENU_NOW() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 Auto Cleaner')
      .addItem('⚡ Clean & Import Month Data', 'startCleaningProcess')
      .addItem('📥 Download Month Data (CSV)', 'downloadMonthDataCSV')
      .addToUi();
}

// 1. CLEANING & IMPORT FUNCTION
function startCleaningProcess() {
  var ui = SpreadsheetApp.getUi();
  var ssCleaned = SpreadsheetApp.getActiveSpreadsheet();
  
  var response = ui.prompt('Data Cleaning Pipeline', 'Raw Master Sheet se kis mahine ka tab clean karke import karna hai?\n(Format: March-2026)', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var targetMonth = response.getResponseText().trim();
  
  ssCleaned.toast("Data fetch aur clean ho raha hai...", "Processing", -1);
  
  try {
    var rawSS = SpreadsheetApp.openById(RAW_MASTER_SHEET_ID);
    var rawSheet = rawSS.getSheetByName(targetMonth);
    
    if (!rawSheet) {
      ui.alert("❌ Error: Raw Master file mein '" + targetMonth + "' naam ka koi tab nahi mila.");
      return;
    }
    
    var data = rawSheet.getDataRange().getValues();
    if (data.length <= 1) {
      ui.alert("⚠️ Data nahi mila ya sheet khali hai.");
      return;
    }

    var finalCleanedData = [];
    var deletedRowsLog = [];
    var titleCleaningLog = [];
    var fixedUrlsLog = [];
    
    // 🚀 NEW: Duplicates ko track karne ke liye Set memory
    var seenRecords = new Set(); 
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rawTitle = row[0] ? row[0].toString().trim() : "";
      var rawLink = row[1] ? row[1].toString().trim() : "";
      var postUrl = row[2] ? row[2].toString().trim() : "";
      var pubDate = row[3] ? row[3].toString().trim() : "";
      var sourceSite = row[4] ? row[4].toString().trim() : "";
      
      // ==========================================
      // 🐛 STEP 2: CLEAR SEPARATED ERROR LOGGING
      // ==========================================
      if (rawTitle === "") {
        deletedRowsLog.push([rawTitle, rawLink, sourceSite, "Step 2: Empty Title", targetMonth]);
        continue;
      }
      if (rawTitle.indexOf('?') !== -1) {
        deletedRowsLog.push([rawTitle, rawLink, sourceSite, "Step 2: Corrupted Text (Question Marks)", targetMonth]);
        continue;
      }
      if (rawLink === "") {
        deletedRowsLog.push([rawTitle, rawLink, sourceSite, "Step 2: Empty Link (No Download URL)", targetMonth]);
        continue;
      }

      // STEP 9: FIX GOOGLE DRIVE URLs
      var cleanedLink = rawLink;
      if (cleanedLink.indexOf('drive.google.com') !== -1) {
        if (cleanedLink.indexOf('id=') !== -1) {
          var match = cleanedLink.match(/id=([a-zA-Z0-9_-]+)/);
          if (match) cleanedLink = "https://drive.google.com/file/d/" + match[1] + "/view";
        } else if (cleanedLink.indexOf('/file/d/') !== -1 && cleanedLink.indexOf('?usp=') !== -1) {
          cleanedLink = cleanedLink.split('?')[0];
        }
      }
      cleanedLink = cleanedLink.replace(/&amp;/g, '&');
      
      if (rawLink !== cleanedLink) {
        fixedUrlsLog.push([rawTitle, rawLink, cleanedLink, targetMonth]);
      }

      // STEP 10: MULTIPLE LINKS CHECK
      var httpMatch = cleanedLink.match(/https?:\/\//gi);
      if (httpMatch && httpMatch.length > 1) {
        deletedRowsLog.push([rawTitle, cleanedLink, sourceSite, "Step 10: Multiple Links in cell", targetMonth]);
        continue;
      }

      // ==========================================
      // 🧹 AGGRESSIVE TITLE CLEANING LOGIC
      // ==========================================
      var cleanedTitle = rawTitle;

      // Fix HTML Entities
      cleanedTitle = cleanedTitle.replace(/&#8211;/g, '-');
      cleanedTitle = cleanedTitle.replace(/&amp;/g, '&');
      cleanedTitle = cleanedTitle.replace(/&#8217;/g, "'");

      // 🚀 Remove 'novel' followed by digits (e.g., novel20712, Novel 12345)
      cleanedTitle = cleanedTitle.replace(/(?:\s|-|_)*novel\s*\d+(?:\s|-|_)*/gi, ' ');

      // Remove specific junk phrases
      cleanedTitle = cleanedTitle.replace(/(?:\s|-)*Urdu Novelians(?:\s|-)*/gi, ' ');
      cleanedTitle = cleanedTitle.replace(/(?:\s|-)*ZNZ(?:\s|-)*/gi, ' ');
      cleanedTitle = cleanedTitle.replace(/(?:\s|-)*Complete\s*PDF(?:\s|-)*/gi, ' ');
      cleanedTitle = cleanedTitle.replace(/(?:\s|-)*Complete(?:\s|-)*/gi, ' ');
      cleanedTitle = cleanedTitle.replace(/(?:\s|-)*Read Online\s*&\s*Download\s*PDF(?:\s|-)*/gi, ' ');
      cleanedTitle = cleanedTitle.replace(/(?:\s|-)*Read Online(?:\s|-)*/gi, ' ');

      // Normal character cleanup
      cleanedTitle = cleanedTitle.replace(/_/g, ' ').replace(/-/g, ' ');
      cleanedTitle = cleanedTitle.replace(/\.(pdf|zip|rar|epub|mobi|txt|docx|html)$/gi, '');
      cleanedTitle = cleanedTitle.replace(/[\[\]\(\)\{\}@#\$%\^&\*\+=\\|<>\/\?]/g, ' ');

      // Deep junk word removal loop
      var prevTitle = "";
      while (cleanedTitle !== prevTitle) {
        prevTitle = cleanedTitle;
        cleanedTitle = cleanedTitle.replace(/(?:\s|\.|\-|_)*(?:pdf|download|free|znz|read\s+online)(?:\s|\.|\-|_)*$/gi, '');
        cleanedTitle = cleanedTitle.replace(/^(?:\s|\.|\-|_)*(?:urdu\s+novel|novel|download)(?:\s|\.|\-|_)*/gi, '');
        cleanedTitle = cleanedTitle.replace(/(?:\s|\.|\-|_)*(?:urdu\s+novel|novel|download)(?:\s|\.|\-|_)*$/gi, '');
        cleanedTitle = cleanedTitle.replace(/[\.\s]+$/, '');
      }

      // Final Polish
      cleanedTitle = cleanedTitle.replace(/\s+/g, ' ').trim();
      cleanedTitle = toTitleCase(cleanedTitle);
      // ==========================================
      
      if (rawTitle !== cleanedTitle) {
        titleCleaningLog.push([rawTitle, cleanedTitle, targetMonth]);
      }

      // STEP 12: WORD COUNT CHECK
      var wordCount = cleanedTitle.split(/\s+/).length;
      if (wordCount <= 2) {
        deletedRowsLog.push([cleanedTitle, cleanedLink, sourceSite, "Step 12: Title has 1 or 2 words only", targetMonth]);
        continue;
      }

      // 🚀 EXACT DUPLICATE CHECK (Step 13)
      var uniqueKey = cleanedTitle.toLowerCase() + "|" + cleanedLink.toLowerCase();
      if (seenRecords.has(uniqueKey)) {
        deletedRowsLog.push([cleanedTitle, cleanedLink, sourceSite, "Step 13: Exact Duplicate Row", targetMonth]);
        continue;
      }
      seenRecords.add(uniqueKey); // Memory mein save kar lo

      finalCleanedData.push([cleanedTitle, cleanedLink, postUrl, pubDate, sourceSite, targetMonth]);
    }

    // SAVE TO TABS
    saveToTab(ssCleaned, "1. Production_Data", ["Cleaned Title", "Fixed Link", "Post URL", "Publish Date", "Source", "Month Added"], finalCleanedData, "#d9ead3");
    saveToTab(ssCleaned, "2. Deleted_Rows_Log", ["Old Title", "Link", "Source", "Reason for Deletion", "Month"], deletedRowsLog, "#f4cccc");
    saveToTab(ssCleaned, "3. Title_Cleaning_Log", ["Original Title", "Cleaned Title", "Month"], titleCleaningLog, "#fff2cc");
    saveToTab(ssCleaned, "4. Fixed_URLs_Log", ["Original Broken URL", "Fixed URL", "Month"], fixedUrlsLog, "#cfe2f3");

    ui.alert("✅ Cleaning Pipeline Complete!\n\nMonth: " + targetMonth + "\n\n📊 SUMMARY:\n- Ready Novels Added: " + finalCleanedData.length + "\n- Rows Deleted (Log): " + deletedRowsLog.length + "\n- Titles Cleaned (Log): " + titleCleaningLog.length + "\n- URLs Fixed (Log): " + fixedUrlsLog.length);

  } catch (e) {
    ui.alert("❌ Error: " + e.message);
  }
}

// 2. 📥 DOWNLOAD SPECIFIC MONTH DATA TO CSV DIRECTLY
function downloadMonthDataCSV() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var productionSheet = ss.getSheetByName("1. Production_Data");
  
  if (!productionSheet) {
    ui.alert("❌ Error: '1. Production_Data' naam ka koi tab nahi mila. Pehle data clean karke import karein.");
    return;
  }
  
  // User se month name lena
  var response = ui.prompt('Download CSV Data', 'Kis mahine ka production data download karna hai?\n(Format: March-2026)', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var targetMonth = response.getResponseText().trim();
  
  var data = productionSheet.getDataRange().getDisplayValues(); 
  
  var csvContent = "";
  var filteredRowsCount = 0;
  
  // CSV Headers
  csvContent += '"Titles","Links","Post URL","Publish Date","Source Website"\r\n';
  
  // Column index jahan month ka naam save hai (Column F yani Index 5)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowMonth = row[5] ? row[5].toString().trim() : "";
    
    // Exact text match
    if (rowMonth.toLowerCase() === targetMonth.toLowerCase()) {
      var title = row[0].toString().replace(/"/g, '""');
      var link = row[1].toString().replace(/"/g, '""');
      var postUrl = row[2].toString().replace(/"/g, '""');
      var pubDate = row[3].toString().replace(/"/g, '""');
      var source = row[4].toString().replace(/"/g, '""');
      
      csvContent += '"' + title + '","' + link + '","' + postUrl + '","' + pubDate + '","' + source + '"\r\n';
      filteredRowsCount++;
    }
  }
  
  if (filteredRowsCount === 0) {
    ui.alert("⚠️ Data Nahi Mila!\n\nProduction Data mein '" + targetMonth + "' ke liye koi records nahi hain.");
    return;
  }
  
  var fileName = "Cleaned_Novels_" + targetMonth.replace("-", "_") + ".csv";
  var base64Content = Utilities.base64Encode(csvContent, Utilities.Charset.UTF_8);
  
  var htmlOutput = HtmlService.createHtmlOutput(
    '<html><body>' +
    '<a id="download_link" download="' + fileName + '" href="data:text/csv;charset=utf-8;base64,' + base64Content + '">Downloading...</a>' +
    '<script>' +
    '  document.getElementById("download_link").click();' +
    '  google.script.host.close();' +
    '</script>' +
    '</body></html>'
  ).setWidth(10).setHeight(10);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, targetMonth + " data download ho raha hai...");
}

// HELPERS
function toTitleCase(str) {
  return str.replace(/\w\S*/g, function(txt) {
    return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
  });
}

function saveToTab(ss, sheetName, headers, dataArray, headerColor) {
  if (dataArray.length === 0) return;
  
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground(headerColor);
    sheet.setFrozenRows(1);
  }
  
  var nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1, dataArray.length, dataArray[0].length).setValues(dataArray);
}