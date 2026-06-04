// ====================================================================
// 🛠️ GLOBAL SETTINGS: AAPKI IDs ADD KAR DI GAYI HAIN
// ====================================================================
var SCRAPER_FOLDER_ID = "10QnfuxgzeYljcqxunHeIMrXI61_L2eZI"; 
var MASTER_SHEET_ID = "1sWEmGf_cJGvqa5Pe6S7OLug9bZ9jpRYuOQJrGbjGJ1g";
// ====================================================================


// 1. AUTOMATIC FUNCTION (Har mahine ki 1st tareekh ko auto chalega)
function compilePreviousMonthData() {
  var date = new Date();
  date.setMonth(date.getMonth() - 1); // Ek mahina peechay (Previous Month)
  
  var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var targetTabName = months[date.getMonth()] + "-" + date.getFullYear();
  
  Logger.log("Auto Trigger Start: Target Month is " + targetTabName);
  runCompilationLogic(targetTabName);
}


// 2. MANUAL FUNCTION (March ya kisi bhi purane month ka data lane ke liye)
function compileSpecificMonthManual() {
  var ui = SpreadsheetApp.getUi();
  
  // User se month ka naam poochna
  var response = ui.prompt(
    'Manual Month Data Compiler',
    'Kripya us mahine aur saal ka naam enter karein jiska data combine karna hai\n(Format: Month-YYYY, Misaal ke taur par: March-2026):',
    ui.ButtonSet.OK_CANCEL
  );
  
  // Agar user OK par click kare
  if (response.getSelectedButton() == ui.Button.OK) {
    var targetTabName = response.getResponseText().trim();
    
    // Format validation check (taake user ghalat spelling ya format na likh de)
    if (!targetTabName.match(/^[a-zA-Z]+-\d{4}$/)) {
      ui.alert("❌ Ghalat Format!\n\nKripya sahi format use karein, jaise: March-2026\n(Pehla letter capital aur beech mein hyphen '-')");
      return;
    }
    
    // ✅ BUG FIXED: ui.toast ki jagah getActiveSpreadsheet().toast use kiya hai
    SpreadsheetApp.getActiveSpreadsheet().toast("'" + targetTabName + "' ka data compile ho raha hai. Kripya thoda intezar karein...", "Processing", -1);
    
    runCompilationLogic(targetTabName);
  }
}


// 3. CORE LOGIC FUNCTION (Jo asal mein saara kaam karta hai)
function runCompilationLogic(targetTabName) {
  var ssMaster = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var masterSheet = ssMaster.getSheetByName(targetTabName);
  
  // Agar is mahine ka tab Master file mein pehle se nahi bana, to naya tab bana lo
  if (!masterSheet) {
    masterSheet = ssMaster.insertSheet(targetTabName);
  } else {
    // Agar pehle se bana hai, to purana data saaf kar do taake double entries (duplicates) na hon
    masterSheet.clearContents();
    masterSheet.clearFormats();
  }
  
  // Headers lagana naye tab mein
  masterSheet.appendRow(["Titles", "Links", "Post URL", "Publish Date", "Source Website"]);
  masterSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#d9ead3"); // Soft Green Header
  
  var folder = DriveApp.getFolderById(SCRAPER_FOLDER_ID);
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var compiledData = [];
  
  // Folder ki sabhi sheets ko scan karna
  while (files.hasNext()) {
    var file = files.next();
    try {
      var ssSource = SpreadsheetApp.openById(file.getId());
      var siteName = file.getName().replace("-MS", "").trim(); 
      var sourceSheet = ssSource.getSheetByName(targetTabName);
      
      // Agar us website ki sheet mein yeh month tab mojood hai
      if (sourceSheet) {
        var data = sourceSheet.getDataRange().getValues();
        if (data.length > 1) {
          for (var i = 1; i < data.length; i++) {
            var row = data[i];
            var title = row[0];       
            var postUrl = row[1];     
            var gDrive = row[2];      
            var mFire = row[3];       
            var pubDate = row[4];     
            
            if (!title || title.toString().trim() === "") continue;
            
            // Drive Link prefer karein, agar nahi hai to Mediafire
            var finalLink = (gDrive && gDrive.toString().trim() !== "") ? gDrive : mFire;
            
            compiledData.push([title, finalLink, postUrl, pubDate, siteName]);
          }
        }
      }
    } catch (e) {
      Logger.log("Error reading file " + file.getName() + ": " + e.message);
    }
  }
  
  // Master Sheet mein data save karna
  if (compiledData.length > 0) {
    var nextRow = masterSheet.getLastRow() + 1;
    masterSheet.getRange(nextRow, 1, compiledData.length, compiledData[0].length).setValues(compiledData);
    
    // Columns ko thoda khoobsurat aur auto-fit karna
    masterSheet.autoResizeColumns(1, 5);
    
    SpreadsheetApp.getUi().alert("✅ Kamyabi!\n\nMonth Tab: '" + targetTabName + "' tayyar ho gaya hai.\nTotal Novels Added: " + compiledData.length);
  } else {
    SpreadsheetApp.getUi().alert("⚠️ Koi data nahi mila!\n\nWebsites ki sheets mein '" + targetTabName + "' ke naam ka koi tab ya data nahi mila.");
  }
}


// 4. AUTOMATIC MONTHLY TRIGGER FUNCTION
function CREATE_MONTHLY_TRIGGER() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'compilePreviousMonthData') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  ScriptApp.newTrigger('compilePreviousMonthData')
    .timeBased()
    .onMonthDay(1)
    .atHour(1)
    .create();
    
  SpreadsheetApp.getUi().alert("✅ Auto-Trigger Set Ho Gaya Hai!\n\nHar mahine ki 1st tareekh ko pichle mahine ka data automatic naye tab mein add ho jaya karega.");
}