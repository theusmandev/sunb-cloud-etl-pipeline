/**
 * Robust Scraper - Last Month Data (Runs on 1st of every month)
 */

function robustScraper() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const startTime = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Pichle mahine (Last Month) ki dates set karna
  const now = new Date();
  // Current month ki 1 tareekh (End date for API)
  const firstDayOfCurrentMonth = new Date(now.getFullYear(), now.getMonth(), 1); 
  // Pichle mahine ki 1 tareekh (Start date for API)
  const firstDayOfLastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1); 

  // Sheet ka naam pichle mahine ke hisaab se (e.g., April-2026)
  const monthName = Utilities.formatDate(firstDayOfLastMonth, Session.getScriptTimeZone(), "MMMM-yyyy");
  
  const afterDate = firstDayOfLastMonth.toISOString();
  const beforeDate = firstDayOfCurrentMonth.toISOString();
  
  // 2. Memory se pichla status hasil karna
  let pageNum = parseInt(scriptProperties.getProperty('CURRENT_PAGE')) || 1;
  let totalAdded = parseInt(scriptProperties.getProperty('TOTAL_ADDED')) || 0;
  
  // 3. Sheet check karna (Agar nahi hai toh nayi banayega)
  let sheet = ss.getSheetByName(monthName);
  if (!sheet) {
    sheet = ss.insertSheet(monthName);
    sheet.appendRow(["Title", "Post URL", "Google Drive Link", "Mediafire Link", "Publish Date"]);
    
    // Formatting
    const headerRange = sheet.getRange("A1:E1");
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#d9ead3");
    
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 300);
    sheet.setColumnWidth(2, 200);
    console.log("New sheet created: " + monthName);
  }

  console.log(`Working on Last Month: ${monthName} | Starting from Page: ${pageNum}`);

  while (true) {
    // Time limit check (5 mins)
    const currentTime = new Date().getTime();
    if (currentTime - startTime > 300000) { 
      console.log("Time limit reached. Saving progress...");
      scriptProperties.setProperty('CURRENT_PAGE', pageNum.toString());
      scriptProperties.setProperty('TOTAL_ADDED', totalAdded.toString());
      createResumeTrigger();
      return; 
    }

    // API URL updated with both 'after' and 'before' for last month
    const apiUrl = `https://baazarofnovels.com/wp-json/wp/v2/posts?after=${afterDate}&before=${beforeDate}&per_page=100&page=${pageNum}`;
    
    try {
      const response = UrlFetchApp.fetch(apiUrl, {"muteHttpExceptions": true});
      
      if (response.getResponseCode() !== 200) {
        console.log("No more data or API error. Finalizing...");
        finalizeScrape(totalAdded, monthName);
        break;
      }

      const posts = JSON.parse(response.getContentText());
      if (posts.length === 0) {
        finalizeScrape(totalAdded, monthName);
        break;
      }

      const existingLinks = sheet.getLastRow() > 1 ? 
          sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().flat() : [];

      posts.forEach(post => {
        if (!existingLinks.includes(post.link)) {
          const content = post.content.rendered;
          const driveLinks = [];
          const mediafireLinks = [];
          
          const driveRegex = /href="(https?:\/\/drive\.google\.com\/[^"\s>]+)"/g;
          const mfRegex = /href="(https?:\/\/(?:www\.)?mediafire\.com\/[^"\s>]+)"/g;
          
          let match;
          while ((match = driveRegex.exec(content)) !== null) driveLinks.push(match[1]);
          while ((match = mfRegex.exec(content)) !== null) mediafireLinks.push(match[1]);

          sheet.appendRow([
            post.title.rendered, 
            post.link, 
            [...new Set(driveLinks)].join("\n"), 
            [...new Set(mediafireLinks)].join("\n"), 
            new Date(post.date).toLocaleDateString()
          ]);
          totalAdded++;
        }
      });

      console.log(`Page ${pageNum} processed.`);
      pageNum++;
      scriptProperties.setProperty('CURRENT_PAGE', pageNum.toString());
      scriptProperties.setProperty('TOTAL_ADDED', totalAdded.toString());
      Utilities.sleep(1000);

    } catch (e) {
      console.log("Stop: " + e.toString());
      break;
    }
  }
}

function createResumeTrigger() {
  deleteTriggers();
  ScriptApp.newTrigger('robustScraper').timeBased().after(60000).create();
}

function finalizeScrape(finalCount, sheetName) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('CURRENT_PAGE');
  scriptProperties.deleteProperty('TOTAL_ADDED');
  deleteTriggers();
  console.log(`Finished! Total added: ${finalCount}`);
}

function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'robustScraper') ScriptApp.deleteTrigger(t);
  });
}