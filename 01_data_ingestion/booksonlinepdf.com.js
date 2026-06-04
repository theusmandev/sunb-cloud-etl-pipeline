/**
 * Robust WordPress Scraper (baazarofnovels.com)
 * یہ اسکرپٹ اب پچھلے مہینے کا ڈیٹا نکالنے کے لیے اپ ڈیٹ کر دیا گیا ہے
 */

function robustScraper() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const startTime = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. پچھلے مہینے کی تاریخوں کا حساب
  const now = new Date();
  const lastMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const monthName = Utilities.formatDate(lastMonthDate, Session.getScriptTimeZone(), "MMMM-yyyy");
  
  // ورڈ پریس API کے لیے تاریخ کی حدود
  const afterDate = new Date(now.getFullYear(), now.getMonth() - 1, 1).toISOString(); // پچھلے مہینے کی 1 تاریخ
  const beforeDate = new Date(now.getFullYear(), now.getMonth(), 1).toISOString();    // اس مہینے کی 1 تاریخ (حد)
  
  // 2. میموری سے پچھلا اسٹیٹس حاصل کرنا
  let pageNum = parseInt(scriptProperties.getProperty('CURRENT_PAGE')) || 1;
  let totalAdded = parseInt(scriptProperties.getProperty('TOTAL_ADDED')) || 0;
  
  // 3. شیٹ کو چیک کرنا یا بنانا
  let sheet = ss.getSheetByName(monthName);
  if (!sheet) {
    sheet = ss.insertSheet(monthName);
    sheet.appendRow(["Title", "Post URL", "Google Drive Link", "Mediafire Link", "Publish Date"]);
    
    const headerRange = sheet.getRange("A1:E1");
    headerRange.setFontWeight("bold").setBackground("#d9ead3");
    
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 300);
    sheet.setColumnWidth(2, 200);
    console.log("New sheet created for previous month: " + monthName);
  }

  console.log(`Working on Sheet: ${monthName} | Starting from Page: ${pageNum}`);

  while (true) {
    // ٹائم چیک (5 منٹ کی حد)
    const currentTime = new Date().getTime();
    if (currentTime - startTime > 300000) { 
      console.log("Time limit reached. Saving progress...");
      scriptProperties.setProperty('CURRENT_PAGE', pageNum.toString());
      scriptProperties.setProperty('TOTAL_ADDED', totalAdded.toString());
      createResumeTrigger();
      return; 
    }

    // ورڈ پریس API یو آر ایل (after اور before کے ساتھ)
    const apiUrl = `http://booksonlinepdf.com/wp-json/wp/v2/posts?after=${afterDate}&before=${beforeDate}&per_page=100&page=${pageNum}`;
    
    try {
      const response = UrlFetchApp.fetch(apiUrl, {"muteHttpExceptions": true});
      
      if (response.getResponseCode() !== 200) {
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

function finalizeScrape(finalCount, monthName) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('CURRENT_PAGE');
  scriptProperties.deleteProperty('TOTAL_ADDED');
  deleteTriggers();
  console.log(`Finished ${monthName}! Total added: ${finalCount}`);
}

function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'robustScraper') ScriptApp.deleteTrigger(t);
  });
}