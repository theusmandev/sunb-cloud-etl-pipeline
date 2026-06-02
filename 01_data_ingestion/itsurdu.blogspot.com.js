/**
 * Robust Blogger Scraper (itsurdu.blogspot.com)
 * یہ اسکرپٹ ہر مہینے کی پہلی تاریخ کو چلے گا اور پچھلے پورے مہینے کا ڈیٹا نکالے گا
 */

function robustBloggerScraper() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const startTime = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. تاریخ کا حساب (پچھلا مہینہ)
  const now = new Date();
  const firstDayOfLastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const firstDayOfCurrentMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  
  const monthName = Utilities.formatDate(firstDayOfLastMonth, Session.getScriptTimeZone(), "MMMM-yyyy");
  
  // بلاگر کے لیے تاریخ کی حدود (ISO Format)
  const publishedMin = firstDayOfLastMonth.toISOString(); 
  const publishedMax = firstDayOfCurrentMonth.toISOString(); 
  
  // 2. میموری سے پچھلا اسٹیٹس حاصل کرنا
  let startIndex = parseInt(scriptProperties.getProperty('START_INDEX')) || 1;
  let totalAdded = parseInt(scriptProperties.getProperty('TOTAL_ADDED')) || 0;
  
  // 3. شیٹ چیک کرنا یا بنانا
  let sheet = ss.getSheetByName(monthName);
  if (!sheet) {
    sheet = ss.insertSheet(monthName);
    sheet.appendRow(["Title", "Post URL", "Google Drive Link", "Mediafire Link", "Publish Date"]);
    const headerRange = sheet.getRange("A1:E1");
    headerRange.setFontWeight("bold").setBackground("#fff2cc");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 300);
    sheet.setColumnWidth(2, 200);
    console.log("New sheet created: " + monthName);
  }

  console.log(`Starting Scrape for: ${monthName} | Start Index: ${startIndex}`);

  while (true) {
    // ٹائم چیک (5 منٹ کی حد تاکہ گوگل اسکرپٹ کریش نہ ہو)
    if (new Date().getTime() - startTime > 300000) { 
      console.log("Time limit reached. Saving Index: " + startIndex);
      scriptProperties.setProperty('START_INDEX', startIndex.toString());
      scriptProperties.setProperty('TOTAL_ADDED', totalAdded.toString());
      createBloggerTrigger();
      return; 
    }

    // بلاگر فیڈ یو آر ایل
    const feedUrl = `https://itsurdu.blogspot.com/feeds/posts/default?alt=json&published-min=${publishedMin}&published-max=${publishedMax}&max-results=50&start-index=${startIndex}`;
    
    try {
      const response = UrlFetchApp.fetch(feedUrl, {"muteHttpExceptions": true});
      if (response.getResponseCode() !== 200) {
        console.log("API Error or End of Data.");
        finalizeBloggerScrape(totalAdded, monthName);
        break;
      }

      const json = JSON.parse(response.getContentText());
      const entries = json.feed.entry;

      // اگر مزید پوسٹس نہیں ملیں تو لوپ بند کر دیں
      if (!entries || entries.length === 0) {
        finalizeBloggerScrape(totalAdded, monthName);
        break;
      }

      const existingLinks = sheet.getLastRow() > 1 ? 
          sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().flat() : [];

      entries.forEach(entry => {
        let postUrl = "";
        for (let i = 0; i < entry.link.length; i++) {
          if (entry.link[i].rel === 'alternate') {
            postUrl = entry.link[i].href;
            break;
          }
        }

        if (!existingLinks.includes(postUrl)) {
          const title = entry.title.$t;
          const content = entry.content.$t;
          const pubDate = new Date(entry.published.$t).toLocaleDateString();

          const driveLinks = [];
          const mediafireLinks = [];
          
          // Regex اپڈیٹ: اب یہ سنگل (') اور ڈبل (") دونوں کوٹس کو کیچ کرے گا
          const driveRegex = /href=["'](https?:\/\/drive\.google\.com\/[^"'\s>]+)["']/g;
          const mfRegex = /href=["'](https?:\/\/(?:www\.)?mediafire\.com\/[^"'\s>]+)["']/g;
          
          let match;
          while ((match = driveRegex.exec(content)) !== null) driveLinks.push(match[1]);
          while ((match = mfRegex.exec(content)) !== null) mediafireLinks.push(match[1]);

          sheet.appendRow([
            title, 
            postUrl, 
            [...new Set(driveLinks)].join("\n"), 
            [...new Set(mediafireLinks)].join("\n"), 
            pubDate
          ]);
          totalAdded++;
        }
      });

      // اگلے پیج کے لیے انڈیکس کو اپڈیٹ کریں
      startIndex += entries.length;
      scriptProperties.setProperty('START_INDEX', startIndex.toString());
      scriptProperties.setProperty('TOTAL_ADDED', totalAdded.toString());
      Utilities.sleep(1000);

    } catch (e) {
      console.log("Error: " + e.toString());
      break;
    }
  }
}

function finalizeBloggerScrape(finalCount, monthName) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('START_INDEX');
  scriptProperties.deleteProperty('TOTAL_ADDED');
  deleteBloggerTriggers();
  console.log(`Finished ${monthName}! Total added: ${finalCount}`);
}

function createBloggerTrigger() {
  deleteBloggerTriggers();
  ScriptApp.newTrigger('robustBloggerScraper').timeBased().after(60000).create();
}

function deleteBloggerTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'robustBloggerScraper') ScriptApp.deleteTrigger(t);
  });
}