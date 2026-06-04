/**
 * Zubi Novels Zone Deep Scraper (Updated for Monthly Run)
 * یہ اسکرپٹ ہر مہینے کی پہلی تاریخ کو پچھلے مہینے کا ڈیٹا ڈیپ اسکین کرتا ہے
 */

function robustZubiDeepScraper() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const startTime = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. پچھلے مہینے کی تاریخوں کا حساب
  const now = new Date();
  const lastMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const monthName = Utilities.formatDate(lastMonthDate, Session.getScriptTimeZone(), "MMMM-yyyy");
  
  // ورڈ پریس API کے لیے تاریخ کی حدود (پچھلا پورا مہینہ)
  const afterDate = new Date(now.getFullYear(), now.getMonth() - 1, 1).toISOString(); // پچھلے مہینے کی 1 تاریخ
  const beforeDate = new Date(now.getFullYear(), now.getMonth(), 1).toISOString();    // اس مہینے کی 1 تاریخ (حد)
  
  // 2. میموری سے پچھلا اسٹیٹس حاصل کرنا
  let currentPage = parseInt(scriptProperties.getProperty('ZUBI_PAGE')) || 1;
  let totalAdded = parseInt(scriptProperties.getProperty('ZUBI_TOTAL')) || 0;
  
  // 3. شیٹ چیک کرنا یا بنانا
  let sheet = ss.getSheetByName(monthName);
  if (!sheet) {
    sheet = ss.insertSheet(monthName);
    sheet.appendRow(["Title", "Post URL", "Google Drive Link", "Mediafire Link", "Publish Date"]);
    sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#ffd966");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 300);
    sheet.setColumnWidth(2, 200);
    console.log("New sheet created for Zubi Deep Scan: " + monthName);
  }

  while (true) {
    // 5 منٹ کا ٹائم لمٹ چیک (ڈیپ اسکین میں ٹائم زیادہ لگتا ہے اس لیے یہ بہت ضروری ہے)
    if (new Date().getTime() - startTime > 300000) { 
      console.log("Time limit reached. Saving Zubi progress at page: " + currentPage);
      scriptProperties.setProperty('ZUBI_PAGE', currentPage.toString());
      scriptProperties.setProperty('ZUBI_TOTAL', totalAdded.toString());
      createZubiTrigger();
      return; 
    }

    // API URL جس میں پچھلے مہینے کی حد (after & before) شامل ہے
    const apiUrl = `https://zubinovelszone.com/wp-json/wp/v2/posts?after=${afterDate}&before=${beforeDate}&per_page=10&page=${currentPage}&_fields=title,link,date`;
    
    try {
      const response = UrlFetchApp.fetch(apiUrl, {"muteHttpExceptions": true});
      if (response.getResponseCode() !== 200) {
        finalizeZubiScrape(totalAdded, monthName);
        break;
      }

      const posts = JSON.parse(response.getContentText());
      if (!posts || posts.length === 0) {
        finalizeZubiScrape(totalAdded, monthName);
        break;
      }

      const existingLinks = sheet.getLastRow() > 1 ? 
          sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().flat() : [];

      posts.forEach(post => {
        if (!existingLinks.includes(post.link)) {
          console.log("Deep Scanning: " + post.link);
          
          // پوسٹ کا مکمل HTML حاصل کرنا
          const pageHtml = fetchFullHtml(post.link);
          
          // ریگولر ایکسپریشن کے ذریعے لنکس نکالنا
          const driveLinks = extractLinksFromHtml(pageHtml, /https?:\/\/drive\.google\.com\/[^\s"'>]+/g);
          const mediafireLinks = extractLinksFromHtml(pageHtml, /https?:\/\/(?:www\.)?mediafire\.com\/[^\s"'>]+/g);

          sheet.appendRow([
            post.title.rendered, 
            post.link, 
            driveLinks.join("\n"), 
            mediafireLinks.join("\n"), 
            new Date(post.date).toLocaleDateString()
          ]);
          totalAdded++;
          
          // سرور کو بلاک ہونے سے بچانے کے لیے تھوڑا وقفہ
          Utilities.sleep(800); 
        }
      });

      console.log(`Zubi Page ${currentPage} processed.`);
      currentPage++;
      scriptProperties.setProperty('ZUBI_PAGE', currentPage.toString());
      scriptProperties.setProperty('ZUBI_TOTAL', totalAdded.toString());

    } catch (e) {
      console.log("Zubi Error: " + e.toString());
      break;
    }
  }
}

// پوسٹ کا مکمل HTML حاصل کرنے کا فنکشن
function fetchFullHtml(url) {
  try {
    const response = UrlFetchApp.fetch(url, {
      "muteHttpExceptions": true,
      "headers": {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
      }
    });
    return response.getContentText();
  } catch (e) {
    return "";
  }
}

// HTML میں سے لنکس صاف کر کے نکالنے کا فنکشن
function extractLinksFromHtml(html, regex) {
  if (!html) return [];
  let links = [];
  let match;
  while ((match = regex.exec(html)) !== null) {
    let link = match[0].split(/[ "'>]/)[0]; 
    if (link.length > 15) {
       links.push(link);
    }
  }
  return [...new Set(links)];
}

function createZubiTrigger() {
  deleteZubiTriggers();
  ScriptApp.newTrigger('robustZubiDeepScraper').timeBased().after(60000).create();
}

function finalizeZubiScrape(finalCount, monthName) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('ZUBI_PAGE');
  scriptProperties.deleteProperty('ZUBI_TOTAL');
  deleteZubiTriggers();
  console.log(`Finished Zubi Deep Scan for ${monthName}! Total added: ${finalCount}`);
}

function deleteZubiTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'robustZubiDeepScraper') ScriptApp.deleteTrigger(t);
  });
}