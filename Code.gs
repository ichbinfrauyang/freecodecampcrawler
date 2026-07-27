/* Triggers:
getArticleLinks()
addADICTsToSheet()
*/
const ss = SpreadsheetApp.getActiveSpreadsheet();
const Urls = ss.getSheetByName('Urls');
const Process = ss.getSheetByName('Process');
const ADICTs = ss.getSheetByName('ADICTs');

const url = "https://www.freecodecamp.org/news";

// Trigger (Every 30 mins ); Extract article Links;
function getArticleLinks() {
  const siteHTML = getHTML(url);
  const $ = Cheerio.load(siteHTML);
  // get @href
  const rowArticleUrls = $('h2').map(function () {
    return 'https://www.freecodecamp.org' + $(this).find('a').attr('href')
  }).get();
  // row URL dedup
  const dedupRowArticleUrls = extractedUrlsDedup(rowArticleUrls);
  // ADD new unique Urls to Processing Sheet
  addUniqueUrlToSheet(dedupRowArticleUrls)
}

// Trigger (Every 5 minutes); Add ADICTs to Sheet 'ADICTs'
function addADICTsToSheet() {
 // Read values in Sheet 'Process' one by one;
 const pendingArticleUrl = Process.getRange(1,1,1,1).getValue();
//  console.log(pendingArticleUrl)
  // getADICTs()
  if (pendingArticleUrl) {
    getADICTs(pendingArticleUrl)
    Process.deleteRow(1)
  }
}

// Extract ADICTs
function getADICTs(pendingArticleUrl) {
  const articleHTML = getHTML(pendingArticleUrl)
  const $ = Cheerio.load(articleHTML);
  const { url, title, date, author, authorDescription, content } = {
    url: $('[property="og:url"]').attr('content'),
    title: $('h1').text(),
    date: $('.post-full-meta-date').text(),
    author: $('.author-card-name').eq(1).text().replace(/\s+/g,''),
    authorDescription: $('[data-test-label="author-bio"]').text(),
    //   sheet.getRange(row,2).setValue(messages[m].getPlainBody().substring(0, 50000));
    content: $('.post-content').text().substring(0,50000).trim()
  }
  // 
  Urls.insertRowBefore(1).getRange(1,1,1,1).setValue([[pendingArticleUrl]])
  // Add adicts to Sheet 'ADICTs'
  ADICTs.insertRowBefore(2).getRange("A2:F2").setValues([[url,title,date,author,authorDescription,content]])
}



function addUniqueUrlToSheet(dedupRowArticleUrls) {
  // 
  const storedUrlInSheetUrls = Process.getDataRange().getValues().map((x) => { return x[0] });
  const urlInSheetProcess = Urls.getDataRange().getValues().map((x) => { return x[0] });
  // Mix value of Sheeet 'Process' & 'Urls';
  const ProcessPlusUrls = [...storedUrlInSheetUrls, ...urlInSheetProcess];
  // Compare values in dedupRowArticleUrls[] and ProcessPlusUrls[], if it's a new value, then, add it to Sheet 'Process';
  for (let i of dedupRowArticleUrls) {
    if (ProcessPlusUrls.indexOf(i) == -1) {
      Process.insertRowBefore(1).getRange(1, 1, 1, 1).setValues([[i]]);
    }
  }
}

// URL Deduplicaion (Array)
function extractedUrlsDedup(urls) {
  const newArr = [... new Set(urls)];
  return newArr
}

// Get site HTML
function getHTML(url) {
  return UrlFetchApp.fetch(url).getContentText()
}
