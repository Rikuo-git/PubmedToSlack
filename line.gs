/**
 * Main function that retrieves RSS feeds and webhook URLs from a Google Spreadsheet,
 * checks for new papers, and sends notifications to the corresponding Slack channels.
 *
 * @function
 */
function pubmedToLINE() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rssSheet = ss.getSheetByName("RSS").getDataRange();
  const tokenData = ss.getSheetByName("tokens").getDataRange().getValues()
    .reduce((dict, row, index) => {
      if (index === 0) return dict;
      dict[row[0]] = row[1];
      return dict;
    }, {});

  const result = rssSheet.getValues()
    .map((rssData, index) => {
      if (index === 0) return rssData;
      rssData[3] = fetchAndPostNewPapersLINE(rssData, tokenData);
      return rssData;
    });

  rssSheet.setValues(result);
}

/**
 * Fetches new papers from an RSS feed and sends the results to LINE notification.
 * 
 * @function
 * @param {Array} rssData - An array containing the keyword, RSS URL, target, and history.
 * @param {Object} tokenData - An object containing the LINE Notify API access token. 
 * @returns {string} - A JSON string representing the history of PubMed IDs that have been posted.
 */
function fetchAndPostNewPapersLINE(rssData, tokenData) {
  const [keyword, rssUrl, target, history] = rssData;
  const token = tokenData[target];
  if (!token) return history;

  const historySet = history ? new Set(JSON.parse(history)) : new Set();
  const rssResponse = UrlFetchApp.fetch(rssUrl);
  if (rssResponse.getResponseCode() !== 200) return history;

  const rssContent = XmlService.parse(rssResponse.getContentText());
  const newPapers = rssContent.getRootElement().getChild('channel').getChildren('item').reverse()
    .filter(item => !historySet.has(item.getChildText('guid')) && historySet.add(item.getChildText('guid')));

  if (!newPapers.length) return history;

  const title = `\nHere are ${newPapers.length} new papers for ${keyword}\n`;
  const blocks = newPapers.map(getInfoFromRSSLINE).reduce(textToBlocks, [title]);
  const lineResponse = blocks.map((text,index) => sendLineNotification(text, token,index !== 0)).every(r => r.getResponseCode() === 200)
  return lineResponse ? JSON.stringify([...historySet]) : history;
}

/**
 * Extracts and formats information from an RSS item element for display.(LINEversion)
 * 
 * @function
 * @param {GoogleAppsScript.XML_Service.Element} item - An XML Element object representing an item from an RSS feed.
 * @returns {string} - A formatted string containing the item's information in mrkdwn format, including a link to the item, title, translated title, author(s), journal, and date.
 */
function getInfoFromRSSLINE(item) {
  const nsDc = XmlService.getNamespace("dc", "http://purl.org/dc/elements/1.1/");
  const pmid = item.getChildText('guid').replace("pubmed:", "");
  const link = `https://pubmed.ncbi.nlm.nih.gov/${pmid}/`;
  const title = lineEscape(item.getChildText('title').replace(/<\/?(em|sup|sub)>/g, ""));
  const titleJa = LanguageApp.translate(title, 'en', 'ja');
  const authors = item.getChildren('creator', nsDc);
  let author = authors.length ? `${authors[0].getText()}${authors.length > 1 ? ", et al." : ""}` : "No Author";
  const journal = item.getChildText('source', nsDc);
  const date = item.getChildText('date', nsDc);
  return `\n${title}\n${titleJa}\n${author} ${journal} ${date} ${link}\n`
}

/**
 * Appends text to the last element of the blocks array. If the last element's length plus the text length
 * exceeds 1000 characters, the text will be added as a new element in the array.
 * 
 * @function
 * @param {Array} blocks - An array of strings, where each element represents a block of text.
 * @param {string} text - The text to be appended to the last element of the blocks array.
 * @returns {Array} - The updated array of blocks with the text appended to the last element or added as a new element.
 */
function textToBlocks(blocks, text) {
  if (blocks[blocks.length - 1].length + text.length > 1000) {
    blocks.push(text);
  } else {
    blocks[blocks.length - 1] += text;
  }
  return blocks;
}


/**
 * Escapes special characters in a string for use in LINE notification.
 * 
 * @function
 * @param {string} string - The input string that contains special characters to be escaped.
 * @returns {string} - A new string with special characters escaped for use in  LINE notification.
 */
function lineEscape(string) {
  return string.replace(/(&amp;|&lt;|&gt;)/g, function (match) {
    return {
      '&amp;': '&',
      '&lt;': '<',
      '&gt;': '>',
    }[match];
  });
}

/**
 * Sends a text message notification to a LINE account using the LINE Notify API.
 * 
 * @function
 * @param {string} text - The text message to be sent as a LINE notification.
 * @param {string} token - The LINE Notify API access token.
 * @returns {HTTPResponse} - The HTTP response returned by the LINE Notify API after sending the notification.
 */
function sendLineNotification(text, token, disableNotification = false) {
  const url = "https://notify-api.line.me/api/notify";
  const options = {
    method: "post",
    headers: {
      Authorization: "Bearer " + token,
    },
    payload: {
      message: text,
      notificationDisabled : disableNotification
    },
    muteHttpExceptions: true
  };
  return UrlFetchApp.fetch(url, options,);
}