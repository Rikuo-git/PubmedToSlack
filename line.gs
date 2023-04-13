/**
 * setting
 */
SLACK_NOTIFY = true
LINE_NOTIFY = true
LINE_TOKEN = ""

/**
 * Main function that retrieves RSS feeds and webhook URLs from a Google Spreadsheet,
 * checks for new papers, and sends notifications to the corresponding Slack channels.
 *
 * @function
 */
function line() {
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
      rssData[3] = linefetchAndPostNewPapers(rssData, tokenData);
      return rssData;
    });

  rssSheet.setValues(result);
}

/**
 * Fetches new papers from an RSS feed and sends the results to a Slack channel.
 * 
 * @function
 * @param {Array} rssData - An array containing the keyword, RSS URL, target, and history.
 * @param {Object} tokenData - An object containing the LINE Notify API access token. 
 * @returns {string} - A JSON string representing the history of PubMed IDs that have been posted.
 */
function linefetchAndPostNewPapers(rssData, tokenData) {
  const [keyword, rssUrl, target, history] = rssData;
  const token = tokenData[target];
  const historySet = history ? new Set(JSON.parse(history)) : new Set()
  const rssResponse = UrlFetchApp.fetch(rssUrl);

  if (rssResponse.getResponseCode() !== 200) return history;

  const rssContent = XmlService.parse(rssResponse.getContentText());
  const newPapers = rssContent.getRootElement().getChild('channel').getChildren('item').reverse()
    .filter(item => !historySet.has(item.getChildText('guid')) && historySet.add(item.getChildText('guid')));

  if (!newPapers.length) return history;

  const title = `\nHere are ${newPapers.length} new papers for ${keyword}\n`;
  const blocks = newPapers.map(linegetInfoFromRSS).reduce(lineappendPapersToBlocks,[title])
  const slackResponse = blocks.map(text => sendLineNotification(text,token)).every(r => r.getResponseCode() === 200)
  return slackResponse ? JSON.stringify([...historySet]) : history;
}

/**
 * Extracts and formats information from an RSS item element for display.
 * 
 * @function
 * @param {GoogleAppsScript.XML_Service.Element} item - An XML Element object representing an item from an RSS feed.
 * @returns {string} - A formatted string containing the item's information in mrkdwn format, including a link to the item, title, translated title, author(s), journal, and date.
 */
function linegetInfoFromRSS(item) {
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
 * Appends a formatted text to blocks and separates sections every 5 items.
 * 
 * @function
 * @param {Array} blocks - An array of block objects to which the formatted text will be appended.
 * @param {string} text - The formatted text that needs to be appended to the blocks.
 * @param {number} index - The current index of the item being processed, used to determine if a new section should be created.
 * @returns {Array} - An updated array of block objects with the formatted text appended.
 */
function lineappendPapersToBlocks(blocks, text, index) {
  if (blocks[blocks.length - 1].length + text.length > 1000) {
    blocks.push(text)
  } else {
    blocks[blocks.length - 1] += text;
  }
  return blocks;
}

/**
 * Escapes special characters in a string for use in mrkdwn format.
 * 
 * @function
 * @param {string} string - The input string that contains special characters to be escaped.
 * @returns {string} - A new string with special characters escaped for use in mrkdwn.
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
function sendLineNotification(text, token) {
  const url = "https://notify-api.line.me/api/notify";
  const options = {
    method: "post",
    headers: {
      Authorization: "Bearer " + token,
    },
    payload: "message=" + text,
    muteHttpExceptions:true
  };
  return UrlFetchApp.fetch(url, options,);
}