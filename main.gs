/**
 * Main function that retrieves RSS feeds and webhook URLs from a Google Spreadsheet,
 * checks for new papers, and sends notifications to the corresponding Slack channels.
 *
 * @function
 */
function pubmedToSlack() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rssSheet = ss.getSheetByName("RSS").getDataRange();
  const webhookData = ss.getSheetByName("webhooks").getDataRange().getValues()
    .reduce((dict, row, index) => {
      if (index === 0) return dict;
      dict[row[0]] = row[1];
      return dict;
    }, {});

  const result = rssSheet.getValues()
    .map((rssData, index) => {
      if (index === 0) return rssData;
      rssData[3] = fetchAndPostNewPapers(rssData, webhookData);
      return rssData;
    });

  rssSheet.setValues(result);
}

/**
 * Fetches new papers from an RSS feed and sends the results to a Slack channel.
 * 
 * @function
 * @param {Array} rssData - An array containing the keyword, RSS URL, target, and history.
 * @param {Object} webhookData - An object containing the webhook URL for the target Slack channel.
 * @returns {string} - A JSON string representing the history of PubMed IDs that have been posted.
 */
function fetchAndPostNewPapers(rssData, webhookData) {
  const [keyword, rssUrl, target, history] = rssData;
  const webhookUrl = webhookData[target];
  if (!webhookUrl) return history;
  
  const historySet = history ? new Set(JSON.parse(history)) : new Set()
  const rssResponse = UrlFetchApp.fetch(rssUrl,{muteHttpExceptions:true});

  if (rssResponse.getResponseCode() !== 200) return history;

  const rssContent = XmlService.parse(rssResponse.getContentText());
  const newPapers = rssContent.getRootElement().getChild('channel').getChildren('item').reverse()
    .filter(item => !historySet.has(item.getChildText('guid')) && historySet.add(item.getChildText('guid')));

  if (!newPapers.length) return history;

  const title = `Here are ${newPapers.length} new papers for *${keyword}* :eyes:`;
  const blocks = [{
    type: "section",
    text: {
      type: "mrkdwn",
      text: title
    }
  }];
  blocks.push(...newPapers.map(getInfoFromRSS).reduce(appendPapersToBlocks, []));
  const slackResponse = sendSlackMsg(title, blocks, webhookUrl);
  return slackResponse.getResponseCode() === 200 ? JSON.stringify([...historySet]) : history;
}

/**
 * Extracts and formats information from an RSS item element for display.
 * 
 * @function
 * @param {GoogleAppsScript.XML_Service.Element} item - An XML Element object representing an item from an RSS feed.
 * @returns {string} - A formatted string containing the item's information in mrkdwn format, including a link to the item, title, translated title, author(s), journal, and date.
 */
function getInfoFromRSS(item) {
  const nsDc = XmlService.getNamespace("dc", "http://purl.org/dc/elements/1.1/");
  const pmid = item.getChildText('guid').replace("pubmed:", "");
  const link = `https://pubmed.ncbi.nlm.nih.gov/${pmid}/`;
  const title = mrkdwnEscape(item.getChildText('title').replace(/<\/?(em|i)>/g, "_").replace(/<\/?(sup|sub)>/g, ""));
  const titleJa = LanguageApp.translate(title, 'en', 'ja');
  const authors = item.getChildren('creator', nsDc);
  let author = authors.length ? `${authors[0].getText()}${authors.length > 1 ? ", _et al._" : ""}` : "No Author";
  const journal = item.getChildText('source', nsDc);
  const date = item.getChildText('date', nsDc);
  return `><${link}|*${title}*>\n>${titleJa}\n>${author} _*${journal}*_ ${date}\n  \n`;
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
function appendPapersToBlocks(blocks, text, index) {
  if (index % 5 === 0) {
    blocks.push({
      type: "section",
      text: {
        type: "mrkdwn",
        text: ""
      }
    });
  }
  blocks[blocks.length - 1].text.text += text;
  return blocks;
}

/**
 * Sends a message with a title and blocks to a Slack channel using a webhook URL.
 *
 * @function
 * @param {string} title - The title of the message to be sent to Slack.
 * @param {Array} blocks - An array of block objects to be included in the Slack message.
 * @param {string} webhookUrl - The webhook URL for the Slack channel where the message will be sent.
 * @returns {HTTPResponse} - The HTTP response object returned by the UrlFetchApp.fetch method.
 */
function sendSlackMsg(title, blocks, webhookUrl) {
  const payload = {
    text: title,
    blocks: blocks
  };
  const options = {
    method: "post",
    headers: { "Content-type": "application/json" },
    payload: JSON.stringify(payload),
    muteHttpExceptions:true
  };
  return UrlFetchApp.fetch(webhookUrl, options);
}

/**
 * Escapes special characters in a string for use in mrkdwn format.
 * 
 * @function
 * @param {string} string - The input string that contains special characters to be escaped.
 * @returns {string} - A new string with special characters escaped for use in mrkdwn.
 */
function mrkdwnEscape(string) {
  return string.replace(/[&<>]/g, function (match) {
    return {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
    }[match];
  });
}
