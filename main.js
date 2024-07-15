function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom scripts')
    .addItem('Import news', 'importRSS')
    .addToUi();
}

function importRSS() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.clear();

  const rssUrl = "https://www.pravda.com.ua/rss/view_news/";

  const response = UrlFetchApp.fetch(rssUrl);
  const blob = response.getBlob();

  const text = blob.setDataFromString(blob.getDataAsString('windows-1251'), 'UTF-8');

  const document = XmlService.parse(text.getDataAsString());
  const root = document.getRootElement();
  const channel = root.getChild("channel");
  const items = channel.getChildren("item");

  if (items.length === 0) return;

  const headers = ["Title", "Link", "Category", "Creator", "Publication Date", "Description", "GUID"];
  sheet.appendRow(headers);

  sheet.setFrozenRows(1);

  const formatRange = sheet.getRange(1, 1, 1, headers.length);
  formatRange.setFontWeight("bold");
  formatRange.setBackground("#D3D3D3");
  formatRange.setFontColor("#000000");
  formatRange.setHorizontalAlignment("center");
  
  items.forEach((item, rowIndex) => {
    let title = item.getChild("title").getText();
    let link = item.getChild("link").getText();
    let category = item.getChild("category") ? item.getChild("category").getText() : "";
    const creatorNamespace = XmlService.getNamespace('dc', 'http://purl.org/dc/elements/1.1/');
    let creator = item.getChild("creator", creatorNamespace) ? item.getChild("creator", creatorNamespace).getText() : "";
    let pubDate = item.getChild("pubDate").getText();
    let description = item.getChild("description").getText();
    let guid = item.getChild("guid").getText();

    let range = sheet.getRange(rowIndex + 2, 1, 1, headers.length);
    range.setValues([[title, link, category, creator, pubDate, description, guid]]);
    range.setWrap(true);

    const richText = SpreadsheetApp.newRichTextValue()
      .setText("Link")
      .setLinkUrl(link)
      .build();
    sheet.getRange(rowIndex + 2, 2).setRichTextValue(richText);

    const richTextGUID = SpreadsheetApp.newRichTextValue()
      .setText("GUID")
      .setLinkUrl(guid)
      .build();
    sheet.getRange(rowIndex + 2, 7).setRichTextValue(richTextGUID);
  });

  SpreadsheetApp.flush();
  sheet.autoResizeColumns(1, headers.length);

  sheet.setColumnWidth(2, 31);
  sheet.setColumnWidth(6, 400);
  sheet.setColumnWidth(7, 40);
}
