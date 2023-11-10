function setHeader(sheet, row, column, columnTitle, columnWidth=120) {
  const headerStyle = SpreadsheetApp.newTextStyle()
    .setBold(true)
    .setFontFamily("Arial")
    .setFontSize(11)
    .build();
  const richText = SpreadsheetApp.newRichTextValue()
    .setText(columnTitle)
    .setTextStyle(headerStyle)
    .build();
  sheet.getRange(row, column)
       .setHorizontalAlignment(DocumentApp.HorizontalAlignment.CENTER)
       .setVerticalAlignment(DocumentApp.VerticalAlignment.TOP)
       .setRichTextValue(richText)
       .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.setColumnWidth(column, columnWidth);
}

function setColumn(sheet, row, column, columnWidth=120) {
  sheet.getRange(row, column)
       .setVerticalAlignment(DocumentApp.VerticalAlignment.TOP)
       .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.setColumnWidth(column, columnWidth);  
}

function setWhenPossible(sheet, row, column, value) {
  const isTextPermanent = sheet.getRange(row, column).getTextStyle().isItalic();
  if (!isTextPermanent) {
    setValue(sheet, row, column, value);
  }
}

function setNote(sheet, row, column, value) {
  sheet.getRange(row, column).setNote(value);
}

function unsetNote(sheet, row, column) {
  sheet.getRange(row, column).clearNote();
}

function setValue(sheet, row, column, value) {
  setValueByRange(sheet.getRange(row, column), value);
}

function setValueByRange(range, value) {
  if (value) {
    range.setFontColor(null)
        .setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT)
        .setVerticalAlignment(DocumentApp.VerticalAlignment.TOP)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        .setValue(value);
  }
}

function setValuesByRow(sheet, row, values, startingColumn=1) {
  setValuesByRange(sheet.getRange(row, startingColumn, values.length, values[0].length), values);
}

function setValuesByColumn(sheet, column, values, startingRow=2) {
  setValuesByRange(sheet.getRange(startingRow, column, values.length, 1), values);
}

function setValuesByRange(range, values) {
  if (values) {
    range.setFontColor(null)
        .setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT)
        .setVerticalAlignment(DocumentApp.VerticalAlignment.TOP)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        .setValues(values);
  }
}

function setRichTextValue(sheet, row, column, value) {
  setRichTextValueByRange(sheet.getRange(row, column), value);
}

function setRichTextValueByRange(range, value) {
  if (value) {
    range.setFontColor(null)
        .setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT)
        .setVerticalAlignment(DocumentApp.VerticalAlignment.TOP)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        .setRichTextValue(value);
  }
}

function setRichTextValuesByRow(sheet, row, values, startingColumn=1) {
  setRichTextValuesByRange(sheet.getRange(row, startingColumn, 1, values[0].length), values);
}

function setRichTextValuesByColumn(sheet, column, values, startingRow=2) {
  setRichTextValuesByRange(sheet.getRange(startingRow, column, values.length, 1), values);
}

function setRichTextValuesByRange(range, values) {
  if (values) {
    range.setFontColor(null)
        .setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT)
        .setVerticalAlignment(DocumentApp.VerticalAlignment.TOP)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        .setRichTextValues(values);
  }
}

function setUrl(sheet, row, column, url, text="") {
  setUrlByRange(sheet.getRange(row, column), url, text);
}

function setUrlByRange(range, url, text="") {
  const richTextUrl = createRichTextUrl(url, (text == "") ? url : text);
  setRichTextUrlByRange(range, richTextUrl);
}

function setRichTextUrl(sheet, row, column, richTextUrl) {
  setRichTextUrlByRange(sheet.getRange(row, column), richTextUrl);
}

function setRichTextUrlByRange(range, richTextUrl) {
  if (richTextUrl) {
    range.setRichTextValue(richTextUrl);
  }
}

function createRichTextUrl(url, text) {
  return SpreadsheetApp.newRichTextValue()
    .setText(text)
    .setLinkUrl(url)
    .build();
}

function setRangeColor(range, colorCode) {
  range.setBackground(colorCode);
}

function setDataValidation(range, valueList, allowInvalid=false) {
  const dataValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(valueList)
      .setAllowInvalid(allowInvalid)
      .build();
  range.setDataValidation(dataValidation);
}

function removeDataValidation(range) {
  range.setDataValidation(null);
}

function setStrictProtection(range, description="Modification is not allowed") {
  const protection = range.protect().setDescription(description);
  const me = Session.getEffectiveUser();
  protection.removeEditors(protection.getEditors().map(user => user.getEmail()))
  protection.addEditor(me)
  protection.setDomainEdit(false);
}

function alert(text) {
  SpreadsheetApp.getUi().alert(text);
}

function prompt(text) {
  return SpreadsheetApp.getUi().prompt(text).getResponseText();
}

function getTimeZone() {
  return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
}

function getValuesFromColumn(sheet, column) {
  const range = sheet.getRange(2, column, sheet.getLastRow(), 1); // skip header
  return range.getDisplayValues().flat();
}

function getFolderId(sheet) {
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const file = DriveApp.getFileById(spreadsheetId);
  return file.getParents().next().getId();
}

function getExportSheetUrl(sheet, format="xlsx") {
  const spreadsheetUrl = sheet.getParent().getUrl().replace(new RegExp('/edit$'), "");
  const sheetId = sheet.getSheetId();
  return `${spreadsheetUrl}/export?gid=${sheetId}&format=${format}`;
}

