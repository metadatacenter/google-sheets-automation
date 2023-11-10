function onOpen(e) {
  const ui = SpreadsheetApp.getUi(); 
  ui.createMenu('🤖 Automation')
    .addItem('New metadata specification...', 'createNew')
    .addSeparator()
    .addItem('↓ Import', 'importFields')
    .addItem('↑ Export', 'exportFields')
    .addSeparator()
    .addItem('Synchronize', 'synchronizeFields')
    .addSeparator()
    .addItem('Rename a field...', 'none')
    .addSeparator()
    .addItem('Convert to CEDAR template...', 'none')
    .addItem('Convert to LinkML schema...', 'generateLinkMl')
    .addSeparator()
    .addItem('Generate SKOS vocabulary...', 'generateSkos')
    .addSeparator()
    .addItem('Publish metadata specification to CEDAR', 'none')
    .addItem('Publish metadata vocabulary to BioPortal', 'none')
    .addItem('Configure...', 'none')
    .addToUi(); 
};

function none() {
  alert('Not implemented yet');
}

function onEdit(e) { 
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getSheetName();

  if (!sheetName.startsWith(".")) {
    const range = e.range;
    const lastColumn = range.getLastColumn();
    if (lastColumn === TEMPLATE_FIELD_NAME) {
      const startColumn = range.getColumn();
      const numRows = range.getNumRows();
      [...Array(numRows).keys()].forEach((rowIndex) => {
        if (startColumn === 1) {
          autoFillOut(range.offset(rowIndex, 1, 1, 1));
        } else if (startColumn === 2) {
          autoFillOut(range.offset(rowIndex, 0, 1, 1));
        }
      });
    }
  }
}

function autoFillOut(cell) {
  const templateColumnIndex = cell.getColumn();
  if (templateColumnIndex != TEMPLATE_FIELD_NAME) {
    return;
  }
  const fieldName = cell.getValue();
  if (fieldName == "") {
    return;
  }
  const rowIndex = cell.getRow();

  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  const templateSheet = ss.getActiveSheet();

  const searchResult = searchField(fieldName);
  if (searchResult.success) {
    Logger.log("Copying '" + fieldName + "' definition to the template sheet");
    const fieldDefinitionRange = searchResult.data.range.offset(0, 1);
    setValuesByRow(templateSheet, rowIndex, fieldDefinitionRange.getValues(), startingColumn=TEMPLATE_FIELD_NAME);
    setRichTextUrl(templateSheet, rowIndex, TEMPLATE_PERMISSIBLE_VALUES, fieldDefinitionRange.offset(0, 2, 1, 1).getRichTextValue()); // view link
    setRichTextUrl(templateSheet, rowIndex, TEMPLATE_VALUE_SET_IRI, fieldDefinitionRange.offset(0, 3, 1, 1).getRichTextValue()); // lookup link
  } else { 
    setValuesByRow(templateSheet, rowIndex, [["","","","","","",""]], startingColumn=TEMPLATE_FIELD_DESCRIPTION);
  }
}

function showReport(title, htmlText) {
  const message = HtmlService.createHtmlOutput(htmlText)
    .setWidth(300)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(message, title);
}

function asList(str) {
  return `<li>${str}</li>`;
}
