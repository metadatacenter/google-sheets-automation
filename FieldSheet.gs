function getFieldSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(".FIELDS");
}

function searchField(fieldName, data=getFieldData(minimal=true)) {
  const fieldSheet = getFieldSheet();
  const matchIndex = data.findIndex(row => row[0] == fieldName);
  const success = matchIndex != -1;
  const row = matchIndex + 2;
  return {
    success,
    data: success ? {
      row,
      range: fieldSheet.getRange(row, FIELD_GLOSSARY_MAPPED_TO, 1, 8),
    } : null,
  }
}

function findFieldId(fieldName, data) {
  const searchResult = searchField(fieldName, data);
  return searchResult.success ? searchResult.data.range.getCell(1, 1).getValue() : null;
}

function getField(whichRow) {
  const fieldSheet = getFieldSheet();
  return fieldSheet.getRange(whichRow, FIELD_GLOSSARY_FIELD_NAME, 1, 7)
      .getDisplayValues()
      .flat();
}

function getFieldList() {
  const fieldSheet = getFieldSheet();
  return fieldSheet.getRange(2, FIELD_GLOSSARY_FIELD_NAME, fieldSheet.getLastRow(), 1)
      .getRichTextValues()
      .map(item => {
        const value = item[0];
        return value.getTextStyle().isStrikethrough() ? "" : value.getText();
      })
      .flat();
}

function getFieldData(minimal=false) {
  const fieldSheet = getFieldSheet();
  const data = fieldSheet.getRange(2, FIELD_GLOSSARY_FIELD_NAME, fieldSheet.getLastRow(), 7).getValues();
  return (minimal) ? data.map(row => [row[0]]) : data;
}

function replaceField(oldName, newName) {
  const fieldSheet = getFieldSheet();
  const result = searchField(oldName);
  if (result.success) {
    const rowIndex = result.data.row;
    const field = result.data.field;
    field.splice(FIELD_GLOSSARY_FIELD_NAME-1, 1, newName);
    fieldSheet.getRange(rowIndex, FIELD_GLOSSARY_REPLACED_BY).setValue(newName);
    fieldSheet.insertRowAfter(rowIndex)
        .getRange(rowIndex + 1,FIELD_GLOSSARY_FIELD_NAME, 1, FIELD_GLOSSARY_COLUMNS)
        .setValues([field]);
  }
}
