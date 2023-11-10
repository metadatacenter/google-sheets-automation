function getValueSetSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(".VALUESETS");
}

function searchValueSet(fieldName, data) {
  const valueSetSheet = getValueSetSheet();
  const matchIndex = data.findIndex(row => row[0] == fieldName);
  const success = matchIndex != -1;
  const row = matchIndex + 2;
  return {
    success,
    data: success ? {
      row,
      link: generateLinkToRange(valueSetSheet.getRange(row, VALUESET_GLOSSARY_CATEGORY)),
      acronym: valueSetSheet.getRange(row, VALUESET_GLOSSARY_MAPPED_TO_ONTOLOGY_ACRONYM).getValue(),
      lookup: expandPrefixedName(valueSetSheet.getRange(row, VALUESET_GLOSSARY_MAPPED_TO_TERM_ID).getValue())
    } : null,
  };
}

function expandPrefixedName(value) {
  if (value == "") {
    return null;
  } else {
    const splitWords = value.split(':');
    const prefix = splitWords[0];
    const label = splitWords[1];
    return getPrefixMap()[prefix] + label;
  }
}

function getValueSetData(minimal=false) {
  const valueSetSheet = getValueSetSheet();
  const data = valueSetSheet.getRange(2, VALUESET_GLOSSARY_CATEGORY, valueSetSheet.getLastRow(), 4).getValues();
  return (minimal) ? data.map(row => [row[0]]) : data;
}

function storeValueSet(fieldName, permissibleValues) {
  const valueSetSheet = getValueSetSheet();
  // Use the field name as the value set's category name and place it in its own row.
  const categoryRow = valueSetSheet.getLastRow() + 1;
  const categoryRange = valueSetSheet.getRange(categoryRow, VALUESET_GLOSSARY_CATEGORY);
  setValueByRange(categoryRange, fieldName);

  // Place each permissible value in the subsequent rows.
  const valueRow = categoryRow + 1;
  const values = permissibleValues.split('\n').map(item => [item]);
  setValuesByColumn(valueSetSheet, VALUESET_GLOSSARY_VALUE_LABEL, values, startingRow=valueRow);

  return generateLinkToRange(categoryRange);
}

function generateLinkToRange(range) {
  const sheet = range.getSheet();
  const sheetId = sheet.getSheetId();
  const spreadsheetUrl = sheet.getParent().getUrl();
  const cellNotation = range.getA1Notation();
  return `${spreadsheetUrl}#gid=${sheetId}&range=${cellNotation}`;
}
