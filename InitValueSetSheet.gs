function initvalueSetSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const valueSetSheet = ss.insertSheet(sheetName, ss.getActiveSheet().getIndex());
  setHeader(valueSetSheet, 1, VALUESET_GLOSSARY_MAPPED_TO_ONTOLOGY_ACRONYM, "", 80);
  setHeader(valueSetSheet, 1, VALUESET_GLOSSARY_MAPPED_TO_TERM_ID, "Mapped To", 150);
  valueSetSheet.getRange(1, 1, 1, 2).mergeAcross();
  setHeader(valueSetSheet, 1, VALUESET_GLOSSARY_CATEGORY, "Category", 150);
  setHeader(valueSetSheet, 1, VALUESET_GLOSSARY_VALUE_LABEL, "Value Label", 150);
  setHeader(valueSetSheet, 1, VALUESET_GLOSSARY_VALUE_DESCRIPTION, "Value Description", 520);
  setHeader(valueSetSheet, 1, VALUESET_GLOSSARY_FORMAL_LABEL, "Concept Label", 150)
  setHeader(valueSetSheet, 1, VALUESET_GLOSSARY_IS_DEPRECATED, "Is Deprecated?", 150);

  const initialColumnSize = 6;
  const initialRowSize = 16;
  valueSetSheet.deleteColumns(initialColumnSize+1, 26-initialColumnSize);
  valueSetSheet.deleteRows(initialRowSize+2, 1000-(initialRowSize+1));
  valueSetSheet.setFrozenRows(1);
  valueSetSheet.setFrozenColumns(2);

  const mappedToRange = valueSetSheet.getRange(2, VALUESET_GLOSSARY_MAPPED_TO_ONTOLOGY_ACRONYM, initialRowSize, 1);
  setRangeColor(mappedToRange, "#ead1dc"); // light red

  const categoryRange = valueSetSheet.getRange(2, VALUESET_GLOSSARY_CATEGORY, initialRowSize, 1);
  setRangeColor(categoryRange, "#ead1dc"); // light red

  const firstColumnHeaderRange = valueSetSheet.getRange(1, VALUESET_GLOSSARY_MAPPED_TO_ONTOLOGY_ACRONYM, 1, 2);
  setRangeColor(firstColumnHeaderRange, "#ead1dc"); // light red

  const valueSetSheetRange = valueSetSheet.getRange(2, VALUESET_GLOSSARY_MAPPED_TO, initialRowSize, initialColumnSize);
  valueSetSheetRange.setHorizontalAlignment("left").setVerticalAlignment("top");

  const protection = valueSetSheet.protect().setDescription('Protected .VALUESET sheet');
  const me = Session.getEffectiveUser();
  protection.removeEditors(protection.getEditors());
  protection.addEditors([me.getEmail()]);

  return valueSetSheet;
}