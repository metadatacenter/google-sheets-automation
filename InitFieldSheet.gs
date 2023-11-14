function initFieldSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fieldSheet = ss.insertSheet(sheetName, ss.getActiveSheet().getIndex());
  setHeader(fieldSheet, 1, FIELD_GLOSSARY_MAPPED_TO, "Mapped To", 150);
  setHeader(fieldSheet, 1, FIELD_GLOSSARY_FIELD_NAME, "Field Name", 175);
  setHeader(fieldSheet, 1, FIELD_GLOSSARY_FIELD_DESCRIPTION, "Field Description", 520);
  setHeader(fieldSheet, 1, FIELD_GLOSSARY_PERMISSIBLE_VALUES, "Enumerated Values", 150);
  setHeader(fieldSheet, 1, FIELD_GLOSSARY_STRING_PATTERN, "String Pattern", 150)
  setHeader(fieldSheet, 1, FIELD_GLOSSARY_NUMBER_RANGE, "Number Range", 150);
  setHeader(fieldSheet, 1, FIELD_GLOSSARY_EXAMPLE, "Example", 150);
  setHeader(fieldSheet, 1, FIELD_GLOSSARY_IS_DEPRECATED, "Is Deprecated?", 150);
  setHeader(fieldSheet, 1, FIELD_GLOSSARY_REPLACED_BY, "Replaced By", 175);

  const initialColumnSize = 10;
  const initialRowSize = 16;
  fieldSheet.deleteColumns(initialColumnSize+1, 26-initialColumnSize);
  fieldSheet.deleteRows(initialRowSize+2, 1000-(initialRowSize+1));
  fieldSheet.setFrozenRows(1);
  fieldSheet.setFrozenColumns(2);

  const mappedToRange = fieldSheet.getRange(2, FIELD_GLOSSARY_MAPPED_TO, initialRowSize, 1);
  setRangeColor(mappedToRange, "#ead1dc"); // light red

  const fieldNameRange = fieldSheet.getRange(2, FIELD_GLOSSARY_FIELD_NAME, initialRowSize, 1);
  setRangeColor(fieldNameRange, "#ead1dc"); // light red

  const firstColumnHeaderRange = fieldSheet.getRange(1, FIELD_GLOSSARY_MAPPED_TO, 1, 2);
  setRangeColor(firstColumnHeaderRange, "#ead1dc"); // light red

  const fieldSheetRange = fieldSheet.getRange(2, FIELD_GLOSSARY_MAPPED_TO, initialRowSize, initialColumnSize);
  fieldSheetRange.setHorizontalAlignment("left").setVerticalAlignment("top");

  const protection = fieldSheet.protect().setDescription('Protected .FIELDS sheet');
  const me = Session.getEffectiveUser();
  protection.removeEditors(protection.getEditors());
  protection.addEditors([me.getEmail()]);

  return fieldSheet;
}