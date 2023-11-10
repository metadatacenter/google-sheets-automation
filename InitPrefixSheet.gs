function initPrefixSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prefixSheet = ss.insertSheet(sheetName, ss.getActiveSheet().getIndex());
  setColumn(prefixSheet, 1, 1, 100);
  setColumn(prefixSheet, 1, 2, 350);

  const initialColumnSize = 2;
  const initialRowSize = 16;
  prefixSheet.deleteColumns(initialColumnSize+1, 26-initialColumnSize);
  prefixSheet.deleteRows(initialRowSize+1, 1000-initialRowSize);

  const PrefixSheetRange = prefixSheet.getRange(1, 1, initialRowSize, initialColumnSize);
  PrefixSheetRange.setHorizontalAlignment("left").setVerticalAlignment("top");

  const protection = prefixSheet.protect().setDescription('Protected .PREFIX sheet');
  const me = Session.getEffectiveUser();
  protection.removeEditors(protection.getEditors());
  protection.addEditors([me.getEmail()]);

  return prefixSheet;
}