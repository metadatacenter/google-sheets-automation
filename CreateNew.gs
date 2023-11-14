function createNew() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetIndex =  ss.getActiveSheet().getIndex();

  /*
   * Check mandatory data storage sheets.
   */
  if (!ss.getSheetByName(".FIELDS")) {
    const progressBar = startProcessing(ss, "Initializing .FIELDS sheet (only once)...")
    initFieldSheet(".FIELDS");
    finishProcessing(ss, progressBar);
  }
  if (!ss.getSheetByName(".VALUESETS")) {
    const progressBar = startProcessing(ss, "Initializing .VALUESETS sheet (only once)...")
    initvalueSetSheet(".VALUESETS");
    finishProcessing(ss, progressBar);
  }
  if (!ss.getSheetByName(".PREFIXES")) {
    const progressBar = startProcessing(ss, "Initializing .PREFIXES sheet (only once)...")
    initPrefixSheet(".PREFIXES");
    finishProcessing(ss, progressBar);
  }

  const specificationName = prompt("Please enter the specification name:");

  const progressBar = startProcessing(ss, "Preparing...")
  const templateSheet = ss.insertSheet(specificationName, sheetIndex);
  setHeader(templateSheet, 1, TEMPLATE_FIELD_REQUIREMENT, "", 80);
  setHeader(templateSheet, 1, TEMPLATE_FIELD_NAME, "Field Name", 175);
  templateSheet.getRange(1, 1, 1, 2).mergeAcross();
  setHeader(templateSheet, 1, TEMPLATE_FIELD_DESCRIPTION, "Field Description", 520);
  setHeader(templateSheet, 1, TEMPLATE_PERMISSIBLE_VALUES, "Enumerated Values", 150);
  setHeader(templateSheet, 1, TEMPLATE_STRING_PATTERN, "String Pattern", 150)
  setHeader(templateSheet, 1, TEMPLATE_NUMBER_RANGE, "Number Range", 150);
  setHeader(templateSheet, 1, TEMPLATE_EXAMPLE, "Example", 150);
  setHeader(templateSheet, 1, TEMPLATE_DEFAULT_VALUE, "Default Value", 150);

  const initialColumnSize = 9;
  const initialRowSize = 16;
  templateSheet.deleteColumns(initialColumnSize+1, 26-initialColumnSize);
  templateSheet.deleteRows(initialRowSize+2, 1000-(initialRowSize+1));
  templateSheet.setFrozenRows(1);
  templateSheet.setFrozenColumns(2);

  const requirementList = ["", "Required", "Optional", "Recommended"];
  const requirementRange = templateSheet.getRange(2, TEMPLATE_FIELD_REQUIREMENT, initialRowSize, 1);
  setDataValidation(requirementRange, requirementList);
  setRangeColor(requirementRange, "#d9d1e9"); // light purple

  const fieldList = getFieldList();
  const fieldNameRange = templateSheet.getRange(2, TEMPLATE_FIELD_NAME, initialRowSize, 1);
  setDataValidation(fieldNameRange, fieldList, allowInvalid=true);
  setRangeColor(fieldNameRange, "#d9d1e9"); // light purple

  const firstColumnHeaderRange = templateSheet.getRange(1, TEMPLATE_FIELD_REQUIREMENT);
  setRangeColor(firstColumnHeaderRange, "#d9d1e9"); // light purple

  const templateRange = templateSheet.getRange(2, TEMPLATE_FIELD_REQUIREMENT, initialRowSize, initialColumnSize);
  templateRange.setHorizontalAlignment("left").setVerticalAlignment("top");

  finishProcessing(ss, progressBar);
}
