function importFields() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  const templateSheet = ss.getActiveSheet();
  const row = templateSheet.getActiveRange().getRow();
  const requirementList = ["", "Required", "Optional", "Hidden"];
  setDataValidation(templateSheet.getRange(row, TEMPLATE_FIELD_REQUIREMENT), requirementList);
  const fieldList = getFieldList();
  setDataValidation(templateSheet.getRange(row, TEMPLATE_FIELD_NAME), fieldList, allowInvalid=true);
}