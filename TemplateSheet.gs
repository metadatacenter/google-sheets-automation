function getActiveTemplateSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

function getTemplateData() {
  const templateSheet = getActiveTemplateSheet();
  const templateData = templateSheet.getRange(2, TEMPLATE_FIELD_REQUIREMENT, templateSheet.getLastRow()-1, 9).getValues();
  const fieldData = getFieldData(minimal=true);
  return templateData.map((row) => [...row, findFieldId(row[1], fieldData)]);
}
