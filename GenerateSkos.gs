function generateSkos() {
  const skosSheet = createSkosSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const progressBar = startProcessing(ss, "Preparing...")

  generateSkosTable(skosSheet);
  SpreadsheetApp.flush();
  const file = generateSkosFile(skosSheet);

  finishProcessing(ss, progressBar);
  showDownloadDialog(file);
}

function createSkosSheet() {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var skosSheet = activeSpreadsheet.getSheetByName(".SKOSPLAY");
  if (skosSheet != null) {
    activeSpreadsheet.deleteSheet(skosSheet);
  }
  skosSheet = activeSpreadsheet.insertSheet(activeSpreadsheet.getNumSheets());
  skosSheet.setName(".SKOSPLAY");
  return skosSheet;
}

function generateSkosTable(sheet) {
  let endRow = 1;
  endRow = createPreamble(sheet, 1);
  endRow = createHeader(sheet, endRow + 1);
  endRow = createConceptRecords(sheet, endRow + 1);
  endRow = createFieldRecords(sheet, endRow + 1);
}

function generateSkosFile(sheet) {
  const baseUrl = "https://xls2rdf.sparna.fr/rest/convert";
  const exportSheetUrl = getExportSheetUrl(sheet);
  const encodedSheetUrl = encodeURIComponent(exportSheetUrl);
  const url = `${baseUrl}?url=${encodedSheetUrl}&lang=en`;
  const response = UrlFetchApp.fetch(url);
  const folderId = getFolderId(sheet);
  const folder = DriveApp.getFolderById(folderId);
  const dateTime = Utilities.formatDate(new Date(), getTimeZone(), "yyyy-MM-dd'T'HHmmss");
  return folder.createFile(`skos-valueset_${dateTime}.ttl`, response);
}

function showDownloadDialog(file) {
  const htmlTemplate = HtmlService.createTemplateFromFile('Download.html');
  htmlTemplate.data = { url: file.getDownloadUrl() };
  const html = htmlTemplate
    .evaluate()
    .setWidth(200)
    .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, 'Download');
}

function createPreamble(sheet, startingRow) {
  setValuesByRow(sheet, startingRow, [['ConceptScheme URI', 'https://purl.humanatlas.io/vocab/hravs']]);
  const prefixes = getPrefixSheet().getDataRange().getValues().map((row) => ["PREFIX"].concat(row));
  setValuesByRow(sheet, startingRow + 1, prefixes);
  return startingRow + prefixes.length + 1;
}
				
function createHeader(sheet, startingRow) {
  setHeader(sheet, startingRow, 1, "URI", 175);
  setHeader(sheet, startingRow, 2, "skos:prefLabel", 250);
  setHeader(sheet, startingRow, 3, "skos:definition@en", 520);
  setHeader(sheet, startingRow, 4, "rdfs:label", 250);
  setHeader(sheet, startingRow, 5, "skos:broader(separator=\",\")", 250);
  setHeader(sheet, startingRow, 6, "rdf:type", 250);
  return startingRow;
}

function createConceptRecords(sheet, startingRow) {
  let categoryId = '';
  const valueSets = getValueSetSheet().getDataRange().getValues()
      .slice(1)  // remove the table header
      .filter((row) => !row[6] || row[6] === '') // not deprecated terms
      .map((row) => {
        categoryId = (row[2] !== '') ? row[1] : categoryId;
        const uri = row[1];
        const prefLabel = (row[3] === '') ? row[2] : row[3]; // if term label is empty then use the category label
        const definition = row[4];
        const label = row[5];
        const broader = (row[2] === '') ? categoryId : '';
        const type = '';
        return [ uri, prefLabel, definition, label, broader, type ];
      })
  setValuesByRow(sheet, startingRow, valueSets);
  return startingRow + valueSets.length - 1;
}

function createFieldRecords(sheet, startingRow) {
  const fields = getFieldSheet().getDataRange().getValues()
      .slice(1)  // remove the table header
      .filter((row) => !row[8] || row[8] === '') // not deprecated terms
      .map((row) => {
        const uri = row[0];
        const prefLabel = toSnakeCase(row[1]);
        const definition = row[2];
        const label = row[1];
        const broader = '';
        const type = 'owl:AnnotationProperty';
        return [ uri, prefLabel, definition, label, broader, type ];
      })
  setValuesByRow(sheet, startingRow, fields);
  return startingRow + fields.length - 1;
}
