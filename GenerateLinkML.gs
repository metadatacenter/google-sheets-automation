function generateLinkMl() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const progressBar = startProcessing(ss, "Preparing...")

  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getSheetName();
  const prefixes = getPrefixData();
  const fields = getTemplateData();
  
  const template = {
    id: `https://w3id.org/linkml/${Utilities.getUuid()}/${toSnakeCase(sheetName)}`,
    name: toSnakeCase(sheetName),
    description: `Schema about ${sheetName}, generated from a Google Sheets.`,
    license: "https://creativecommons.org/licenses/by/4.0/",
    default_curi_maps: [ "semweb_context "],
    imports: [ "linkml:types" ],
    prefixes: Object.fromEntries(prefixes),
    default_range: "string",
    classes: generateClasses(fields),
    slots: generateSlots(fields)
  }

  const folderId = getFolderId(sheet);
  const folder = DriveApp.getFolderById(folderId);

  const dateTime = Utilities.formatDate(new Date(), getTimeZone(), "yyyy-MM-dd'T'HHmmss");
  const fileName = `${toSnakeCase(sheetName)}_${dateTime}.yaml`;
  
  const output = yaml.dump(template);
  const file = folder.createFile(fileName, output);

  finishProcessing(ss, progressBar);
  showDownloadDialog(file);
}

function generateClasses(fields) {
  return {
    "MetadataInstance": {
      "slots": fields.map((field) => toSnakeCase(field[1]))
    }
  };
}

function generateSlots(fields) {
  return fields.reduce((acc, value) => ({
    ...acc,
    [toSnakeCase(value[1])]: cleanEmpty({
      required: value[0] ? (value[0] === 'Required' ? true : false) : false,
      aliases: [ value[1] ],
      description: value[2],
      pattern: value[5] || null,
      range: checkType(value[7]),
      slot_uri: value[9]
    })
  }), {});
}

function checkType(value) {
  const typeOf = typeof(value);
  if (typeOf === "number") {
    return (Number.isInteger(value)) ? "integer" : "float";
  } else if (typeOf === "boolean") {
    return "boolean";
  } else if (typeOf === "string") {
    if (Date.parse(value)) {
      const dt = /^\d{4}-[01]\d-[0-3]\dT[0-2]\d:[0-5]\d:[0-5]\d(?:\.\d+)?Z?$/;
      return (dt.test(value)) ? "datatime" : "date";
    } else {
      return "string";
    }
  } else if (typeOf === "object") {
    const instanceOf = Object.prototype.toString.call(value);
    if (instanceOf === '[object Date]') {
      return "date";
    } else {
      return "string";
    }
  } else {
    return "string";
  }
}

function cleanEmpty(obj) {
  if (Array.isArray(obj)) { 
    return obj
        .map(v => (v && typeof v === 'object') ? cleanEmpty(v) : v)
        .filter(v => !(v == null)); 
  } else { 
    return Object.entries(obj)
        .map(([k, v]) => [k, v && typeof v === 'object' ? cleanEmpty(v) : v])
        .reduce((a, [k, v]) => (v == null ? a : (a[k]=v, a)), {});
  }
}

Map.prototype.inspect = function() {
  return `Map(${mapEntriesToString(this.entries())})`
}

function mapEntriesToString(entries) {
  return Array
    .from(entries, ([k, v]) => `\n  ${k}: ${v}`)
    .join("") + "\n";
}
