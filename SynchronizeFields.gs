function synchronizeFields() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  const templateSheet = ss.getActiveSheet();
  const fieldSheet = getFieldSheet();

  const progressBar = startProcessing(ss, "Synchronizing fields...");

  const matchingFields = [];
  const matchingValueSets = [];

  const fieldNameRange = templateSheet.getRange(2, TEMPLATE_FIELD_NAME, templateSheet.getLastRow(), 1);
  const fieldNames = fieldNameRange.getDisplayValues().flat();

  // Cache data
  const fieldData = getFieldData(minimal=true);
  const valueSetData = getValueSetData(minimal=true);

  /*
   * Iterate through each field name mentioned in the template sheet. If the field name
   * is present in the .FIELDS sheet, update the field definition. Similarly, if the
   * field name is present in the .VALUESET sheet, update the link to the value set.
   */
  for (let index = 0; index < fieldNames.length; index++) {
    const fieldName = fieldNames[index];
    if (fieldName != "") {
      const rowIndex = index + 2;
      const fieldSearchResult = searchField(fieldName, fieldData);
      if (fieldSearchResult.success) {
        const fieldSearchResultRange = fieldSearchResult.data.range;
        const effectiveFieldName = fieldSearchResultRange.getCell(1, 2).getValue();
        const reportingFieldName = `${fieldName}${fieldName != effectiveFieldName ? ` -> ${effectiveFieldName}` : ''}`;

        // Update the view link and lookup URL when a value set is present for the field name.
        const valueSetSearchResult = searchValueSet(effectiveFieldName, valueSetData);
        if (valueSetSearchResult.success) {
          const valueSetLink = valueSetSearchResult.data.link;
          const valueSetName = valueSetSearchResult.data.name;
          setUrl(fieldSheet, fieldSearchResult.data.row, FIELD_GLOSSARY_PERMISSIBLE_VALUES, valueSetLink, valueSetName);
          matchingValueSets.push(matchingValueSets.includes(reportingFieldName) ? `${reportingFieldName} (duplicate)` : reportingFieldName);
        }

        Logger.log("Fetching field definition: " + effectiveFieldName);
        const fieldDefinitionRange = fieldSearchResultRange.offset(0, 1, 1, 6);  // exclude example  
        setValuesByRow(templateSheet, rowIndex, fieldDefinitionRange.getValues(), startingColumn=TEMPLATE_FIELD_NAME);
        setRichTextValue(templateSheet, rowIndex, TEMPLATE_PERMISSIBLE_VALUES, fieldDefinitionRange.offset(0, 2, 1, 1).getRichTextValue()); // value set text
        matchingFields.push(matchingFields.includes(reportingFieldName) ? `${reportingFieldName} (duplicate)` : reportingFieldName);
      
        templateSheet.getRange(rowIndex, TEMPLATE_FIELD_NAME).clearDataValidations();
      }
    }
  }
  finishProcessing(ss, progressBar);
  showUpdateFieldsReport(matchingFields, matchingValueSets);
}

function showUpdateFieldsReport(matchingFields, matchingValueSets) {
  const text = `
  <div style="font-family:Arial;font-size:10pt">
    ${matchingFields.length > 0 ?
      `<p>Successfully updating ${matchingFields.length} field(s):</p>
      <ul>
        ${matchingFields.map(asList).join('')}
      </ul>` : `<p>No matching field found</p>`
    }
    ${matchingValueSets.length > 0 ?
      `<p>Successfully updating ${matchingValueSets.length} value set(s):</p>
      <ul>
        ${matchingValueSets.map(asList).join('')}
      </ul>` : `<p>No matching value set found</p>`
    }
  </div>`;
  showReport("Synchronization Report", text);
}