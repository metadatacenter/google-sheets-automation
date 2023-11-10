function exportFields() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  const templateSheet = ss.getActiveSheet();

  let selectedRange = templateSheet.getActiveRange();

  // When the selected range includes the header, it automatically refers to the row immediately after the header.
  const startingRow = (selectedRange.getRow() == 1) ? 2 : selectedRange.getRow();
  const numberOfRows = selectedRange.getLastRow() - startingRow + 1;

  const progressBar = startProcessing(ss, "Exporting new fields...");

  const newFieldNamesRange = templateSheet.getRange(startingRow, TEMPLATE_FIELD_NAME, numberOfRows, 1);
  const newFieldNames = newFieldNamesRange.getDisplayValues().flat();

  const storedFields = [];
  const storedValueSets = [];

  // Cache data
  const fieldSheet = getFieldSheet();
  const fieldData = getFieldData(minimal=true);

  /*
   * Iterate through each new field name discovered in the template sheet. If the field name
   * is not present in the .FIELDS sheet, include it. Similarly, if the field name is absent
   * in the .VALUESET sheet, add the user-provided permissible values there as well.
   */
  for (let index = 0; index < newFieldNames.length; index++) {
    const fieldName = newFieldNames[index];
    if (fieldName != "") {
      const rowIndex = startingRow + index;
      const fieldSearchResult = searchField(fieldName, fieldData);
      if (!fieldSearchResult.success) {
        if (!storedFields.includes(fieldName)) {
          /*
           * Store the new field from the template sheet into the .FIELDS sheet
           */
          Logger.log("Storing new field: " + fieldName);
          const fieldDefinitionRange = templateSheet.getRange(rowIndex, TEMPLATE_FIELD_NAME, 1, 7);
          const fieldDefinition = fieldDefinitionRange.getDisplayValues();
          const insertionIndex = fieldSheet.getLastRow() + 1;
          setValuesByRow(fieldSheet, insertionIndex, fieldDefinition, startingColumn=FIELD_GLOSSARY_FIELD_NAME);
          storedFields.push(fieldName);

          /*
           * Store the new value set from the template sheet into the .VALUESET sheet, if any
           */
          const permissibleValues = fieldDefinition[0][2].trim();
          if (permissibleValues != "") {
            Logger.log("Storing new value set: " + fieldName);
            const valueSetLink = storeValueSet(fieldName, permissibleValues);
            const templateSheetPermissibleValuesRange = templateSheet.getRange(rowIndex, TEMPLATE_PERMISSIBLE_VALUES);
            setUrlByRange(templateSheetPermissibleValuesRange, valueSetLink, "View");
            const fieldSheetPermissibleValueRange = fieldSheet.getRange(insertionIndex, FIELD_GLOSSARY_PERMISSIBLE_VALUES);
            setUrlByRange(fieldSheetPermissibleValueRange, valueSetLink, "View");
            storedValueSets.push(fieldName);
          }

          /*
           * Remove data dropdown for the exported field.
           */
          const exportedFieldRange = templateSheet.getRange(rowIndex, TEMPLATE_FIELD_NAME);
          exportedFieldRange.clearDataValidations();
          setStrictProtection(exportedFieldRange, `Protected`);
        } else {
          Logger.log("Duplicate new field detected: " + fieldName);
          // TODO: Flag the cell
        }
      }
    }
  }
  finishProcessing(ss, progressBar)
  showExportFieldsReport(storedFields, storedValueSets);
}

function showExportFieldsReport(storedFields, storedValueSets) {
  const text = `
  <div style="font-family:Arial;font-size:10pt">
    ${storedFields.length > 0 ? 
      `<p>Successfully adding new ${storedFields.length} field(s):</p>
      <ul>
        ${storedFields.map(asList).join('')}
      </ul>` : `<p>No new field found.</p>`
    }
    ${storedValueSets.length > 0 ?
      `<p>Successfully adding new ${storedValueSets.length} value set(s):</p>
      <ul>
        ${storedValueSets.map(asList).join('')}
      </ul>` : `<p>No new value set found.</p>`
    }
  </div>`;
  showReport("Export Fields Report", text);
}