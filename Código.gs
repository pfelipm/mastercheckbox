/**
 * Uses an onEdit() simple trigger to monitor the status of checkboxes in the spreadsheet and
 * toggles all contiguous checkboxes immediately below (same column) accordingly.
 * 
 * Copyright (C) Pablo Felip (@pfelipm) · Se distribuye bajo licencia MIT.
 * 
 * @OnlyCurrentDoc
 */

function onEdit(e) {

  // Get edited range

  const range = e.range;
  const rows = range.getNumRows();
  const cols = range.getNumColumns();

  // Single cell edit?

  if (rows == 1 && cols == 1) {
    
    // Is it a checkbox?

    let dataValidation = range.getDataValidation(); // getDataValidation() may return null!
    const isCheckbox = dataValidation ? dataValidation.getCriteriaType() == SpreadsheetApp.DataValidationCriteria.CHECKBOX : false;

    if (isCheckbox) {
    
      // Has this checkbox another one immediately above?

      let isUpperCheckbox;
      const masterRow = range.getRow();
      if (masterRow > 1) {
        dataValidation = range.offset(-1, 0).getDataValidation();
        isUpperCheckbox = dataValidation ?
                          !range.offset(-1, 0).getDataValidation().getCriteriaType() == SpreadsheetApp.DataValidationCriteria.CHECKBOX :
                          true;

      } else { // row = 1
        isUpperCheckbox = true;
      }

      if (isUpperCheckbox) {

        // All conditions cleared, let's find how many checkboxes immediately below.

        const value = range.getValue();
        const expandedRange = range.getDataRegion(SpreadsheetApp.Dimension.ROWS);
        const lowerRow = expandedRange.getNumRows();

        // Grow range to include all contiguous checkboxes immediately below the one that has been edited, if any

        if (lowerRow > masterRow) {

          const dataValidations = expandedRange.getDataValidations();          
          let lastCheckboxFound = false;
          let actualLowerRow;
          for (actualLowerRow = masterRow + 1; actualLowerRow <= lowerRow; actualLowerRow++) {
            dataValidation = dataValidations[actualLowerRow - 1][0];
            const lastCheckboxFound = dataValidation ?
                                      dataValidation.getCriteriaType() != SpreadsheetApp.DataValidationCriteria.CHECKBOX :
                                      true;
            if (lastCheckboxFound) break;
          }

          // Check or uncheck accordingly

          const numCheckboxes = actualLowerRow - 1 - masterRow;
          if (numCheckboxes > 0) {
            
            console.info(masterRow + 1, masterRow + numCheckboxes);
            SpreadsheetApp.getActive().toast('Conmutando casillas de verificación...','',2);
            range.offset(1, 0, actualLowerRow - 1 - masterRow, 1).setValue(value);
            // SpreadsheetApp.flush();
          
          }
        }
      }
    }
  }
}