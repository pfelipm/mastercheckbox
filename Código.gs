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

  // Single cell edit?
  if (range.getNumRows() == 1 && range.getNumColumns() == 1) {
    
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

      // If all conditions met, let's find how many checkboxes immediately below.
      if (isUpperCheckbox) {

        // Grow range vertically to contain all contiguous non-blank cells  
        const expandedRange = range.getDataRegion(SpreadsheetApp.Dimension.ROWS);

        // Calculate position of masterRow and last row inside expanded range to check downwards
        const expandedRangeRow = expandedRange.getRow();
        const startRow = masterRow - expandedRangeRow + 2;
        const lowerRow = expandedRange.getNumRows();
        
        if (lowerRow >= startRow) {

          const dataValidations = expandedRange.getDataValidations();        
          let actualLowerRow;
          for (actualLowerRow = startRow; actualLowerRow <= lowerRow; actualLowerRow++) {

            dataValidation = dataValidations[actualLowerRow - 1][0];
            const lastCheckboxFound = dataValidation ?
                                      dataValidation.getCriteriaType() != SpreadsheetApp.DataValidationCriteria.CHECKBOX :
                                      true;
            if (lastCheckboxFound) break;

          }

          // Check or uncheck accordingly
          const numCheckboxes = actualLowerRow - startRow;
          if (numCheckboxes > 0) {
            
            SpreadsheetApp.getActive().toast(`Conmutando casillas de verificación (${numCheckboxes}).`,'',2);
            range.offset(1, 0, actualLowerRow - startRow, 1).setValue(range.getValue());
            // SpreadsheetApp.flush();
          
          }
        }
      }
    }
  }
}