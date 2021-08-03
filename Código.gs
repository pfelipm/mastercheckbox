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

  // Single cell?

  if (rows == 1 && cols == 1) {
    
    // Is it a checkbox?

    const isCheckbox = range.getDataValidation().getCriteriaType() == SpreadsheetApp.DataValidationCriteria.CHECKBOX;

    if (isCheckbox) {
    
      // Is this checkbox at the top row of a 1-column checkbox interval?

      let isUpperCheckbox;
      const masterRow = range.getRow();
      const masterCol = range.getColumn();
      if (masterRow > 1) {
        const dataValidation = range.offset(-1, 0).getDataValidation();
        if (dataValidation) {
          isUpperCheckbox = !range.offset(-1, 0).getDataValidation().getCriteriaType() == SpreadsheetApp.DataValidationCriteria.CHECKBOX;
        }
        else {
          isUpperCheckbox = true;
        }
      } else { // row = 1
        isUpperCheckbox = true;
      }

      if (isUpperCheckbox) {

        // All conditions cleared, go ahead!

        const value = range.getValue();
        const expandedRange = range.getDataRegion(SpreadsheetApp.Dimension.ROWS);
        const lowerRow = masterRow + expandedRange.getNumRows() - 1;
        let actualLowerRow = masterRow + 1;

        // Grow range to include all contiguous checkboxes immediately below the one that has been edited, if any

        if (lowerRow > masterRow) {

          const dataValidations = expandedRange.getDataValidations();
          while (actualLowerRow <=  lowerRow && dataValidations[actualLowerRow - masterRow][0].getCriteriaType() == SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
            actualLowerRow++;
          };

          SpreadsheetApp.getActive().toast(`Master checkbox change detected at R${masterRow}C${masterCol} Lower row: ${actualLowerRow - 1} value ${value}`);

          // Check / uncheck accordingly

          expandedRange.offset(1, 0, actualLowerRow - masterRow - 1, 1).setValue(value);
          // SpreadsheetApp.flush();

        }
      }
    }
  }
}