(function () {
  let currentDialog = null;

  Office.onReady((info) => {
      if (info.host === Office.HostType.Excel) {
          const openRuleEditor = document.getElementById('openRuleEditor');
          openRuleEditor.addEventListener('click', editRuleConditions)
      }
  });
  
  function editRuleConditions() {
    const url = "https://bluesage-dev.bluesageusa.com/droolsrules/RuleEditor-Ex.html"
    Office.context.ui.displayDialogAsync(url, { height: 50, width: 50 }, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
          console.error('Failed to open dialog: ' + result.error.message);
      } else {
          currentDialog = result.value;
          console.log('Dialog opened successfully.');
          currentDialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessageFromDialog);
          currentDialog.addEventHandler(Office.EventType.DialogEventReceived, handleDialogEvent);
      }
    });
  }

  function getColumnDictionary(aData) {
    let columnDictionary = {}; // Use an object instead of an array for the dictionary
    let sColumnName = null;

    for (let iCol = 0; iCol < aData[0].length; iCol++) {
        sColumnName = aData[0][iCol].trim(); // Use the JavaScript trim method

        if (sColumnName !== '') {
            columnDictionary[sColumnName] = iCol;
        }
    }
    return columnDictionary;
  
  }

  async function getRule() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const headerRange = sheet.getRange("1:1");
            headerRange.load("values");
            await context.sync();

            // Use getColumnDictionary to parse headers
            const dicColumn = getColumnDictionary(headerRange.values);

            // Get the active cell row
            const activeCell = context.workbook.getActiveCell();
            activeCell.load("rowIndex");
            await context.sync();

            // Get the column index for 'When', adjusting for Excel's 0-based index
            const whenColumnIndex = dicColumn['When'];

            // Fetch the value from the 'When' column in the active row
            const ruleRange = sheet.getRangeByIndexes(activeCell.rowIndex, whenColumnIndex, 1, 1);
            ruleRange.load("values");
            await context.sync();

            // Return the value from the cell
            const sRule = ruleRange.values[0][0];
            console.log(sRule);
            return sRule;
        });
    } catch (error) {
        console.error("Error: " + error);
    }
  }

})();