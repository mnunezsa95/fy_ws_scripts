function populateInstructionalDays() {
  const frequencyMapping = [];
  const ui = SpreadsheetApp.getUi();
  const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeTargetColumn = targetSpreadsheet.getActiveCell().getColumn();
  const activeColumnLetter = String.fromCharCode(64 + activeTargetColumn);
  const frquencyValueArray = targetSpreadsheet.getRange(`${activeColumnLetter}4`).getValue().split("-");
  const frequencyPart = frquencyValueArray[1];
  const frequencyValues = frequencyPart.split(",").map((value) => value.trim());

  frequencyValues.forEach((freqValue) => {
    switch (freqValue) {
      case "M":
        frequencyMapping.push("Mon");
        break;
      case "T":
        frequencyMapping.push("Tue");
        break;
      case "W":
        frequencyMapping.push("Wed");
        break;
      case "Th":
        frequencyMapping.push("Thu");
        break;
      case "F":
        frequencyMapping.push("Fri");
        break;
      default:
        frequencyMapping.push("Unknown");
    }
  });

  const urlResponse = ui.prompt("Enter", "Enter the spreadsheet URL:", ui.ButtonSet.OK_CANCEL);
  const tabResponse = ui.prompt("Enter", "Enter the tab name:", ui.ButtonSet.OK_CANCEL);
  const columnResponse = ui.prompt("Enter", "Enter the column letter:", ui.ButtonSet.OK_CANCEL);

  if (urlResponse.getSelectedButton() == ui.Button.OK) {
    const url = urlResponse.getResponseText();
    const spreadsheetId = url.match(/[-\w]{25,}/);

    if (!spreadsheetId) {
      ui.alert("Invalid URL. Please provide a valid spreadsheet URL.");
      return;
    }

    const calendarSpreadsheet = SpreadsheetApp.openById(spreadsheetId[0]);

    if (tabResponse.getSelectedButton() == ui.Button.OK) {
      const tabName = tabResponse.getResponseText();
      const sheet = calendarSpreadsheet.getSheetByName(tabName);

      if (!sheet) {
        ui.alert("Tab not found. Please provide a valid tab name.");
        return;
      }

      if (columnResponse.getSelectedButton() == ui.Button.OK) {
        const columnLetter = columnResponse.getResponseText().toUpperCase();
        const lastRow = sheet.getLastRow();
        const range = `${columnLetter}2:${columnLetter}${lastRow}`;
        const targetRangeValues = sheet.getRange(range).getValues();
        const columnBValues = sheet
          .getRange(`B2:B${lastRow}`)
          .getValues()
          .map((date) => String(date).split(" ")[0]);

        let results = [];
        for (let i = 0; i < targetRangeValues.length; i++) {
          let targetValue = targetRangeValues[i][0];
          let columnBValue = columnBValues[i];

          if (frequencyMapping.includes(columnBValue)) {
            if (/(?:\w+\s*\/\s*\d+|\b\d+\b)/.test(targetValue)) {
              results.push(targetValue);
            }
          }
        }

        const processedResults = results.map((val) => {
          if (String(val).includes("/")) {
            return val.split("/")[1].trim();
          } else {
            return String(val);
          }
        });

        const resultRange = targetSpreadsheet.getRange(6, activeTargetColumn, processedResults.length, 1);
        const resultRangeTwo = targetSpreadsheet.getRange(6, activeTargetColumn + 2, processedResults.length, 1);
        const resultValues = processedResults.map((val) => [val]);
        resultRange.setValues(resultValues);
        resultRangeTwo.setValues(resultValues);
      }
    }
  }
}
