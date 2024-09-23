function populateInstructionalDays() {
  const frequencyMapping = [];
  const ui = SpreadsheetApp.getUi();
  const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeTargetColumn = targetSpreadsheet.getActiveCell().getColumn();
  const activeColumnLetter = getColumnLetter(activeTargetColumn);
  const frequencyValueArray = targetSpreadsheet.getRange(`${activeColumnLetter}4`).getValue().split("-");
  const frequencyPart = frequencyValueArray[1];
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

  const urlResponse = ui.prompt(
    "Input Required",
    "Enter the Academic Calendar URL:\n• Copy it directly from your browser's address bar.",
    ui.ButtonSet.OK_CANCEL
  );
  if (cancelPromptAlert(urlResponse)) return;

  const tabResponse = ui.prompt(
    "Input Required",
    "Specify the name of the tab:\n• Make sure it matches the name in the Academic Calendar.",
    ui.ButtonSet.OK_CANCEL
  );
  if (cancelPromptAlert(tabResponse)) return;

  const columnResponse = ui.prompt(
    "Input Required",
    "Specify the column letter:\n• Use a single uppercase letter that corresponds to the desired column.",
    ui.ButtonSet.OK_CANCEL
  );
  if (cancelPromptAlert(columnResponse)) return;

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

        const headerRange = targetSpreadsheet.getRange(`${activeColumnLetter}1`);
        let headerCellValue;

        if (headerRange.isPartOfMerge()) {
          const mergedRange = headerRange.getMergedRanges()[0];
          headerCellValue = mergedRange.getCell(1, 1).getValue().toUpperCase();
        } else {
          headerCellValue = headerRange.getValue().toUpperCase();
        }

        if (headerCellValue.includes("NUMERACY") || headerCellValue.includes("LITERACY")) {
          resultRangeTwo.setValues(resultValues);
        }

        resultRange.setValues(resultValues);
      }
    }
  }
}
