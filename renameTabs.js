function renameTabs() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const overviewSheet = spreadsheet.getSheetByName("Overview");
  const classification = overviewSheet.getRange("D1").getValue();
  const sheetNames = ["X1", "X2", "X3", "X4", "X5", "X6", "X7", "X12", "X34", "X56"];

  let prefix;
  switch (classification) {
    case "Grade":
      prefix = "G";
      break;
    case "Primary":
      prefix = "P";
      break;
    case "Class":
      prefix = "C";
      break;
    case "Standard":
      prefix = "Standard";
      break;
    default:
      SpreadsheetApp.getUi().alert("Invalid classification in D1");
      return;
  }

  const newSheetNames = [
    `${prefix}1`,
    `${prefix}2`,
    `${prefix}3`,
    `${prefix}4`,
    `${prefix}5`,
    `${prefix}6`,
    `${prefix}7`,
    `${prefix}12`,
    `${prefix}34`,
    `${prefix}56`,
  ];

  for (let i = 0; i < sheetNames.length; i++) {
    const sheet = spreadsheet.getSheetByName(sheetNames[i]);
    if (sheet) {
      sheet.setName(newSheetNames[i]);
    }
  }

  SpreadsheetApp.getUi().alert("Sheets renamed to: " + newSheetNames.join(", "));
}
