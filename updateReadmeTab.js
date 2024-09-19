function updateReadmeTab() {
  const activeCell = SpreadsheetApp.getActiveSpreadsheet().getRange("L4");
  const date = new Date();

  const options = {
    weekday: "short",
    year: "numeric",
    month: "short",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    timeZoneName: "short",
  };
  const dateString = date.toLocaleDateString("en-US", options);
  console.log(dateString);
  activeCell.setValue(dateString);
}
