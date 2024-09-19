function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Custom Actions")
    .addItem("Update Overview", "updateOverview")
    .addItem("Rename Tabs", "renameTabs")
    .addItem("Find Frequencies & DsOW", "findFrequency")
    .addItem("Populate Instructional Days", "populateInstructionalDays")
    .addItem("Update README Tab", "updateReadmeTab")
    .addToUi();
}
