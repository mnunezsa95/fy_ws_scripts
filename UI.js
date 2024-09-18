function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Custom Actions")
    .addItem("Update Overview", "updateOverview")
    .addItem("Rename Tabs", "renameTabs")
    .addItem("Populate Instructional Days", "populateInstructionalDays")
    .addItem("Find Frequencies & DOW", "findFrequency")
    .addToUi();
}
