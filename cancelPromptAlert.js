function cancelPromptAlert(responseVar) {
  const ui = SpreadsheetApp.getUi();

  if (responseVar.getSelectedButton() == ui.Button.CANCEL) {
    ui.alert("Action canceled.");
    return true;
  }

  return false; // Return false if not canceled
}
