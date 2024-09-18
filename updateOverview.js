function updateOverview() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const overviewSheet = spreadsheet.getSheetByName("Overview");
  const gradeClassification = overviewSheet.getRange("J1").getValue();
  const programAbbreviation = overviewSheet.getRange("D1").getValue();

  if (
    !programAbbreviation ||
    !["AP", "BA", "KW", "BEN", "Edo", "EKO", "GAM", "KE", "LR", "MEG", "MN", "NG", "Oyo", "RCA", "RW", "UG"].includes(
      programAbbreviation
    )
  ) {
    SpreadsheetApp.getUi().alert("Invalid or missing program abbreviation in D1");
    return;
  }

  let numSubjectName, litSubjectName, lanSubjectName, sciSubjectName;
  if (programAbbreviation == "AP") {
    numSubjectName = "Supplementary Maths";
    litSubjectName = "Supplementary English 1 and 2";
    lanSubjectName = "Fluency";
    sciSubjectName = "EVS";
  } else if (programAbbreviation == "BA" || programAbbreviation == "KW") {
    numSubjectName = "Mathematics 1 and 2";
    litSubjectName = "English Studies - Reading 1 and 2";
    lanSubjectName = "English Studies - Language";
    sciSubjectName = "Basic Science and Technology";
  } else if (programAbbreviation == "Edo") {
    numSubjectName = "Mathematics 1 and 2 (P1-P2)\nPreparatory Maths (P3-P6, JSS)";
    litSubjectName = "English Studies - Reading 1 and 2 (P1-P6)\nPreparatory English 1 and 2 (JSS)";
    lanSubjectName = "English Studies - Language";
    sciSubjectName = "Basic Science and Technology";
  } else if (programAbbreviation == "EKO") {
    numSubjectName = "Mathematics 1 and 2";
    litSubjectName = "English Studies - Reading 1 and 2";
    lanSubjectName = "English Studies - Language";
    sciSubjectName = "Basic Science";
  } else if (programAbbreviation == "KE") {
    numSubjectName = "Numeracy";
    litSubjectName = "Literacy Revision 1 and 2";
    lanSubjectName = "Supplemental Language";
    sciSubjectName = "N/A - Not Running";
  } else if (programAbbreviation == "LR") {
    numSubjectName = "Mathematics 1 and 2";
    litSubjectName = "English Studies - Reading 1 and 2";
    lanSubjectName = "English Studies - Language";
    sciSubjectName = "Science";
  } else if (programAbbreviation == "MN") {
    numSubjectName = "Supplemental Mathematics";
    litSubjectName = "Supplemental English";
    lanSubjectName = "N/A - Not Running";
    sciSubjectName = "N/A - Not Running";
  } else if (programAbbreviation == "NG") {
    numSubjectName = "Mathematics 1 and 2 (P1-P3)\nNumeracy (P4-P6)";
    litSubjectName = "English Studies - Reading 1 and 2";
    lanSubjectName = "English Studies - Language";
    sciSubjectName = "Basic Science";
  } else if (programAbbreviation == "RW") {
    numSubjectName = "Mathematics 1 and 2";
    litSubjectName = "English Studies - Reading 1 and 2";
    lanSubjectName = "English Studies - Language";
    sciSubjectName = "Science and Elementary Technology";
  } else if (programAbbreviation == "UG") {
    numSubjectName = "Supplemental Numeracy";
    litSubjectName = "English Literacy Revision 1 and 2";
    lanSubjectName = "N/A - Not Running";
    sciSubjectName = "N/A - Not Running";
  }

  overviewSheet.getRange("B5").setValue(numSubjectName);
  overviewSheet.getRange("E5").setValue(litSubjectName);
  overviewSheet.getRange("H5").setValue(lanSubjectName);
  overviewSheet.getRange("K5").setValue(sciSubjectName);

  let gradePrefix;
  switch (gradeClassification) {
    case "Grade":
      gradePrefix = "G";
      break;
    case "Primary":
      gradePrefix = "P";
      break;
    case "Class":
      gradePrefix = "C";
      break;
    case "Standard":
      gradePrefix = "Standard";
      break;
    default:
      SpreadsheetApp.getUi().alert("Invalid classification in J1");
      return;
  }

  const rangesToUpdate = ["A8:A14", "A17:A19"];

  rangesToUpdate.forEach((range) => {
    const rangeObj = overviewSheet.getRange(range);
    const cells = rangeObj.getValues();

    for (let i = 0; i < cells.length; i++) {
      if (cells[i][0] && cells[i][0].startsWith("X")) {
        cells[i][0] = `${gradePrefix}${cells[i][0].slice(1)}`;
      }
    }
    rangeObj.setValues(cells);
    rangeObj.setFontColor("black");
  });

  SpreadsheetApp.getUi().alert("Overview updated successfully!");
}
