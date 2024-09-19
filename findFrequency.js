function findFrequency() {
  const ui = SpreadsheetApp.getUi();
  const programName = SpreadsheetApp.getActiveSpreadsheet()?.getSheetByName("Overview")?.getRange("D1").getValue();
  const userEmail = Session.getActiveUser().getEmail();
  const [usernamePart, remainder] = userEmail.split("@");
  const [firstName, lastName] = usernamePart.split(".");
  const username =
    firstName.charAt(0).toUpperCase() + firstName.slice(1) + " " + lastName.charAt(0).toUpperCase() + lastName.slice(1);

  const courses = ["Numeracy", "Literacy", "Language", "Science"];

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

  let courseMap = {};
  let data = {};
  const dayAbbreviations = {
    Monday: "M",
    Tuesday: "T",
    Wednesday: "W",
    Thursday: "Th",
    Friday: "F",
  };

  const numberOfCyclesResponse = ui.prompt(
    "Input Required",
    "Enter the grades to parse: \n • Separate each grade using a comma & space (i.e: G1, G2, G12):",
    ui.ButtonSet.OK_CANCEL
  );

  if (cancelPromptAlert(numberOfCyclesResponse)) return;

  const numberOfCycles = numberOfCyclesResponse
    .getResponseText()
    .split(", ")
    .map((cycle) => cycle.trim());

  const timetableURLResponse = ui.prompt("Input Required", "Enter the timetable URL:", ui.ButtonSet.OK_CANCEL);
  if (cancelPromptAlert(timetableURLResponse)) return;

  const url = timetableURLResponse.getResponseText();
  const spreadsheetId = url.match(/[-\w]{25,}/);

  if (!spreadsheetId) {
    ui.alert("Invalid URL. Please provide a valid spreadsheet URL.");
    return;
  }

  const timetableSpreadsheet = SpreadsheetApp.openById(spreadsheetId[0]);

  courses.forEach((course) => {
    const subjectResponse = ui.prompt(
      `Input Required`,
      `What is/are the subject name(s) for ${course}?:\n\nSeparate each subject name using a comma & space, i.e: \n• Mathematics 1, Mathematics 2\n• English Studies - Reading 1, English Studies - Reading 2\n• English Studies - Language\n• Science`,
      ui.ButtonSet.OK_CANCEL
    );

    if (cancelPromptAlert(subjectResponse)) return;

    const subjectNames = subjectResponse
      .getResponseText()
      .split(",")
      .map((name) => name.trim());
    courseMap[course] = subjectNames;
  });

  numberOfCycles.forEach((cycle) => {
    const currentSheet = timetableSpreadsheet.getSheetByName(cycle);

    if (!currentSheet) {
      ui.alert(`Sheet ${cycle} not found.`);
      return;
    }

    const lastColumnResponse = ui.prompt(
      "Enter",
      `${cycle} - Enter the last COLUMN of the timetable:`,
      ui.ButtonSet.OK_CANCEL
    );
    const lastRowResponse = ui.prompt(
      "Enter",
      `${cycle} - Enter the last ROW of the timetable:`,
      ui.ButtonSet.OK_CANCEL
    );

    const lastColumn = lastColumnResponse.getResponseText();
    const lastRow = lastRowResponse.getResponseText();

    if (!lastColumn || !lastRow) {
      ui.alert(`Invalid column or row input for sheet ${cycle}.`);
      return;
    }

    const currentRange = currentSheet.getRange(`A1:${lastColumn}${lastRow}`);
    const values = currentRange.getValues();
    const days = values[0];

    data[cycle] = {};

    courses.forEach((course) => {
      data[cycle][course] = { frequency: 0, days: [] };
    });

    for (let row = 1; row < values.length; row++) {
      for (let col = 0; col < values[row].length; col++) {
        const cellValue = values[row][col].toString().trim();

        for (const course in courseMap) {
          if (courseMap[course].includes(cellValue)) {
            data[cycle][course].frequency += 1;
            const day = days[col];
            const abbreviatedDay = dayAbbreviations[day];
            if (abbreviatedDay && !data[cycle][course].days.includes(abbreviatedDay)) {
              data[cycle][course].days.push(abbreviatedDay);
            }
          }
        }
      }
    }
  });

  let resultString = "";
  for (const cycle in data) {
    resultString += `\n${cycle}:\n`;
    for (const course in data[cycle]) {
      resultString += ` • ${course}: [ Freq: ${data[cycle][course].frequency} ], [ Days: ${data[cycle][
        course
      ].days.join(", ")} ]\n`;
    }
  }

  ui.alert("Frequency Data", resultString, ui.ButtonSet.OK);

  MailApp.sendEmail({
    to: userEmail,
    subject: `Frequency Data Results for ${programName}`,
    body: `Script Executed by: ${username}\nScript Executed on: ${dateString}\n\nHere are the frequency data results for ${programName}: \n${resultString}`,
  });

  ui.alert("Script Finished", "The script has finished running. The results will be emailed to you.", ui.ButtonSet.OK);

  return data;
}
