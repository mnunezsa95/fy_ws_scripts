# Full Year Worksheets Scripts

# Google Sheets Script Functions

This repository contains custom Google Apps Script functions that help automate tasks within Google Sheets. These scripts are used for renaming sheets, updating overview information, populating instructional days, and finding frequencies in a timetable.

## Functions

### 1. `renameTabs()`
This function renames specific sheets in the active spreadsheet based on a classification found in cell `D1` of the "Overview" sheet. The classification determines the prefix (e.g., G, P, C, Standard) for the new sheet names.

#### Sheet Names Mapping:
- X1 → G1, P1, C1, or Standard1 (based on classification)
- X2 → G2, P2, C2, or Standard2
- And so on for X3, X4, X5, X6, X7, X12, X34, and X56

The function will:
- Check the value of `D1` in the "Overview" sheet.
- Assign a prefix based on the classification.
- Rename the sheets accordingly.
- Alert the user when renaming is complete.

### 2. `updateOverview()`
This function updates specific cells in the "Overview" sheet based on the program abbreviation in `D1` and grade classification in `J1`.

#### It updates:
- Cell `B5` with the subject name for numeracy.
- Cell `E5` with the subject name for literacy.
- Cell `H5` with the subject name for language.
- Cell `K5` with the subject name for science.

The program abbreviation must match one of the valid options (e.g., AP, BA, KW, etc.). Based on the abbreviation, it fills in the correct subject names for each of the above fields.

Additionally, the function updates certain ranges of the sheet (A8:A14, A17:A19) by replacing any "X" at the beginning of the values with the appropriate grade prefix.

### 3. `populateInstructionalDays()`
This function populates instructional days based on the frequency defined in a specific cell. The process involves:
- Reading the frequency value from the active column in the sheet (e.g., "Mon", "Tue").
- Prompting the user to provide the URL, tab name, and column letter of a calendar sheet.
- Extracting and populating the instructional days into the active sheet.

The results are displayed in two target columns in the active sheet.

### 4. `findFrequency()`
This function helps parse timetable frequencies and match subject days with instructional cycles. The user is prompted to:
1. Enter the grades to parse (e.g., G1, G2).
2. Provide the timetable URL.
3. Specify the subject names for courses like Numeracy, Literacy, Language, and Science.

The function maps the subjects to their respective days and cycles, allowing you to analyze the instructional days and frequencies of each subject.

## Setup and Usage
1. Copy the script files into the Google Apps Script editor associated with your Google Sheets file.
2. Run the appropriate functions as needed, ensuring you provide valid inputs when prompted.

## Prerequisites
Ensure you have Google Apps Script access enabled in your Google Sheets file to run these custom functions.

### Notes:
- Ensure the "Overview" sheet contains valid data in cells `D1` and `J1` for the `renameTabs()` and `updateOverview()` functions to work correctly.
- The URLs entered for the `populateInstructionalDays()` and `findFrequency()` functions must be valid Google Sheets URLs.
