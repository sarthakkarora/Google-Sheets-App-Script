function isEligibleForCompany(criteria, studentData) {
  const { year, dropYear, backlogStatus, cgpa, department } = studentData;

  if (!criteria) return [false, "Invalid Company Criteria"];

  const { backlogLimit, minCGPA, validYears, allowedDepartments, dropYearAllowed } = criteria;

  if (
    (backlogLimit === "No" ? backlogStatus === "No" : backlogStatus <= backlogLimit) &&
    cgpa >= minCGPA &&
    (!validYears || validYears.includes(year)) &&
    (!allowedDepartments || allowedDepartments.includes(department)) &&
    (dropYearAllowed || dropYear === "No")
  ) {
    return [true, `Eligible for ${criteria.companyName}`];
  }

  return [false, `Not eligible for ${criteria.companyName} (Check criteria)`];
}

function fetchCompanyCriteria(companyName) {
  const criteriaSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CompanyCriteria");
  if (!criteriaSheet) return null;

  const criteriaRows = criteriaSheet.getDataRange().getValues();
  for (let i = 1; i < criteriaRows.length; i++) {
    if (criteriaRows[i][0] === companyName) {
      return {
        companyName: criteriaRows[i][0],
        backlogLimit: criteriaRows[i][1],
        minCGPA: criteriaRows[i][2],
        validYears: criteriaRows[i][3] ? criteriaRows[i][3].split(",").map(Number) : null,
        allowedDepartments: criteriaRows[i][4] ? criteriaRows[i][4].split(",") : null,
        dropYearAllowed: criteriaRows[i][5] === "Yes",
      };
    }
  }

  return null;
}

function validateSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rows = sheet.getDataRange().getValues();
  const lastCol = rows[0].length;

  for (let i = 1; i < rows.length; i++) {
    const studentData = {
      year: rows[i][1],
      dropYear: rows[i][2],
      backlogStatus: rows[i][3],
      cgpa: rows[i][4],
      department: rows[i][5],
    };

    const companyCriteria = fetchCompanyCriteria(rows[i][6]);
    const eligibility = isEligibleForCompany(companyCriteria, studentData);

    rows[i][lastCol - 2] = eligibility[0] ? "Eligible" : "Not Eligible"; 
    rows[i][lastCol - 1] = eligibility[1]; 
  }

  sheet.getDataRange().setValues(rows);

  SpreadsheetApp.getActiveSpreadsheet().toast("Eligibility updated successfully", "Success", 3);
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const editedRow = e.range.getRow();
  const editedColumn = e.range.getColumn();

  if (sheet.getName() === "Main" && editedColumn >= 1 && editedColumn <= 6) {
    const row = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    const studentData = {
      year: row[1],
      dropYear: row[2],
      backlogStatus: row[3],
      cgpa: row[4],
      department: row[5],
    };

    const companyCriteria = fetchCompanyCriteria(row[6]);
    const eligibility = isEligibleForCompany(companyCriteria, studentData);

    sheet.getRange(editedRow, sheet.getLastColumn() - 1).setValue(eligibility[0] ? "Eligible" : "Not Eligible");
    sheet.getRange(editedRow, sheet.getLastColumn()).setValue(eligibility[1]);

    logChange(e, eligibility[1]);
  }
}

function logChange(e, message) {
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AuditLog") || 
                   SpreadsheetApp.getActiveSpreadsheet().insertSheet("AuditLog");

  const timestamp = new Date();
  const user = Session.getActiveUser().getEmail();
  const editedRange = e.range.getA1Notation();

  logSheet.appendRow([timestamp, user, editedRange, message]);
}

function setup() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  if (!headers.includes("Eligibility Status")) {
    sheet.getRange(1, headers.length + 1).setValue("Eligibility Status");
    sheet.getRange(1, headers.length + 2).setValue("Eligibility Message");
  }


  const criteriaSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CompanyCriteria") || 
                        SpreadsheetApp.getActiveSpreadsheet().insertSheet("CompanyCriteria");

  if (criteriaSheet.getLastRow() === 0) {
    criteriaSheet.appendRow([
      "Company Name",
      "Backlog Limit",
      "Minimum CGPA",
      "Valid Years (comma-separated)",
      "Allowed Departments (comma-separated)",
      "Drop Year Allowed (Yes/No)",
    ]);
  }
}

function applyConditionalFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange(2, sheet.getLastColumn() - 1, sheet.getLastRow() - 1);

  const rules = sheet.getConditionalFormatRules();
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Eligible")
      .setBackground("#C8E6C9") 
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Not Eligible")
      .setBackground("#FFCDD2") 
      .setRanges([range])
      .build()
  );


  sheet.setConditionalFormatRules(rules);
}
