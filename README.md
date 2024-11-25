
# Student Eligibility Validator

This Google Apps Script project validates student eligibility for various companies based on predefined criteria. It automates the eligibility checking process in a Google Sheet, providing real-time updates, conditional formatting, and logging changes.

---

## Features

- **Dynamic Eligibility Check**: Validates student data against company-specific criteria.
- **Real-Time Updates**: Automatically updates eligibility status and message when student data is edited.
- **Company Criteria Management**: Easily manage company eligibility rules via a separate sheet.
- **Conditional Formatting**: Highlights eligibility status with color coding (e.g., green for eligible, red for not eligible).
- **Audit Logging**: Logs all changes in an `AuditLog` sheet with timestamps, user information, and messages.

---

## Setup Instructions

1. **Clone or Copy the Script**  
   Open the Google Apps Script editor for your spreadsheet:
   - In Google Sheets, click on **Extensions > Apps Script**.
   - Replace the default code with the provided script.

2. **Run the `setup()` Function**  
   - Creates required headers (`Eligibility Status` and `Eligibility Message`) in the main sheet.
   - Sets up the `CompanyCriteria` sheet for managing company eligibility rules.

3. **Define Company Criteria**  
   Populate the `CompanyCriteria` sheet with the following columns:  
   | Company Name | Backlog Limit | Minimum CGPA | Valid Years | Allowed Departments | Drop Year Allowed |
   - Example:  
     ```
     Company 1 | No | 7.5 | 3,4 | CSE,CSBS| No
     Company 2 | 1 | 8.0 | 4 | All | Yes
     ```

4. **Populate Main Sheet**  
   Ensure your main sheet has the following columns:
   | Student Name | Academic Year | Drop Year | Backlog Status | CGPA | Department | Company Criteria | Eligibility Status | Eligibility Message |

5. **Validate Data**  
   - Run the `validateSheet()` function to check eligibility for all students in the sheet.

6. **Real-Time Updates**  
   - Add an **onEdit trigger**:
     - Go to the Apps Script editor.
     - Click on **Triggers > Add Trigger**.
     - Select the `onEdit` function and set it to trigger on "Spreadsheet edit."

7. **Apply Conditional Formatting**  
   - Run the `applyConditionalFormatting()` function to highlight eligibility status in the sheet.

---

## How It Works

1. **Eligibility Logic**:  
   - Validates student data against rules defined in the `CompanyCriteria` sheet.
   - Checks criteria like backlog limit, CGPA, academic year, department, and drop year.

2. **Real-Time Updates**:  
   - Listens for changes in the first six columns of the main sheet.
   - Revalidates and updates eligibility status and message.

3. **Audit Logging**:  
   - Logs every change in the `AuditLog` sheet with details.

---

## Functions Overview

| Function                  | Description                                                                                          |
|---------------------------|------------------------------------------------------------------------------------------------------|
| `setup()`                 | Sets up required sheets and headers.                                                                |
| `validateSheet()`         | Validates all rows in the main sheet and updates eligibility status and messages.                    |
| `onEdit(e)`               | Automatically checks eligibility when a relevant column is edited.                                  |
| `fetchCompanyCriteria()`  | Retrieves eligibility rules for a specific company from the `CompanyCriteria` sheet.                |
| `isEligibleForCompany()`  | Checks student data against eligibility rules.                                                      |
| `applyConditionalFormatting()` | Applies color coding to eligibility status in the sheet.                                         |
| `logChange()`             | Logs edits to the `AuditLog` sheet.                                                                 |

---

## Example Workflow

1. **Add a New Company**  
   In the `CompanyCriteria` sheet:
   ```
   Company 3 | 0 | 6.5 | 4 | CSE,IT | No
   ```

2. **Add Student Data**  
   In the main sheet:
   ```
   Ravi Kiran | 4 | No | 0 | 7.0 | CSE | Company 3
   ```

3. **Run `validateSheet()`**  
   Output:
   ```
   Eligible | Eligible for Company 3
   ```

4. **Edit Data**  
   Change `CGPA` to 6.0.  
   Output:
   ```
   Not Eligible | Not eligible for Company 3 (Check criteria)
   ```

---

## Sheet Structure

### Main Sheet
| Column               | Description                                                     |
|-----------------------|-----------------------------------------------------------------|
| `Student Name`       | Name of the student.                                           |
| `Academic Year`      | Student's academic year (e.g., 1, 2, 3, 4).                    |
| `Drop Year`          | Whether the student has a drop year (`Yes`/`No`).              |
| `Backlog Status`     | Backlogs (`Yes`, `No`, or count of backlogs).                  |
| `CGPA`               | Student's CGPA.                                                |
| `Department`         | Department (e.g., `CSE`, `CSBS`).                              |
| `Company Criteria`   | Name of the company the student is applying for.              |
| `Eligibility Status` | Automatically updated status (`Eligible`/`Not Eligible`).     |
| `Eligibility Message`| Automatically updated message explaining the decision.        |

### CompanyCriteria Sheet
| Column               | Description                                                     |
|-----------------------|-----------------------------------------------------------------|
| `Company Name`       | Name of the company.                                           |
| `Backlog Limit`      | Allowed backlogs (`No` or max number of backlogs).             |
| `Minimum CGPA`       | Minimum required CGPA.                                         |
| `Valid Years`        | Eligible academic years (comma-separated).                    |
| `Allowed Departments`| Eligible departments (comma-separated or `All`).              |
| `Drop Year Allowed`  | Whether drop year is allowed (`Yes`/`No`).                     |

---

## Future Enhancements

- Add email notifications to inform students about eligibility updates.
- Integrate a dashboard to visualize eligibility statistics.
- Allow admin users to dynamically update criteria via a custom menu.

---

## License

This project is open-source and free to use under the MIT License.
