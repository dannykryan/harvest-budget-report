# Harvest Budget Report Automation

This Google Apps Script automates the generation of budget and rework reports for projects using data from the Harvest API. It creates a multi-tabbed Google Sheet summarizing project budgets, rework, client totals, and more, and emails a link to the recipients. The script can be run manually or scheduled to run automatically.

The project has now been updated to work with dummy data for testing when a Harvest account ID and access token has not been supplied.

View an example output sheet: [HERE](https://docs.google.com/spreadsheets/d/1FTmQwZVPsT-qDZld-Hjd87sc1HViZaC6-WuhwkAu3Dg/edit?usp=sharing)

---
Screenshot:
<img width="1869" height="686" alt="image" src="https://github.com/user-attachments/assets/17ae874c-35d0-44d5-9474-842913dbe0fa" />

---

## Table of Contents

- [Breakdown of spreadsheets created](#breakdown-of-spreadsheets-created)
- [Features](#features)
- [How It Works](#how-it-works)
- [Configuration](#configuration)
- [Manual Execution](#manual-execution)
- [Automated Scheduling](#automated-scheduling)
- [Code Structure](#code-structure)
- [Adding/Removing Recipients](#addingremoving-recipients)
- [API Integration](#api-integration)
- [Error Handling & Logging](#error-handling--logging)
- [Extending the Script](#extending-the-script)
- [Support](#support)

---

## Breakdown of spreadsheets created

Running the script will produce a multi-tabbed Google Sheet with the following reports:

- **Over Budget Projects:**  
  Projects that have exceeded their allocated budget (budget remaining is negative and budget is not monthly).

- **Nearly Over Budget Projects:**  
  Projects that are not over budget but have used at least 80% of their allocated budget (i.e., 20% or less budget remaining, but not negative, and budget is not monthly).

- **Profitable Projects:**  
  Projects that have more than 20% of their allocated budget remaining and are not currently active.

- **Reworked Projects:**  
  Projects from the Over Budget and Nearly Over Budget lists that have logged any time entries with the word "rework" in the task name. Only projects with nonzero rework hours are included.

- **Missing Budget Projects:**  
  Billable projects that do not have a budget set in Harvest (i.e., budget is missing, zero, or set to "none").

Additional sheets include:

- **Current FY Client Totals:**  
  A breakdown of billable, non-billable, and rework hours (and their values) by client and by month for the current financial year.

- **Current FY Rework by Team:**  
  Total and average rework costs for each team/role, calculated across all projects for the current financial year.

- **Summary:**  
  Key metrics and totals from the above sheets, including:
    - **Total Wastage Over Budget:** Total cost overrun for all projects that have exceeded their budget in the current financial year.
    - **Active Project Over Budget # Projects:** Number of active projects that are over budget.
    - **Active Project Wastage Over Budget £:** Total cost overrun for active projects that are over budget.
    - **Total Wastage Rework Cost:** Total cost of rework across all projects.
    - **Active Project Wastage Rework # Projects:** Number of active projects with logged rework.
    - **Active Project Wastage Rework £:** Total rework cost for active projects.
    - **Total Overall Wastage:** Combined total of over-budget and rework costs.
    - **Total Billable Hours / Total Hours:** Aggregate hours for all projects.
    - **AEDR (from budget):** Average earned daily rate, calculated from budget data.

A project is included in the report if **any** of the following are true:

- The project is currently active.
- The project is archived but its start date is within the current financial year.
- The project is archived but its end date is within the current financial year.
- The project (active or archived) has time logged to it during the current financial year.

---

## Features

- Fetches project, user, role, invoice, and time entry data from the Harvest API.
- Aggregates and categorizes projects by budget status:
  - Over Budget
  - Nearly Over Budget
  - Profitable
  - Missing Budget
- Calculates rework hours and costs by team/role.
- Generates a Google Sheet with multiple tabs:
  - Over Budget Projects
  - Nearly Over Budget Projects
  - Reworked Projects
  - Current FY Client Totals
  - Current FY Rework by Team
  - Summary
  - Missing Budget Projects
  - Profitable Projects
- Emails a link to the generated report to configured recipients.
- Can be run manually or on a schedule (e.g., monthly).

---

## How It Works

1. **Fetches Data:**  
   - Retrieves all projects, users, roles, invoices, and time entries from Harvest for the current financial year.

2. **Processes Data:**  
   - Categorizes projects by budget status and calculates rework by role and client totals.

3. **Generates Reports:**  
   - Creates a Google Sheet with multiple tabs, each summarizing a different aspect of project budgets and rework.

4. **Sends Email:**  
   - Emails the link to the generated report to the configured recipients.

---

## Configuration

### Script Properties

Set the following script properties in Google Apps Script (`File > Project properties > Script properties`):

- `harvestAccessToken`: Your Harvest API access token.
- `harvestAccountID`: Your Harvest account ID.
- `recipientEmail`: Your email/the email you want to receive the report.
- `outputFolderId`: the ID of a Googel Drive folder you have editing rights to where the report can be created. 
- `harvestBaseId`: Use the base Url of your harvest account to get correct links to your projects in the Harvest dashboard

### Folder and Recipients

- **Use Dummy Data:**  
  The project is now supplied with Dummy data which can be used for testing when you dont have access to a Harvest account ID or access token.
  You will still need to provide a Google Drive folder ID and email address which can be set in the 'Script Properties' section of the script near the top of the main 'productionScript.gs' file.    

---

## Manual Execution

### From Google Apps Script

To run the script manually:

1. Open the script in [Google Apps Script](https://script.google.com/).
2. Select the `createBudgetReport` function.
3. Click the **Run** button.

The script will fetch data, generate the report, and send the email as described above.

---

## Automated Scheduling

The script can be scheduled to run automatically (e.g., monthly):

1. In Google Apps Script, go to **Triggers** (clock icon in the left sidebar).
2. Click **Add Trigger**.
3. Choose the `createBudgetReport` function.
4. Set the event source to **Time-driven**.
5. Choose a schedule (e.g., "Month timer" > "Before end of month").
6. Save the trigger.

---

## Code Structure

- **API URLs & Headers:**  
  Defined at the top for Harvest endpoints.

- **Main Logic:**  
  - `createBudgetReport()`: Orchestrates fetching, processing, report generation, and emailing.
  - `fetchFromApi()`, `fetchAllPages()`: Handle API requests and pagination.
  - `fetchAllProjects()`, `fetchAllTimeEntriesFY()`: Fetch all projects and time entries for the financial year.
  - `getBudgetReportData()`, `getUsers()`, `getRoles()`: Fetch report, user, and role data.
  - `processBudgetReport()`, `getProjectDetails()`: Process and enrich project data for reporting.

- **Sheet Writers:**  
  - `createCombinedGoogleSheet()`: Creates and populates the multi-tabbed report.
  - `populateSheet()`: Populates Over Budget, Nearly Over Budget, and Profitable Projects tabs.
  - `populateReworkedProjectsSheet()`: Populates the Reworked Projects tab.
  - `populateClientTotalsSheet()`: Populates the Client Totals tab.
  - `populateReworkByTeamSheet()`: Populates the Rework by Team tab.
  - `populateSummarySheet()`: Populates the Summary tab.
  - `populateMissingBudgetSheet()`: Populates the Missing Budget Projects tab.

- **Utilities:**  
  - `calculateTimeByRole()`: Calculates hours and rework by role.
  - `mergeClientBalances()`: Aggregates client totals.
  - `cropSheet()`: Trims and formats sheets.
  - `projectTypeRename()`: Utility for budget type display.
  - `fetchInvoicesByProjectId()`: Fetches invoices for a project.

- **Email:**  
  - `sendEmail()`: Sends the report link to recipients.

---

## API Integration

- **Harvest:**  
  Used for all project, user, role, invoice, and time entry data.

All API calls use secure tokens stored in script properties.

---

## Error Handling & Logging

- The script uses `Logger.log()` for debugging and error reporting.
- API calls have retry logic and log failures.
- All errors are logged to the Apps Script log, accessible via **View > Logs** in the Apps Script editor.

---

## Extending the Script

- **To add new report categories or sheets:**  
  Create new functions similar to the existing `populate*Sheet` functions.
- **To change email recipients or folder:**  
  Update the `recipientEmail` and/or `outputFolderId` script properties in the project settings once copied over to Google Apps Script.
- **To change the reporting period:**  
  Adjust the date logic near the top of the script.
