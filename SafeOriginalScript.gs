// // create an alert email 
// async function budgetReportEmail() {

//   const currentDate = new Date();
//   const timeZone = "Europe/London"; // London timezone
//   const formattedDate = Utilities.formatDate(currentDate, timeZone, "yyyy/MM/dd HH:mm");

//   try {
//     const report = await getBudgetReport();
//     const processedBudgetReport = await processBudgetReport(report);

//     const sheetUrl = await createGoogleSheet(processedBudgetReport, formattedDate);

//     // Send Email
//     const subject = "Automated Harvest Budget Report " + formattedDate;
//     const body = "Hi Everyone,\n\nPlease find the attached budget report in the Google Sheet: " + sheetUrl + "\n\nAs this is an automated report, let me know if you'd like any changes.\n\nThanks,\n\nRobot Adam";
//     MailApp.sendEmail({
//       to: "danny.ryan@marketingpod.com",
//       subject: subject,
//       body: body
//     });
//   } catch (e) {
//     console.error(`Error in budgetReportEmail: ${e.message}`, e.stack);
//   }
// }

// async function fetchFromApi(url, options) {
//   const response = UrlFetchApp.fetch(url, options);
//   if (response.getResponseCode() !== 200) {
//     throw new Error(`Error fetching data from ${url}: ${response.getContentText()}`);
//   }
//   return JSON.parse(response.getContentText());
// }

// async function getBudgetReport() {
//   const apiUrl = "https://api.harvestapp.com/v2/reports/project_budget?is_active=true&per_page=2000";
//   const scriptProperties = PropertiesService.getScriptProperties();
//   const accessToken = scriptProperties.getProperty('harvestAccessToken');
//   console.log(accessToken)
//   const accountId = scriptProperties.getProperty('harvestAccountID');
//   console.log(accountId)

//   const headers = {
//     "Harvest-Account-Id": accountId,
//     "Authorization": `Bearer ${accessToken}`,
//     "User-Agent": "Budget Report AppScript Integration"
//   };

//   return fetchFromApi(apiUrl, { method: "get", headers, muteHttpExceptions: true });
// }

// async function processBudgetReport(budgetReport) {
//   try {
//     const filteredResults = [];
//     const users = await getUsers();

//     for (const item of budgetReport.results) {
//       if (item && item.budget_remaining < 0 && !item.budget_is_monthly && item.is_active) {
//         const projectDetails = await getProjectDetails(item.project_id, users);
//         let budgetPercentRemaining = item.budget === 0 ? 0 : (item.budget_remaining / item.budget) * 100;

//         // Update for roleBreakdownArray
//         const roleBreakdownArray = [];
//         for (const [role, details] of Object.entries(projectDetails.time_percentages_by_role)) {
//           roleBreakdownArray.push({
//             role: role,
//             totalHours: details.totalHours
//           });
//         }

//         const enrichedItem = {
//           ...item,
//           project_code: projectDetails.code,
//           budget_percent_remaining: budgetPercentRemaining.toFixed(2),
//           role_breakdown: roleBreakdownArray
//         };

//         filteredResults.push(enrichedItem);
//       }
//     }

//     filteredResults.sort((a, b) => a.budget_remaining - b.budget_remaining);

//     return filteredResults;
//   } catch (e) {
//     throw new Error(`Error processing budget report: ${e.message}`);
//   }
// }


// async function getProjectDetails(projectId, users) {
//   const projectApiUrl = `https://api.harvestapp.com/v2/projects/${projectId}`;
//   const scriptProperties = PropertiesService.getScriptProperties();
//   const accessToken = scriptProperties.getProperty('harvestAccessToken');
//   const accountId = scriptProperties.getProperty('harvestAccountID');

//   const headers = {
//     "Harvest-Account-Id": accountId,
//     "Authorization": `Bearer ${accessToken}`,
//     "User-Agent": "Budget Report AppScript Integration"
//   };

//   const projectDetails = await fetchFromApi(projectApiUrl, { method: "get", headers, muteHttpExceptions: true });
//   const timeEntries = await getTimeEntries(projectId);
//   const timePercentagesByRole = calculateTimeByRole(timeEntries, users);

//   projectDetails['time_percentages_by_role'] = timePercentagesByRole;

//   return projectDetails;
// }

// async function getTimeEntries(projectId) {
//   const timeEntriesApiUrl = `https://api.harvestapp.com/v2/time_entries?project_id=${projectId}`;
//   const scriptProperties = PropertiesService.getScriptProperties();
//   const accessToken = scriptProperties.getProperty('harvestAccessToken');
//   const accountId = scriptProperties.getProperty('harvestAccountID');

//   const headers = {
//     "Harvest-Account-Id": accountId,
//     "Authorization": `Bearer ${accessToken}`,
//     "User-Agent": "Budget Report AppScript Integration"
//   };

//   return fetchFromApi(timeEntriesApiUrl, { method: "get", headers, muteHttpExceptions: true }).then(data => data.time_entries);
// }

// function projectTypeRename(projectBudgetBy) {
//   let projectBudgetByRenamed;
//   switch (projectBudgetBy) {
//     case "project":
//       projectBudgetByRenamed = "hours";
//       break;
//     case "project_cost":
//       projectBudgetByRenamed = "£";
//       break;
//     default:
//       projectBudgetByRenamed = "Error";
//   }
//   return projectBudgetByRenamed;
// }

// function calculateTimeByRole(timeEntries, users) {
//   const timeByRole = {};
//   const userRoles = users.reduce((acc, user) => {
//     acc[user.id] = user.roles[0] || 'Unknown'; // Assuming each user has at least one role
//     return acc;
//   }, {});

//   timeEntries.forEach(entry => {
//     const role = userRoles[entry.user.id] || 'Unknown';

//     if (!timeByRole[role]) {
//       timeByRole[role] = { totalHours: 0, entries: [] };
//     }

//     timeByRole[role].entries.push({
//       user_id: entry.user.id,
//       hours: entry.hours
//     });

//     timeByRole[role].totalHours += entry.hours;
//   });

//   return timeByRole;
// }

// async function getUsers() {
//   const usersApiUrl = "https://api.harvestapp.com/v2/users";
//   const scriptProperties = PropertiesService.getScriptProperties();
//   const accessToken = scriptProperties.getProperty('harvestAccessToken');
//   const accountId = scriptProperties.getProperty('harvestAccountID');

//   const headers = {
//     "Harvest-Account-Id": accountId,
//     "Authorization": `Bearer ${accessToken}`,
//     "User-Agent": "Budget Report AppScript Integration"
//   };

//   return fetchFromApi(usersApiUrl, { method: "get", headers, muteHttpExceptions: true }).then(data => data.users);
// }

// async function createGoogleSheet(data, formattedDate) {

//   const ss = SpreadsheetApp.create("Harvest Budget Report " + formattedDate);

//   // Move the spreadsheet to the specified folder
//   const folderId = "1QylzuhBPTSst8o6mSQhqru0PsrOz1iu9";
//   const file = DriveApp.getFileById(ss.getId());
//   const folder = DriveApp.getFolderById(folderId);
//   file.moveTo(folder);

//   const sheet = ss.getActiveSheet();

//   // Identify all unique roles
//   const allRoles = new Set();
//   data.forEach(item => {
//     item.role_breakdown.forEach(role => {
//       allRoles.add(role.role);
//     });
//   });

//   // Set headers
//   const headers = ["Job number", "Client name", "Project name", "Budgeted by", "Budget total", "Budget spent", "Over budget by", "Over budget by %", ...allRoles];
//   sheet.appendRow(headers);

//   // Populate data
//   data.forEach(item => {
//     const rowData = new Array(headers.length).fill(""); // Initialize row data
//     rowData[0] = item.project_code;
//     rowData[1] = item.client_name;
//     rowData[2] = item.project_name;
//     rowData[3] = projectTypeRename(item.budget_by);
//     rowData[4] = item.budget;
//     rowData[5] = item.budget_spent;
//     rowData[6] = Math.abs(item.budget_remaining);
//     rowData[7] = Math.abs(Number(item.budget_percent_remaining)) + "%";

//     item.role_breakdown.forEach(role => {
//       const roleIndex = headers.indexOf(role.role);
//       rowData[roleIndex] = parseFloat(role.totalHours).toFixed(2) + " hours";
//     });

//     sheet.appendRow(rowData);
//   });

//   // Sort the spreadsheet from low to high (most to least budget lost)
//   sheet.getRange("A2:Z").sort({column: 1, ascending: true});

//   // Set frozen first row and format as bold
//   sheet.setFrozenRows(1);
//   sheet.getRange("A1:1").setFontWeight('bold');

//   // Set format of the % budget remaining
//   sheet.getRange("D2:D").setNumberFormat("#0.##%");

//   // Auto-resize columns for readability
//   for (let i = 1; i <= headers.length; i++) {
//     sheet.autoResizeColumn(i);
//   }

//   // Crop the spreadsheet
//   cropSheet(sheet);

//   return ss.getUrl();
// }

// // Crop the spreadsheet from: https://developers.google.com/apps-script/add-ons/clean-sheet
// function cropSheet(sheettocrop) {
//   const dataRange = sheettocrop.getDataRange();
//   sheettocrop = dataRange.getSheet();

//   let numRows = dataRange.getNumRows();
//   let numColumns = dataRange.getNumColumns();

//   const maxRows = sheettocrop.getMaxRows();
//   const maxColumns = sheettocrop.getMaxColumns();

//   const numFrozenRows = sheettocrop.getFrozenRows();
//   const numFrozenColumns = sheettocrop.getFrozenColumns();

//   // If last data row is less than maximium row, then deletes rows after the last data row.
//   if (numRows < maxRows) {
//     numRows = Math.max(numRows, numFrozenRows + 1); // Don't crop empty frozen rows.
//     sheettocrop.deleteRows(numRows + 1, maxRows - numRows);
//   }

//   // If last data column is less than maximium column, then deletes columns after the last data column.
//   if (numColumns < maxColumns) {
//     numColumns = Math.max(numColumns, numFrozenColumns + 1); // Don't crop empty frozen columns.
//     sheettocrop.deleteColumns(numColumns + 1, maxColumns - numColumns);
//   }
// }
