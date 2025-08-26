// This script should be ran in Google Apps Script

// Development Settings
let folderId = "1234uhBPTSst8o6mSQhqru0PsrOzABCD"; // Replace with your Google Drive folder ID for testing
let recipientEmails = [
  "dannykryan@gmail.com" // Replace with your email for testing
];

// Api urls
const budgetReportApi = "https://api.harvestapp.com/v2/reports/project_budget?&per_page=2000";
const usersApiUrl = "https://api.harvestapp.com/v2/users";
const rolesApiUrl = "https://api.harvestapp.com/v2/roles";
const invoicesApiUrl = `https://api.harvestapp.com/v2/invoices`;
// projectApiUrl and timeEntriesApiUrl will require a project ID appended
const projectApiUrl = "https://api.harvestapp.com/v2/projects/";
const timeEntriesApiUrl = "https://api.harvestapp.com/v2/time_entries?project_id=";
const invoiceApiUrl = "https://api.harvestapp.com/v2/invoices?project_id=";

// Script Properties include Harvest token and ID needed for API calls
const scriptProperties = PropertiesService.getScriptProperties();
const accessToken = scriptProperties.getProperty("harvestAccessToken");
const accountId = scriptProperties.getProperty("harvestAccountID");

// Create API headers object using Harvest token and ID
const headers = {
  "Harvest-Account-Id": accountId,
  Authorization: `Bearer ${accessToken}`,
  "User-Agent": "Budget Report AppScript Integration",
};

// Get and format the date
const currentDate = new Date();
const timeZone = "Europe/London"; // London timezone
const formattedDate = Utilities.formatDate(currentDate, timeZone, "yyyy/MM/dd HH:mm");

// Determine the start and end dates of the financial year
const currentYear = currentDate.getFullYear();
const currentMonth = currentDate.getMonth();
const currentDay = currentDate.getDate();

const startYear = currentMonth > 3 || (currentMonth === 3 && currentDay >= 6) ? currentYear : currentYear - 1;
const endYear = startYear + 1;

const currentFY = {
  start: new Date(startYear, 3, 1), // April 1st
  // start: new Date(startYear, 10, 6), // reduced timeframe for development, remember to change back after development
  end: new Date(endYear, 3, 1), // April 1st
};

let currentFYClientTotals = [];
let allTimeEntriesFY = {};
let totalReworkTimeByRole = {};
let totalReworkCost = 0;
let totalBillableHours = 0;
let totalHours = 0;
let totalCostOverBudget = 0;
let billableRateByRole = {};
// The billable rate by role is populated in the calculateTimeByRole() function
let defaultDayRate = 975;
let defaultHourlyRate = defaultDayRate / 7.5; // Default hourly rate based on a 7.5-hour workday

// Main function
async function createBudgetReport() {
  try {
    const [rawBudgetData, users, roles, allProjects] = await Promise.all([
      getBudgetReportData(),
      getUsers(),
      getRoles(),
      fetchAllProjects(), // Fetch all projects
    ]);

    // get list of projects which are:
    // 1. Active
    // 2. Archived but have a start or end date within the financial year
    // 3. Archived but have time logged to it during the financial year
    // 4. Not including Pod (12129858) or Pod Marketing (13980295) and

    const activeProjects = Object.values(allProjects).filter((project) => project.is_active);
    const archivedProjects = Object.values(allProjects).filter((project) => !project.is_active);
    const archivedProjectsWithStartDate = archivedProjects.filter((project) => {
      const startDate = new Date(project.starts_on);
      return startDate >= currentFY.start && startDate <= currentFY.end;
    });
    const archivedProjectsWithEndDate = archivedProjects.filter((project) => {
      const endDate = new Date(project.ends_on);
      return endDate >= currentFY.start && endDate <= currentFY.end;
    });

    await fetchAllTimeEntriesFY();

    // Logger.log(`allTimeEntriesFY count: ${Object.keys(allTimeEntriesFY).length}`);

    // get list of projects which have time logged to it during the financial year
    const projectsWithTimeLogged = Object.values(allTimeEntriesFY).reduce((acc, entries) => {
      if (entries.length > 0) {
        acc.push(entries[0].project);
      }
      return acc;
    }, []);

    // create unique list of projects
    const reportableProjectsIncludingDuplicates = activeProjects.concat(archivedProjectsWithStartDate, archivedProjectsWithEndDate, projectsWithTimeLogged);
    const reportableProjectsUnique = Array.from(new Set(reportableProjectsIncludingDuplicates.map((project) => project.id))).map((id) => {
      return reportableProjectsIncludingDuplicates.find((project) => project.id === id);
    });

    // Filter out marketing pod and pod marketing projects
    const reportableProjects = reportableProjectsUnique.filter((project) => project.client && project.client.id !== 12129858 && project.client.id !== 13980295);

    // filter to ambion id 12124635 or HIVE 12129899 for testing
    // const reportableProjects = reportableProjectsUnique.filter((project) => project.client && project.client.id === 12129899);

    Logger.log(`Reportable projects count: ${reportableProjects.length}`);
    Logger.log(`Reportable projects example: ${JSON.stringify(reportableProjects[0])}`);

    const { masterReport } = await processBudgetReport(rawBudgetData, users, roles, reportableProjects);

    // Filter `masterReport` for reportable projects

    // Filter `masterReport` for specific reports:
    const overBudgetReport = masterReport.filter((item) => item.remaining_budget < 0 && !item.budget_is_monthly);
    const nearlyOverBudgetReport = masterReport.filter((item) => {
      const budgetPercentRemaining = item.budget === 0 ? 0 : (item.remaining_budget / item.budget) * 100;
      return budgetPercentRemaining <= 20 && budgetPercentRemaining >= 0 && !item.budget_is_monthly;
    });
    const profitableProjectsReport = masterReport.filter((item) => {
      const budgetPercentRemaining = item.budget > 0 ? (item.remaining_budget / item.budget) * 100 : 0;
      return budgetPercentRemaining > 20 && !item.is_active;
    });

    const missingBudgetReport = reportableProjects.filter((item) => {
      return (item.cost_budget === null || item.cost_budget === 0 || item.budget_by === "none") && item.is_billable === true;
    });

    const combinedSheetUrl = createCombinedGoogleSheet(masterReport, overBudgetReport, nearlyOverBudgetReport, profitableProjectsReport, currentFYClientTotals, roles, missingBudgetReport);
    Logger.log(`Combined report created at: ${combinedSheetUrl}`);

    sendEmail(combinedSheetUrl);
    Logger.log("Email sent");
  } catch (e) {
    Logger.log(`Error in createBudgetReport: ${e.message}`);
  }
  // createBudgetReport();
}

// Reusable function to fetch data from an API - includes automatic retry logic
async function fetchFromApi(url, options, retries = 3, delay = 1000) {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      // Wrap the synchronous UrlFetchApp.fetch in a Promise
      const response = await new Promise((resolve, reject) => {
        try {
          const result = UrlFetchApp.fetch(url, options);
          resolve(result);
        } catch (error) {
          reject(error);
        }
      });

      // Check response code
      if (response.getResponseCode() === 200) {
        return JSON.parse(response.getContentText());
      } else {
        throw new Error(`Attempt ${attempt} failed: ${response.getResponseCode()} ${response.getContentText()}`);
      }
    } catch (error) {
      Logger.log(`Attempt ${attempt} failed: ${error.message}`);
      if (attempt < retries) {
        Logger.log(`Retrying in ${delay} ms...`);
        Utilities.sleep(delay); // Apps Script utility to add delay
      } else {
        throw new Error(`All ${retries} attempts failed.`);
      }
    }
  }
}

// Fetch all projects
async function fetchAllProjects() {
  try {
    let allProjects = {};
    let page = 1;

    while (true) {
      const projectsUrl = `https://api.harvestapp.com/v2/projects?page=${page}`;
      const response = await fetchFromApi(projectsUrl, {
        method: "GET",
        headers,
        muteHttpExceptions: true,
      });

      if (!response.projects.length) break;

      response.projects.forEach((project) => {
        allProjects[project.id] = project;
      });

      page++;
    }

    return allProjects;
  } catch (error) {
    Logger.log("Error fetching all projects: " + error.message);
    return {};
  }
}

// Fetch all pages
async function fetchAllPages({ url, method = "GET", muteHttpExceptions = true } = {}) {
  let pages = [];
  let page = 1;

  Logger.log(`Fetching Data from Harvest. Please wait...`);
  try {
    while (url) {
      const response = await fetchFromApi(url, {
        method,
        headers,
        muteHttpExceptions,
      });

      pages.push(response);
      url = response.links.next;
      page++;
    }
  } catch (error) {
    Logger.log("Error fetching all time entries: " + error.message);
  }

  return pages;
}

// Fetch time entries for projects with start or end date within FY
async function fetchAllTimeEntriesFY() {
  try {
    const startFYFormatted = Utilities.formatDate(currentFY.start, timeZone, "yyyy-MM-dd");
    const endFYFormatted = Utilities.formatDate(currentDate, timeZone, "yyyy-MM-dd");

    let allTimeEntriesFYData = [];
    let nextPageUrl = `https://api.harvestapp.com/v2/time_entries?from=${startFYFormatted}&to=${endFYFormatted}&page=1`;

    allTimeEntriesFYData = (await fetchAllPages({ url: nextPageUrl })).reduce((previous, current) => previous.concat(current.time_entries), []);

    allTimeEntriesFYData.forEach((entry) => {
      if (!allTimeEntriesFY[entry.project.id]) {
        allTimeEntriesFY[entry.project.id] = [];
      }
      allTimeEntriesFY[entry.project.id].push(entry);
    });
  } catch (error) {
    Logger.log("Error fetching all time entries: " + error.message);
  }
}

// Get the Budget Report data
async function getBudgetReportData() {
  try {
    return (await fetchAllPages({ url: budgetReportApi })).reduce((previous, current) => ({ results: previous.results.concat(current.results) }), { results: [] });
  } catch (error) {
    Logger.log("Error fetching Budget Report Data: " + error.message);
    return null;
  }
}

// Get user data
async function getUsers() {
  const usersResponse = await fetchFromApi(usersApiUrl, {
    method: "get",
    headers,
    muteHttpExceptions: true,
  });
  // Ensure 'users' is an array
  return usersResponse.users || [];
}

// Get role data
async function getRoles() {
  const rolesResponse = await fetchFromApi(rolesApiUrl, {
    method: "get",
    headers,
    muteHttpExceptions: true,
  });
  // Ensure 'roles' is an array
  return rolesResponse.roles || [];
}

// Process the Budget Report
async function processBudgetReport(rawBudgetData, users, roles, reportableProjects) {
  Logger.log("Processing all projects to compile a master report...");

  try {
    const masterReport = [];
    const clientProjectsMap = new Map();

    if (!rawBudgetData || !rawBudgetData.results || !users) {
      throw new Error("Invalid input data");
    }

    // filter out projects that are not reportable
    const reportableProjectsBudgetData = rawBudgetData.results.filter((item) => {
      return reportableProjects.some((project) => project.id === item.project_id);
    });

    for (const item of reportableProjectsBudgetData) {
      // Get detailed project information for reportable projects
      const projectDetails = await getProjectDetails(item.project_id, users, roles, reportableProjects, "master");

      if (!projectDetails) continue;

      const budgetPercentRemaining = item.budget === 0 ? 0 : (item.budget_remaining / item.budget) * 100;

      // Update for roleBreakdownArray
      const roleBreakdownArray = [];
      for (const [role, details] of Object.entries(projectDetails.total_time_by_role)) {
        roleBreakdownArray.push({
          role: role,
          totalHours: details.totalHours,
        });
      }

      const reworkRoleBreakdownArray = [];
      for (const [role, details] of Object.entries(projectDetails.rework_time_by_role)) {
        if (details.reworkHours > 0) {
          reworkRoleBreakdownArray.push({
            role: role,
            totalHours: details.reworkHours, // Corrected
            totalReworkCost: details.reworkCost, // Corrected
          });
        }
      }

      const formattedStartsOn = projectDetails.starts_on ? Utilities.formatDate(new Date(projectDetails.starts_on), timeZone, "dd/MM/yyyy") : null;
      const formattedEndsOn = projectDetails.ends_on ? Utilities.formatDate(new Date(projectDetails.ends_on), timeZone, "dd/MM/yyyy") : null;

      const enrichedItem = {
        ...item,
        project_code: projectDetails.code,
        budget_percent_remaining: budgetPercentRemaining.toFixed(2),
        role_breakdown: roleBreakdownArray,
        rework_role_breakdown: reworkRoleBreakdownArray,
        starts_on: formattedStartsOn,
        ends_on: formattedEndsOn,
        first_time_entry: projectDetails.first_time_entry,
        last_time_entry: projectDetails.last_time_entry,
        total_logged_hours: projectDetails.total_logged_hours,
        billable_logged_hours: projectDetails.billable_logged_hours,
        total_balance: item.budget_remaining, // Add total balance for all time
        id: projectDetails.id,
        associated_invoices: projectDetails.associated_invoices,
        time_entries: projectDetails.time_entries || [],
        remaining_budget: item.budget_remaining,
        budget_utilization: item.budget === 0 ? "N/A" : ((item.budget_spent / item.budget) * 100).toFixed(2) + "%",
      };

      masterReport.push(enrichedItem);

      // Group projects by client
      if (!clientProjectsMap.has(enrichedItem.client_name)) {
        clientProjectsMap.set(enrichedItem.client_name, []);
      }
      clientProjectsMap.get(enrichedItem.client_name).push(enrichedItem);
    }

    Logger.log("Master report compilation complete.");
    return { masterReport, clientProjectsMap };
  } catch (e) {
    throw new Error(`Error processing master budget report: ${e.message}`);
  }
}

// Get project details
async function getProjectDetails(projectId, users, roles, reportableProjects) {
  try {
    const projectDetails = reportableProjects.find((project) => project.id === projectId);
    if (!projectDetails) {
      Logger.log(`Project with ID ${projectId} not found in cached data.`);
      return null;
    }

    // Retrieve time entries for each reportable project - by calling the timeEntries API for each project id
    // Do we still want to get the totals for the financial year as well? YES

    const timeEntriesData = await fetchAllPages({ url: `https://api.harvestapp.com/v2/time_entries?project_id=${projectId}` });
    const timeEntries = timeEntriesData[0].time_entries || [];

    if (timeEntries === null) return null;

    const projectInvoices = await fetchInvoicesByProjectId(projectId);

    const totalLoggedHours = timeEntries.reduce((total, entry) => total + entry.hours, 0);
    const billableHours = timeEntries.filter((entry) => entry.billable).reduce((total, entry) => total + entry.hours, 0);

    // Calculate time by role in a single pass
    const timeByRole = calculateTimeByRole(timeEntries, roles);

    // Populate total_time_by_role and rework_time_by_role
    projectDetails["total_time_by_role"] = timeByRole;
    projectDetails["rework_time_by_role"] = Object.entries(timeByRole).reduce((acc, [role, data]) => {
      acc[role] = {
        reworkHours: data.totalReworkHours,
        reworkCost: data.totalReworkCost,
      };
      return acc;
    }, {});

    // Set up first and last time entries
    if (timeEntries.length === 0) {
      projectDetails["first_time_entry"] = "N/A";
      projectDetails["last_time_entry"] = "N/A";
    } else {
      const timeEntryDates = timeEntries.map((entry) => new Date(entry.spent_date));
      const firstTimeEntry = new Date(Math.min(...timeEntryDates));
      const lastTimeEntry = new Date(Math.max(...timeEntryDates));

      projectDetails["first_time_entry"] = Utilities.formatDate(firstTimeEntry, timeZone, "dd/MM/yyyy");
      projectDetails["last_time_entry"] = Utilities.formatDate(lastTimeEntry, timeZone, "dd/MM/yyyy");
    }

    projectDetails["associated_invoices"] = projectInvoices;
    projectDetails["total_logged_hours"] = totalLoggedHours;
    projectDetails["billable_logged_hours"] = billableHours;

    // Create the clientBalancesFYTD object
    const clientBalancesFYTD = {};
    timeEntries.forEach((entry) => {
      const entryDate = new Date(entry.spent_date);
      if (entryDate >= currentFY.start && entryDate <= currentFY.end) {
        const monthKey = `${entryDate.getFullYear()}-${entryDate.getMonth() + 1}`;
        if (!clientBalancesFYTD[monthKey]) {
          clientBalancesFYTD[monthKey] = { billable_total: 0, non_billable_total: 0, rework_total: 0 };
        }

        if (entry.task.name.toLowerCase().includes("rework")) {
          clientBalancesFYTD[monthKey].rework_total += entry.hours;
        } else if (entry.billable) {
          clientBalancesFYTD[monthKey].billable_total += entry.hours;
        } else {
          clientBalancesFYTD[monthKey].non_billable_total += entry.hours;
        }
      }
    });

    const clientBalanceArray = [
      {
        client: {
          [projectDetails.client.name]: clientBalancesFYTD,
        },
      },
    ];

    mergeClientBalances(clientBalanceArray);

    return projectDetails;
  } catch (error) {
    Logger.log(`Error in getProjectDetails: ${error.message}`);
    return null;
  }
}

// Helper function to merge project balance into global totals
function mergeClientBalances(newClientBalanceArray) {
  // Mapping month indices to month names
  const monthNames = {
    4: "April",
    5: "May",
    6: "June",
    7: "July",
    8: "August",
    9: "September",
    10: "October",
    11: "November",
    12: "December",
    1: "January",
    2: "February",
    3: "March",
  };

  // Function to convert month key from "YYYY-M" to "Month Name"
  function convertMonthKey(monthKey) {
    const [year, month] = monthKey.split("-").map(Number);
    return `${monthNames[month]} ${year}`;
  }

  // Function to sort the month keys from April to March, considering the year
  function sortMonthKeys(keys) {
    const monthOrder = ["April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March"];

    return keys.sort((a, b) => {
      const [monthA, yearA] = a.split(" ");
      const [monthB, yearB] = b.split(" ");

      if (yearA !== yearB) {
        return yearA - yearB;
      }

      return monthOrder.indexOf(monthA) - monthOrder.indexOf(monthB);
    });
  }

  newClientBalanceArray.forEach((entry) => {
    const clientName = Object.keys(entry.client)[0];
    const clientBalance = entry.client[clientName];

    // Find existing client entry in currentFYClientTotals
    let clientEntry = currentFYClientTotals.find((item) => Object.keys(item)[0] === clientName);

    if (!clientEntry) {
      // If client not found, add new entry
      clientEntry = { [clientName]: {} };
      currentFYClientTotals.push(clientEntry);
    }

    const clientData = clientEntry[clientName];

    // Merge balances by month with conversion
    Object.keys(clientBalance).forEach((monthKey) => {
      const monthName = convertMonthKey(monthKey);

      if (!clientData[monthName]) {
        clientData[monthName] = { billable_total: 0, non_billable_total: 0, rework_total: 0 };
      }

      clientData[monthName].billable_total += clientBalance[monthKey].billable_total;
      clientData[monthName].non_billable_total += clientBalance[monthKey].non_billable_total;
      clientData[monthName].rework_total += clientBalance[monthKey].rework_total;
    });

    // Sort the months within the client data
    const sortedMonthKeys = sortMonthKeys(Object.keys(clientData));
    const sortedClientData = {};
    sortedMonthKeys.forEach((monthKey) => {
      sortedClientData[monthKey] = clientData[monthKey];
    });
    clientEntry[clientName] = sortedClientData;
  });
}

// Calculate time by role
function calculateTimeByRole(timeEntries, roles) {
  const calculatedTimes = {};
  const reworkEntries = []; // Separate storage for rework entries
  const userRoles = roles.reduce((acc, role) => {
    role.user_ids.forEach((userId) => {
      acc[userId] = role.name; // Map user ID to the role name
    });
    return acc;
  }, {});

  // First pass: Process non-rework entries
  timeEntries.forEach((entry) => {
    const role = userRoles[entry.user.id] || "Unknown";
    const isRework = entry.task.name.toLowerCase().includes("rework");

    // Store rework entries for later
    if (isRework) {
      reworkEntries.push(entry);
      return; // Skip rework entries for now
    }

    // Initialize calculatedTimes for the role if not already present
    if (!calculatedTimes[role]) {
      calculatedTimes[role] = {
        totalHours: 0,
        totalReworkHours: 0,
        totalReworkCost: 0,
        entries: [],
      };
    }

    // Add entry and update totals
    calculatedTimes[role].entries.push(entry);
    calculatedTimes[role].totalHours += entry.hours;

    // Populate billableRateByRole if not already set
    if (!billableRateByRole[role] && entry.billable_rate) {
      billableRateByRole[role] = entry.billable_rate;
    }
    if (entry.billable) {
      totalBillableHours += entry.hours;
    }
    totalHours += entry.hours;
  });

  // Second pass: Process rework entries
  // Rework entries are processed separately to calculate rework billable rates
  reworkEntries
    .filter(entry => {
      const entryDate = new Date(entry.spent_date);
      return entryDate >= currentFY.start && entryDate <= currentFY.end;
    })
    .forEach((entry) => {
      const role = userRoles[entry.user.id] || "Unknown";
      const hourlyRate = billableRateByRole[role] || 0; // Default rate to 0 if not set

    // Initialize calculatedTimes for the role if not already present
    if (!calculatedTimes[role]) {
      calculatedTimes[role] = {
        totalHours: 0,
        totalReworkHours: 0,
        totalReworkCost: 0,
        entries: [],
      };
    }

    // Add entry and update totals directly within `calculatedTimes`
    calculatedTimes[role].entries.push(entry);
    calculatedTimes[role].totalReworkHours += entry.hours;
    calculatedTimes[role].totalReworkCost += entry.hours * hourlyRate;
  });

  // Now update totalReworkTimeByRole with calculatedTimes
  Object.keys(calculatedTimes).forEach((role) => {
    const roleData = calculatedTimes[role];
    if (!totalReworkTimeByRole[role]) {
      totalReworkTimeByRole[role] = {
        totalReworkCost: 0,
        totalReworkHours: 0,
      };
    }
    totalReworkTimeByRole[role].totalReworkCost += roleData.totalReworkCost;
    totalReworkTimeByRole[role].totalReworkHours += roleData.totalReworkHours;
  });

  return calculatedTimes;
}

// Create the Google sheet with multiple tabs for different reports
function createCombinedGoogleSheet(masterReport, overBudgetData, nearlyOverBudgetData, profitableProjectsReport, currentFYClientTotals, roles, missingBudgetReport) {
  const ss = SpreadsheetApp.create("Harvest Budget Report " + formattedDate);

  // Move the spreadsheet to the specified folder Id
  const file = DriveApp.getFileById(ss.getId());
  const folder = DriveApp.getFolderById(folderId);
  file.moveTo(folder);

  // Remove default empty sheet if it exists
  const sheets = ss.getSheets();
  if (sheets.length > 1) {
    ss.deleteSheet(sheets[0]);
  }

  // Create and populate the 'Over Budget' sheet
  const overBudgetSheet = ss.getSheetByName("Sheet1") || ss.insertSheet(">100% Budget Projects");
  overBudgetSheet.setName(">100% Budget Projects");
  populateSheet(overBudgetSheet, overBudgetData, "overBudget");

  // Create and populate the 'Nearly Over Budget' sheet
  const nearlyOverBudgetSheet = ss.getSheetByName(">80% Budget Projects") || ss.insertSheet(">80% Budget Projects");
  nearlyOverBudgetSheet.setName(">80% Budget Projects");
  populateSheet(nearlyOverBudgetSheet, nearlyOverBudgetData, "nearlyOverBudget");

  const reworkedProjectsData = masterReport.filter(item => {
    if (!item.rework_role_breakdown || !Array.isArray(item.rework_role_breakdown)) {
      return false;
    }
    const totalReworkHours = item.rework_role_breakdown.reduce((sum, role) => sum + (role.totalHours || 0), 0);
    return totalReworkHours > 0;
  });

  // Create and populate the 'Reworked Projects' sheet
  const reworkedProjectsSheet = ss.getSheetByName("Reworked Projects") || ss.insertSheet("Reworked Projects");
  populateReworkedProjectsSheet(reworkedProjectsSheet, reworkedProjectsData);

  // Create and populate the 'Current FY Client Totals' sheet
  const clientTotalsSheet = ss.getSheetByName("Current FY Client Totals") || ss.insertSheet("Current FY Client Totals");
  populateClientTotalsSheet(clientTotalsSheet, currentFYClientTotals);

  // Create and populate the 'Current FY Rework by Team' sheet
  const reworkByTeamSheet = ss.getSheetByName("Current FY Rework by Team") || ss.insertSheet("Current FY Rework by Team");
  populateReworkByTeamSheet(reworkByTeamSheet, currentFYClientTotals, roles);

  // Create and populate the 'Summary' sheet
  const summarySheet = ss.getSheetByName("Summary") || ss.insertSheet("Summary");
  populateSummarySheet(summarySheet, overBudgetData);

  // Move Summary sheet to first position
  ss.setActiveSheet(summarySheet);
  ss.moveActiveSheet(0);

  // Create and populate the 'Missing Budget Projects' sheet
  const missingBudgetSheet = ss.getSheetByName("Missing Budget Projects") || ss.insertSheet("Missing Budget Projects");
  populateMissingBudgetSheet(missingBudgetSheet, missingBudgetReport);

  // Create and populate the 'Profitable project' sheet
  const profitableProjectSummarySheet = ss.getSheetByName("Profitable Projects") || ss.insertSheet("Profitable Projects");
  populateSheet(profitableProjectSummarySheet, profitableProjectsReport, "profitable");

  // Delete the default empty sheet if necessary
  const remainingSheets = ss.getSheets();
  if (remainingSheets.length > 8) {
    ss.deleteSheet(remainingSheets[0]);
  }

  return ss.getUrl();
}

function populateReworkedProjectsSheet(sheet, data) {
  try {
    Logger.log("Populating 'Reworked Projects' spreadsheet");

    // Identify all unique roles
    const allRoles = data.reduce((roles, item) => {
      if (item.role_breakdown) {
        item.role_breakdown.forEach((role) => roles.add(role.role));
      }
      return roles;
    }, new Set());

    const headers = [
      "Job number",
      "Client name",
      "Project name",
      "Status",
      "Rework Hours",
      "Rework Cost (£)",
      "Over budget by %",
      "Harvest Start Date",
      "Harvest End Date",
      "Start Billable Logged Date",
      "Latest Billable Logged Date",
      "Associated Invoices",
      "Link to Project in Harvest",
      "Rework Reason",
      ...Array.from(allRoles).map((role) => `${role} Rework Hrs`), // Add all roles as headers
    ];

    sheet.appendRow(headers);

    data.forEach((item) => {
      // Calculate total rework hours and cost
      let totalReworkHours = 0;
      let totalReworkCost = 0;
      if (Array.isArray(item.rework_role_breakdown)) {
        totalReworkHours = item.rework_role_breakdown.reduce((sum, role) => sum + (role.totalHours || 0), 0);
        totalReworkCost = item.rework_role_breakdown.reduce((sum, role) => sum + (role.totalReworkCost || 0), 0);
      }

      // Skip appending rows if no rework hours
      if (totalReworkHours === 0) return;

      const rowData = new Array(headers.length).fill(""); // Initialize row data
      rowData[0] = item.project_code || "";
      rowData[1] = item.client_name || "";
      rowData[2] = item.project_name || "";
      rowData[3] = item.is_active ? "Active" : "Archived";
      rowData[4] = totalReworkHours.toFixed(2); // Rework Hours
      rowData[5] = totalReworkCost.toFixed(2); // Rework Cost (£)

      // Update rowData[6] with calculated over-budget percentage
      if (item.budget > 0) {
        // Avoid division by zero
        const overBudgetPercentage = ((item.budget_spent - item.budget) / item.budget) * 100;
        rowData[6] = `${overBudgetPercentage.toFixed(2)}%`; // Format to 2 decimal places
      } else {
        rowData[6] = "N/A"; // Handle cases where budget is zero or missing
      }

      rowData[7] = item.starts_on || ""; // Harvest Start Date
      rowData[8] = item.ends_on || ""; // Harvest End Date
      rowData[9] = item.first_time_entry || ""; // Start Billable Logged Date
      rowData[10] = item.last_time_entry || ""; // Latest Billable Logged Date
      rowData[11] = Array.isArray(item.associated_invoices) ? item.associated_invoices.join(", ") : item.associated_invoices || ""; // Associated Invoices
      rowData[12] = `https://marketingpod.harvestapp.com/projects/${item.id || ""}`; // Link to Project in Harvest
      rowData[13] = item.rework_reason || ""; // Rework Reason
      // Ensure all role columns in rowData are initialized to 0
      Array.from(allRoles).forEach((role) => {
        const roleColumnName = `${role} Rework Hrs`;
        const roleIndex = headers.indexOf(roleColumnName);
        if (roleIndex !== -1) {
          rowData[roleIndex] = "0.00"; // Initialize with 0.00
        }
      });

      if (Array.isArray(item.rework_role_breakdown)) {
        item.rework_role_breakdown.forEach((role) => {
          const roleColumnName = `${role.role} Rework Hrs`; // Construct the correct column name
          const roleIndex = headers.indexOf(roleColumnName); // Find the correct header index
          const roleHours = parseFloat(role.totalHours) || 0; // Default to 0 if undefined or NaN
          if (roleIndex !== -1) {
            rowData[roleIndex] = roleHours.toFixed(2); // Overwrite only if role is present
          }
        });
      } else {
        Logger.log(`Item with missing or invalid rework_role_breakdown: ${JSON.stringify(item)}`);
      }
      sheet.appendRow(rowData);
    });
  } catch (error) {
    Logger.log(`Error in populateReworkedProjectsSheet: ${error.message}`);
  }
  // Set frozen first row and format as bold
  sheet.setFrozenRows(1);
  sheet.getRange("A1:1").setFontWeight("bold");

  // Set date columns to be right-aligned
  sheet.getRange("H2:K").setHorizontalAlignment("right");

  // Auto-resize columns for readability
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
  cropSheet(sheet, headers);
}

// Function to populate a sheet with data
function populateSheet(sheet, data, reportType) {
  try {
    Logger.log(`Populating '${reportType}' spreadsheet`);

    const overBudgetHeaders = ["Job number", "Client name", "Project name", "Status", "Budgeted by", "Budget total", "Budget spent", "Over budget by", "Over budget by %", "Harvest Start Date", "Harvest End Date", "Start Billable Logged Date", "Latest Billable Logged Date", "Associated Invoices", "Link to Project in Harvest", "Why has it gone Over?", "What is the action?", "Wastage Approved By", "Total Project Hours Logged", "Billable Project Hours Logged"];

    const inBudgetHeaders = ["Job number", "Client name", "Project name", "Status", "Budgeted by", "Budget total", "Budget spent", "Remaining Budget", "Budget Utilization", "Harvest Start Date", "Harvest End Date", "Start Billable Logged Date", "Latest Billable Logged Date", "Associated Invoices", "Link to Project in Harvest", "What is the action?", "Wastage Approved By", "Total Project Hours Logged", "Billable Project Hours Logged"];

    // Set headers based on report type
    const headers = reportType === "overBudget" ? overBudgetHeaders : inBudgetHeaders;

    sheet.appendRow(headers);

    // Populate data
    data.forEach((item) => {
      const rowData = new Array(headers.length).fill(""); // Initialize row data
      const budgetType = projectTypeRename(item.budget_by);
      rowData[0] = item.project_code || "";
      rowData[1] = item.client_name || "";
      rowData[2] = item.project_name || "";
      rowData[3] = item.is_active ? "Active" : "Archived";
      rowData[4] = budgetType || "";
      // Budget Total
      rowData[5] = budgetType === "£" ? item.budget || 0 : (item.budget * defaultHourlyRate).toFixed(2);
      // Budget Spent
      rowData[6] = budgetType === "£" ? item.budget_spent || 0 : (item.budget_spent * defaultHourlyRate).toFixed(2);

      if (reportType === "overBudget") {
        // Calculate and add to totalCostOverBudget (new)
        const overBudgetBy = budgetType === "£" ? Math.abs(item.budget_remaining) : Math.abs(item.budget_remaining) * defaultHourlyRate;
        totalCostOverBudget += overBudgetBy;

        rowData[7] = budgetType === "£" ? Math.abs(item.budget_remaining) : (Math.abs(item.budget_remaining) * defaultHourlyRate).toFixed(2); // Over budget by
        const overBudgetPercentage = item.budget === 0 ? "No budget amount set" : ((item.budget_spent - item.budget) / item.budget) * 100;

        rowData[8] = typeof overBudgetPercentage === "string" ? overBudgetPercentage : `${overBudgetPercentage.toFixed(2)}%`; // Over budget by %
      } else if (reportType !== "overBudget") {
        rowData[7] = budgetType === "£" ? item.budget_remaining || 0 : (Math.abs(item.remaining_budget) * defaultHourlyRate).toFixed(2); // Remaining Budget
        rowData[8] = item.budget_utilization || ""; // Budget Utilization
      } else {
        Logger.log("Invalid report type");
      }

      rowData[9] = item.starts_on || "";
      rowData[10] = item.ends_on || "";
      rowData[11] = item.first_time_entry || "";
      rowData[12] = item.last_time_entry || "";
      rowData[13] = Array.isArray(item.associated_invoices) ? item.associated_invoices.join(", ") : item.associated_invoices || "";
      rowData[14] = "https://marketingpod.harvestapp.com/projects/" + (item.id || "");
      rowData[15] = ""; // Blank cell - Why has it gone over?
      rowData[16] = ""; // What is the action?

      if (reportType === "overBudget") {
        rowData[17] = ""; // Wastage Approved By
        rowData[18] = item.total_logged_hours || "Not Available";
        rowData[19] = item.billable_logged_hours || "Not Available";
      } else {
        rowData[17] = item.total_logged_hours || "Not Available";
        rowData[18] = item.billable_logged_hours || "Not Available";
      }

      sheet.appendRow(rowData);
    });

    // Add data validation for the 'Action' column
    const actionColumnIndex = headers.indexOf("What is the action?") + 1; // Get the amount of columns and add 1
    const rangeAction = sheet.getRange(2, actionColumnIndex, sheet.getLastRow() - 1);
    const ruleAction = SpreadsheetApp.newDataValidation().requireValueInList(["", "Invoiced", "Bill Client", "Write-Off"], true).setAllowInvalid(false).build();
    rangeAction.setDataValidation(ruleAction);

    // Add data validation for the 'Approved By' column
    const approvedByColumnIndex = headers.indexOf("Wastage Approved By") + 1; // Get the amount of columns and add 1
    const rangeApprovedBy = sheet.getRange(2, approvedByColumnIndex, sheet.getLastRow() - 1);
    const ruleApprovedBy = SpreadsheetApp.newDataValidation().requireValueInList(["", "Jenny Hughes", "Jodie Williams", "Adam Leach", "Kate Garratt", "Emma Crofts"], true).setAllowInvalid(false).build();
    rangeApprovedBy.setDataValidation(ruleApprovedBy);

    // Set frozen first row and format as bold
    sheet.setFrozenRows(1);
    sheet.getRange("A1:1").setFontWeight("bold");

    // Set format of the % budget remaining
    sheet.getRange("D2:D").setNumberFormat("#0.##%");

    // Set currency format for 'Budget total', 'Budget spent' and 'Over budget by' columns
    sheet.getRange("F2:H" + sheet.getLastRow()).setNumberFormat("£#,##0.00");

    if (reportType === "overBudget") {
      // Sort the overBudget spreadsheet by 'Over budget By' column  (highest to lowest)
      sheet.getRange("A2:T" + sheet.getLastRow()).sort({ column: 8, ascending: false });
    } else if (reportType === "nearlyOverBudget") {
      // Sort the spreadsheet from low to high (most to least budget lost)
      sheet.getRange("A2:S").sort({ column: 1, ascending: true });
    }

    // Add totals row for 'Over budget' and 'Nearly over budget' reports
    if (reportType !== "currentFYClientTotal") {
      // Calculate the row for the totals
      const lastRow = sheet.getLastRow();
      const totalsRow = lastRow + 1;

      // Add totals for columns F, G, and H
      sheet.getRange(`F${totalsRow}`).setFormula(`=SUM(F2:F${lastRow})`);
      sheet.getRange(`G${totalsRow}`).setFormula(`=SUM(G2:G${lastRow})`);
      sheet.getRange(`H${totalsRow}`).setFormula(`=SUM(H2:H${lastRow})`);

      // Set the totals row to bold
      sheet.getRange(`F${totalsRow}:H${totalsRow}`).setFontWeight("bold");
    }

    // Set date columns to be right-aligned
    sheet.getRange("J2:L").setHorizontalAlignment("right");

    // Auto-resize columns for readability
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }

    Logger.log("Sheet populated successfully.");
  } catch (error) {
    Logger.log(`Error populating sheet: ${error.message}`);
  }
}

function populateClientTotalsSheet(clientTotalsSheet, currentFYClientTotals) {
  Logger.log("Populating 'Client Totals' spreadsheet");
  const headers = ["Month", "Client name", "Billable Worked FYTD £", "Non-billable Worked FYTD £", "Total Rework FYTD £", "AEDR (from hours / budget)"];

  // Clear the sheet first to avoid appending to old data
  clientTotalsSheet.clear();

  // Set the headers
  clientTotalsSheet.appendRow(headers);

  // Collect data rows in an array for sorting
  const dataRows = [];

  currentFYClientTotals.forEach((clientData) => {
    const clientName = Object.keys(clientData)[0];
    const monthlyTotals = clientData[clientName];

    for (const [month, totals] of Object.entries(monthlyTotals)) {
      const billableWorked = totals.billable_total ? (totals.billable_total * defaultHourlyRate).toFixed(2) : "0.00";
      const nonBillableWorked = totals.non_billable_total ? (totals.non_billable_total * defaultHourlyRate).toFixed(2) : "0.00";
      const totalRework = totals.rework_total ? (totals.rework_total * defaultHourlyRate).toFixed(2) : "0.00";
      const totalTimeBooked = totals.billable_total + totals.non_billable_total;
      const aedr = totalTimeBooked ? ((totals.billable_total / totalTimeBooked) * 7.5 * defaultHourlyRate).toFixed(2) : "0.00";

      const row = [month, clientName, billableWorked, nonBillableWorked, totalRework, aedr];
      dataRows.push(row);
    }
  });

  // Sort dataRows by Month (first column) and then by Client name (second column)
  const monthOrder = ["April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March"];

  dataRows.sort((a, b) => {
    // Extract month and year from the date strings
    const [aMonth, aYear] = a[0].split(" ");
    const [bMonth, bYear] = b[0].split(" ");

    // Get the month indices from the monthOrder array
    const aMonthIndex = monthOrder.indexOf(aMonth);
    const bMonthIndex = monthOrder.indexOf(bMonth);

    // Adjust year for April at the end of the financial year (Corrected logic)
    const aYearAdjusted = aMonth === "April" && aMonthIndex > monthOrder.indexOf("May") ? parseInt(aYear) + 1 : parseInt(aYear);
    const bYearAdjusted = bMonth === "April" && bMonthIndex > monthOrder.indexOf("May") ? parseInt(bYear) + 1 : parseInt(bYear);

    // Compare adjusted years first
    if (aYearAdjusted !== bYearAdjusted) {
      return aYearAdjusted - bYearAdjusted;
    }

    // If adjusted years are the same, compare months based on the monthOrder
    if (aMonthIndex !== bMonthIndex) {
      return aMonthIndex - bMonthIndex;
    }

    // If months are equal, sort by Client name
    return a[1].localeCompare(b[1]);
  });

  // Append sorted rows to the sheet
  dataRows.forEach((row) => clientTotalsSheet.appendRow(row));

  // Set frozen first row and format as bold
  clientTotalsSheet.setFrozenRows(1);
  clientTotalsSheet.getRange("A1:F1").setFontWeight("bold");

  // Set currency format for columns C, D, E, and F
  clientTotalsSheet.getRange("C2:F" + clientTotalsSheet.getLastRow()).setNumberFormat("£#,##0.00");

  // Auto-resize columns for readability
  for (let i = 1; i <= headers.length; i++) {
    clientTotalsSheet.autoResizeColumn(i);
  }

  // Crop the spreadsheet
  cropSheet(clientTotalsSheet, headers);
}

function populateReworkByTeamSheet(reworkByTeamSheet, currentFYClientTotals, roles) {
  Logger.log("Populating 'Rework by Team' spreadsheet");
  Logger.log("Nearly there now!");
  const headers = [
    "Team", // This is the role
    "Client Total Rework FYTD £",
    "Average Rework per Person FYTD £",
  ];

  // Clear the sheet first to avoid appending to old data
  reworkByTeamSheet.clear();

  // Set the headers
  reworkByTeamSheet.appendRow(headers);

  // Collect data rows in an array for sorting
  const dataRows = [];

  // Iterate through each role in the roles object
  roles.forEach((role) => {
    const teamName = role.name;
    const totalReworkCost = totalReworkTimeByRole[teamName] ? totalReworkTimeByRole[teamName].totalReworkCost : 0;

    // Calculate the average rework cost per person
    const teamSize = role.user_ids.length;
    const averageReworkCost = totalReworkCost / teamSize;

    // Add the data row to the array
    dataRows.push([teamName, totalReworkCost, averageReworkCost]);
  });

  // Sort dataRows by Team name (first column)
  dataRows.sort((a, b) => a[0].localeCompare(b[0]));

  // Append sorted rows to the sheet
  dataRows.forEach((row) => reworkByTeamSheet.appendRow(row));

  // Set frozen first row and format as bold
  reworkByTeamSheet.setFrozenRows(1);
  reworkByTeamSheet.getRange("A1:C1").setFontWeight("bold");

  // Set currency format for columns B and C
  reworkByTeamSheet.getRange("B2:C" + reworkByTeamSheet.getLastRow()).setNumberFormat("£#,##0.00");

  // Auto-resize columns for readability
  reworkByTeamSheet.autoResizeColumns(1, headers.length);

  // Crop the spreadsheet
  cropSheet(reworkByTeamSheet, headers);
}

function populateSummarySheet(summarySheet, overBudgetData) {
  Logger.log("Populating 'Summary' spreadsheet");

  // Clear the sheet first to avoid appending to old data
  summarySheet.clear();

  const introduction = "This report is filtered to projects that are either a) active b) start or end within FY c) activity within FY (excluding Pod or Pod Marketing)";
  summarySheet.getRange("A1").setValue(introduction);

  // Get today's date in DD/MM/YYYY format
  const today = new Date();
  const day = String(today.getDate()).padStart(2, "0");
  const month = String(today.getMonth() + 1).padStart(2, "0");
  const year = today.getFullYear();
  const dateString = `${day}/${month}/${year}`;

  // Add the date created row
  summarySheet.getRange("A3").setValue("Date Created").setFontWeight("bold");
  summarySheet.getRange("B3").setValue(dateString);

  // Calculate totalReworkCost
  let totalReworkCost = 0;
  for (const role in totalReworkTimeByRole) {
    totalReworkCost += totalReworkTimeByRole[role].totalReworkCost;
  }

  // Calculate number of projects over budget using overBudgetData
  const activeProjectsOverBudgetCount = overBudgetData.filter((item) => item.is_active).length;

  // Calculate active project wastage over budget
  const activeProjectWastageOverBudget = overBudgetData
    .filter((item) => item.is_active)
    .reduce((total, item) => {
      const budgetType = projectTypeRename(item.budget_by);
      const overBudgetBy = budgetType === "£" ? Math.abs(item.budget_remaining) : Math.abs(item.budget_remaining) * defaultHourlyRate;
      return total + overBudgetBy;
    }, 0);

  // Calculate number of active projects with rework logged
  const activeProjectsWithReworkCount = overBudgetData.filter((item) => item.is_active && item.rework_role_breakdown && item.rework_role_breakdown.length > 0).length;

  const activeProjectsWastageRework = overBudgetData
    .filter((item) => item.is_active && item.rework_role_breakdown && item.rework_role_breakdown.length > 0)
    .reduce((total, item) => {
      return total + item.rework_role_breakdown.reduce((sum, role) => sum + role.totalReworkCost, 0);
    }, 0);

  // Calculate total overall wastage FYTD
  const totalOverallWastageFYTD = totalCostOverBudget + totalReworkCost;

  // Add the data rows
  summarySheet.getRange("A4").setValue("Total Wastage Over Budget £").setFontWeight("bold");
  summarySheet.getRange("B4").setValue(totalCostOverBudget).setNumberFormat("£#,##0.00");

  summarySheet.getRange("A5").setValue("Active Project Over Budget # Projects").setFontWeight("bold");
  summarySheet.getRange("B5").setValue(activeProjectsOverBudgetCount);

  summarySheet.getRange("A6").setValue("Active Project Wastage Over Budget £").setFontWeight("bold");
  summarySheet.getRange("B6").setValue(activeProjectWastageOverBudget).setNumberFormat("£#,##0.00");

  summarySheet.getRange("A7").setValue("Total Wastage Rework Cost £").setFontWeight("bold");
  summarySheet.getRange("B7").setValue(totalReworkCost).setNumberFormat("£#,##0.00");

  summarySheet.getRange("A8").setValue("Active Project Wastage Rework # Projects").setFontWeight("bold");
  summarySheet.getRange("B8").setValue(activeProjectsWithReworkCount);

  summarySheet.getRange("A9").setValue("Active Project Wastage Rework £").setFontWeight("bold");
  summarySheet.getRange("B9").setValue(activeProjectsWastageRework).setNumberFormat("£#,##0.00");

  summarySheet.getRange("A10").setValue("Total Overall Wastage").setFontWeight("bold");
  summarySheet.getRange("B10").setValue(totalOverallWastageFYTD).setNumberFormat("£#,##0.00");

  summarySheet.getRange("A11").setValue("Total Billable Hours").setFontWeight("bold");
  summarySheet.getRange("B11").setValue(totalBillableHours);

  summarySheet.getRange("A12").setValue("Total Hours").setFontWeight("bold");
  summarySheet.getRange("B12").setValue(totalHours);

  const aedrFromBudget = calculateAEDRFromBudget(overBudgetData);
  summarySheet.getRange("A13").setValue("AEDR (from budget)").setFontWeight("bold");
  summarySheet.getRange("B13").setValue(aedrFromBudget).setNumberFormat("£#,##0.00");

  // Auto-resize columns for readability
  summarySheet.autoResizeColumns(1, 2);

  // Crop the spreadsheet
  cropSheet(summarySheet);
}

function populateMissingBudgetSheet(missingBudgetSheet, missingBudgetReport) {
  Logger.log("Populating 'missing Budget' spreadsheet");

  // Clear the sheet first to avoid appending to old data
  missingBudgetSheet.clear();

  const missingBudgetHeaders = ["Job number", "Client name", "Project name", "Status", "Budgeted by", "Budget total", "Harvest Start Date", "Harvest End Date", "Link to Project in Harvest", "Total Project Hours Logged"];

  missingBudgetSheet.appendRow(missingBudgetHeaders);

  missingBudgetReport.forEach((item) => {
    const rowData = new Array(headers.length).fill(""); // Initialize row data
    const budgetType = projectTypeRename(item.budget_by);
    rowData[0] = item.code || "";
    rowData[1] = item.client?.name || "";
    rowData[2] = item.name || "";
    rowData[3] = item.is_active ? "Active" : "Archived";
    rowData[4] = item.budget_by || "";
    // Budget Total
    rowData[5] = budgetType === "£" ? item.budget || 0 : (item.budget * defaultHourlyRate).toFixed(2);
    rowData[6] = item.starts_on || "";
    rowData[7] = item.ends_on || "";
    rowData[8] = "https://marketingpod.harvestapp.com/projects/" + (item.id || "");
    rowData[9] = item.total_logged_hours || "Not Available";

    missingBudgetSheet.appendRow(rowData);
  });
}

// Calculate AEDR (from budget)
function calculateAEDRFromBudget(overBudgetData) {
  let totalValueOfBilledTime = 0;
  let totalAmountOverBudget = 0;
  let totalTimeBooked = 0;

  overBudgetData.forEach((item) => {
    const budgetType = projectTypeRename(item.budget_by);
    const overBudgetBy = budgetType === "£" ? Math.abs(item.budget_remaining) : Math.abs(item.budget_remaining) * defaultHourlyRate;

    totalValueOfBilledTime += item.budget_spent;
    totalAmountOverBudget += overBudgetBy;
    totalTimeBooked += item.total_logged_hours;
  });

  const aedr = ((totalValueOfBilledTime - totalAmountOverBudget) / totalTimeBooked) * 7.5;
  return aedr;
}

// Crop the spreadsheet from: https://developers.google.com/apps-script/add-ons/clean-sheet
function cropSheet(sheetToCrop, headers) {
  const dataRange = sheetToCrop.getDataRange();
  sheetToCrop = dataRange.getSheet();

  let numRows = dataRange.getNumRows();
  let numColumns = dataRange.getNumColumns();

  const maxRows = sheetToCrop.getMaxRows();
  const maxColumns = sheetToCrop.getMaxColumns();

  const numFrozenRows = sheetToCrop.getFrozenRows();
  const numFrozenColumns = sheetToCrop.getFrozenColumns();

  // If last data row is less than maximum row, then delete rows after the last data row.
  if (numRows < maxRows) {
    numRows = Math.max(numRows, numFrozenRows + 1); // Don't crop empty frozen rows.
    sheetToCrop.deleteRows(numRows + 1, maxRows - numRows);
  }

  // If last data column is less than maximum column, then delete columns after the last data column.
  if (numColumns < maxColumns) {
    numColumns = Math.max(numColumns, numFrozenColumns + 1); // Don't crop empty frozen columns.
    sheetToCrop.deleteColumns(numColumns + 1, maxColumns - numColumns);
  }

  if (headers) {
    // Auto-resize columns for readability
    for (let i = 1; i <= headers.length; i++) {
      sheetToCrop.autoResizeColumn(i);
    }
  }
}

function projectTypeRename(projectBudgetBy) {
  let projectBudgetByRenamed;
  switch (projectBudgetBy) {
    case "project":
      projectBudgetByRenamed = "hours";
      break;
    case "project_cost":
      projectBudgetByRenamed = "£";
      break;
    case "task_fees":
      projectBudgetByRenamed = "£";
      break;
    case "none":
      projectBudgetByRenamed = "none";
      break;
    default:
      console.log(`Unknown project budget type: ${projectBudgetBy}`);
      projectBudgetByRenamed = "Error";
  }
  return projectBudgetByRenamed;
}

function fetchInvoicesByProjectId(projectId) {
  try {
    const url = invoiceApiUrl + projectId;
    const options = {
      method: "get",
      headers: headers,
      muteHttpExceptions: true,
    };

    // Use UrlFetchApp.fetch for synchronous behavior
    const response = UrlFetchApp.fetch(url, options);

    // Parse the response as JSON
    const jsonResponse = JSON.parse(response.getContentText());

    const projectInvoices = jsonResponse.invoices.map((invoice) => invoice.number);

    return projectInvoices;
  } catch (error) {
    Logger.log(`Error fetching invoices for project ${projectId}: ${error.message}`);
    // Instead of returning null, return an empty array so the rest of the code can continue.
    return [];
  }
}

function sendEmail(combinedSheetUrl) {
  const subject = "Automated Harvest Budget Report " + formattedDate;
  const body = "Hi Everyone,\n\nPlease find the attached budget report in the Google Sheet: " + combinedSheetUrl + "\n\nAs this is an automated report, please book in any changes.";
  MailApp.sendEmail({
    to: recipientEmails.join(","),
    subject: subject,
    body: body,
  });
}

// createBudgetReport();
