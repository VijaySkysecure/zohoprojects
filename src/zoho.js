const axios = require("axios");
const config = require("./config");
const {
  getUserToken: mongoGetUserToken,
  storeUserToken: mongoStoreUserToken,
  updateUserToken: mongoUpdateUserToken
} = require("./mongodb");
const moment = require("moment");

const { zohoApiBaseUrl, zohoPortalId } = config;

let lastApiCall = 0;
const API_CALL_DELAY = 1000; // 1 second between calls
const MAX_REQUESTS_PER_WINDOW = 90; // Stay under 100 limit
const WINDOW_DURATION = 120000; // 2 minutes in milliseconds
let requestCount = 0;
let windowStart = Date.now();

// -------------------------
// AUTH HELPERS
// -------------------------

async function refreshAccessToken(teamsChatId, refreshToken = null) {
  try {
    if (!refreshToken) {
      const existingToken = await mongoGetUserToken(teamsChatId);
      if (!existingToken || !existingToken.refreshToken) {
        throw new Error("No refresh token found in database");
      }
      refreshToken = existingToken.refreshToken;
    }

    const body = new URLSearchParams({
      grant_type: "refresh_token",
      client_id: config.zohoClientId,
      client_secret: config.zohoClientSecret,
      refresh_token: refreshToken,
    });

    // Send POST with body
    response = await axios.post(
      "https://accounts.zoho.in/oauth/v2/token",
      body.toString(),
      { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
    );


    const { access_token, expires_in } = response.data;

    console.log("response.data", response.data)
    if (!access_token || !expires_in) {
      throw new Error("Invalid refresh token response");
    }

    const updatedToken = await mongoUpdateUserToken(teamsChatId, {
      accessToken: access_token,
      refreshToken,
      expiresAt: Date.now() + expires_in * 1000,
    });

    return updatedToken;
  } catch (error) {
    console.error(`[TOKEN REFRESH] Error refreshing token: ${error}`);
    throw error;
  }
}

async function getUserToken(teamsChatId) {
  try {
    const token = await mongoGetUserToken(teamsChatId);
    if (!token) {
      throw new Error("Token not found");
    }

    if (token.expiresAt < Date.now()) {
      console.log(
        `Access token expired. Refreshing for teamsChatId: ${teamsChatId}`
      );
      return await refreshAccessToken(teamsChatId, token.refreshToken);
    }

    return token;
  } catch (error) {
    console.error(`Error fetching token: ${error.message}`);
    throw error;
  }
}

// -------------------------
// GENERIC API CALL
// -------------------------

async function makeZohoAPICall(
  endpoint,
  token,
  method = "GET",
  data = null,
  params = {},
  teamsChatId = null,
  portalId = null
) {
  try {
    // Rate limiting logic
    const now = Date.now();
    
    // Reset window if 2 minutes have passed
    if (now - windowStart > WINDOW_DURATION) {
      requestCount = 0;
      windowStart = now;
    }
    
    // Check if we're approaching the limit
    if (requestCount >= MAX_REQUESTS_PER_WINDOW) {
      const waitTime = WINDOW_DURATION - (now - windowStart);
      console.log(`Rate limit approaching. Waiting ${Math.ceil(waitTime / 1000)} seconds...`);
      await new Promise(resolve => setTimeout(resolve, waitTime));
      requestCount = 0;
      windowStart = Date.now();
    }
    
    // Add delay between calls
    const timeSinceLastCall = now - lastApiCall;
    if (timeSinceLastCall < API_CALL_DELAY) {
      const delay = API_CALL_DELAY - timeSinceLastCall;
      await new Promise(resolve => setTimeout(resolve, delay));
    }
    
    lastApiCall = Date.now();
    requestCount++;
    
    console.log(`API Call #${requestCount} in current window`);

    const url = `${zohoApiBaseUrl.endsWith("/") ? zohoApiBaseUrl : zohoApiBaseUrl + "/"
      }${endpoint}`;

    const headers = {
      Authorization: `Zoho-oauthtoken ${token}`,
      "Content-Type": "application/json",
    };

    if (portalId) {
      headers["X-ZOHO-PROJECTS"] = portalId;
    }

    const configs = {
      method,
      url,
      headers,
      params,
    };

    if (["POST", "PUT", "PATCH"].includes(method.toUpperCase())) {
      configs.data = data;
    }

    let retries = 3;
    while (retries > 0) {
      try {
        const response = await axios(configs);
        return response;
      } catch (error) {
        // Handle rate limiting specifically
        if (error.response?.status === 429) {
          const retryAfter = error.response.headers['retry-after'];
          const waitTime = retryAfter ? parseInt(retryAfter) * 1000 : Math.pow(2, 3 - retries) * 1000;
          console.log(`Rate limited. Waiting ${waitTime / 1000} seconds...`);
          await new Promise((resolve) => setTimeout(resolve, waitTime));
          retries--;
          continue;
        }
        
        // Handle 400 errors that might be rate limiting
        if (error.response?.status === 400) {
          const errorData = error.response.data;
          if (errorData?.error?.error_type === 'OPERATIONAL_VALIDATION_ERROR' && 
              errorData?.error?.title === 'URL_ROLLING_THROTTLES_LIMIT_EXCEEDED') {
            console.log('Rate limit exceeded. Waiting 2 minutes...');
            await new Promise((resolve) => setTimeout(resolve, 120000)); // Wait 2 minutes
            retries--;
            continue;
          }
        }
        
        if (error.response?.status === 401 && teamsChatId) {
          try {
            const refreshedToken = await refreshAccessToken(teamsChatId, null);
            if (refreshedToken && refreshedToken.accessToken) {
              configs.headers.Authorization = `Zoho-oauthtoken ${refreshedToken.accessToken}`;
              const response = await axios(configs);
              return response;
            }
          } catch (refreshError) {
            console.error(`[ZOHO API] Token refresh failed:`, refreshError.message);
            throw error;
          }
        }
        
        throw error;
      }
    }
    throw new Error("Max retries reached for Zoho API");
  } catch (error) {
    console.error("Error in makeZohoAPICall:", error);
    throw error;
  }
}

// -------------------------
// OWNER RESOLUTION
// -------------------------

async function resolveOwnerId(teamsChatId, portalId, ownerName) {
  try {
    const token = await getUserToken(teamsChatId); // fetch token for the actual Teams user
    if (!token) throw new Error("Token not found");

    console.log("ownerNameownerName:", ownerName)
    const response = await makeZohoAPICall(
      `portal/${portalId}/users`,
      token.accessToken,
      "GET",
      null,
      {},
      teamsChatId,
      portalId
    );

    // console.log("responserespos:", response.data)
    const users = response?.data?.users || [];


    if (!users.length) return null;

    let owner = users.find(
      u => u.full_name.toLowerCase() === ownerName.toLowerCase()
    );
    if (!owner) {
      owner = users.find(u =>
        u.full_name.toLowerCase().includes(ownerName.toLowerCase())
      );
    }
    if (!owner) return null;

    return { id: owner.zpuid || owner.id, name: owner.full_name };
  } catch (err) {
    console.error("[Zoho] Error in resolveOwnerId:", err);
    return null;
  }
}




// -------------------------
// ZOHO PROJECTS FUNCTIONS
// -------------------------

async function getProjects(token) {
  try {
    const response = await makeZohoAPICall(
      `portal/${zohoPortalId}/projects`,
      token,
      "GET",
      null,
      { per_page: 200 }
    );
    return response.data?.projects || [];
  } catch (error) {
    console.error("Error fetching Zoho Projects:", error.message);
    throw error;
  }
}



async function getPendingTasksByOwner(context, state, ownerName) {
  const teamsChatId = "11111";
  const portalId = config.zohoPortalId;

  const resolvedOwner = await resolveOwnerId(teamsChatId, portalId, ownerName);
  if (!resolvedOwner) return [];

  const token = await getUserToken(teamsChatId);
  
  console.log(`Searching for pending tasks for owner: ${ownerName}`);

  try {
    // Function to get the last 5 pages of tasks only - FIXED VERSION
    async function getLatestTasks() {
      console.log(`\n=== FETCHING LATEST TASKS (LAST 5 PAGES) ===`);
      
      let allTasks = [];
      const perPage = 100;
      
      // Start from a high page number and work backwards to find the last 5 pages
      let currentPage = 90; // Start high
      let foundPages = 0;
      let lastValidPage = 0;
      
      // First, find the last page by starting high and working down
      console.log(`Finding the last page...`);
      while (currentPage > 0 && foundPages < 7) {
        try {
          const response = await makeZohoAPICall(
            `portal/${portalId}/tasks`,
            token.accessToken,
            "GET",
            null,
            { 
              per_page: perPage,
              page: currentPage
            },
            teamsChatId,
            portalId
          );

          const tasks = response?.data?.tasks || [];
          if (tasks.length > 0) {
            lastValidPage = currentPage;
            foundPages++;
            console.log(`  Page ${currentPage}: Found ${tasks.length} tasks`);
            allTasks = allTasks.concat(tasks);
          }
          currentPage--;
          
          // Small delay to avoid rate limiting
          await new Promise(resolve => setTimeout(resolve, 50));
        } catch (error) {
          console.log(`  Page ${currentPage} failed: ${error.message}`);
          currentPage--;
        }
      }

      console.log(`Found last valid page: ${lastValidPage}`);
      console.log(`Total latest tasks fetched: ${allTasks.length}`);
      return allTasks;
    }

    // Get latest tasks only
    const allTasks = await getLatestTasks();

    // Filter for pending tasks assigned to the owner
    const pendingTasks = allTasks.filter(task => {
      if (task.is_completed === true) return false;
      if (!task.owners_and_work?.owners || task.owners_and_work.owners.length === 0) return false;

      return task.owners_and_work.owners.some(owner => {
        const ownerFullName = `${owner.first_name} ${owner.last_name}`.toLowerCase().trim();
        const ownerFirstName = owner.first_name.toLowerCase().trim();
        const ownerLastName = owner.last_name.toLowerCase().trim();
        const searchName = ownerName.toLowerCase().trim();
        const resolvedOwnerName = resolvedOwner.name.toLowerCase().trim();
        
        return ownerFullName.includes(searchName) || 
               ownerFirstName.includes(searchName) || 
               ownerLastName.includes(searchName) ||
               owner.name.toLowerCase().includes(searchName) ||
               ownerFullName.includes(resolvedOwnerName) ||
               ownerFirstName.includes(resolvedOwnerName) ||
               ownerLastName.includes(resolvedOwnerName) ||
               owner.name.toLowerCase().includes(resolvedOwnerName);
      });
    });

    // Limit to latest 15 tasks only
    const limitedPendingTasks = pendingTasks.slice(0, 15);
    
    console.log(`Found ${pendingTasks.length} total pending tasks for ${ownerName}`);
    console.log(`Returning latest ${limitedPendingTasks.length} pending tasks`);

    // Format tasks for response with proper formatting
    const formattedTasks = limitedPendingTasks.map(task => ({
      id: task.id,
      name: task.name,
      description: task.description,
      status: task.status?.name || 'Unknown',
      priority: task.priority || 'none',
      project: task.project?.name || 'Unknown Project',
      startDate: task.start_date,
      endDate: task.end_date,
      dueDate: task.end_date || 'N/A',
      owners: task.owners_and_work?.owners?.map(owner => ({
        name: owner.name,
        email: owner.email,
        firstName: owner.first_name,
        lastName: owner.last_name
      })) || [],
      completionPercentage: task.completion_percentage || 0,
      isCompleted: task.is_completed || false
    }));

    return formattedTasks;

  } catch (error) {
    console.error(`Error fetching pending tasks for ${ownerName}:`, error.message);
    return [];
  }
}




async function getProjectByName(teamsChatId, portalId, projectName) {
  const token = await getUserToken(teamsChatId);
  if (!token) throw new Error("No token found for user");

  console.log(`\n=== PROJECT SEARCH DIAGNOSTIC ===`);
  console.log(`Searching for project: "${projectName}"`);

  try {
    // Function to get ALL projects with pagination
    async function getAllProjects() {
      console.log(`\n=== FETCHING ALL PROJECTS WITH PAGINATION ===`);
      
      let allProjects = [];
      let page = 1;
      const perPage = 100;
      let hasMore = true;

      while (hasMore) {
        console.log(`Fetching projects page ${page}...`);
        
        try {
          const response = await makeZohoAPICall(
            `portal/${portalId}/projects`,
            token.accessToken,
            "GET",
            null,
            { 
              per_page: perPage,
              page: page
            },
            teamsChatId,
            portalId
          );

          // FIX: Projects are in the root array, not response.data.projects
          const projects = response?.data || [];
          console.log(`  Page ${page}: Found ${projects.length} projects`);
          
          if (projects.length === 0) {
            hasMore = false;
          } else {
            allProjects = allProjects.concat(projects);
            page++;
            
            // Small delay to avoid rate limiting
            await new Promise(resolve => setTimeout(resolve, 100));
          }
        } catch (error) {
          console.log(`  Projects page ${page} failed: ${error.message}`);
          hasMore = false;
        }
      }

      console.log(`Total projects fetched: ${allProjects.length}`);
      
      // Show all project names for debugging
      console.log(`\nAll projects found:`);
      allProjects.forEach((project, index) => {
        console.log(`  ${index + 1}. "${project.name}" (ID: ${project.id})`);
      });
      
      return allProjects;
    }

    // Get all projects
    const allProjects = await getAllProjects();

    // Search for matching projects
    const matched = allProjects.filter((p) =>
      p.name.toLowerCase().includes(projectName.toLowerCase())
    );

    console.log(`\nMatching projects for "${projectName}": ${matched.length}`);
    matched.forEach((project, index) => {
      console.log(`  ${index + 1}. "${project.name}" (ID: ${project.id})`);
    });

    if (matched.length === 0) {
      console.log(`No projects found matching "${projectName}"`);
      return { notFound: true };
    }
    if (matched.length > 1) {
      console.log(`Multiple projects found: ${matched.map(p => p.name).join(', ')}`);
      return { multiple: matched.map((p) => p.name) };
    }

    const project = matched[0];
    console.log(`Found project: "${project.name}" (ID: ${project.id})`);

    // Extract project details from the response structure
    let percent = project.percent_complete || "-";
    let openTasks = project.tasks?.open_count || "-";
    let closedTasks = project.tasks?.closed_count || "-";
    let tag = "-"; // Not available in this response
    let issuesList = [];
    let ownerName = project.owner?.full_name || project.owner?.name || "-";
    let statusName = project.status?.name || project.status || "-";
    let startDate = project.start_date
      ? moment(project.start_date).format("DD MMM YYYY")
      : "-";
    let endDate = project.end_date
      ? moment(project.end_date).format("DD MMM YYYY")
      : "-";

    // Try to get additional project details
    try {
      const projectDetail = await makeZohoAPICall(
        `portal/${portalId}/projects/${project.id}`,
        token.accessToken,
        "GET",
        null,
        {},
        teamsChatId,
        portalId
      );

      if (projectDetail?.data) {
        const detail = projectDetail.data;
        // Update with more detailed info if available
        percent = detail.percent_complete || percent;
        openTasks = detail.tasks?.open_count || openTasks;
        closedTasks = detail.tasks?.closed_count || closedTasks;
        tag = detail.tag || tag;
      }
    } catch (error) {
      console.log(`Error fetching additional project details: ${error.message}`);
    }

// Try to get project issues
// Try to get project issues
try {
  console.log(`\n=== FETCHING PROJECT ISSUES ===`);
  
  // Try different possible endpoints
  const possibleEndpoints = [
    `portal/${portalId}/projects/${project.id}/issues`,
    `portal/${portalId}/projects/${project.id}/bugs`,
    `portal/${portalId}/issues`,
    `portal/${portalId}/bugs`
  ];
  
  let issuesFound = false;
  
  for (const endpoint of possibleEndpoints) {
    try {
      console.log(`Trying endpoint: ${endpoint}`);
      const issuesResponse = await makeZohoAPICall(
        endpoint,
        token.accessToken,
        "GET",
        null,
        { per_page: 10 },
        teamsChatId,
        portalId
      );
      
      if (issuesResponse?.data) {
        const issues = issuesResponse.data.issues || issuesResponse.data.bugs || issuesResponse.data || [];
        if (Array.isArray(issues) && issues.length > 0) {
          issuesList = issues.slice(0, 5).map((issue) => ({
            title: issue.title || issue.name || issue.subject || "Untitled Issue",
            status: issue.status?.name || issue.status || "Unknown",
            priority: issue.priority || issue.severity || "Unknown",
          }));
          console.log(`Found ${issuesList.length} issues using ${endpoint}`);
          issuesFound = true;
          break;
        }
      }
    } catch (error) {
      console.log(`Endpoint ${endpoint} failed: ${error.message}`);
    }
  }
  
  if (!issuesFound) {
    console.log(`No working issues endpoint found`);
    issuesList = [];
  }
  
} catch (error) {
  console.log(`Error fetching project issues: ${error.message}`);
  issuesList = [];
}

    console.log(`\nProject details retrieved successfully`);
    console.log(`Project: ${project.name}`);
    console.log(`Owner: ${ownerName}`);
    console.log(`Status: ${statusName}`);
    console.log(`Completion: ${percent}%`);
    console.log(`Open Tasks: ${openTasks}, Closed Tasks: ${closedTasks}`);

    return {
      id: project.id,
      id_string: project.key || project.id,
      name: project.name,
      description: project.description || "-",
      owner: ownerName,
      status: statusName,
      percent: percent,
      openTasks: openTasks,
      closedTasks: closedTasks,
      tag: tag,
      startDate: startDate,
      endDate: endDate,
      issues: issuesList,
    };
  } catch (error) {
    console.error(`Error in getProjectByName: ${error.message}`);
    throw error;
  }
}




// Get list of users for dropdown
async function getUsers(teamsChatId) {
  const token = await getUserToken(teamsChatId);
  const resp = await makeZohoAPICall(
    `portal/${zohoPortalId}/users`,
    token.accessToken,
    "GET",
    null,
    {},
    teamsChatId,
    zohoPortalId
  );
  const users = resp?.data?.users || [];
  return users.map(u => ({
    id: u.zpuid || u.id || u.id_string,
    name: u.full_name || u.name
  }));
}



async function getAllTimeLogs(teamsChatId, portalId) {
  try {
    const token = await getUserToken(teamsChatId);
    if (!token || !token.accessToken) {
      throw new Error("No valid token found");
    }

    console.log("\n=== FETCHING ALL TIME LOGS WITH PAGINATION (INDIA REGION) ===");
    
    let allTimeLogs = [];
    
    // Get current month's data
    const currentDate = new Date();
    const currentMonth = String(currentDate.getMonth() + 1).padStart(2, '0');
    const currentYear = currentDate.getFullYear();
    const formattedDate = `${currentMonth}-01-${currentYear}`;
    
    try {
      const params = {
        users_list: 'all',
        view_type: 'month',
        date: formattedDate,
        bill_status: 'All',
        component_type: 'task'
      };

      console.log(`Fetching time logs for India region with params:`, params);

      const resp = await makeZohoAPICall(
        `logs`, // This will use the India REST API endpoint
        token.accessToken,
        "GET",
        null,
        params,
        teamsChatId,
        portalId
      );
      
      console.log(`Response structure:`, Object.keys(resp?.data || {}));
      console.log(`Response data sample:`, JSON.stringify(resp?.data, null, 2).substring(0, 500));
      
      // Parse the response according to the documented structure
      const timelogsData = resp?.data?.timelogs || {};
      const dateEntries = timelogsData.date || [];
      
      console.log(`Found ${dateEntries.length} date entries`);
      
      // Extract logs from the nested structure
      dateEntries.forEach(dateEntry => {
        const tasklogs = dateEntry.tasklogs || [];
        tasklogs.forEach(log => {
          const processedLog = {
            date: dateEntry.date,
            hours: log.hours || 0,
            minutes: log.minutes || 0,
            userName: log.owner_name || "Unknown User",
            projectName: log.project?.name || "Unknown Project",
            taskName: log.task?.name || "Unknown Task",
            description: log.notes || "",
            billable: log.bill_status === "Billable"
          };
          allTimeLogs.push(processedLog);
        });
      });

      console.log(`Total time logs processed: ${allTimeLogs.length}`);
      
    } catch (error) {
      console.error(`Time logs fetch failed for India region: ${error.message}`);
      console.error(`Error response:`, error.response?.data);
    }

    return allTimeLogs;

  } catch (error) {
    console.error("[getAllTimeLogs] Error:", error);
    throw error;
  }
}



// Get timelogs for user between two dates
async function getTimeLogsForUser(teamsChatId, portalId, userId, fromDate, toDate) {
  try {
    const token = await getUserToken(teamsChatId);
    if (!token || !token.accessToken) {
      throw new Error("No valid token found");
    }

    console.log(`\n=== FETCHING TIME LOGS FOR USER ${userId} ===`);
    console.log(`Date range: ${fromDate} to ${toDate}`);

    // Try multiple API endpoints
    const endpoints = [
      {
        name: 'timelogs',
        endpoint: `portal/${portalId}/timelogs`,
        params: {
          users_list: userId,
          view_type: "custom_date",
          date: `${moment(fromDate).format("MM-DD-YYYY")} to ${moment(toDate).format("MM-DD-YYYY")}`,
          per_page: 200
        }
      },
      {
        name: 'timesheet',
        endpoint: `portal/${portalId}/timesheet`,
        params: {
          users_list: userId,
          view_type: "custom_date",
          date: `${moment(fromDate).format("MM-DD-YYYY")} to ${moment(toDate).format("MM-DD-YYYY")}`,
          per_page: 200
        }
      },
      {
        name: 'logs (REST API)',
        endpoint: 'logs',
        params: {
          users_list: userId,
          view_type: "custom_date",
          date: `${moment(fromDate).format("MM-DD-YYYY")} to ${moment(toDate).format("MM-DD-YYYY")}`,
          bill_status: 'All',
          component_type: 'task'
        }
      }
    ];

    for (const endpointConfig of endpoints) {
      try {
        console.log(`Trying ${endpointConfig.name} endpoint...`);
        
        const resp = await makeZohoAPICall(
          endpointConfig.endpoint,
          token.accessToken,
          "GET",
          null,
          endpointConfig.params,
          teamsChatId,
          portalId
        );

        console.log(`${endpointConfig.name} response:`, Object.keys(resp?.data || {}));

        // Try different response structures
        let entries = [];
        if (resp?.data?.timelogs) {
          entries = resp.data.timelogs;
        } else if (resp?.data?.logs) {
          entries = resp.data.logs;
        } else if (resp?.data?.timesheet) {
          entries = resp.data.timesheet;
        } else if (Array.isArray(resp?.data)) {
          entries = resp.data;
        } else if (resp?.data?.date) {
          // Handle nested structure from logs endpoint
          const dateEntries = resp.data.date || [];
          entries = [];
          dateEntries.forEach(dateEntry => {
            const tasklogs = dateEntry.tasklogs || [];
            tasklogs.forEach(log => {
              entries.push({
                work_date: dateEntry.date,
                hours: log.hours || 0,
                owner: { name: log.owner_name || "Unknown User" },
                project: { name: log.project?.name || "Unknown Project" }
              });
            });
          });
        }

        if (entries && entries.length > 0) {
          console.log(`✅ ${endpointConfig.name} found ${entries.length} entries`);
          return entries.map(e => ({
            date: e.work_date || e.date,
            hours: Number(e.hours || e.time_spent || 0),
            userName: e.owner?.name || "Unknown User",
            projectName: e.project?.name || "Unknown Project"
          }));
        }
      } catch (error) {
        console.log(`❌ ${endpointConfig.name} failed:`, error.message);
      }
    }

    console.log("All endpoints failed or returned no data");
    return [];

  } catch (error) {
    console.error("[getTimeLogsForUser] Error:", error);
    throw error;
  }
}



// -------------------------
// ISSUES FUNCTIONS
// -------------------------

async function getProjectIssues(teamsChatId, portalId, projectName) {
  try {
    const token = await getUserToken(teamsChatId);
    if (!token) throw new Error("No token found for user");

    console.log(`\n=== FETCHING ISSUES FOR PROJECT: ${projectName} ===`);

    // Function to get ALL issues with pagination
    async function getAllIssues() {
      console.log(`\n=== FETCHING ALL ISSUES WITH PAGINATION ===`);
      
      let allIssues = [];
      let page = 1;
      const perPage = 100;
      let hasMore = true;

      while (hasMore) {
        console.log(`Fetching issues page ${page}...`);
        
        try {
          const response = await makeZohoAPICall(
            `portal/${portalId}/issues`,
            token.accessToken,
            "GET",
            null,
            { 
              per_page: perPage,
              page: page
            },
            teamsChatId,
            portalId
          );

          const issues = response?.data?.issues || [];
          console.log(`  Page ${page}: Found ${issues.length} issues`);
          
          if (issues.length === 0) {
            hasMore = false;
          } else {
            allIssues = allIssues.concat(issues);
            page++;
            
            // Small delay to avoid rate limiting
            await new Promise(resolve => setTimeout(resolve, 100));
          }
        } catch (error) {
          console.log(`  Issues page ${page} failed: ${error.message}`);
          hasMore = false;
        }
      }

      console.log(`Total issues fetched: ${allIssues.length}`);
      return allIssues;
    }

    // Get all issues
    const allIssues = await getAllIssues();

    // Debug: Log all unique project names found in issues
    const uniqueProjectNames = [...new Set(allIssues.map(issue => 
      issue.project?.name || issue.project_name || "No Project"
    ))];
    console.log(`\n=== DEBUG: All unique project names in issues ===`);
    uniqueProjectNames.forEach((name, index) => {
      console.log(`  ${index + 1}. "${name}"`);
    });
    console.log(`Searching for: "${projectName}"`);

    // Filter issues by project name with improved matching
    const projectIssues = allIssues.filter(issue => {
      const issueProjectName = issue.project?.name || issue.project_name || "";
      
      // Normalize both strings: remove extra spaces, convert to lowercase
      const normalizedIssueProject = issueProjectName.toLowerCase().replace(/\s+/g, ' ').trim();
      const normalizedSearchProject = projectName.toLowerCase().replace(/\s+/g, ' ').trim();
      
      // Check for exact match or if the issue project contains the search term
      const matches = normalizedIssueProject.includes(normalizedSearchProject) || 
                     normalizedSearchProject.includes(normalizedIssueProject);
      
      if (matches) {
        console.log(`✅ Match found: "${issueProjectName}" matches "${projectName}"`);
      }
      return matches;
    });

    console.log(`Found ${projectIssues.length} issues for project: ${projectName}`);


    // Format issues for response
    const formattedIssues = projectIssues.map(issue => ({
      name: issue.title || issue.name || issue.subject || "Untitled Issue",
      project: issue.project?.name || issue.project_name || "Unknown Project",
      reporter: issue.created_by?.first_name || issue.created_by?.name || issue.reporter?.full_name || issue.reporter?.name || "Unknown",
      createdTime: issue.created_time ? moment(issue.created_time).format("DD MMM YYYY HH:mm") : "N/A",
      assignee: issue.assignee?.first_name || issue.assignee?.full_name || issue.assignee?.name || "Unassigned",
      lastClosed: issue.last_closed ? moment(issue.last_closed).format("DD MMM YYYY HH:mm") : "N/A",
      lastModified: issue.last_updated_time ? moment(issue.last_updated_time).format("DD MMM YYYY HH:mm") : "N/A",
      dueDate: issue.due_date ? moment(issue.due_date).format("DD MMM YYYY HH:mm") : "N/A",
      status: issue.status?.name || issue.status || "Unknown",
      severity: issue.severity?.value || issue.severity?.name || issue.severity?.type || issue.severity || "None",
      description: issue.description || "No description available"
    }));

    return formattedIssues;

  } catch (error) {
    console.error(`Error in getProjectIssues: ${error.message}`);
    throw error;
  }
}

// -------------------------
// EXPORTS
// -------------------------

module.exports = {
  getUserToken,
  storeUserToken: mongoStoreUserToken,
  getPendingTasksByOwner,
  getProjects,
  getProjectByName,
  resolveOwnerId,
  getUsers,
  getTimeLogsForUser,
  getAllTimeLogs,
  getProjectIssues
};
