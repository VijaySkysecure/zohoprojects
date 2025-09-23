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
      let currentPage = 100; // Start high
      let foundPages = 0;
      let lastValidPage = 0;
      
      // First, find the last page by starting high and working down
      console.log(`Finding the last page...`);
      while (currentPage > 0 && foundPages < 5) {
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


// Function to get ALL time logs with pagination
// Function to get ALL time logs with pagination
async function getAllTimeLogs(teamsChatId, portalId) {
  try {
    const token = await getUserToken(teamsChatId);
    if (!token || !token.accessToken) {
      throw new Error("No valid token found");
    }

    console.log("\n=== FETCHING ALL TIME LOGS WITH PAGINATION ===");
    
    let allTimeLogs = [];
    let page = 1;
    const perPage = 100;
    let hasMore = true;

    while (hasMore) {
      console.log(`Fetching time logs page ${page}...`);
      
      try {
        const params = {
          per_page: perPage,
          page: page
        };

        let resp = null;
        let workingEndpoint = null;
        
        // Approach 1: Try timesheet API with module in request body
        try {
          console.log(`  Trying timesheet API...`);
          resp = await makeZohoAPICall(
            `portal/${portalId}/timesheet`,
            token.accessToken,
            "GET",
            null,
            { per_page: params.per_page, page: params.page },
            teamsChatId,
            portalId
          );
          workingEndpoint = 'timesheet';
          console.log(`  Using timesheet endpoint successfully`);
        } catch (error) {
          console.log(`  Timesheet endpoint failed: ${error.message}`);
          
          // Approach 2: Try with different parameters
          try {
            console.log(`  Trying timesheet with different params...`);
            resp = await makeZohoAPICall(
              `portal/${portalId}/timesheet`,
              token.accessToken,
              "GET",
              null,
              { per_page: 50, page: 1, sort: 'created_time' },
              teamsChatId,
              portalId
            );
            workingEndpoint = 'timesheet-alt';
            console.log(`  Using timesheet with alt params`);
          } catch (error2) {
            console.log(`  Timesheet alt params failed: ${error2.message}`);
            
            // Approach 3: Try tasks API without timelogs include
            try {
              console.log(`  Trying tasks API...`);
              resp = await makeZohoAPICall(
                `portal/${portalId}/tasks`,
                token.accessToken,
                "GET",
                null,
                { per_page: params.per_page, page: params.page },
                teamsChatId,
                portalId
              );
              workingEndpoint = 'tasks';
              console.log(`  Using tasks API`);
            } catch (error3) {
              console.log(`  Tasks API failed: ${error3.message}`);
              throw new Error(`All time log endpoints failed`);
            }
          }
        }

        // Try different possible data structures
        const logs = resp?.data?.timesheet || 
                    resp?.data?.timelogs || 
                    resp?.data?.logs || 
                    resp?.data?.time_logs || 
                    resp?.data || [];
        
        console.log(`  Page ${page}: Found ${logs.length} time logs`);
        console.log(`  Response structure:`, Object.keys(resp?.data || {}));
        
        if (logs.length === 0) {
          hasMore = false;
        } else {
          // Process and add logs
          const processedLogs = logs.map(log => {
            // Handle different possible field names
            const workDate = log.work_date || log.date || log.log_date || log.created_time;
            const hours = Number(log.hours || log.time_spent || log.duration || log.work_hours || 0);
            const userName = log.user?.name || log.user_name || log.user?.full_name || log.owner?.name || "Unknown User";
            const projectName = log.project?.name || log.project_name || log.project?.title || "Unknown Project";
            const taskName = log.task?.name || log.task_name || log.task?.title || "Unknown Task";
            const description = log.description || log.notes || log.comments || "";
            const billable = log.bill_status === "billable" || log.is_billable === true;
            
            return {
              date: workDate,
              hours: hours,
              userName: userName,
              projectName: projectName,
              taskName: taskName,
              description: description,
              billable: billable
            };
          });
          
          allTimeLogs = allTimeLogs.concat(processedLogs);
          page++;
          
          // Small delay to avoid rate limiting
          await new Promise(resolve => setTimeout(resolve, 200));
        }
      } catch (error) {
        console.log(`  Time logs page ${page} failed: ${error.message}`);
        hasMore = false;
      }
    }

    console.log(`Total time logs fetched: ${allTimeLogs.length}`);
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

    const params = {
      users_list: userId,
      view_type: "custom_date",
      custom_date: `{start_date:${moment(fromDate).format("DD-MM-YYYY")}, end_date:${moment(toDate).format("DD-MM-YYYY")}}`,
      bill_status: "all",
      per_page: 200
    };

    console.log("Fetching time logs with params:", params);

    const resp = await makeZohoAPICall(
      `portal/${portalId}/logs`,
      token.accessToken,
      "GET",
      null,
      params,
      teamsChatId,
      portalId
    );

    console.log("Time logs API response:", resp?.data);

    const entries = resp?.data?.logs || resp?.data?.time_logs || [];
    return entries.map(e => ({
      date: e.work_date || e.date,
      hours: Number(e.hours || e.time_spent || 0)
    }));
  } catch (error) {
    console.error("[getTimeLogsForUser] Error:", error);
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
  getAllTimeLogs
};
