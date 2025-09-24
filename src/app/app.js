const { MemoryStorage, MessageFactory, CardFactory } = require("botbuilder");
const path = require("path");
const config = require("../config");
const moment = require("moment");

// Teams AI library
const { Application, ActionPlanner, OpenAIModel, PromptManager } = require("@microsoft/teams-ai");

// Import Zoho Projects helpers
const {
  getUserToken,
  getPendingTasksByOwner,
  getProjectByName,
  getProjects,
  resolveOwnerId,
  getUsers,
  getTimeLogsForUser,
  getAllTimeLogs,
  getProjectIssues
} = require("../zoho");



// -------------------------
// AI MODEL CONFIG
// -------------------------
const model = new OpenAIModel({
  azureApiKey: config.azureOpenAIKey,
  azureDefaultDeployment: config.azureOpenAIDeploymentName,
  azureEndpoint: config.azureOpenAIEndpoint,
  useSystemMessages: true,
  logRequests: true,
  azureApiVersion: "2024-12-01-preview",
  stream: false,
});

// -------------------------
// PROMPTS
// -------------------------
const prompts = new PromptManager({
  promptsFolder: path.join(__dirname, "../prompts"),
});



// helper function for IST date-time
prompts.addFunction("currentDateTime", async () => {
  const options = {
    timeZone: "Asia/Kolkata",
    year: "numeric",
    month: "long",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
    hour12: true,
  };
  return new Date().toLocaleString("en-US", options);
});

// -------------------------
// ACTION PLANNER
// -------------------------
const planner = new ActionPlanner({
  model,
  prompts,
  defaultPrompt: "chat",
  actions: path.join(__dirname, "../actions.json"),
});

// -------------------------
// APP + STORAGE
// -------------------------
const storage = new MemoryStorage();

const app = new Application({
  storage,
  ai: {
    planner,
    feedbackLoopEnabled: true,
  },
});


// -------------------------
// DEFAULT CONVERSATION STATE
// -------------------------
const defaultConversationState = {
  isAuthenticated: false,
  lastTasksData: null,
  tasksCount: 0,
  formattedTasks: null,
  lastRawTasksResponse: null,
  userId: null,
};

async function initializeConversationState(state) {
  if (!state.conversation) state.conversation = { ...defaultConversationState };
  for (const key of Object.keys(defaultConversationState)) {
    if (state.conversation[key] === undefined) {
      state.conversation[key] = defaultConversationState[key];
    }
  }
}



const teamsChatId="11111"
app.ai.action("GetPendingTasksByOwner", async (context, state, parameters) => {
  try {
    const { ownerName } = parameters;
    if (!ownerName) {
      await context.sendActivity("❌ Please provide an owner's name to find pending tasks.");
      return "Missing required parameter: ownerName";
    }

    // const teamsChatId = context.activity.from.id;

    // 1️⃣ Get user token
    let tokenDoc;
    try {
      tokenDoc = await getUserToken(teamsChatId);
    } catch (err) {
      console.error("[GetPendingTasksByOwner] Token error:", err.message);
      await context.sendActivity("🔒 You need to authenticate with Zoho Projects first.");
      return "Authentication required";
    }
    const token = tokenDoc.accessToken;

    console.log(`Searching for pending tasks for owner: ${ownerName}`);

    // 2️⃣ Resolve Zoho ownerId with fuzzy matching
    const owner = await resolveOwnerId(teamsChatId, config.zohoPortalId, ownerName);
    if (!owner) {
      await context.sendActivity(`❌ Could not resolve owner for name: **${ownerName}**.`);
      return `Error: Owner resolution failed for ${ownerName}`;
    }

    // 3️⃣ Fetch pending tasks by ownerName
    const tasks = await getPendingTasksByOwner(context, state, owner.name);
    if (!tasks || tasks.length === 0) {
      await context.sendActivity(`📊 No pending tasks found for **${owner.name}**.`);
      return `No pending tasks for ${owner.name}`;
    }

    // 4️⃣ Format tasks with proper line breaks
    const formattedTasks = tasks.map((task, i) => {
      const dueDate = task.dueDate ? moment(task.dueDate).format("DD MMM YYYY") : "N/A";
      const projectName = task.project || "-";
      const status = task.status || "-";
      const priority = task.priority || "none";
      
      return `${i + 1}. **${task.name}**\n   Project: ${projectName}\n   Status: ${status}\n   Priority: ${priority}\n   Due: ${dueDate}`;
    });

    const message = `   Found ${tasks.length} pending task(s) for **${owner.name}**:\n\n${formattedTasks.join("\n\n")}`;
    await context.sendActivity(message);

    return `Successfully retrieved ${tasks.length} pending tasks for ${owner.name}`;

  } catch (error) {
    console.error("[GetPendingTasksByOwner] Unexpected error:", error);
    await context.sendActivity("⚠️ An unexpected error occurred while retrieving tasks. Please try again.");
    return "Unexpected error";
  }
});



// Action to get project details
app.ai.action("GetProjectDetails", async (context, state, params) => {
  try {
    const projectName = params.projectName;
    // const teamsChatId = context.activity.from.id;

    if (!projectName) {
      await context.sendActivity(MessageFactory.text("❌ Missing required parameter: projectName."));
      return "Missing required parameters";
    }

    const projectResult = await getProjectByName("11111", config.zohoPortalId, projectName);

    if (projectResult.notFound) {
      await context.sendActivity(MessageFactory.text("❌ No project found."));
      return "No project found";
    }

    if (projectResult.multiple) {
      const projectsList = projectResult.multiple.join(", ");
      await context.sendActivity(MessageFactory.text(`⚠️ Multiple projects match your query: ${projectsList}. Please specify.`));
      return "Multiple projects found";
    }

    // Handle specific field query
    if (params.fields && params.fields.length > 0) {
      const field = params.fields[0];
      const fieldValue = projectResult[field] || "-";
      await context.sendActivity(MessageFactory.text(`📌 The ${field} for **${projectResult.name}** is: **${fieldValue}**`));
    }

    // Send the Adaptive Card
    const card = {
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        {
          type: "TextBlock",
          text: `📌 ${projectResult.name}`,
          weight: "Bolder",
          size: "Large",
          wrap: true,
        },
        {
          type: "FactSet",
          facts: [
            { title: "Owner", value: projectResult.owner || "-" },
            { title: "Status", value: projectResult.status || "-" },
            { title: "% Completed", value: projectResult.percent || "-" },
            { title: "Open Tasks", value: projectResult.openTasks || "-" },
            { title: "Closed Tasks", value: projectResult.closedTasks || "-" },
            {
              title: "Issues",
              value: Array.isArray(projectResult.issues) && projectResult.issues.length > 0
                ? projectResult.issues.map((issue, index) => 
                    `${index + 1}. ${issue.title} (${issue.status})`
                  ).join("\n")
                : "No issues found"
            },
            { title: "Start Date", value: projectResult.startDate || "-" },
            { title: "End Date", value: projectResult.endDate || "-" },
            // { title: "Tag", value: projectResult.tag || "-" },
          ],
        },
      ],
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    };

    await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
    return "Successfully retrieved project details";

  } catch (err) {
    console.error("[PROJECT DETAILS ERROR]", err);
    await context.sendActivity(MessageFactory.text("❌ I couldn’t fetch the data, please try again."));
    return "Error occurred";
  }
});


// -------------------------
// ShowTimeLogs action - now handles natural language queries
app.ai.action("ShowTimeLogs", async (context, state, parameters) => {
  console.log("\n=== SHOW TIME LOGS ACTION CALLED ===");
  console.log("Parameters:", parameters);
  

  const teamsChatId = "11111";

  try {
    console.log("Fetching all time logs...");
    const allTimeLogs = await getAllTimeLogs(teamsChatId, config.zohoPortalId);
    console.log("Total time logs fetched:", allTimeLogs.length);
    
    if (!allTimeLogs || allTimeLogs.length === 0) {
      await context.sendActivity(MessageFactory.text("❌ No time logs found in the system."));
      return;
    }

    // Get users for reference
    const users = await getUsers(teamsChatId);
    
    // Show summary of available data
    const uniqueUsers = [...new Set(allTimeLogs.map(log => log.userName))];
    const dateRange = {
      earliest: moment.min(allTimeLogs.map(log => moment(log.date))).format("DD MMM YYYY"),
      latest: moment.max(allTimeLogs.map(log => moment(log.date))).format("DD MMM YYYY")
    };
    
    const message = `�� **Time Logs Available**\n\n` +
      `**Users with time logs:** ${uniqueUsers.length}\n` +
      `**Date range:** ${dateRange.earliest} to ${dateRange.latest}\n` +
      `**Total entries:** ${allTimeLogs.length}\n\n` +
      `**Available users:**\n${uniqueUsers.map((name, index) => `${index + 1}. ${name}`).join('\n')}\n\n` +
      `**Query examples:**\n` +
      `• "time logs for Divakar in May"\n` +
      `• "show me Rajat's time logs for first week of April"\n` +
      `• "time logs for Anuj from 2025-01-01 to 2025-01-31"`;
    
    await context.sendActivity(MessageFactory.text(message));
    console.log("Time logs summary sent successfully");
  } catch (error) {
    console.error("Error in ShowTimeLogs:", error);
    await context.sendActivity(MessageFactory.text(`❌ Error loading time logs data: ${error.message}`));
  }
});


// Debug function to test user resolution
async function debugUserResolution(teamsChatId, userName) {
  try {
    const users = await getUsers(teamsChatId);
    console.log("All available users:");
    users.forEach((user, index) => {
      console.log(`  ${index + 1}. ID: ${user.id}, Name: "${user.name}"`);
    });
    
    console.log(`\nSearching for: "${userName}"`);
    const foundUser = users.find(u => 
      u.name.toLowerCase().includes(userName.toLowerCase()) ||
      u.name.toLowerCase().split(' ')[0] === userName.toLowerCase()
    );
    
    if (foundUser) {
      console.log(`Found user: ID: ${foundUser.id}, Name: "${foundUser.name}"`);
    } else {
      console.log("User not found");
    }
    
    return foundUser;
  } catch (error) {
    console.error("Debug user resolution error:", error);
    return null;
  }
}

// -------------------------
// GetTimeLogs action
app.ai.action("GetTimeLogs", async (context, state, parameters) => {
  console.log("\n=== GET TIME LOGS ACTION CALLED ===");
  console.log("Parameters:", parameters);
  
  const teamsChatId = "11111";
  
  const { userInput } = parameters || {};
  console.log("User input:", userInput);
  
  if (!userInput) {
    await context.sendActivity(MessageFactory.text("Please specify which user's time logs you want to see."));
    return;
  }
  
  try {
    // Get users first for debugging
    const users = await getUsers(teamsChatId);
    console.log("Available users:", users.map(u => `${u.name} (ID: ${u.id})`));
    
    // Parse user input first
    const { userName, startDate, endDate, period } = parseTimeLogQuery(userInput, users);
    
    if (!userName) {
      const availableUsers = users.map(u => u.name).join(', ');
      await context.sendActivity(MessageFactory.text(`Could not identify the user. Available users: ${availableUsers}`));
      return;
    }

    if (!startDate || !endDate) {
      await context.sendActivity(MessageFactory.text("Could not identify the date range. Please specify dates clearly."));
      return;
    }

    // Find the user ID for the specific user
    const targetUser = users.find(u => u.name.toLowerCase() === userName.toLowerCase());
    if (!targetUser) {
      await context.sendActivity(MessageFactory.text(`User "${userName}" not found in the system.`));
      return;
    }

    console.log(`Fetching time logs for ${userName} (ID: ${targetUser.id}) from ${startDate} to ${endDate}`);

    // Use getTimeLogsForUser instead of getAllTimeLogs
    const timeLogs = await getTimeLogsForUser(teamsChatId, config.zohoPortalId, targetUser.id, startDate, endDate);
    console.log(`Found ${timeLogs.length} time log entries`);
    
    if (timeLogs.length === 0) {
      await context.sendActivity(
        MessageFactory.text(`No time logs found for **${userName}** from ${moment(startDate).format("DD MMM YYYY")} to ${moment(endDate).format("DD MMM YYYY")}.`)
      );
      return;
    }
    
    // Calculate statistics and format response
    const stats = calculateTimeLogStats(timeLogs, startDate, endDate);
    
    const formattedLogs = timeLogs
      .sort((a, b) => moment(b.date).diff(moment(a.date)))
      .slice(0, 20) // Limit to 20 entries
      .map(log => {
        const date = moment(log.date).format("DD MMM YYYY (ddd)");
        const hours = Number(log.hours) || 0;
        const project = log.projectName || "Unknown Project";
        return `${date}: ${hours} hours - ${project}`;
      })
      .join("\n");
    
    const periodText = period ? ` (${period})` : "";
    const message = `Time Logs for ${userName}${periodText}\n` +
      `**Period:** ${moment(startDate).format("DD MMM YYYY")} to ${moment(endDate).format("DD MMM YYYY")}\n\n` +
      `**Summary:**\n` +
      `• Total Hours: ${stats.totalHours}\n` +
      `• Work Days: ${stats.workDays}\n` +
      `• Average per day: ${stats.averageHours.toFixed(1)} hours\n\n` +
      `**Daily Logs:**\n${formattedLogs}`;
    
    await context.sendActivity(MessageFactory.text(message));
    console.log("Time logs response sent successfully");
    
  } catch (error) {
    console.error("[GetTimeLogs] Error:", error);
    await context.sendActivity(
      MessageFactory.text(`Error fetching time logs: ${error.message}`)
    );
  }
});

// -------------------------
// GetProjectIssues action
app.ai.action("GetProjectIssues", async (context, state, parameters) => {
  console.log("\n=== GET PROJECT ISSUES ACTION CALLED ===");
  console.log("Parameters:", parameters);
  
  const teamsChatId = "11111";
  
  const { projectName } = parameters || {};
  console.log("Project name:", projectName);
  
  if (!projectName) {
    await context.sendActivity(MessageFactory.text("❌ Please specify which project's issues you want to see."));
    return;
  }
  
  try {
    // Get user token
    let tokenDoc;
    try {
      tokenDoc = await getUserToken(teamsChatId);
    } catch (err) {
      console.error("[GetProjectIssues] Token error:", err.message);
      await context.sendActivity("🔒 You need to authenticate with Zoho Projects first.");
      return;
    }

    console.log(`Fetching issues for project: ${projectName}`);
    
    // Fetch issues for the project
    const issues = await getProjectIssues(teamsChatId, config.zohoPortalId, projectName);
    console.log(`Found ${issues.length} issues for project: ${projectName}`);
    
    // Debug: Log the first issue to see the data structure
    if (issues.length > 0) {
      console.log(`\n=== DEBUG: First issue data ===`);
      console.log(JSON.stringify(issues[0], null, 2));
    }
    
    if (issues.length === 0) {
      await context.sendActivity(
        MessageFactory.text(`📊 No issues found for project **${projectName}**.`)
      );
      return;
    }

    // Helper function to sanitize values for adaptive cards
    const sanitizeValue = (value) => {
      if (!value || value === "N/A") return "N/A";
      return String(value).replace(/[^\w\s\-.,:()]/g, '').trim() || "N/A";
    };

    // Create adaptive card for issues
    const card = {
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        {
          type: "TextBlock",
          text: `📋 Issues in ${projectName}`,
          weight: "Bolder",
          size: "Large",
          wrap: true,
        },
        {
          type: "TextBlock",
          text: `Found ${issues.length} issue(s)`,
          size: "Medium",
          color: "Accent",
          wrap: true,
        },
        {
          type: "Container",
          items: issues.slice(0, 10).map((issue, index) => ({
            type: "Container",
            style: "emphasis",
            items: [
              {
                type: "ColumnSet",
                columns: [
                  {
                    type: "Column",
                    width: "stretch",
                    items: [
                      {
                        type: "TextBlock",
                        text: `${sanitizeValue(issue.name)}`,
                        weight: "Bolder",
                        wrap: true,
                        size: "Medium"
                      },
                      {
                        type: "FactSet",
                        facts: [
                          { title: "Project", value: sanitizeValue(issue.project) },
                          { title: "Reporter", value: sanitizeValue(issue.reporter) },
                          { title: "Assignee", value: sanitizeValue(issue.assignee) },
                          { title: "Status", value: sanitizeValue(issue.status) },
                          { title: "Severity", value: sanitizeValue(issue.severity) },
                          { title: "Created", value: sanitizeValue(issue.createdTime) },
                          { title: "Due Date", value: sanitizeValue(issue.dueDate) },
                          { title: "Last Modified", value: sanitizeValue(issue.lastModified) }
                        ]
                      }
                    ]
                  }
                ]
              }
            ]
          }))
        }
      ],
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    };

    // Add pagination info if there are more than 10 issues
    if (issues.length > 10) {
      card.body.push({
        type: "TextBlock",
        text: `*Showing first 10 of ${issues.length} issues*`,
        size: "Small",
        color: "Default",
        wrap: true,
      });
    }

    await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
    console.log("Project issues response sent successfully");
    
  } catch (error) {
    console.error("[GetProjectIssues] Error:", error);
    await context.sendActivity(
      MessageFactory.text(`❌ Error fetching issues for project ${projectName}: ${error.message}`)
    );
  }
});

function parseTimeLogQuery(userInput, users) {
  const moment = require("moment");
  const input = userInput.toLowerCase();
  let userName = null;
  let startDate = null;
  let endDate = null;
  let period = null;
  
  console.log("Input:", input);
  console.log("Available users:", users.map(u => u.name));
  
  // Extract the name part from the input - FIXED LOGIC
  let nameQuery = input
    .replace(/time logs for /gi, '')
    .replace(/show me /gi, '')
    .replace(/'s time logs/gi, '')
    .replace(/\s+(in|for|during)\s+.*/gi, '') // Remove everything after "in", "for", "during"
    .trim();
  
  console.log("Extracted name query:", nameQuery);
  
  // Find user name by matching with available users
  for (const user of users) {
    const userFullName = user.name.toLowerCase();
    const parts = userFullName.split(' ');
    const userFirstName = parts[0] || "";
    const userLastName = parts[parts.length - 1] || "";
    
    console.log(`Checking user: ${user.name} (${userFirstName}, ${userLastName})`);
    
    // Check exact matches first, then partial matches
    if (userFullName === nameQuery || 
        userFirstName === nameQuery || 
        userLastName === nameQuery ||
        userFullName.includes(nameQuery) ||
        nameQuery.includes(userFirstName)) {
      userName = user.name;
      console.log(`Found matching user: ${userName}`);
      break;
    }
  }
  
  // If no exact match, try more flexible matching
  if (!userName) {
    console.log("No exact match found, trying flexible matching...");
    for (const user of users) {
      const userFullName = user.name.toLowerCase();
      const parts = userFullName.split(' ');
      
      // Check if any word in the query matches any word in the user name
      const queryWords = nameQuery.split(' ');
      const nameWords = parts;
      
      const hasMatch = queryWords.some(qWord => 
        nameWords.some(nWord => 
          qWord.length > 2 && nWord.includes(qWord)
        )
      );
      
      if (hasMatch) {
        userName = user.name;
        console.log(`Found flexible matching user: ${userName}`);
        break;
      }
    }
  }
  
  // Parse date patterns - keeping 2025 as requested
  const currentYear = 2025;
  
  const monthMatch = input.match(/(?:in|for|during)\s+(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)/i);
  if (monthMatch) {
    const monthName = monthMatch[1];
    const monthNum = moment().month(monthName).month();
    if (monthNum >= 0) {
      const targetYear = currentYear;
      startDate = moment().year(targetYear).month(monthNum).startOf('month').format('YYYY-MM-DD');
      endDate = moment().year(targetYear).month(monthNum).endOf('month').format('YYYY-MM-DD');
      period = `${monthName} ${targetYear}`;
      console.log(`Parsed month: ${monthName} -> ${monthNum} -> ${startDate} to ${endDate}`);
    }
  }
  
  // Other date parsing patterns...
  const weekMatch = input.match(/(?:first|second|third|fourth|1st|2nd|3rd|4th)\s+week\s+of\s+(\w+)/);
  if (weekMatch) {
    const monthName = weekMatch[1];
    const monthNum = moment().month(monthName).month();
    if (monthNum >= 0) {
      const monthStart = moment().year(currentYear).month(monthNum).startOf('month');
      startDate = monthStart.format('YYYY-MM-DD');
      endDate = monthStart.add(6, 'days').format('YYYY-MM-DD');
      period = `first week of ${monthName}`;
    }
  }
  
  const dateRangeMatch = input.match(/from\s+(\d{4}-\d{2}-\d{2})\s+to\s+(\d{4}-\d{2}-\d{2})/);
  if (dateRangeMatch) {
    startDate = dateRangeMatch[1];
    endDate = dateRangeMatch[2];
  }
  
  if (input.includes('last 7 days')) {
    endDate = moment().format('YYYY-MM-DD');
    startDate = moment().subtract(7, 'days').format('YYYY-MM-DD');
    period = 'last 7 days';
  } else if (input.includes('last month')) {
    startDate = moment().subtract(1, 'month').startOf('month').format('YYYY-MM-DD');
    endDate = moment().subtract(1, 'month').endOf('month').format('YYYY-MM-DD');
    period = 'last month';
  }
  
  console.log(`Parsed result - User: ${userName}, Start: ${startDate}, End: ${endDate}`);
  return { userName, startDate, endDate, period };
}

// Add this test function to debug what endpoints are available
async function testAvailableEndpoints(teamsChatId, portalId) {
  const token = await getUserToken(teamsChatId);
  
  const endpointsToTest = [
    'timelogs',
    'timesheet', 
    'logs',
    'projects',
    'tasks'
  ];
  
  for (const endpoint of endpointsToTest) {
    try {
      console.log(`Testing endpoint: ${endpoint}`);
      const resp = await makeZohoAPICall(
        `portal/${portalId}/${endpoint}`,
        token.accessToken,
        "GET",
        null,
        { per_page: 1 },
        teamsChatId,
        portalId
      );
      console.log(`✅ ${endpoint} works - response keys:`, Object.keys(resp?.data || {}));
    } catch (error) {
      console.log(`❌ ${endpoint} failed:`, error.message);
    }
  }
}

// Helper function to calculate time log statistics
function calculateTimeLogStats(logs, startDate, endDate) {
  
  let totalHours = 0;
  let billableHours = 0;
  
  // Calculate work days in the period
  let workDays = 0;
  let current = moment(startDate);
  const end = moment(endDate);
  
  while (current.isSameOrBefore(end)) {
    if (current.isoWeekday() <= 5) workDays++;
    current.add(1, "day");
  }
  
  const totalWorkHours = workDays * 8;
  
  // Calculate hours from logs
  logs.forEach(log => {
    const hours = Number(log.hours) || 0;
    totalHours += hours;
    
    const logDate = moment(log.date);
    if (logDate.isoWeekday() <= 5) {
      billableHours += hours;
    }
  });
  
  const averageHours = logs.length > 0 ? totalHours / logs.length : 0;
  
  return {
    totalHours,
    billableHours,
    totalWorkHours,
    workDays,
    averageHours
  };
}






// -------------------------
// FEEDBACK LOOP
// -------------------------
app.feedbackLoop(async (context, state, feedbackLoopData) => {
  console.log("Feedback received: " + JSON.stringify(context.activity.value));
});

module.exports = app;