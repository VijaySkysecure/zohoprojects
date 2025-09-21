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
  resolveOwnerId
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

// helper function for prompt template
prompts.addFunction("isAuthenticated", async (context, memory) => {
  const teamsChatId = context.activity.channelData?.teamsChatId || context.activity.from.id;
  try {
    const token = await getUserToken(teamsChatId);
    return !!(token && token.accessToken);
  } catch {
    return false;
  }
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

// -------------------------
// AI ACTIONS
// -------------------------

// Action to get pending tasks by owner
// app.ai.action("GetPendingTasksByOwner", async (context, state, parameters) => {
//   try {
//     await initializeConversationState(state);

//     const teamsChatId = context.activity.from.id;
//     const token = await getUserToken(teamsChatId);

//     if (!token || !token.accessToken) {
//       state.conversation.isAuthenticated = false;
//       await context.sendActivity(MessageFactory.text("🔒 You need to authenticate with Zoho Projects first."));
//       return "Authentication required";
//     }

//     if (!parameters.ownerName) {
//       await context.sendActivity(MessageFactory.text("❌ Missing required parameter: ownerName."));
//       return "Missing required parameters";
//     }

//     const resolvedOwner = await resolveOwnerName(context, token.accessToken, parameters.ownerName, teamsChatId, config.zohoPortalId);
//     if (!resolvedOwner) {
//       await context.sendActivity(`❌ I couldn't find a user named **${parameters.ownerName}** in your Zoho Projects portal.`);
//       return "User not found";
//     }

//     const tasks = await getPendingTasksByOwner(
//       context,
//       state,
//       parameters.ownerName,
//       parameters.limit || 15
//     );

//     if (!tasks || tasks.length === 0) {
//       await context.sendActivity(MessageFactory.text(`📊 No pending tasks found for **${resolvedOwner.name}**.`));
//       return `No pending tasks found for ${resolvedOwner.name}`;
//     }

//     const formattedTasks = tasks.map((task, i) =>
//       `✅ ${i + 1}. **${task.name}** (${task.projectName}) – Due: ${
//         task.dueDate ? moment(task.dueDate).format("DD MMM YYYY") : "N/A"
//       }`
//     );

//     const message = `📊 Found ${tasks.length} pending task(s) for **${resolvedOwner.name}**:\n\n` +
//       formattedTasks.join("\n");

//     await context.sendActivity(MessageFactory.text(message));
//     return `Successfully retrieved tasks.`;
//   } catch (error) {
//     console.error("Error in GetPendingTasksByOwner:", error);
//     await context.sendActivity(MessageFactory.text("❌ I couldn’t fetch the data, please try again."));
//     return `Error occurred: ${error.message}`;
//   }
// });


app.ai.action("GetPendingTasksByOwner", async (context, state, parameters) => {
  try {
    const { ownerName } = parameters;
    if (!ownerName) {
      await context.sendActivity("❌ Please provide an owner's name to find pending tasks.");
      return "Missing required parameter: ownerName";
    }

    const teamsChatId = context.activity.from.id;

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

    // 4️⃣ Format tasks nicely
    const formattedTasks = tasks.map((task, i) => {
      const dueDate = task.dueDate ? moment(task.dueDate).format("DD MMM YYYY") : "N/A";
      const projectName = task.projectName || "-";
      const status = task.status || "-";
      return `✅ ${i + 1}. **${task.name}** (${projectName}) – Status: ${status}, Due: ${dueDate}`;
    });

    const message = `📊 Found ${tasks.length} pending task(s) for **${owner.name}**:\n\n${formattedTasks.join("\n")}`;
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
    const teamsChatId = context.activity.from.id;

    if (!projectName) {
      await context.sendActivity(MessageFactory.text("❌ Missing required parameter: projectName."));
      return "Missing required parameters";
    }

    const projectResult = await getProjectByName(teamsChatId, config.zohoPortalId, projectName);

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
                ? projectResult.issues.join(", ")
                : "-"
            },
            { title: "Start Date", value: projectResult.startDate || "-" },
            { title: "End Date", value: projectResult.endDate || "-" },
            { title: "Tag", value: projectResult.tag || "-" },
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
// FEEDBACK LOOP
// -------------------------
app.feedbackLoop(async (context, state, feedbackLoopData) => {
  console.log("Feedback received: " + JSON.stringify(context.activity.value));
});

module.exports = app;