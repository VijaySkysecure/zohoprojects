const axios = require("axios");
const config = require("./config");
const {
  getUserToken: mongoGetUserToken,
  storeUserToken: mongoStoreUserToken,
  updateUserToken: mongoUpdateUserToken
} = require("./mongodb");
const moment = require("moment");

const { zohoApiBaseUrl, zohoPortalId } = config;

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
        if (error.response?.status === 429 && retries > 0) {
          const delay = Math.pow(2, 3 - retries) * 1000;
          await new Promise((resolve) => setTimeout(resolve, delay));
          retries--;
          continue;
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
  if (!resolvedOwner) {
    console.warn(`[ZOHO API] Could not resolve owner for name: ${ownerName}`);
    return [];
  }

  const token = await getUserToken(teamsChatId);
  const tasksResponse = await makeZohoAPICall(
    `portal/${portalId}/tasks`,
    token.accessToken,
    "GET",
    null,
    { owner: resolvedOwner.id, status: "Open,In Progress,To be Tested" },
    teamsChatId,
    portalId
  );

  return tasksResponse?.data?.tasks || [];
}



async function getProjectByName(teamsChatId, portalId, projectName) {
  const token = await getUserToken(teamsChatId);
  if (!token) throw new Error("No token found for user");

  const url = `portal/${portalId}/projects`;
  const projectsResponse = await makeZohoAPICall(
    url,
    token.accessToken,
    "GET",
    null,
    {},
    teamsChatId,
    portalId
  );

  const projects = projectsResponse?.data?.projects || [];

  const matched = projects.filter((p) =>
    p.name.toLowerCase().includes(projectName.toLowerCase())
  );

  if (matched.length === 0) return { notFound: true };
  if (matched.length > 1) return { multiple: matched.map((p) => p.name) };

  const project = matched[0];

  let percent = "-";
  let openTasks = "-";
  let closedTasks = "-";
  let tag = "-";
  let issuesList = [];
  let ownerName = project.owner?.name || "-";
  let statusName = project.status?.name || project.status || "-";
  let startDate = project.start_date
    ? moment(project.start_date).format("DD MMM YYYY")
    : "-";
  let endDate = project.end_date
    ? moment(project.end_date).format("DD MMM YYYY")
    : "-";

  try {
    // Fetch project details
    const projectDetail = await makeZohoAPICall(
      `/projects/${project.id_string}`,
      token.accessToken,
      "GET",
      null,
      {},
      teamsChatId,
      portalId
    );
    if (projectDetail?.data?.projects?.length > 0) {
      const detail = projectDetail.data.projects[0];
      percent = detail.percent_complete?.toString() || "-";
      tag = detail?.custom_status?.name || "-";
    }

    // Fetch tasks
    const tasksData = await makeZohoAPICall(
      `/projects/${project.id_string}/tasks`,
      token.accessToken,
      "GET",
      null,
      {},
      teamsChatId,
      portalId
    );
    if (tasksData?.data?.tasks) {
      const tasks = tasksData.data.tasks;
      openTasks = tasks
        .filter((t) =>
          ["Open", "In Progress", "To be Tested"].includes(t.status?.name)
        )
        .length.toString();
      closedTasks = tasks
        .filter((t) => t.status?.name === "Closed")
        .length.toString();
    }

    // Fetch issues
    const issuesData = await makeZohoAPICall(
      `/projects/${project.id_string}/issues`,
      token.accessToken,
      "GET",
      null,
      {},
      teamsChatId,
      portalId
    );
    if (issuesData?.data?.issues) {
      issuesList = issuesData.data.issues.map((i) => i.title);
    }
  } catch (err) {
    console.error("[PROJECT DETAILS API CALL ERROR]", err.message);
  }

  return {
    name: project.name,
    owner: ownerName,
    status: statusName,
    percent,
    openTasks,
    closedTasks,
    issues: issuesList.length > 0 ? issuesList : ["-"],
    startDate,
    endDate,
    tag,
  };
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
};
