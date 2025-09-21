// Import required packages
const express = require("express");

// This agent's adapter
const adapter = require("./adapter");

// This agent's main dialog.
const app = require("./app/app");

// Import MongoDB functions
const { 
  storeUserToken, 
  getUserToken, 
  deleteUserToken, 
  updateUserToken 
} = require("./mongodb");

// Create express application.
const expressApp = express();
expressApp.use(express.json());

const server = expressApp.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nAgent started, ${expressApp.name} listening to`, server.address());
});

// -------------------------
// API ENDPOINTS FOR USER CREDENTIALS
// -------------------------

// Store user credentials/token
expressApp.post("/api/auth/token", async (req, res) => {
  try {
    const { teamsChatId, userId, accessToken, refreshToken, expiresIn } = req.body;

    // Validate required fields
    if (!teamsChatId || !userId || !accessToken || !refreshToken || !expiresIn) {
      return res.status(400).json({
        success: false,
        error: "Missing required fields: teamsChatId, userId, accessToken, refreshToken, expiresIn"
      });
    }

    // Store the token
    const result = await storeUserToken(teamsChatId, userId, accessToken, refreshToken, expiresIn);
    
    res.json({
      success: true,
      message: "User token stored successfully",
      data: {
        teamsChatId: result.teamsChatId,
        userId: result.userId,
        expiresAt: result.expiresAt
      }
    });
  } catch (error) {
    console.error("Error storing user token:", error);
    res.status(500).json({
      success: false,
      error: "Failed to store user token",
      details: error.message
    });
  }
});

// Get user token
expressApp.get("/api/auth/token/:teamsChatId", async (req, res) => {
  try {
    const { teamsChatId } = req.params;
    
    if (!teamsChatId) {
      return res.status(400).json({
        success: false,
        error: "teamsChatId is required"
      });
    }

    const token = await getUserToken(teamsChatId);
    
    if (!token) {
      return res.status(404).json({
        success: false,
        error: "Token not found"
      });
    }

    res.json({
      success: true,
      data: {
        teamsChatId: token.teamsChatId,
        userId: token.userId,
        expiresAt: token.expiresAt,
        isExpired: token.expiresAt < Date.now()
      }
    });
  } catch (error) {
    console.error("Error getting user token:", error);
    res.status(500).json({
      success: false,
      error: "Failed to get user token",
      details: error.message
    });
  }
});

// Update user token
expressApp.put("/api/auth/token/:teamsChatId", async (req, res) => {
  try {
    const { teamsChatId } = req.params;
    const updateData = req.body;
    
    if (!teamsChatId) {
      return res.status(400).json({
        success: false,
        error: "teamsChatId is required"
      });
    }

    // Remove fields that shouldn't be updated directly
    delete updateData.teamsChatId;
    delete updateData.createdAt;

    const result = await updateUserToken(teamsChatId, updateData);
    
    if (!result) {
      return res.status(404).json({
        success: false,
        error: "Token not found"
      });
    }

    res.json({
      success: true,
      message: "User token updated successfully",
      data: {
        teamsChatId: result.teamsChatId,
        userId: result.userId,
        expiresAt: result.expiresAt,
        updatedAt: result.updatedAt
      }
    });
  } catch (error) {
    console.error("Error updating user token:", error);
    res.status(500).json({
      success: false,
      error: "Failed to update user token",
      details: error.message
    });
  }
});

// Delete user token
expressApp.delete("/api/auth/token/:teamsChatId", async (req, res) => {
  try {
    const { teamsChatId } = req.params;
    
    if (!teamsChatId) {
      return res.status(400).json({
        success: false,
        error: "teamsChatId is required"
      });
    }

    const deleted = await deleteUserToken(teamsChatId);
    
    if (!deleted) {
      return res.status(404).json({
        success: false,
        error: "Token not found"
      });
    }

    res.json({
      success: true,
      message: "User token deleted successfully"
    });
  } catch (error) {
    console.error("Error deleting user token:", error);
    res.status(500).json({
      success: false,
      error: "Failed to delete user token",
      details: error.message
    });
  }
});

// Health check endpoint
expressApp.get("/api/health", (req, res) => {
  res.json({
    success: true,
    message: "API is running",
    timestamp: new Date().toISOString()
  });
});

// Test endpoint to verify token storage
expressApp.get("/api/test/token/:teamsChatId", async (req, res) => {
  try {
    const { teamsChatId } = req.params;
    console.log(`[TEST] Checking token for teamsChatId: ${teamsChatId}`);
    
    const token = await getUserToken(teamsChatId);
    
    res.json({
      success: true,
      teamsChatId,
      token: token ? {
        userId: token.userId,
        expiresAt: token.expiresAt,
        isExpired: token.expiresAt < Date.now(),
        hasAccessToken: !!token.accessToken,
        hasRefreshToken: !!token.refreshToken
      } : null,
      message: token ? "Token found" : "Token not found"
    });
  } catch (error) {
    console.error("[TEST] Error:", error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Check current token details
expressApp.get("/api/debug/token/:teamsChatId", async (req, res) => {
  try {
    const { teamsChatId } = req.params;
    const token = await getUserToken(teamsChatId);
    
    if (!token) {
      return res.json({
        success: false,
        message: "No token found"
      });
    }
    
    res.json({
      success: true,
      token: {
        teamsChatId: token.teamsChatId,
        userId: token.userId,
        accessToken: token.accessToken,
        refreshToken: token.refreshToken,
        expiresAt: token.expiresAt,
        isExpired: token.expiresAt < Date.now(),
        tokenLength: token.accessToken?.length || 0,
        tokenPrefix: token.accessToken?.substring(0, 10) || 'N/A'
      }
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Test Zoho API with stored token
expressApp.get("/api/test/zoho/:teamsChatId", async (req, res) => {
  try {
    const { teamsChatId } = req.params;
    console.log(`[TEST ZOHO] Testing Zoho API for teamsChatId: ${teamsChatId}`);
    
    const token = await getUserToken(teamsChatId);
    if (!token) {
      return res.status(404).json({
        success: false,
        error: "Token not found"
      });
    }

    const axios = require('axios');
    const config = require('./config');
    
    // Test multiple endpoints to find what works
    const tests = [
      {
        name: "Tasks endpoint (current)",
        url: `${config.zohoApiBaseUrl}/portal/${config.zohoPortalId}/tasks`
      },
      {
        name: "Tasks endpoint (with slash)",
        url: `${config.zohoApiBaseUrl}/portal/${config.zohoPortalId}/tasks`
      },
      {
        name: "Tasks endpoint (v1)",
        url: `https://projectsapi.zoho.in/portal/${config.zohoPortalId}/tasks`
      },
      {
        name: "Tasks endpoint (v3)",
        url: `https://projectsapi.zoho.in/api/v3/portal/${config.zohoPortalId}/tasks`
      },
      {
        name: "Users endpoint (working)",
        url: `${config.zohoApiBaseUrl}/portal/${config.zohoPortalId}/users`
      }
    ];
    
    const results = [];
    
    for (const test of tests) {
      try {
        console.log(`[TEST ZOHO] Testing ${test.name}: ${test.url}`);
        const response = await axios.get(test.url, {
          headers: {
            'Authorization': `Zoho-oauthtoken ${token.accessToken}`,
            'Content-Type': 'application/json'
          }
        });
        
        results.push({
          name: test.name,
          success: true,
          status: response.status,
          data: response.data
        });
        
        // If one works, return success
        res.json({
          success: true,
          message: `Zoho API test successful with ${test.name}`,
          data: {
            workingEndpoint: test.name,
            status: response.status,
            response: response.data
          }
        });
        return;
        
      } catch (error) {
        console.log(`[TEST ZOHO] ${test.name} failed:`, error.response?.status, error.response?.data?.error?.message);
        results.push({
          name: test.name,
          success: false,
          error: error.response?.data?.error?.message || error.message,
          status: error.response?.status
        });
      }
    }
    
    // If all tests failed
    res.status(500).json({
      success: false,
      message: "All Zoho API tests failed",
      results: results
    });
    
  } catch (error) {
    console.error("[TEST ZOHO] Error:", error.response?.data || error.message);
    res.status(500).json({
      success: false,
      error: error.response?.data || error.message,
      status: error.response?.status
    });
  }
});

// Quick endpoint to copy token from 11111 to another teamsChatId
expressApp.post("/api/copy-token/:fromTeamsChatId/:toTeamsChatId", async (req, res) => {
  try {
    const { fromTeamsChatId, toTeamsChatId } = req.params;
    console.log(`[COPY] Copying token from ${fromTeamsChatId} to ${toTeamsChatId}`);
    
    // Get the source token
    const sourceToken = await getUserToken(fromTeamsChatId);
    if (!sourceToken) {
      return res.status(404).json({
        success: false,
        error: `No token found for source teamsChatId: ${fromTeamsChatId}`
      });
    }
    
    // Store it with the new teamsChatId
    const result = await storeUserToken(
      toTeamsChatId,
      sourceToken.userId,
      sourceToken.accessToken,
      sourceToken.refreshToken,
      Math.floor((sourceToken.expiresAt - Date.now()) / 1000) // Convert back to seconds
    );
    
    res.json({
      success: true,
      message: `Token copied from ${fromTeamsChatId} to ${toTeamsChatId}`,
      data: {
        fromTeamsChatId,
        toTeamsChatId,
        userId: result.userId,
        expiresAt: result.expiresAt
      }
    });
  } catch (error) {
    console.error("[COPY] Error:", error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Listen for incoming requests.
expressApp.post("/api/messages", async (req, res) => {
  // Route received a request to adapter for processing
  await adapter.process(req, res, async (context) => {
    // Dispatch to application for routing
    await app.run(context);
  });
});
