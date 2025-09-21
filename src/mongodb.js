const mongoose = require('mongoose');
const config = require('./config');
console.log(config);

// -------------------------
// USER TOKEN SCHEMA
// -------------------------
const userTokenSchema = new mongoose.Schema({
  teamsChatId: {
    type: String,
    required: true,
    unique: true,
    index: true
  },
  userId: {
    type: String,
    required: true
  },
  accessToken: {
    type: String,
    required: true
  },
  refreshToken: {
    type: String,
    required: true
  },
  expiresAt: {
    type: Number,
    required: true
  },
  createdAt: {
    type: Date,
    default: Date.now
  },
  updatedAt: {
    type: Date,
    default: Date.now
  }
});

// Update the updatedAt field before saving
userTokenSchema.pre('save', function (next) {
  this.updatedAt = Date.now();
  next();
});

// Create the model
const UserToken = mongoose.model('UserToken', userTokenSchema);



/**
 * Store or update user token in MongoDB
 * @param {string} teamsChatId - Teams chat ID
 * @param {string} userId - User ID
 * @param {string} accessToken - Zoho access token
 * @param {string} refreshToken - Zoho refresh token
 * @param {number} expiresIn - Token expiration time in seconds
 * @returns {Object} Stored token data
 */
async function storeUserToken(teamsChatId, userId, accessToken, refreshToken, expiresIn) {
  try {

    const expiresAt = Date.now() + (expiresIn * 1000);

    const tokenData = {
      teamsChatId,
      userId,
      accessToken,
      refreshToken,
      expiresAt
    };

    // Use upsert to either update existing or create new
    const result = await UserToken.findOneAndUpdate(
      { teamsChatId },
      tokenData,
      {
        upsert: true,
        new: true
      }
    );

    console.log(`Stored/Updated Zoho token for teamsChatId: ${teamsChatId}`);
    return result;
  } catch (error) {
    console.error('Error storing user token:', error);
    throw error;
  }
}

/**
 * Get user token from MongoDB
 * @param {string} teamsChatId - Teams chat ID
 * @returns {Object|null} Token data or null if not found
 */
async function getUserToken(teamsChatId) {
  try {
    console.log(`[MongoDB] Getting token for teamsChatId: ${teamsChatId}`);
   

    const token = await UserToken.findOne({ teamsChatId });
    console.log(`[MongoDB] Token query result:`, token ? 'Found' : 'Not found');

    if (!token) {
      console.log(`[MongoDB] No token found for teamsChatId: ${teamsChatId}`);
      return null;
    }

    return token;
  } catch (error) {
    console.error('[MongoDB] Error getting user token:', error);
    throw error;
  }
}


/**
 * Delete user token from MongoDB
 * @param {string} teamsChatId - Teams chat ID
 * @returns {boolean} True if deleted, false if not found
 */
async function deleteUserToken(teamsChatId) {
  try {

    const result = await UserToken.findOneAndDelete({ teamsChatId });
    return result;
  } catch (error) {
    console.error('Error deleting user token:', error);
    throw error;
  }
}

/**
 * Update user token (useful for refresh scenarios)
 * @param {string} teamsChatId - Teams chat ID
 * @param {Object} updateData - Data to update
 * @returns {Object|null} Updated token data or null if not found
 */
async function updateUserToken(teamsChatId, updateData) {
  try {

    const result = await UserToken.findOneAndUpdate(
      { teamsChatId },
      { ...updateData, updatedAt: Date.now() },
      { new: true }
    );

    if (result) {
      console.log(`Updated token for teamsChatId: ${teamsChatId}`);
    }

    return result;
  } catch (error) {
    console.error('Error updating user token:', error);
    throw error;
  }
}

// -------------------------
// EXPORTS
// -------------------------
module.exports = {
  storeUserToken,
  getUserToken,
  deleteUserToken,
  updateUserToken,
  UserToken
};
