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
userTokenSchema.pre('save', function(next) {
  this.updatedAt = Date.now();
  next();
});

// Create the model
const UserToken = mongoose.model('UserToken', userTokenSchema);

// -------------------------
// CONNECTION MANAGEMENT
// -------------------------
let isConnected = false;

async function connectToMongoDB() {
  if (isConnected) {
    return;
  }

  try {
    if (!config.mongoDBConnectionString) {
      throw new Error('MongoDB connection string is not configured');
    }

    await mongoose.connect(config.mongoDBConnectionString, {
      useNewUrlParser: true,
      useUnifiedTopology: true,
    });

    isConnected = true;
    console.log('Connected to MongoDB successfully');
  } catch (error) {
    console.error('MongoDB connection error:', error);
    throw error;
  }
}

async function disconnectFromMongoDB() {
  if (isConnected) {
    await mongoose.disconnect();
    isConnected = false;
    console.log('Disconnected from MongoDB');
  }
}

// -------------------------
// TOKEN MANAGEMENT FUNCTIONS
// -------------------------

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
    await connectToMongoDB();

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
        new: true,
        runValidators: true
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
    await connectToMongoDB();

    const token = await UserToken.findOne({ teamsChatId });
    console.log(`[MongoDB] Token query result:`, token ? 'Found' : 'Not found');

    if (!token) {
      console.log(`[MongoDB] No token found for teamsChatId: ${teamsChatId}`);
      return null;
    }

    // Log expiration state but return the token so caller can handle refresh logic
    const now = Date.now();
    const isExpired = token.expiresAt < now;
    console.log(`[MongoDB] Token expiration check - Now: ${now}, Expires: ${token.expiresAt}, IsExpired: ${isExpired}`);

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
    await connectToMongoDB();

    const result = await UserToken.deleteOne({ teamsChatId });
    console.log(`Deleted token for teamsChatId: ${teamsChatId}`);
    return result.deletedCount > 0;
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
    await connectToMongoDB();

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
  connectToMongoDB,
  disconnectFromMongoDB,
  storeUserToken,
  getUserToken,
  deleteUserToken,
  updateUserToken,
  UserToken
};
