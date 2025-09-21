// Load environment variables from .env file
require('dotenv').config({ path: ['env/.env'] });

const config = {
  MicrosoftAppId: process.env.BOT_ID,
  MicrosoftAppType: process.env.BOT_TYPE,
  MicrosoftAppTenantId: process.env.BOT_TENANT_ID,
  MicrosoftAppPassword: process.env.BOT_PASSWORD,
  azureOpenAIKey: process.env.AZURE_OPENAI_API_KEY,
  azureOpenAIEndpoint: process.env.AZURE_OPENAI_ENDPOINT,
  azureOpenAIDeploymentName: process.env.AZURE_OPENAI_DEPLOYMENT_NAME,
  mongoDBConnectionString: process.env.MONGODB_URL,
  zohoClientId: process.env.ZOHO_CLIENT_ID,
  zohoClientSecret: process.env.ZOHO_CLIENT_SECRET,
  zohoPortalId: process.env.ZOHO_PORTAL_ID,
  zohoApiBaseUrl: process.env.ZOHO_API_BASE_URL || "https://projectsapi.zoho.in/api/v3",
};

module.exports = config;
