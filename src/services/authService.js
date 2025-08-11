const { ConfidentialClientApplication } = require('@azure/msal-node');
const logger = require('../utils/logger');

class AuthService {
  constructor() {
    this.msalConfig = {
      auth: {
        clientId: process.env.AZURE_CLIENT_ID,
        clientSecret: process.env.AZURE_CLIENT_SECRET,
        authority: process.env.AZURE_AUTHORITY
      },
      system: {
        loggerOptions: {
          loggerCallback(loglevel, message, containsPii) {
            if (!containsPii) {
              logger.debug('MSAL:', message);
            }
          },
          piiLoggingEnabled: false,
          logLevel: 'Info'
        }
      }
    };

    if (!process.env.AZURE_CLIENT_ID || !process.env.AZURE_CLIENT_SECRET) {
      logger.warn('⚠️ Azure AD credentials not configured. Real Teams integration disabled.');
      this.cca = null;
      return;
    }

    try {
      this.cca = new ConfidentialClientApplication(this.msalConfig);
      logger.info('✅ Microsoft Graph authentication initialized');
    } catch (error) {
      logger.error('❌ Failed to initialize Microsoft Graph auth:', error);
      this.cca = null;
    }

    // Token cache (in production, use Redis or database)
    this.tokenCache = new Map();
  }

  // Check if authentication is available
  isAvailable() {
    return this.cca !== null;
  }

  // Get application-only access token (for daemon apps)
  async getAppOnlyToken() {
    if (!this.isAvailable()) {
      throw new Error('Microsoft Graph authentication not available');
    }

    try {
      // Check cache first
      const cachedToken = this.tokenCache.get('app_token');
      if (cachedToken && cachedToken.expiresOn > new Date()) {
        return cachedToken.accessToken;
      }

      // Get new token
      const clientCredentialRequest = {
        scopes: ['https://graph.microsoft.com/.default'],
      };

      const response = await this.cca.acquireTokenByClientCredential(clientCredentialRequest);
      
      if (!response || !response.accessToken) {
        throw new Error('Failed to acquire access token');
      }

      // Cache the token
      this.tokenCache.set('app_token', {
        accessToken: response.accessToken,
        expiresOn: response.expiresOn
      });

      logger.info('✅ App-only access token acquired');
      return response.accessToken;

    } catch (error) {
      logger.error('❌ Failed to get app-only token:', error);
      throw error;
    }
  }

  // Get auth URL for user sign-in
  getAuthUrl() {
    if (!this.isAvailable()) {
      throw new Error('Microsoft Graph authentication not available');
    }

    const authCodeUrlParameters = {
      scopes: ['User.Read', 'OnlineMeetings.ReadWrite', 'Calendars.ReadWrite'],
      redirectUri: process.env.REDIRECT_URI || 'http://localhost:5000/api/auth/callback',
      prompt: 'consent'
    };

    return this.cca.getAuthCodeUrl(authCodeUrlParameters);
  }

  // Exchange auth code for token
  async getTokenFromCode(authCode) {
    if (!this.isAvailable()) {
      throw new Error('Microsoft Graph authentication not available');
    }

    try {
      const tokenRequest = {
        code: authCode,
        scopes: ['User.Read', 'OnlineMeetings.ReadWrite', 'Calendars.ReadWrite'],
        redirectUri: process.env.REDIRECT_URI || 'http://localhost:5000/api/auth/callback'
      };

      const response = await this.cca.acquireTokenByCode(tokenRequest);
      
      if (!response || !response.accessToken) {
        throw new Error('Failed to exchange code for token');
      }

      // Cache user token
      if (response.account) {
        this.tokenCache.set(response.account.homeAccountId, {
          accessToken: response.accessToken,
          refreshToken: response.refreshToken,
          expiresOn: response.expiresOn,
          account: response.account
        });
      }

      logger.info('✅ User access token acquired');
      return response;

    } catch (error) {
      logger.error('❌ Failed to exchange code for token:', error);
      throw error;
    }
  }

  // Get user token (for demo, we'll use app-only token)
  async getUserToken(userId = 'demo-user') {
    // For now, return app-only token
    // In production, you'd get user-specific token
    return await this.getAppOnlyToken();
  }

  // Clear token cache
  clearTokenCache() {
    this.tokenCache.clear();
    logger.info('🗑️ Token cache cleared');
  }
}

// Create singleton instance
const authService = new AuthService();

module.exports = authService;