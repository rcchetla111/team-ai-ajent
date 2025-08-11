const { CosmosClient } = require('@azure/cosmos');
const logger = require('../utils/logger');

class CosmosDbClient {
  constructor() {
    this.endpoint = process.env.COSMOS_ENDPOINT;
    this.key = process.env.COSMOS_KEY;
    this.databaseId = process.env.COSMOS_DATABASE_ID || 'agent365db';
    
    this.containers = {
      meetings: process.env.COSMOS_CONTAINER_MEETINGS || 'meetings',
      users: process.env.COSMOS_CONTAINER_USERS || 'users',
      chats: process.env.COSMOS_CONTAINER_CHATS || 'chats',
      summaries: process.env.COSMOS_CONTAINER_SUMMARIES || 'summaries',
      notifications: process.env.COSMOS_CONTAINER_NOTIFICATIONS || 'notifications',
      reminders: process.env.COSMOS_CONTAINER_REMINDERS || 'reminders'
    };

    if (!this.endpoint || !this.key) {
      throw new Error('Cosmos DB endpoint and key must be provided in environment variables');
    }

    this.client = new CosmosClient({
  endpoint: this.endpoint,
  key: this.key,
  userAgentSuffix: 'Agent365DigitalWorker',
  // ADD THESE LINES for local emulator:
  connectionPolicy: {
    DisableSSLVerification: true
  },
  // Alternative approach:
  agent: process.env.NODE_ENV === 'development' ? 
    require('https').Agent({ rejectUnauthorized: false }) : undefined
});

    this.database = null;
    this.containerClients = {};
  }

  async createDatabaseIfNotExists() {
    try {
      const { database } = await this.client.databases.createIfNotExists({
        id: this.databaseId
      });
      
      this.database = database;
      logger.info(`✅ Database '${this.databaseId}' ready`);
      return database;
    } catch (error) {
      logger.error('❌ Error creating database:', error);
      throw error;
    }
  }

  async createContainersIfNotExists() {
  const containerDefinitions = [
    {
      id: this.containers.meetings,
      partitionKey: '/userId',
      indexingPolicy: {
        automatic: true,
        indexingMode: 'consistent',
        includedPaths: [
          { path: '/*' },  // This is the required "/" path
          { path: '/meetingId/?' },
          { path: '/userId/?' },
          { path: '/startTime/?' },
          { path: '/status/?' }
        ],
        excludedPaths: [
          { path: '/"_etag"/?' }
        ]
      }
    },
    {
      id: this.containers.users,
      partitionKey: '/userId',
      indexingPolicy: {
        automatic: true,
        indexingMode: 'consistent',
        includedPaths: [
          { path: '/*' },  // This is the required "/" path
          { path: '/userId/?' },
          { path: '/email/?' },
          { path: '/lastLogin/?' }
        ],
        excludedPaths: [
          { path: '/"_etag"/?' }
        ]
      }
    },
    // Add other containers with simple indexing
    {
      id: this.containers.chats,
      partitionKey: '/meetingId',
      indexingPolicy: {
        automatic: true,
        indexingMode: 'consistent',
        includedPaths: [{ path: '/*' }],
        excludedPaths: [{ path: '/"_etag"/?' }]
      }
    },
    {
      id: this.containers.summaries,
      partitionKey: '/meetingId',
      indexingPolicy: {
        automatic: true,
        indexingMode: 'consistent',
        includedPaths: [{ path: '/*' }],
        excludedPaths: [{ path: '/"_etag"/?' }]
      }
    },
    {
      id: this.containers.notifications,
      partitionKey: '/userId',
      indexingPolicy: {
        automatic: true,
        indexingMode: 'consistent',
        includedPaths: [{ path: '/*' }],
        excludedPaths: [{ path: '/"_etag"/?' }]
      }
    },
    {
      id: this.containers.reminders,
      partitionKey: '/userId',
      indexingPolicy: {
        automatic: true,
        indexingMode: 'consistent',
        includedPaths: [{ path: '/*' }],
        excludedPaths: [{ path: '/"_etag"/?' }]
      }
    }
  ];

  try {
    for (const containerDef of containerDefinitions) {
      const { container } = await this.database.containers.createIfNotExists(containerDef);
      this.containerClients[containerDef.id] = container;
      logger.info(`✅ Container '${containerDef.id}' ready`);
    }
  } catch (error) {
    logger.error('❌ Error creating containers:', error);
    throw error;
  }
}

  getContainer(containerName) {
    const container = this.containerClients[containerName];
    if (!container) {
      throw new Error(`Container '${containerName}' not found`);
    }
    return container;
  }

  // Generic CRUD operations
  async createItem(containerName, item) {
    try {
      const container = this.getContainer(containerName);
      const { resource } = await container.items.create(item);
      logger.debug(`✅ Created item in ${containerName}:`, resource.id);
      return resource;
    } catch (error) {
      logger.error(`❌ Error creating item in ${containerName}:`, error);
      throw error;
    }
  }

  async getItem(containerName, id, partitionKey) {
    try {
      const container = this.getContainer(containerName);
      const { resource } = await container.item(id, partitionKey).read();
      return resource;
    } catch (error) {
      if (error.code === 404) {
        return null;
      }
      logger.error(`❌ Error getting item from ${containerName}:`, error);
      throw error;
    }
  }

  async updateItem(containerName, id, partitionKey, updates) {
    try {
      const container = this.getContainer(containerName);
      const { resource: existing } = await container.item(id, partitionKey).read();
      
      const updated = {
        ...existing,
        ...updates,
        updatedAt: new Date().toISOString()
      };
      
      const { resource } = await container.item(id, partitionKey).replace(updated);
      logger.debug(`✅ Updated item in ${containerName}:`, resource.id);
      return resource;
    } catch (error) {
      logger.error(`❌ Error updating item in ${containerName}:`, error);
      throw error;
    }
  }

  async deleteItem(containerName, id, partitionKey) {
    try {
      const container = this.getContainer(containerName);
      await container.item(id, partitionKey).delete();
      logger.debug(`✅ Deleted item from ${containerName}:`, id);
    } catch (error) {
      logger.error(`❌ Error deleting item from ${containerName}:`, error);
      throw error;
    }
  }

  async queryItems(containerName, query, parameters = []) {
    try {
      const container = this.getContainer(containerName);
      const { resources } = await container.items.query({
        query,
        parameters
      }).fetchAll();
      
      return resources;
    } catch (error) {
      logger.error(`❌ Error querying items from ${containerName}:`, error);
      throw error;
    }
  }

  // Meeting-specific methods
  async createMeeting(meetingData) {
    const meeting = {
      id: meetingData.id || require('uuid').v4(),
      meetingId: meetingData.meetingId,
      userId: meetingData.userId,
      subject: meetingData.subject,
      startTime: meetingData.startTime,
      endTime: meetingData.endTime,
      joinUrl: meetingData.joinUrl,
      webUrl: meetingData.webUrl,
      status: meetingData.status || 'scheduled',
      attendees: meetingData.attendees || [],
      agenda: meetingData.agenda,
      agentAttended: false,
      graphEventId: meetingData.graphEventId,
      isRecurring: meetingData.isRecurring || false,
      recurrencePattern: meetingData.recurrencePattern,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    return await this.createItem(this.containers.meetings, meeting);
  }

  async getMeetingsByUser(userId) {
    const query = 'SELECT * FROM c WHERE c.userId = @userId ORDER BY c.startTime DESC';
    return await this.queryItems(this.containers.meetings, query, [
      { name: '@userId', value: userId }
    ]);
  }

  async getMeetingsByDateRange(userId, startDate, endDate) {
    const query = `
      SELECT * FROM c 
      WHERE c.userId = @userId 
      AND c.startTime >= @startDate 
      AND c.startTime <= @endDate 
      ORDER BY c.startTime ASC
    `;
    return await this.queryItems(this.containers.meetings, query, [
      { name: '@userId', value: userId },
      { name: '@startDate', value: startDate },
      { name: '@endDate', value: endDate }
    ]);
  }

  // User-specific methods
  async createOrUpdateUser(userData) {
    const user = {
      id: userData.id || userData.userId,
      userId: userData.userId,
      email: userData.email,
      name: userData.name,
      tenantId: userData.tenantId,
      preferences: userData.preferences || {},
      lastLogin: new Date().toISOString(),
      createdAt: userData.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    // Try to update first, create if doesn't exist
    try {
      const existing = await this.getItem(this.containers.users, user.id, user.userId);
      if (existing) {
        return await this.updateItem(this.containers.users, user.id, user.userId, {
          lastLogin: user.lastLogin,
          name: user.name,
          preferences: user.preferences
        });
      }
    } catch (error) {
      // Item doesn't exist, create new
    }

    return await this.createItem(this.containers.users, user);
  }
}

// Create singleton instance
const cosmosClient = new CosmosDbClient();

module.exports = cosmosClient;