const axios = require('axios');
const authService = require('./authService'); // FIXED: Remove destructuring
const cosmosClient = require('../config/cosmosdb');
const logger = require('../utils/logger');

class MeetingAttendanceService {
  constructor() {
    this.graphEndpoint = process.env.GRAPH_API_ENDPOINT || 'https://graph.microsoft.com/v1.0';
    this.activeMeetings = new Map(); // Track meetings the agent is attending
    this.chatMonitors = new Map(); // Track chat monitoring for each meeting
  }

  // Join a meeting as an AI agent
  async joinMeeting(meetingId, userId) {
    try {
      logger.info('🤖 Agent attempting to join meeting', { meetingId, userId });

      // Get meeting details from database
      const meetings = await cosmosClient.queryItems('meetings',
        'SELECT * FROM c WHERE c.meetingId = @meetingId',
        [{ name: '@meetingId', value: meetingId }]
      );

      if (!meetings || meetings.length === 0) {
        throw new Error('Meeting not found');
      }

      const meeting = meetings[0];

      // Check if meeting is scheduled to start soon or already started
      const now = new Date();
      const startTime = new Date(meeting.startTime);
      const endTime = new Date(meeting.endTime);
      const bufferTime = 15 * 60 * 1000; // 15 minutes before

      if (now < startTime.getTime() - bufferTime) {
        throw new Error('Meeting has not started yet (joins 15 minutes before start time)');
      }

      if (now > endTime) {
        throw new Error('Meeting has already ended');
      }

      // Simulate agent joining (in real implementation, this would use Teams SDK)
      const attendanceRecord = {
        meetingId: meetingId,
        agentJoinedAt: now.toISOString(),
        agentStatus: 'attending',
        monitoringStarted: now.toISOString(),
        chatCaptureEnabled: true,
        transcriptCaptureEnabled: true
      };

      // Store in active meetings
      this.activeMeetings.set(meetingId, attendanceRecord);

      // Update meeting record in database - use the database ID, not meetingId
      await cosmosClient.updateItem('meetings', meeting.id, userId, {
        agentAttended: true,
        agentJoinedAt: now.toISOString(),
        agentStatus: 'attending'
      });

      // Start chat monitoring
      await this.startChatMonitoring(meetingId, meeting);

     logger.info('✅ Agent successfully joined meeting', {
        meetingId,
        joinedAt: attendanceRecord.agentJoinedAt
      });

      // 🆕 ADD THIS CODE HERE:
      // Start auto-insights
      const chatCaptureService = require('./chatCaptureService');
      await chatCaptureService.startAutoInsights(meetingId);

      return {
        success: true,
        message: 'Agent joined meeting successfully',
        attendanceRecord: attendanceRecord,
        capabilities: {
          chatMonitoring: true,
          transcriptCapture: true,
          realTimeAnalysis: true
        }
      };

    } catch (error) {
      logger.error('❌ Failed to join meeting:', error);
      throw error;
    }
  }

  // Leave a meeting
  async leaveMeeting(meetingId, userId) {
    try {
      logger.info('🤖 Agent leaving meeting', { meetingId });

      const attendanceRecord = this.activeMeetings.get(meetingId);
      if (!attendanceRecord) {
        throw new Error('Agent is not currently attending this meeting');
      }

      const now = new Date();
      attendanceRecord.agentLeftAt = now.toISOString();
      attendanceRecord.agentStatus = 'left';

      // Stop chat monitoring
      await this.stopChatMonitoring(meetingId);

      // Get meeting to update the correct database record
      const meetings = await cosmosClient.queryItems('meetings',
        'SELECT * FROM c WHERE c.meetingId = @meetingId',
        [{ name: '@meetingId', value: meetingId }]
      );

      if (meetings && meetings.length > 0) {
        const meeting = meetings[0];
        await cosmosClient.updateItem('meetings', meeting.id, userId, {
          agentLeftAt: now.toISOString(),
          agentStatus: 'left'
        });
      }

      const chatCaptureService = require('./chatCaptureService');
      await chatCaptureService.stopAutoInsights(meetingId);

      // Remove from active meetings
      this.activeMeetings.delete(meetingId);

      logger.info('✅ Agent left meeting successfully', {
        meetingId,
        leftAt: attendanceRecord.agentLeftAt
      });

      return {
        success: true,
        message: 'Agent left meeting successfully',
        attendanceRecord: attendanceRecord
      };

    } catch (error) {
      logger.error('❌ Failed to leave meeting:', error);
      throw error;
    }
  }

  async startSimulatedMonitoring(meetingId, meeting) {
  try {
    logger.info('🔄 Starting simulated chat monitoring', { meetingId });

    const chatMonitor = {
      meetingId: meetingId,
      startedAt: new Date().toISOString(),
      lastMessageCheck: new Date().toISOString(),
      messagesCount: 0,
      isActive: true,
      simulated: true
    };

    this.chatMonitors.set(meetingId, chatMonitor);

    // Start simulated chat checking (every 45 seconds)
    const chatInterval = setInterval(async () => {
      try {
        await this.checkForNewMessages(meetingId, null); // null for simulated
      } catch (error) {
        logger.warn('Simulated chat monitoring error:', error);
      }
    }, 45000); // 45 seconds

    chatMonitor.interval = chatInterval;

    logger.info('✅ Simulated chat monitoring started', { meetingId });
    return chatMonitor;

  } catch (error) {
    logger.error('❌ Failed to start simulated monitoring:', error);
    throw error;
  }
}

  // Start monitoring chat for a meeting
 async startChatMonitoring(meetingId, meeting) {
  try {
    logger.info('🔍 Starting chat monitoring', { meetingId });

    // ⭐ CRITICAL: Add null check for authService
    if (!authService || !authService.isAvailable || !authService.isAvailable()) {
      logger.warn('⚠️ Auth service not available, using simulated monitoring');
      return this.startSimulatedMonitoring(meetingId, meeting);
    }

    let accessToken;
    try {
      accessToken = await authService.getAppOnlyToken();
    } catch (authError) {
      logger.warn('⚠️ Failed to get auth token, using simulated monitoring:', authError.message);
      return this.startSimulatedMonitoring(meetingId, meeting);
    }

    // Initialize chat monitoring record
    const chatMonitor = {
      meetingId: meetingId,
      startedAt: new Date().toISOString(),
      lastMessageCheck: new Date().toISOString(),
      messagesCount: 0,
      isActive: true
    };

    this.chatMonitors.set(meetingId, chatMonitor);

    // Start periodic chat checking (every 30 seconds)
    const chatInterval = setInterval(async () => {
      try {
        await this.checkForNewMessages(meetingId, accessToken);
      } catch (error) {
        logger.warn('Chat monitoring error:', error);
      }
    }, 30000); // 30 seconds

    chatMonitor.interval = chatInterval;

    logger.info('✅ Chat monitoring started', { meetingId });
    return chatMonitor;

  } catch (error) {
    logger.error('❌ Failed to start chat monitoring:', error);
    // Fall back to simulated monitoring
    return this.startSimulatedMonitoring(meetingId, meeting);
  }
}

  // Start simulated monitoring when auth is not available
  async startSimulatedMonitoring(meetingId, meeting) {
    try {
      logger.info('🔄 Starting simulated chat monitoring', { meetingId });

      const chatMonitor = {
        meetingId: meetingId,
        startedAt: new Date().toISOString(),
        lastMessageCheck: new Date().toISOString(),
        messagesCount: 0,
        isActive: true,
        simulated: true
      };

      this.chatMonitors.set(meetingId, chatMonitor);

      // Start simulated chat checking (every 45 seconds)
      const chatInterval = setInterval(async () => {
        try {
          await this.checkForNewMessages(meetingId, null);
        } catch (error) {
          logger.warn('Simulated chat monitoring error:', error);
        }
      }, 45000); // 45 seconds

      chatMonitor.interval = chatInterval;

      logger.info('✅ Simulated chat monitoring started', { meetingId });
      return chatMonitor;

    } catch (error) {
      logger.error('❌ Failed to start simulated monitoring:', error);
      throw error;
    }
  }

  // Stop chat monitoring
  async stopChatMonitoring(meetingId) {
    try {
      const chatMonitor = this.chatMonitors.get(meetingId);
      if (chatMonitor) {
        chatMonitor.isActive = false;
        chatMonitor.stoppedAt = new Date().toISOString();

        // Clear interval
        if (chatMonitor.interval) {
          clearInterval(chatMonitor.interval);
        }

        this.chatMonitors.delete(meetingId);

        logger.info('✅ Chat monitoring stopped', { meetingId });
      }
    } catch (error) {
      logger.error('❌ Failed to stop chat monitoring:', error);
    }
  }

  // Check for new messages in a meeting
  async checkForNewMessages(meetingId, accessToken) {
    try {
      const chatMonitor = this.chatMonitors.get(meetingId);
      if (!chatMonitor || !chatMonitor.isActive) {
        return;
      }

      // In a real implementation, this would call Teams Graph API to get new messages
      // For now, we'll simulate message detection and processing
      
      // Simulated message detection (replace with actual Graph API calls)
      const simulatedMessages = await this.simulateMessageDetection(meetingId);

      if (simulatedMessages.length > 0) {
        logger.info(`📝 Detected ${simulatedMessages.length} new messages`, { meetingId });

        // Process each message
        for (const message of simulatedMessages) {
          await this.processNewMessage(meetingId, message);
        }

        // Update monitoring stats
        chatMonitor.lastMessageCheck = new Date().toISOString();
        chatMonitor.messagesCount += simulatedMessages.length;
      }

    } catch (error) {
      logger.warn('Error checking for new messages:', error);
    }
  }

  // Process a new message (categorize and store)
  async processNewMessage(meetingId, message) {
    try {
      // Categorize the message using AI
      const category = await this.categorizeMessage(message.content);
      
      // Enhanced message object
      const enhancedMessage = {
        ...message,
        meetingId: meetingId,
        category: category,
        processedAt: new Date().toISOString(),
        aiAnalysis: {
          isQuestion: category.includes('question'),
          isActionItem: category.includes('action'),
          isDecision: category.includes('decision'),
          urgency: this.detectUrgency(message.content),
          mentions: this.extractMentions(message.content),
          sentiment: this.analyzeSentiment(message.content)
        }
      };

      // Save to Cosmos DB
      await cosmosClient.createItem('chats', enhancedMessage);

      logger.debug('✅ Message processed and stored', {
        meetingId,
        category,
        sender: message.sender
      });

    } catch (error) {
      logger.error('❌ Failed to process message:', error);
    }
  }

  // Rest of the methods remain the same...
  async categorizeMessage(content) {
    const contentLower = content.toLowerCase();
    const categories = [];

    // Question detection
    if (contentLower.includes('?') || 
        contentLower.startsWith('what') || 
        contentLower.startsWith('how') || 
        contentLower.startsWith('when') || 
        contentLower.startsWith('where') || 
        contentLower.startsWith('why') ||
        contentLower.startsWith('can we') ||
        contentLower.startsWith('should we')) {
      categories.push('question');
    }

    // Action item detection
    if (contentLower.includes('action item') ||
        contentLower.includes('todo') ||
        contentLower.includes('need to') ||
        contentLower.includes('will do') ||
        contentLower.includes('by friday') ||
        contentLower.includes('by next week') ||
        contentLower.includes('deadline') ||
        contentLower.includes('assigned to')) {
      categories.push('action_item');
    }

    // Decision detection
    if (contentLower.includes('decided') ||
        contentLower.includes('agree') ||
        contentLower.includes('approved') ||
        contentLower.includes('we will') ||
        contentLower.includes('let\'s go with') ||
        contentLower.includes('final decision')) {
      categories.push('decision');
    }

    // Link/file sharing
    if (contentLower.includes('http') || 
        contentLower.includes('www.') ||
        contentLower.includes('shared a file') ||
        contentLower.includes('attachment')) {
      categories.push('resource_sharing');
    }

    return categories.length > 0 ? categories : ['general'];
  }

  detectUrgency(content) {
    const contentLower = content.toLowerCase();
    
    if (contentLower.includes('urgent') || 
        contentLower.includes('asap') || 
        contentLower.includes('emergency') ||
        contentLower.includes('critical') ||
        contentLower.includes('immediately')) {
      return 'high';
    }
    
    if (contentLower.includes('soon') || 
        contentLower.includes('quickly') ||
        contentLower.includes('priority')) {
      return 'medium';
    }
    
    return 'low';
  }

  extractMentions(content) {
    const mentionRegex = /@(\w+)/g;
    const mentions = [];
    let match;
    
    while ((match = mentionRegex.exec(content)) !== null) {
      mentions.push(match[1]);
    }
    
    return mentions;
  }

  analyzeSentiment(content) {
    const contentLower = content.toLowerCase();
    const positiveWords = ['good', 'great', 'excellent', 'perfect', 'awesome', 'happy', 'agree'];
    const negativeWords = ['bad', 'terrible', 'wrong', 'problem', 'issue', 'concerned', 'disagree'];
    
    const positiveCount = positiveWords.filter(word => contentLower.includes(word)).length;
    const negativeCount = negativeWords.filter(word => contentLower.includes(word)).length;
    
    if (positiveCount > negativeCount) return 'positive';
    if (negativeCount > positiveCount) return 'negative';
    return 'neutral';
  }

async simulateMessageDetection(meetingId) {
  // 🚫 DISABLED: Only real messages allowed
  return [];
}

  getActiveMeetings() {
    return Array.from(this.activeMeetings.entries()).map(([meetingId, record]) => ({
      meetingId,
      ...record
    }));
  }

  getChatMonitoringStatus() {
    return Array.from(this.chatMonitors.entries()).map(([meetingId, monitor]) => ({
      meetingId,
      ...monitor,
      interval: !!monitor.interval
    }));
  }

  isAttendingMeeting(meetingId) {
    return this.activeMeetings.has(meetingId);
  }

  async getAttendanceSummary(meetingId) {
    try {
      const attendanceRecord = this.activeMeetings.get(meetingId);
      const chatMonitor = this.chatMonitors.get(meetingId);

      // Get stored messages for this meeting
      const messages = await cosmosClient.queryItems('chats',
        'SELECT * FROM c WHERE c.meetingId = @meetingId ORDER BY c.timestamp ASC',
        [{ name: '@meetingId', value: meetingId }]
      );

      // Categorize messages
      const categorizedMessages = {
        questions: messages.filter(m => m.category && m.category.includes('question')),
        actionItems: messages.filter(m => m.category && m.category.includes('action_item')),
        decisions: messages.filter(m => m.category && m.category.includes('decision')),
        resources: messages.filter(m => m.category && m.category.includes('resource_sharing')),
        general: messages.filter(m => !m.category || m.category.includes('general'))
      };

      return {
        meetingId,
        isActive: !!attendanceRecord,
        attendanceRecord,
        chatMonitoring: chatMonitor,
        messagesSummary: {
          total: messages.length,
          categorized: {
            questions: categorizedMessages.questions.length,
            actionItems: categorizedMessages.actionItems.length,
            decisions: categorizedMessages.decisions.length,
            resources: categorizedMessages.resources.length,
            general: categorizedMessages.general.length
          }
        },
        messages: categorizedMessages
      };

    } catch (error) {
      logger.error('❌ Failed to get attendance summary:', error);
      throw error;
    }
  }
}

// Create singleton instance
const meetingAttendanceService = new MeetingAttendanceService();

module.exports = meetingAttendanceService;