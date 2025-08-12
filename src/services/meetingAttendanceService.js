const axios = require('axios');
const authService = require('./authService');
const cosmosClient = require('../config/cosmosdb');
const logger = require('../utils/logger');

class MeetingAttendanceService {
  constructor() {
    this.graphEndpoint = process.env.GRAPH_API_ENDPOINT || 'https://graph.microsoft.com/v1.0';
    this.activeMeetings = new Map(); // Track meetings the agent is attending
    this.botAppId = process.env.MICROSOFT_APP_ID;
  }

  // Join meeting as AI agent (POC Feature 1.4)
  async joinMeeting(meetingId, userId) {
    try {
      logger.info('🤖 AI Agent joining meeting', { meetingId, userId });

      // Get meeting details from database
      const meetings = await cosmosClient.queryItems('meetings',
        'SELECT * FROM c WHERE c.meetingId = @meetingId',
        [{ name: '@meetingId', value: meetingId }]
      );

      if (!meetings || meetings.length === 0) {
        throw new Error('Meeting not found');
      }

      const meeting = meetings[0];

      // Check meeting timing
      const now = new Date();
      const startTime = new Date(meeting.startTime);
      const endTime = new Date(meeting.endTime);
      const bufferTime = 15 * 60 * 1000; // 15 minutes buffer

      if (now < startTime.getTime() - bufferTime) {
        throw new Error('Meeting has not started yet (joins 15 minutes before start time)');
      }

      if (now > endTime) {
        throw new Error('Meeting has already ended');
      }

      // Join meeting using Microsoft Graph API
      if (authService.isAvailable()) {
        await this.joinMeetingViaGraph(meeting);
      }

      // Create attendance record
      const attendanceRecord = {
        meetingId: meetingId,
        agentJoinedAt: now.toISOString(),
        agentStatus: 'attending',
        monitoringStarted: now.toISOString(),
        chatCaptureEnabled: true,
        transcriptCaptureEnabled: true,
        agentVisible: true,
        agentName: 'AI Meeting Assistant'
      };

      this.activeMeetings.set(meetingId, attendanceRecord);

      // Update meeting record
      await cosmosClient.updateItem('meetings', meeting.id, userId, {
        agentAttended: true,
        agentJoinedAt: now.toISOString(),
        agentStatus: 'attending',
        agentVisible: true
      });

      // Start chat monitoring and transcript capture
      await this.startMeetingMonitoring(meetingId, meeting);

      logger.info('✅ AI Agent successfully joined meeting', {
        meetingId,
        joinedAt: attendanceRecord.agentJoinedAt,
        agentVisible: true
      });

      return {
        success: true,
        message: 'AI Agent joined meeting and is now visible to participants',
        attendanceRecord: attendanceRecord,
        capabilities: {
          chatMonitoring: true,
          transcriptCapture: true,
          realTimeAnalysis: true,
          interactiveChat: true
        }
      };

    } catch (error) {
      logger.error('❌ Failed to join meeting:', error);
      throw error;
    }
  }

  // Join meeting via Microsoft Graph API (makes agent visible)
  async joinMeetingViaGraph(meeting) {
    try {
      if (!authService.isAvailable()) {
        logger.warn('⚠️ Graph API not available, using simulated join');
        return;
      }

      const accessToken = await authService.getAppOnlyToken();

      // Add AI agent as a meeting participant
      if (meeting.graphEventId) {
        const agentAttendee = {
          emailAddress: {
            address: `ai-agent@${process.env.AZURE_TENANT_ID || 'company.com'}`,
            name: 'AI Meeting Assistant'
          },
          type: 'required'
        };

        // Update meeting to include AI agent as attendee
        const updatePayload = {
          attendees: [
            ...(meeting.attendees || []).map(email => ({
              emailAddress: { address: email, name: email.split('@')[0] },
              type: 'required'
            })),
            agentAttendee
          ]
        };

        const usersResponse = await axios.get(
          `${this.graphEndpoint}/users?$top=1&$select=id`,
          { headers: { 'Authorization': `Bearer ${accessToken}` } }
        );

        if (usersResponse.data.value && usersResponse.data.value.length > 0) {
          const userId = usersResponse.data.value[0].id;
          
          await axios.patch(
            `${this.graphEndpoint}/users/${userId}/events/${meeting.graphEventId}`,
            updatePayload,
            {
              headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
              }
            }
          );

          logger.info('✅ AI Agent added to meeting attendee list');
        }
      }

    } catch (error) {
      logger.warn('⚠️ Could not add agent via Graph API:', error.message);
    }
  }

  // Start meeting monitoring (POC Feature 1.4)
  async startMeetingMonitoring(meetingId, meeting) {
    try {
      logger.info('🔍 Starting meeting monitoring', { meetingId });

      // Initialize monitoring systems
      const chatCaptureService = require('./chatCaptureService');
      
      // Start chat capture
      await chatCaptureService.initiateAutomaticCapture(meeting);

      // Send welcome message to meeting
      await this.sendWelcomeMessage(meetingId);

      logger.info('✅ Meeting monitoring started');

    } catch (error) {
      logger.error('❌ Failed to start monitoring:', error);
    }
  }

  // Send welcome message when AI agent joins
  async sendWelcomeMessage(meetingId) {
    try {
      const chatCaptureService = require('./chatCaptureService');
      
      const welcomeMessage = `🤖 **AI Meeting Assistant has joined**\n\n` +
        `✅ Now monitoring meeting for insights\n` +
        `💬 You can chat with me during the meeting\n` +
        `📋 I'll generate a summary at the end\n` +
        `🚨 I'll track action items and decisions\n\n` +
        `*Type @AI or mention me to interact directly*`;

      await chatCaptureService.sendToMeetingChat(meetingId, welcomeMessage);
      
    } catch (error) {
      logger.warn('⚠️ Could not send welcome message:', error.message);
    }
  }

  // Leave meeting
  async leaveMeeting(meetingId, userId) {
    try {
      logger.info('🤖 AI Agent leaving meeting', { meetingId });

      const attendanceRecord = this.activeMeetings.get(meetingId);
      if (!attendanceRecord) {
        throw new Error('Agent is not currently attending this meeting');
      }

      const now = new Date();
      attendanceRecord.agentLeftAt = now.toISOString();
      attendanceRecord.agentStatus = 'left';

      // Stop monitoring
      const chatCaptureService = require('./chatCaptureService');
      await chatCaptureService.stopChatCapture(meetingId);

      // Generate final summary
      await this.generateFinalSummary(meetingId);

      // Update meeting record
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

      this.activeMeetings.delete(meetingId);

      logger.info('✅ AI Agent left meeting successfully');

      return {
        success: true,
        message: 'AI Agent left meeting successfully',
        attendanceRecord: attendanceRecord
      };

    } catch (error) {
      logger.error('❌ Failed to leave meeting:', error);
      throw error;
    }
  }

  // Generate final summary when leaving
  async generateFinalSummary(meetingId) {
    try {
      const meetingSummaryService = require('./meetingSummaryService');
      const chatCaptureService = require('./chatCaptureService');
      
      logger.info('📋 Generating final meeting summary', { meetingId });

      // Generate comprehensive summary
      const summary = await meetingSummaryService.generateMeetingSummary(meetingId);

      // Send summary to meeting chat
      const summaryMessage = `📋 **Meeting Summary Generated**\n\n` +
        `📝 **Key Points:** ${summary.executiveSummary}\n\n` +
        `🎯 **Action Items:** ${summary.actionItems.length} identified\n` +
        `✅ **Decisions:** ${summary.metrics.decisionsTracked}\n` +
        `❓ **Questions:** ${summary.metrics.questionsAsked}\n` +
        `💬 **Total Messages:** ${summary.metrics.totalMessages}\n\n` +
        `📈 **Meeting Quality:** ${summary.qualityScores.overall}/10\n\n` +
        `*Full detailed summary available in dashboard*`;

      await chatCaptureService.sendToMeetingChat(meetingId, summaryMessage);

    } catch (error) {
      logger.warn('⚠️ Could not generate final summary:', error.message);
    }
  }

  // Get attendance summary
  async getAttendanceSummary(meetingId) {
    try {
      const attendanceRecord = this.activeMeetings.get(meetingId);
      const chatCaptureService = require('./chatCaptureService');

      // Get chat analysis
      const chatAnalysis = await chatCaptureService.getChatAnalysis(meetingId);

      return {
        meetingId,
        isActive: !!attendanceRecord,
        attendanceRecord,
        chatAnalysis,
        agentCapabilities: {
          chatMonitoring: true,
          transcriptCapture: true,
          realTimeInsights: true,
          interactiveChat: true,
          summaryGeneration: true
        }
      };

    } catch (error) {
      logger.error('❌ Failed to get attendance summary:', error);
      throw error;
    }
  }

  // Get all active meetings
  getActiveMeetings() {
    return Array.from(this.activeMeetings.entries()).map(([meetingId, record]) => ({
      meetingId,
      ...record,
      agentVisible: true
    }));
  }

  // Check if agent is attending a specific meeting
  isAttendingMeeting(meetingId) {
    return this.activeMeetings.has(meetingId);
  }
}

// Create singleton instance
const meetingAttendanceService = new MeetingAttendanceService();
module.exports = meetingAttendanceService;