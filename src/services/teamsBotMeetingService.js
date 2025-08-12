// src/services/teamsBotMeetingService.js
const { ActivityHandler, MessageFactory } = require('botbuilder');
const axios = require('axios');
const logger = require('../utils/logger');

class TeamsBotMeetingService {
  constructor(adapter, botAppId, botAppPassword) {
    this.adapter = adapter;
    this.botAppId = botAppId;
    this.botAppPassword = botAppPassword;
    this.activeMeetings = new Map(); // Track active meeting sessions
  }

  // Join a Teams meeting using the bot framework
  async joinTeamsMeeting(meetingJoinUrl, meetingId) {
    try {
      logger.info('🤖 Bot attempting to join Teams meeting', {
        meetingId,
        joinUrl: meetingJoinUrl.substring(0, 50) + '...'
      });

      // For now, we'll simulate successful join since actual bot framework
      // joining requires more complex setup with proper authentication flow
      logger.info('✅ Simulated bot join (full implementation requires Teams app manifest)');
      
      // Store the "meeting session" for tracking
      this.activeMeetings.set(meetingId, {
        joinedAt: new Date().toISOString(),
        joinUrl: meetingJoinUrl,
        status: 'simulated_join'
      });

      return {
        success: true,
        message: 'Bot join simulated successfully',
        joinedAt: new Date().toISOString(),
        note: 'Full bot join requires Teams app deployment'
      };

    } catch (error) {
      logger.error('❌ Failed to join Teams meeting with bot:', error);
      throw error;
    }
  }

  // Send message to active meeting
  async sendMessageToMeeting(meetingId, message) {
    try {
      const meetingSession = this.activeMeetings.get(meetingId);
      
      if (!meetingSession) {
        throw new Error('No active meeting session found');
      }

      logger.info('📝 Simulated message send to Teams meeting', { meetingId, message });
      
      // In full implementation, this would send via bot framework
      return { 
        success: true, 
        message: 'Message simulation completed',
        simulatedMessage: message
      };

    } catch (error) {
      logger.error('❌ Failed to send message to meeting:', error);
      throw error;
    }
  }

  // Leave Teams meeting
  async leaveTeamsMeeting(meetingId) {
    try {
      const meetingSession = this.activeMeetings.get(meetingId);
      
      if (!meetingSession) {
        logger.warn('⚠️ No active meeting session to leave', { meetingId });
        return { success: true, message: 'No active session' };
      }

      // Remove from active meetings
      this.activeMeetings.delete(meetingId);

      logger.info('✅ Bot left Teams meeting (simulated)', { meetingId });
      return {
        success: true,
        message: 'Bot successfully left Teams meeting',
        leftAt: new Date().toISOString()
      };

    } catch (error) {
      logger.error('❌ Failed to leave Teams meeting:', error);
      throw error;
    }
  }

  // Get status of active meetings
  getActiveMeetings() {
    return Array.from(this.activeMeetings.entries()).map(([meetingId, session]) => ({
      meetingId,
      joinedAt: session.joinedAt,
      status: session.status
    }));
  }

  // Handle incoming messages from Teams meetings
  async handleMeetingMessage(context) {
    try {
      const messageText = context.activity.text || '';
      const conversationId = context.activity.conversation.id;
      
      logger.info('💬 Received Teams message', {
        conversationId,
        message: messageText.substring(0, 50)
      });

      // Check if message is directed at the bot
      const isDirectedAtBot = messageText.toLowerCase().includes('@ai') ||
                             messageText.toLowerCase().includes('ai assistant');

      if (isDirectedAtBot) {
        const response = await this.processAIRequest(messageText, conversationId);
        
        const replyMessage = MessageFactory.text(
          `🤖 **AI Assistant**: ${response}`
        );
        
        await context.sendActivity(replyMessage);
        
        logger.info('✅ Responded to bot mention in meeting');
      }

    } catch (error) {
      logger.error('❌ Error handling meeting message:', error);
    }
  }

  // Process AI requests from meeting participants
  async processAIRequest(message, conversationId) {
    try {
      const geminiAI = require('./geminiAI');
      
      if (!geminiAI.isAvailable()) {
        return "I'm here monitoring the meeting, but AI processing is temporarily unavailable.";
      }

      const prompt = `
        You are an AI Meeting Assistant in a live Microsoft Teams meeting. 
        A participant said: "${message}"
        
        Respond helpfully and concisely (1-2 sentences max). You can:
        - Answer questions about the meeting
        - Track action items and decisions
        - Provide meeting insights
        - Help with meeting facilitation
        
        Keep your response professional and meeting-appropriate.
      `;

      const result = await geminiAI.model.generateContent(prompt);
      const response = await result.response;
      return response.text();

    } catch (error) {
      logger.error('❌ Error processing AI request:', error);
      return "I'm here and monitoring the meeting! Feel free to ask me about action items, decisions, or key points.";
    }
  }
}

module.exports = TeamsBotMeetingService;