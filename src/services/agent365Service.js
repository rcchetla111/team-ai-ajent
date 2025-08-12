// src/services/agent365Service.js - Microsoft 365 Agent SDK Integration
const { Agent365Client, MeetingAgent } = require('@microsoft/agent365-sdk');
const authService = require('./authService');
const cosmosClient = require('../config/cosmosdb');
const geminiAI = require('./geminiAI');
const logger = require('../utils/logger');

class Agent365Service {
  constructor() {
    this.client = null;
    this.meetingAgent = null;
    this.activeMeetings = new Map();
    this.isInitialized = false;
  }

  // Initialize the Agent 365 SDK
  async initialize() {
    try {
      if (this.isInitialized) {
        return;
      }

      logger.info('🤖 Initializing Microsoft 365 Agent SDK');

      // Get access token for Agent SDK
      const accessToken = await authService.getAppOnlyToken();

      // Initialize Agent 365 client
      this.client = new Agent365Client({
        tenantId: process.env.AZURE_TENANT_ID,
        clientId: process.env.AZURE_CLIENT_ID,
        clientSecret: process.env.AZURE_CLIENT_SECRET,
        accessToken: accessToken,
        agentConfig: {
          name: 'AI Meeting Assistant',
          description: 'Intelligent meeting assistant for Teams',
          capabilities: ['meetings', 'chat', 'summarization', 'actionItems'],
          version: '1.0.0'
        }
      });

      // Initialize Meeting Agent
      this.meetingAgent = new MeetingAgent(this.client, {
        autoJoin: true,
        autoTranscribe: true,
        autoSummarize: true,
        realTimeAnalysis: true,
        chatParticipation: true
      });

      // Set up event handlers
      this.setupEventHandlers();

      this.isInitialized = true;
      logger.info('✅ Microsoft 365 Agent SDK initialized successfully');

    } catch (error) {
      logger.error('❌ Failed to initialize Agent 365 SDK:', error);
      throw error;
    }
  }

  // Set up event handlers for the meeting agent
  setupEventHandlers() {
    // Meeting join event
    this.meetingAgent.on('meetingJoined', async (meetingInfo) => {
      logger.info('🎉 Agent successfully joined Teams meeting', {
        meetingId: meetingInfo.id,
        subject: meetingInfo.subject,
        participantCount: meetingInfo.participants.length
      });

      // Store meeting session
      this.activeMeetings.set(meetingInfo.id, {
        meetingInfo,
        joinedAt: new Date().toISOString(),
        status: 'active',
        participants: meetingInfo.participants,
        messages: []
      });

      // Send welcome message
      await this.meetingAgent.sendMessage(meetingInfo.id, {
        type: 'text',
        content: '🤖 **AI Meeting Assistant** has joined the meeting!\n\n' +
                'I\'m here to help with:\n' +
                '• Real-time meeting notes\n' +
                '• Action item tracking\n' +
                '• Q&A assistance\n' +
                '• Meeting summarization\n\n' +
                'Just mention me (@AI Assistant) for assistance!'
      });

      // Update database
      await this.updateMeetingStatus(meetingInfo.id, 'agent_joined');
    });

    // Meeting message event
    this.meetingAgent.on('messageReceived', async (message) => {
      logger.info('💬 Received meeting message', {
        meetingId: message.meetingId,
        sender: message.sender.name,
        isDirectedAtAgent: message.mentionsAgent
      });

      // Store message
      const meetingSession = this.activeMeetings.get(message.meetingId);
      if (meetingSession) {
        meetingSession.messages.push(message);
      }

      // Store in database for analysis
      await this.storeMeetingMessage(message);

      // Respond if agent is mentioned
      if (message.mentionsAgent) {
        await this.handleAgentMention(message);
      }

      // Analyze for action items and key points
      await this.analyzeMessage(message);
    });

    // Meeting participant joined/left events
    this.meetingAgent.on('participantJoined', async (participant, meetingId) => {
      logger.info('👤 Participant joined meeting', {
        meetingId,
        participant: participant.name
      });

      const meetingSession = this.activeMeetings.get(meetingId);
      if (meetingSession) {
        meetingSession.participants.push(participant);
      }
    });

    this.meetingAgent.on('participantLeft', async (participant, meetingId) => {
      logger.info('👋 Participant left meeting', {
        meetingId,
        participant: participant.name
      });

      const meetingSession = this.activeMeetings.get(meetingId);
      if (meetingSession) {
        meetingSession.participants = meetingSession.participants.filter(
          p => p.id !== participant.id
        );
      }
    });

    // Meeting ended event
    this.meetingAgent.on('meetingEnded', async (meetingInfo) => {
      logger.info('📊 Meeting ended, generating summary', {
        meetingId: meetingInfo.id
      });

      await this.handleMeetingEnd(meetingInfo);
    });

    // Transcription events
    this.meetingAgent.on('transcriptionReceived', async (transcription) => {
      logger.info('🎤 Received meeting transcription', {
        meetingId: transcription.meetingId,
        speaker: transcription.speaker,
        duration: transcription.duration
      });

      await this.processTranscription(transcription);
    });
  }

  // Join a Teams meeting using Agent SDK
  async joinMeeting(meetingId, meetingJoinUrl) {
    try {
      if (!this.isInitialized) {
        await this.initialize();
      }

      logger.info('🤖 Joining Teams meeting via Agent 365 SDK', { meetingId });

      // Use Agent SDK to join meeting
      const joinResult = await this.meetingAgent.joinMeeting({
        meetingId: meetingId,
        joinUrl: meetingJoinUrl,
        agentConfig: {
          displayName: 'AI Meeting Assistant',
          role: 'assistant',
          capabilities: ['chat', 'transcription', 'analysis'],
          visibility: 'visible' // Agent appears in participant list
        }
      });

      logger.info('✅ Successfully joined meeting via Agent SDK', {
        meetingId,
        agentId: joinResult.agentId
      });

      return {
        success: true,
        method: 'agent365_sdk',
        agentId: joinResult.agentId,
        joinedAt: new Date().toISOString(),
        capabilities: joinResult.capabilities,
        visibleToParticipants: true
      };

    } catch (error) {
      logger.error('❌ Failed to join meeting via Agent SDK:', error);
      throw error;
    }
  }

  // Leave a Teams meeting
  async leaveMeeting(meetingId) {
    try {
      logger.info('🚪 Leaving Teams meeting via Agent SDK', { meetingId });

      // Generate final summary before leaving
      await this.generateMeetingSummary(meetingId);

      // Send goodbye message
      await this.meetingAgent.sendMessage(meetingId, {
        type: 'text',
        content: '🤖 **AI Meeting Assistant** is leaving the meeting.\n\n' +
                '📋 Meeting summary and action items will be sent shortly.\n' +
                'Thank you for letting me assist with your meeting!'
      });

      // Leave the meeting
      const leaveResult = await this.meetingAgent.leaveMeeting(meetingId);

      // Clean up local state
      this.activeMeetings.delete(meetingId);

      // Update database
      await this.updateMeetingStatus(meetingId, 'agent_left');

      logger.info('✅ Successfully left meeting via Agent SDK', { meetingId });

      return {
        success: true,
        leftAt: new Date().toISOString(),
        finalSummaryGenerated: true
      };

    } catch (error) {
      logger.error('❌ Failed to leave meeting via Agent SDK:', error);
      throw error;
    }
  }

  // Handle when agent is mentioned in chat
  async handleAgentMention(message) {
    try {
      logger.info('🎯 Processing agent mention', {
        meetingId: message.meetingId,
        sender: message.sender.name
      });

      // Generate AI response using Gemini
      let response = 'I\'m here and monitoring the meeting!';
      
      if (geminiAI.isAvailable()) {
        const prompt = `
          You are an AI Meeting Assistant in a live Teams meeting.
          Participant "${message.sender.name}" said: "${message.content}"
          
          Respond helpfully as a meeting assistant (1-2 sentences max):
          - Answer questions about the meeting
          - Track action items and decisions  
          - Provide meeting insights
          - Help with facilitation
          
          Be professional and concise.
        `;

        try {
          const result = await geminiAI.model.generateContent(prompt);
          const aiResponse = await result.response;
          response = aiResponse.text();
        } catch (aiError) {
          logger.warn('AI response generation failed, using fallback');
        }
      }

      // Send response in meeting
      await this.meetingAgent.sendMessage(message.meetingId, {
        type: 'text',
        content: `@${message.sender.name} ${response}`,
        replyToMessageId: message.id
      });

      logger.info('✅ Responded to agent mention');

    } catch (error) {
      logger.error('❌ Failed to handle agent mention:', error);
    }
  }

  // Analyze messages for action items and key points
  async analyzeMessage(message) {
    try {
      if (!geminiAI.isAvailable()) {
        return;
      }

      const prompt = `
        Analyze this meeting message for key information:
        
        Message: "${message.content}"
        Speaker: ${message.sender.name}
        
        Determine:
        1. Is this an action item? (true/false)
        2. Is this a decision? (true/false)  
        3. Is this a question? (true/false)
        4. Urgency level (low/medium/high)
        5. Key topics mentioned
        
        Respond in JSON format.
      `;

      const result = await geminiAI.model.generateContent(prompt);
      const response = await result.response;
      const analysis = JSON.parse(response.text());

      // Store analysis
      await this.storeMessageAnalysis(message.id, analysis);

      // Handle action items
      if (analysis.isActionItem) {
        await this.handleActionItem(message, analysis);
      }

      // Handle important decisions
      if (analysis.isDecision && analysis.urgency === 'high') {
        await this.highlightDecision(message);
      }

    } catch (error) {
      logger.error('❌ Message analysis failed:', error);
    }
  }

  // Process meeting transcription
  async processTranscription(transcription) {
    try {
      // Store transcription in database
      await cosmosClient.createItem('transcriptions', {
        id: require('uuid').v4(),
        meetingId: transcription.meetingId,
        speaker: transcription.speaker,
        content: transcription.text,
        timestamp: transcription.timestamp,
        confidence: transcription.confidence,
        language: transcription.language
      });

      // Analyze transcription for insights
      if (geminiAI.isAvailable()) {
        await this.analyzeTranscription(transcription);
      }

    } catch (error) {
      logger.error('❌ Failed to process transcription:', error);
    }
  }

  // Generate meeting summary when meeting ends
  async generateMeetingSummary(meetingId) {
    try {
      const meetingSession = this.activeMeetings.get(meetingId);
      if (!meetingSession) {
        logger.warn('No meeting session found for summary generation');
        return;
      }

      logger.info('📋 Generating comprehensive meeting summary via Agent SDK');

      // Use Agent SDK's built-in summarization
      const summary = await this.meetingAgent.generateSummary(meetingId, {
        includeTranscript: true,
        includeChat: true,
        includeActionItems: true,
        includeDecisions: true,
        includeParticipantInsights: true
      });

      // Store summary in database
      await cosmosClient.createItem('summaries', {
        id: require('uuid').v4(),
        meetingId: meetingId,
        summary: summary,
        generatedAt: new Date().toISOString(),
        method: 'agent365_sdk',
        participants: meetingSession.participants.length,
        messageCount: meetingSession.messages.length
      });

      // Send summary to meeting participants via email
      await this.sendSummaryEmail(meetingId, summary);

      logger.info('✅ Meeting summary generated and sent');

      return summary;

    } catch (error) {
      logger.error('❌ Failed to generate meeting summary:', error);
      throw error;
    }
  }

  // Store meeting message in database
  async storeMeetingMessage(message) {
    try {
      await cosmosClient.createItem('chats', {
        id: message.id,
        meetingId: message.meetingId,
        sender: message.sender.name,
        senderId: message.sender.id,
        content: message.content,
        timestamp: message.timestamp,
        messageType: message.type,
        mentionsAgent: message.mentionsAgent,
        capturedViaAgent365: true
      });
    } catch (error) {
      logger.error('❌ Failed to store meeting message:', error);
    }
  }

  // Update meeting status in database
  async updateMeetingStatus(meetingId, status) {
    try {
      const meetings = await cosmosClient.queryItems('meetings',
        'SELECT * FROM c WHERE c.meetingId = @meetingId',
        [{ name: '@meetingId', value: meetingId }]
      );

      if (meetings.length > 0) {
        const meeting = meetings[0];
        await cosmosClient.updateItem('meetings', meeting.id, meeting.userId, {
          agentStatus: status,
          agentJoinedAt: status === 'agent_joined' ? new Date().toISOString() : meeting.agentJoinedAt,
          agentLeftAt: status === 'agent_left' ? new Date().toISOString() : null,
          agentAttended: true
        });
      }
    } catch (error) {
      logger.error('❌ Failed to update meeting status:', error);
    }
  }

  // Get service status
  getStatus() {
    return {
      initialized: this.isInitialized,
      activeMeetings: this.activeMeetings.size,
      sdkVersion: '1.0.0',
      capabilities: {
        realTimeMeetingJoin: true,
        liveTranscription: true,
        chatParticipation: true,
        automaticSummarization: true,
        actionItemTracking: true,
        participantAnalysis: true,
        visibleParticipation: true
      },
      message: this.isInitialized 
        ? 'Agent 365 SDK ready for meeting participation'
        : 'Agent 365 SDK not initialized'
    };
  }

  // Get active meetings
  getActiveMeetings() {
    return Array.from(this.activeMeetings.entries()).map(([meetingId, session]) => ({
      meetingId,
      subject: session.meetingInfo.subject,
      joinedAt: session.joinedAt,
      participantCount: session.participants.length,
      messageCount: session.messages.length,
      status: session.status
    }));
  }
}

// Create singleton instance
const agent365Service = new Agent365Service();

module.exports = agent365Service;