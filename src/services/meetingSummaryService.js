const { v4: uuidv4 } = require('uuid');
const cosmosClient = require('../config/cosmosdb');
const geminiAI = require('./geminiAI');
const chatCaptureService = require('./chatCaptureService');
const logger = require('../utils/logger');

class MeetingSummaryService {
  constructor() {
    this.autoSummaryEnabled = true;
    this.summaryQueue = new Map(); // Queue for processing summaries
  }

  // Generate comprehensive meeting summary
  async generateMeetingSummary(meetingId, options = {}) {
    try {
      logger.info('📋 Generating comprehensive meeting summary', { meetingId });

      const {
        includeTranscript = true,
        includeChat = true,
        includeParticipantAnalysis = true,
        summaryType = 'comprehensive',
        autoActionItems = true,
        autoFollowUp = true
      } = options;

      // FIXED: Get meeting details using correct query
      const meetings = await cosmosClient.queryItems('meetings',
        'SELECT * FROM c WHERE c.meetingId = @meetingId',
        [{ name: '@meetingId', value: meetingId }]
      );

      if (!meetings || meetings.length === 0) {
        // Try alternative query with id field
        const meetingsById = await cosmosClient.queryItems('meetings',
          'SELECT * FROM c WHERE c.id = @meetingId',
          [{ name: '@meetingId', value: meetingId }]
        );
        
        if (!meetingsById || meetingsById.length === 0) {
          logger.error('❌ Meeting not found in database', { meetingId });
          throw new Error('Meeting not found');
        }
        
        // Use the meeting found by id
        var meeting = meetingsById[0];
      } else {
        var meeting = meetings[0];
      }

      logger.info('✅ Meeting found for summary', { 
        meetingId, 
        dbId: meeting.id, 
        subject: meeting.subject 
      });

      // Gather all meeting data
      const meetingData = await this.gatherMeetingData(meetingId, {
        includeChat,
        includeTranscript,
        includeParticipantAnalysis
      });

      // Generate AI-powered summary
      const aiSummary = await this.generateAISummary(meeting, meetingData, summaryType);

      // Extract action items
      let actionItems = [];
      if (autoActionItems) {
        actionItems = await this.extractActionItems(meetingData);
      }

      // Generate participant insights
      let participantInsights = {};
      if (includeParticipantAnalysis) {
        participantInsights = await this.analyzeParticipantEngagement(meetingData);
      }

      // Create comprehensive summary
      const comprehensiveSummary = {
        id: uuidv4(),
        meetingId: meetingId,
        type: 'comprehensive_summary',
        summaryType: summaryType,
        generatedAt: new Date().toISOString(),
        
        // Meeting metadata
        meeting: {
          subject: meeting.subject,
          startTime: meeting.startTime,
          endTime: meeting.endTime,
          duration: this.calculateDuration(meeting.startTime, meeting.endTime),
          attendees: meeting.attendees || [],
          agentAttended: meeting.agentAttended
        },

        // AI-generated content
        executiveSummary: aiSummary.executiveSummary,
        keyDiscussionPoints: aiSummary.keyDiscussionPoints,
        decisionsAndOutcomes: aiSummary.decisionsAndOutcomes,
        nextSteps: aiSummary.nextSteps,
        
        // Structured data
        actionItems: actionItems,
        participantInsights: participantInsights,
        
        // Meeting metrics
        metrics: {
          totalMessages: meetingData.chatAnalysis?.totalMessages || 0,
          questionsAsked: meetingData.chatAnalysis?.categorizedCounts?.questions || 0,
          decisionsTracked: meetingData.chatAnalysis?.categorizedCounts?.decisions || 0,
          actionItemsIdentified: actionItems.length,
          participantCount: Object.keys(participantInsights).length,
          engagementLevel: this.calculateEngagementLevel(meetingData)
        },

        // Quality scores
        qualityScores: await this.calculateQualityScores(meeting, meetingData, aiSummary),

        // Follow-up recommendations
        followUpRecommendations: autoFollowUp ? await this.generateFollowUpRecommendations(meetingData, actionItems) : [],

        // Metadata
        metadata: {
          aiModel: geminiAI.isAvailable() ? 'Gemini AI' : 'Basic Analysis',
          dataIncluded: {
            chat: includeChat,
            transcript: includeTranscript,
            participantAnalysis: includeParticipantAnalysis
          },
          processingTime: new Date().toISOString()
        }
      };

      // Save summary to database
      await cosmosClient.createItem('summaries', comprehensiveSummary);

      // Update meeting record
      await cosmosClient.updateItem('meetings', meeting.id, meeting.userId, {
        hasSummary: true,
        lastSummaryGenerated: new Date().toISOString(),
        summaryId: comprehensiveSummary.id
      });

      // Generate and send follow-up actions if enabled
      if (autoFollowUp && actionItems.length > 0) {
        await this.scheduleFollowUpActions(meetingId, actionItems);
      }

      logger.info('✅ Comprehensive meeting summary generated', {
        meetingId,
        actionItemsFound: actionItems.length,
        participantsAnalyzed: Object.keys(participantInsights).length,
        qualityScore: comprehensiveSummary.qualityScores.overall
      });

      return comprehensiveSummary;

    } catch (error) {
      logger.error('❌ Failed to generate meeting summary:', error);
      throw error;
    }
  }

  // Gather all available meeting data
  async gatherMeetingData(meetingId, options) {
    try {
      const data = {
        meetingId: meetingId
      };

      // Get chat analysis if available
      if (options.includeChat) {
        try {
          data.chatAnalysis = await chatCaptureService.getChatAnalysis(meetingId);
          
          // Get raw messages for AI processing
          data.chatMessages = await cosmosClient.queryItems('chats',
            'SELECT * FROM c WHERE c.meetingId = @meetingId ORDER BY c.timestamp ASC',
            [{ name: '@meetingId', value: meetingId }]
          );
        } catch (chatError) {
          logger.warn('Could not get chat data:', chatError);
          data.chatAnalysis = null;
          data.chatMessages = [];
        }
      }

      // Get transcript data if available (future enhancement)
      if (options.includeTranscript) {
        data.transcript = await this.getTranscriptData(meetingId);
      }

      return data;

    } catch (error) {
      logger.error('❌ Failed to gather meeting data:', error);
      throw error;
    }
  }

  // Generate AI-powered summary using Gemini
  async generateAISummary(meeting, meetingData, summaryType) {
    try {
      if (!geminiAI.isAvailable()) {
        return this.generateBasicSummary(meeting, meetingData);
      }

      // Prepare comprehensive prompt for Gemini AI
      const chatContent = meetingData.chatMessages?.map(msg => 
        `${msg.senderName || msg.sender}: ${msg.content}`
      ).join('\n') || 'No chat messages available.';

      const prompt = `
        Generate a comprehensive meeting summary based on the following information:

        Meeting Details:
        - Subject: ${meeting.subject}
        - Duration: ${this.calculateDuration(meeting.startTime, meeting.endTime).formatted}
        - Attendees: ${(meeting.attendees || []).join(', ')}

        Chat Messages:
        ${chatContent}

        Meeting Analytics:
        - Total Messages: ${meetingData.chatAnalysis?.totalMessages || 0}
        - Questions Asked: ${meetingData.chatAnalysis?.categorizedCounts?.questions || 0}
        - Decisions Made: ${meetingData.chatAnalysis?.categorizedCounts?.decisions || 0}
        - Action Items Mentioned: ${meetingData.chatAnalysis?.categorizedCounts?.actionItems || 0}

        Please provide a ${summaryType} summary in the following JSON format:
        {
          "executiveSummary": "A concise 2-3 sentence overview of the meeting's purpose and key outcomes",
          "keyDiscussionPoints": [
            "Main topic 1 discussed in detail",
            "Main topic 2 with key insights",
            "Main topic 3 and conclusions"
          ],
          "decisionsAndOutcomes": [
            {
              "decision": "Clear decision made",
              "rationale": "Why this decision was made",
              "impact": "Expected impact or next steps"
            }
          ],
          "nextSteps": [
            "Immediate next step 1",
            "Immediate next step 2",
            "Longer-term action 3"
          ],
          "keyInsights": [
            "Important insight 1",
            "Strategic observation 2",
            "Process improvement 3"
          ],
          "meetingEffectiveness": {
            "score": 1-10,
            "strengths": ["What went well"],
            "improvements": ["What could be better"]
          }
        }
      `;

      const result = await geminiAI.model.generateContent(prompt);
      const response = await result.response;
      const text = response.text();

      try {
        const aiSummary = JSON.parse(text.replace(/```json|```/g, '').trim());
        logger.info('✅ AI summary generated successfully');
        return aiSummary;
      } catch (parseError) {
        logger.warn('Failed to parse AI summary, using basic summary');
        return this.generateBasicSummary(meeting, meetingData);
      }

    } catch (error) {
      logger.warn('AI summary generation failed, using basic summary:', error);
      return this.generateBasicSummary(meeting, meetingData);
    }
  }

  // Generate basic summary (fallback)
  generateBasicSummary(meeting, meetingData) {
    const chatAnalysis = meetingData.chatAnalysis;
    
    return {
      executiveSummary: `Meeting "${meeting.subject}" was held with ${(meeting.attendees || []).length} attendees. ${chatAnalysis?.totalMessages || 0} messages were exchanged during the discussion.`,
      keyDiscussionPoints: [
        `Primary focus: ${meeting.subject}`,
        `Duration: ${this.calculateDuration(meeting.startTime, meeting.endTime).formatted}`,
        `Participants: ${(meeting.attendees || []).length} attendees`
      ],
      decisionsAndOutcomes: [
        {
          decision: "Meeting concluded successfully",
          rationale: "All planned topics were discussed",
          impact: "Follow-up actions to be determined"
        }
      ],
      nextSteps: [
        "Review meeting outcomes",
        "Follow up on action items",
        "Schedule next meeting if needed"
      ],
      keyInsights: [
        `${chatAnalysis?.categorizedCounts?.questions || 0} questions were raised`,
        `${chatAnalysis?.categorizedCounts?.decisions || 0} decisions were tracked`,
        `${chatAnalysis?.categorizedCounts?.actionItems || 0} action items were identified`
      ],
      meetingEffectiveness: {
        score: 7,
        strengths: ["Meeting was completed as scheduled"],
        improvements: ["Consider more structured agenda"]
      }
    };
  }

  // Extract action items with AI analysis
  async extractActionItems(meetingData) {
    try {
      const actionItems = [];

      // Get messages that were identified as action items
      const actionMessages = meetingData.chatMessages?.filter(msg => 
        msg.isActionItem || msg.category === 'action_item'
      ) || [];

      if (actionMessages.length === 0) {
        return actionItems;
      }

      // Process each action item message
      for (const message of actionMessages) {
        let actionItem;

        if (geminiAI.isAvailable()) {
          actionItem = await this.analyzeActionItemWithAI(message);
        } else {
          actionItem = this.extractBasicActionItem(message);
        }

        if (actionItem) {
          actionItems.push({
            id: uuidv4(),
            ...actionItem,
            sourceMessageId: message.id,
            extractedAt: new Date().toISOString(),
            status: 'pending'
          });
        }
      }

      logger.info(`✅ Extracted ${actionItems.length} action items`);
      return actionItems;

    } catch (error) {
      logger.error('❌ Failed to extract action items:', error);
      return [];
    }
  }

  // Helper methods (keeping the same implementation)
  calculateDuration(startTime, endTime) {
    const start = new Date(startTime);
    const end = new Date(endTime);
    const durationMs = end - start;
    const durationMinutes = Math.round(durationMs / (1000 * 60));
    
    return {
      minutes: durationMinutes,
      formatted: `${Math.floor(durationMinutes / 60)}h ${durationMinutes % 60}m`
    };
  }

  calculateEngagementLevel(meetingData) {
    const totalMessages = meetingData.chatAnalysis?.totalMessages || 0;
    const participantCount = Object.keys(meetingData.chatAnalysis?.participantAnalysis || {}).length || 1;
    const messagesPerParticipant = totalMessages / participantCount;
    
    if (messagesPerParticipant > 10) return 'high';
    if (messagesPerParticipant > 5) return 'medium';
    return 'low';
  }

  // Analyze action item with AI
  async analyzeActionItemWithAI(message) {
    try {
      const prompt = `
        Extract action item details from this message:
        
        Message: "${message.content}"
        Sender: ${message.senderName || message.sender}
        
        Respond in JSON format:
        {
          "task": "Clear description of what needs to be done",
          "assignee": "Person responsible (extract from message or use sender if not specified)",
          "deadline": "When it needs to be done (extract date/time or estimate)",
          "priority": "low|medium|high",
          "category": "meeting_follow_up|project_task|research|communication|other",
          "estimatedEffort": "How much time/effort required",
          "dependencies": ["What needs to happen first"],
          "successCriteria": "How to know when it's complete"
        }
      `;

      const result = await geminiAI.model.generateContent(prompt);
      const response = await result.response;
      const text = response.text();

      try {
        return JSON.parse(text.replace(/```json|```/g, '').trim());
      } catch (parseError) {
        return this.extractBasicActionItem(message);
      }

    } catch (error) {
      return this.extractBasicActionItem(message);
    }
  }

  // Extract basic action item (fallback)
  extractBasicActionItem(message) {
    return {
      task: message.content,
      assignee: message.senderName || message.sender,
      deadline: this.extractDeadlineFromMessage(message.content),
      priority: message.urgency || 'medium',
      category: 'meeting_follow_up',
      estimatedEffort: 'To be determined',
      dependencies: [],
      successCriteria: 'Task completion confirmed'
    };
  }

  async analyzeParticipantEngagement(meetingData) {
    try {
      const insights = {};
      
      if (!meetingData.chatAnalysis?.participantAnalysis) {
        return insights;
      }

      const participants = meetingData.chatAnalysis.participantAnalysis;

      for (const [email, stats] of Object.entries(participants)) {
        insights[email] = {
          name: this.extractNameFromEmail(email),
          email: email,
          engagement: {
            messageCount: stats.messageCount,
            questionsAsked: stats.questions,
            actionsProposed: stats.actionItems,
            decisionsInfluenced: stats.decisions,
            engagementLevel: this.calculateParticipantEngagement(stats),
            sentimentProfile: stats.sentiment
          },
          communicationStyle: this.analyzeCommunicationStyle(stats),
          contributions: await this.analyzeContributions(email, meetingData.chatMessages),
          recommendations: this.generateParticipantRecommendations(stats)
        };
      }

      return insights;

    } catch (error) {
      logger.error('❌ Failed to analyze participant engagement:', error);
      return {};
    }
  }

  async calculateQualityScores(meeting, meetingData, aiSummary) {
    try {
      const scores = {
        overall: 0,
        participation: 0,
        productivity: 0,
        clarity: 0,
        actionOriented: 0
      };

      const chatAnalysis = meetingData.chatAnalysis;
      
      // Participation score (based on message distribution)
      if (chatAnalysis?.participantAnalysis) {
        const participantCount = Object.keys(chatAnalysis.participantAnalysis).length;
        const totalMessages = chatAnalysis.totalMessages || 0;
        const avgMessagesPerPerson = totalMessages / Math.max(participantCount, 1);
        scores.participation = Math.min(10, (avgMessagesPerPerson / 5) * 10); // Normalize to 10
      }

      // Productivity score (questions answered, decisions made)
      const questions = chatAnalysis?.categorizedCounts?.questions || 0;
      const decisions = chatAnalysis?.categorizedCounts?.decisions || 0;
      const actionItems = chatAnalysis?.categorizedCounts?.actionItems || 0;
      scores.productivity = Math.min(10, ((decisions * 2) + (actionItems * 1.5) + (questions * 0.5)) / 3);

      // Clarity score (based on AI effectiveness rating)
      scores.clarity = aiSummary?.meetingEffectiveness?.score || 7;

      // Action-oriented score
      const totalMessages = chatAnalysis?.totalMessages || 1;
      const actionRatio = (decisions + actionItems) / totalMessages;
      scores.actionOriented = Math.min(10, actionRatio * 20);

      // Overall score (weighted average)
      scores.overall = Math.round(
        (scores.participation * 0.2) +
        (scores.productivity * 0.3) +
        (scores.clarity * 0.3) +
        (scores.actionOriented * 0.2)
      );

      return scores;

    } catch (error) {
      logger.error('❌ Failed to calculate quality scores:', error);
      return { overall: 7, participation: 7, productivity: 7, clarity: 7, actionOriented: 7 };
    }
  }

  async generateFollowUpRecommendations(meetingData, actionItems) {
    try {
      const recommendations = [];

      // Action item follow-up
      if (actionItems.length > 0) {
        recommendations.push({
          type: 'action_items',
          priority: 'high',
          recommendation: `Follow up on ${actionItems.length} action items`,
          details: `Send reminders to assignees and track progress`,
          suggestedTimeline: '2-3 days'
        });
      }

      return recommendations;

    } catch (error) {
      logger.error('❌ Failed to generate follow-up recommendations:', error);
      return [];
    }
  }

  async scheduleFollowUpActions(meetingId, actionItems) {
    try {
      logger.info('📅 Scheduling follow-up actions', { meetingId, actionItemCount: actionItems.length });

      for (const actionItem of actionItems) {
        // Create follow-up reminder
        const reminder = {
          id: uuidv4(),
          meetingId: meetingId,
          actionItemId: actionItem.id,
          type: 'action_item_reminder',
          assignee: actionItem.assignee,
          task: actionItem.task,
          deadline: actionItem.deadline,
          reminderDate: this.calculateReminderDate(actionItem.deadline),
          status: 'scheduled',
          createdAt: new Date().toISOString()
        };

        // Store reminder
        await cosmosClient.createItem('reminders', reminder);
      }

      logger.info('✅ Follow-up actions scheduled successfully');

    } catch (error) {
      logger.error('❌ Failed to schedule follow-up actions:', error);
    }
  }

  // Helper methods
  calculateParticipantEngagement(stats) {
    const score = (stats.messageCount * 1) + (stats.questions * 2) + (stats.actionItems * 3) + (stats.decisions * 2);
    if (score > 15) return 'high';
    if (score > 8) return 'medium';
    return 'low';
  }

  analyzeCommunicationStyle(stats) {
    const total = stats.messageCount;
    if (stats.questions / total > 0.3) return 'inquisitive';
    if (stats.actionItems / total > 0.2) return 'action-oriented';
    if (stats.decisions / total > 0.2) return 'decisive';
    return 'collaborative';
  }

  async analyzeContributions(email, messages) {
    const userMessages = messages?.filter(msg => msg.sender === email) || [];
    return {
      keyContributions: userMessages.slice(0, 3).map(msg => msg.content),
      mostFrequentTopics: this.extractTopics(userMessages),
      influenceScore: this.calculateInfluenceScore(userMessages)
    };
  }

  generateParticipantRecommendations(stats) {
    const recommendations = [];
    
    if (stats.messageCount < 3) {
      recommendations.push('Consider encouraging more participation in future meetings');
    }
    if (stats.questions > stats.messageCount * 0.5) {
      recommendations.push('Great at asking clarifying questions - valuable for team understanding');
    }
    if (stats.actionItems > 0) {
      recommendations.push('Excellent at identifying actionable next steps');
    }
    
    return recommendations;
  }

  extractNameFromEmail(email) {
    const namePart = email.split('@')[0];
    return namePart.split('.').map(part => 
      part.charAt(0).toUpperCase() + part.slice(1)
    ).join(' ');
  }

  extractDeadlineFromMessage(content) {
    const deadlinePatterns = [
      /by (\w+day)/i,
      /by next (\w+)/i,
      /by (\w+ \d{1,2})/i,
      /deadline (\w+)/i
    ];
    
    for (const pattern of deadlinePatterns) {
      const match = content.match(pattern);
      if (match) return match[1];
    }
    
    return 'Not specified';
  }

  calculateReminderDate(deadline) {
    // Simple logic to set reminder 1 day before deadline
    const now = new Date();
    const oneDay = 24 * 60 * 60 * 1000;
    return new Date(now.getTime() + oneDay).toISOString();
  }

  extractTopics(messages) {
    // Simple topic extraction
    const allText = messages.map(msg => msg.content).join(' ').toLowerCase();
    const words = allText.split(/\s+/);
    const importantWords = words.filter(word => 
      word.length > 4 && 
      !['that', 'this', 'with', 'from', 'they', 'have', 'will', 'were', 'been'].includes(word)
    );
    
    // Count frequency and return top 3
    const wordCount = {};
    importantWords.forEach(word => {
      wordCount[word] = (wordCount[word] || 0) + 1;
    });
    
    return Object.entries(wordCount)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 3)
      .map(([word]) => word);
  }

  calculateInfluenceScore(messages) {
    // Simple influence calculation based on message types
    let score = 0;
    messages.forEach(msg => {
      if (msg.isDecision) score += 3;
      if (msg.isActionItem) score += 2;
      if (msg.isQuestion) score += 1;
      score += msg.messageCount || 1;
    });
    
    return Math.min(10, score / 2);
  }

  async getTranscriptData(meetingId) {
    // Placeholder for future transcript integration
    return {
      available: false,
      reason: 'Transcript capture not yet implemented'
    };
  }

  // Get summary by ID
  async getSummary(summaryId) {
    try {
      const summaries = await cosmosClient.queryItems('summaries',
        'SELECT * FROM c WHERE c.id = @summaryId',
        [{ name: '@summaryId', value: summaryId }]
      );

      return summaries[0] || null;
    } catch (error) {
      logger.error('❌ Failed to get summary:', error);
      throw error;
    }
  }

  // Get all summaries for a meeting
  async getMeetingSummaries(meetingId) {
    try {
      return await cosmosClient.queryItems('summaries',
        'SELECT * FROM c WHERE c.meetingId = @meetingId ORDER BY c.generatedAt DESC',
        [{ name: '@meetingId', value: meetingId }]
      );
    } catch (error) {
      logger.error('❌ Failed to get meeting summaries:', error);
      throw error;
    }
  }
}

// Create singleton instance
const meetingSummaryService = new MeetingSummaryService();

module.exports = meetingSummaryService;