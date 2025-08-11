// Create this as src/bot/bot.js
const { 
  TeamsActivityHandler, 
  MessageFactory, 
  CardFactory,
  TurnContext
} = require('botbuilder');
const axios = require('axios');
const moment = require('moment');
const geminiAI = require('../services/geminiAI');
const logger = require('../utils/logger');

class Agent365Bot extends TeamsActivityHandler {
  constructor() {
    super();

    // Handle when someone messages the bot
    this.onMessage(async (context, next) => {
      try {
        const userMessage = context.activity.text?.trim();
        const userId = context.activity.from.id;
        const userName = context.activity.from.name;

        logger.info('🤖 Bot received message', { 
          message: userMessage, 
          user: userName 
        });

        // Process the user's request
        const response = await this.processUserRequest(userMessage, userId, userName, context);
        
        // Send response back to user
        await context.sendActivity(response);

      } catch (error) {
        logger.error('❌ Bot error:', error);
        await context.sendActivity('Sorry, I encountered an error. Please try again.');
      }

      await next();
    });

    // Handle when bot is added to a team
    this.onMembersAdded(async (context, next) => {
      const welcomeText = `👋 Hi! I'm **Agent 365**, your AI-powered meeting assistant!

**What I can do:**
🗓️ Create smart meetings with AI enhancements
🤖 Join meetings and monitor conversations  
📊 Provide real-time insights and summaries
📋 Extract action items and track decisions

**Try saying:**
• "Create a meeting about quarterly review tomorrow at 2 PM"
• "Schedule a brainstorming session for next Friday"
• "Help me plan a project kickoff meeting"
• "Show me my recent meetings"

Let's get started! 🚀`;

      for (const member of context.activity.membersAdded) {
        if (member.id !== context.activity.recipient.id) {
          await context.sendActivity(MessageFactory.text(welcomeText));
        }
      }
      await next();
    });
  }

  async processUserRequest(message, userId, userName, context) {
    try {
      // Use AI to understand the user's intent
      const intent = await this.analyzeUserIntent(message);
      
      logger.info('🧠 AI Intent Analysis', { intent });

      switch (intent.action) {
        case 'create_meeting':
          return await this.handleCreateMeeting(intent, userId, userName, context);
        
        case 'list_meetings':
          return await this.handleListMeetings(userId);
        
        case 'meeting_status':
          return await this.handleMeetingStatus(intent.meetingId, userId);
        
        case 'help':
          return await this.handleHelp();
        
        case 'greeting':
          return await this.handleGreeting(userName);
        
        default:
          return await this.handleGeneralQuery(message, userId);
      }

    } catch (error) {
      logger.error('❌ Error processing user request:', error);
      return MessageFactory.text('I had trouble understanding that. Could you try rephrasing?');
    }
  }

  async analyzeUserIntent(message) {
    try {
      if (!geminiAI.isAvailable()) {
        return this.basicIntentAnalysis(message);
      }

      const prompt = `
        Analyze this user message and determine their intent:
        
        Message: "${message}"
        
        Possible intents:
        - create_meeting: User wants to create/schedule a meeting
        - list_meetings: User wants to see their meetings
        - meeting_status: User asking about a specific meeting
        - help: User needs help or instructions
        - greeting: User is greeting the bot
        - general_query: Other questions about meetings/functionality
        
        Extract meeting details if present:
        - subject/topic
        - date/time
        - attendees
        - duration
        
        Respond in JSON format:
        {
          "action": "create_meeting|list_meetings|meeting_status|help|greeting|general_query",
          "confidence": 0.0-1.0,
          "meetingDetails": {
            "subject": "extracted subject",
            "dateTime": "ISO date if mentioned",
            "attendees": ["email1", "email2"],
            "duration": "minutes if mentioned",
            "description": "additional context"
          }
        }
      `;

      const result = await geminiAI.model.generateContent(prompt);
      const response = await result.response;
      const text = response.text();
      
      const analysis = JSON.parse(text.replace(/```json|```/g, '').trim());
      return analysis;

    } catch (error) {
      logger.warn('AI intent analysis failed, using basic analysis');
      return this.basicIntentAnalysis(message);
    }
  }

  basicIntentAnalysis(message) {
    const lowerMessage = message.toLowerCase();
    
    if (lowerMessage.includes('create') || lowerMessage.includes('schedule') || lowerMessage.includes('meeting')) {
      return {
        action: 'create_meeting',
        confidence: 0.8,
        meetingDetails: {
          subject: this.extractSubject(message),
          dateTime: this.extractDateTime(message),
          attendees: [],
          duration: 30,
          description: message
        }
      };
    }
    
    if (lowerMessage.includes('list') || lowerMessage.includes('show') || lowerMessage.includes('my meetings')) {
      return { action: 'list_meetings', confidence: 0.9 };
    }
    
    if (lowerMessage.includes('help') || lowerMessage.includes('what can you do')) {
      return { action: 'help', confidence: 0.9 };
    }
    
    if (lowerMessage.includes('hello') || lowerMessage.includes('hi') || lowerMessage.includes('hey')) {
      return { action: 'greeting', confidence: 0.9 };
    }
    
    return { action: 'general_query', confidence: 0.5 };
  }

  async handleCreateMeeting(intent, userId, userName, context) {
    try {
      const details = intent.meetingDetails;
      
      // Generate smart meeting details with AI
      const enhancedDetails = await this.enhanceMeetingDetails(details);
      
      // Create the meeting via your existing API
      const meetingData = {
        subject: enhancedDetails.subject,
        description: enhancedDetails.description,
        startTime: enhancedDetails.startTime,
        endTime: enhancedDetails.endTime,
        attendees: enhancedDetails.attendees,
        useAI: true,
        autoJoinAgent: true,
        enableChatCapture: true
      };

      const response = await axios.post('http://localhost:5000/api/meetings/create', meetingData);
      const meeting = response.data.meeting;

      // Create a rich Teams card to show the meeting
      const card = this.createMeetingCard(meeting, response.data);
      
      return MessageFactory.attachment(CardFactory.adaptiveCard(card));

    } catch (error) {
      logger.error('❌ Error creating meeting:', error);
      return MessageFactory.text(`Sorry, I couldn't create the meeting. Error: ${error.message}`);
    }
  }

  async enhanceMeetingDetails(details) {
    // Use AI to enhance the meeting details
    const now = moment();
    
    // Smart date/time parsing
    let startTime = details.dateTime || now.add(1, 'day').hour(14).minute(0).toISOString();
    let endTime = moment(startTime).add(details.duration || 30, 'minutes').toISOString();
    
    // Smart subject generation
    let subject = details.subject || 'AI-Generated Meeting';
    if (geminiAI.isAvailable() && details.description) {
      try {
        subject = await geminiAI.generateMeetingTitle(details.description, details.attendees);
      } catch (error) {
        logger.warn('Failed to generate AI title');
      }
    }

    return {
      subject: subject,
      description: details.description || `Meeting created via Agent 365 bot`,
      startTime: startTime,
      endTime: endTime,
      attendees: details.attendees || []
    };
  }

  createMeetingCard(meeting, responseData) {
    return {
      type: "AdaptiveCard",
      version: "1.3",
      body: [
        {
          type: "TextBlock",
          text: "✅ Meeting Created Successfully!",
          weight: "Bolder",
          size: "Medium",
          color: "Good"
        },
        {
          type: "ColumnSet",
          columns: [
            {
              type: "Column",
              width: "auto",
              items: [
                {
                  type: "TextBlock",
                  text: "📅",
                  size: "Large"
                }
              ]
            },
            {
              type: "Column",
              width: "stretch",
              items: [
                {
                  type: "TextBlock",
                  text: meeting.subject,
                  weight: "Bolder",
                  wrap: true
                },
                {
                  type: "TextBlock",
                  text: `🕐 ${moment(meeting.startTime).format('MMMM Do, YYYY [at] h:mm A')}`,
                  spacing: "Small",
                  wrap: true
                }
              ]
            }
          ]
        },
        {
          type: "FactSet",
          facts: [
            {
              title: "🤖 AI Enhanced:",
              value: meeting.aiEnhanced ? "✅ Yes" : "❌ No"
            },
            {
              title: "🟢 Real Teams:",
              value: meeting.isRealTeamsMeeting ? "✅ Yes" : "❌ No"
            },
            {
              title: "👥 Attendees:",
              value: (meeting.attendees || []).length.toString()
            },
            {
              title: "🤖 Agent:",
              value: responseData.agentConfig?.autoJoin ? "Will auto-join" : "Manual join"
            }
          ]
        },
        {
          type: "TextBlock",
          text: "💡 **AI Features Active:**\n• Smart agenda generation\n• Real-time chat monitoring\n• Automatic summary generation\n• Action item tracking",
          wrap: true,
          spacing: "Medium"
        }
      ],
      actions: [
        {
          type: "Action.OpenUrl",
          title: "🔗 Join Meeting",
          url: meeting.joinUrl || "#"
        },
        {
          type: "Action.Submit",
          title: "📊 View Status",
          data: {
            action: "meeting_status",
            meetingId: meeting.meetingId
          }
        }
      ]
    };
  }

  async handleListMeetings(userId) {
    try {
      const response = await axios.get(`http://localhost:5000/api/meetings?limit=5`);
      const meetings = response.data.meetings;

      if (meetings.length === 0) {
        return MessageFactory.text("You don't have any recent meetings. Say 'create a meeting' to get started!");
      }

      let text = "📅 **Your Recent Meetings:**\n\n";
      meetings.forEach((meeting, index) => {
        const time = moment(meeting.startTime).format('MMM DD, h:mm A');
        const status = meeting.status === 'completed' ? '✅' : 
                      meeting.status === 'in_progress' ? '🔄' : '⏳';
        text += `${status} **${meeting.subject}**\n`;
        text += `   📅 ${time} • ${meeting.agentAttended ? '🤖 Agent attended' : '👤 Manual only'}\n\n`;
      });

      return MessageFactory.text(text);

    } catch (error) {
      return MessageFactory.text("Sorry, I couldn't retrieve your meetings right now.");
    }
  }

  async handleMeetingStatus(meetingId, userId) {
    try {
      // Get meeting details and analysis
      const [meetingResponse, analysisResponse] = await Promise.all([
        axios.get(`http://localhost:5000/api/meetings/${meetingId}`),
        axios.get(`http://localhost:5000/api/meetings/${meetingId}/chat-analysis`)
      ]);

      const meeting = meetingResponse.data;
      const analysis = analysisResponse.data.analysis;

      let statusText = `📊 **Meeting Status: ${meeting.subject}**\n\n`;
      statusText += `📅 **Time:** ${moment(meeting.startTime).format('MMMM Do, YYYY [at] h:mm A')}\n`;
      statusText += `🔄 **Status:** ${meeting.status}\n`;
      statusText += `🤖 **Agent:** ${meeting.agentAttended ? 'Attended' : 'Not attended'}\n\n`;

      if (analysis && analysis.totalMessages > 0) {
        statusText += `💬 **Chat Analysis:**\n`;
        statusText += `• Total messages: ${analysis.totalMessages}\n`;
        statusText += `• Questions: ${analysis.categorizedCounts?.questions || 0}\n`;
        statusText += `• Action items: ${analysis.categorizedCounts?.actionItems || 0}\n`;
        statusText += `• Decisions: ${analysis.categorizedCounts?.decisions || 0}\n\n`;
      }

      if (meeting.hasSummary) {
        statusText += `📋 **Summary:** Available (generated with AI)\n`;
      }

      return MessageFactory.text(statusText);

    } catch (error) {
      return MessageFactory.text("Sorry, I couldn't get the meeting status right now.");
    }
  }

  async handleHelp() {
    const helpText = `🤖 **Agent 365 - Your AI Meeting Assistant**

**🗓️ Create Meetings:**
• "Create a meeting about quarterly review tomorrow at 2 PM"
• "Schedule a brainstorming session next Friday"  
• "Set up a project kickoff meeting"

**📊 View Information:**
• "Show my meetings"
• "List recent meetings"
• "What's the status of my last meeting?"

**🤖 AI Features:**
• Smart agenda generation
• Real-time chat monitoring
• Automatic summaries
• Action item tracking
• Participant analysis

**💡 Pro Tips:**
• Be specific about dates and times
• Mention attendees if you want to invite them
• I'll automatically enhance your meetings with AI!

Try asking me to create a meeting now! 🚀`;

    return MessageFactory.text(helpText);
  }

  async handleGreeting(userName) {
    const greetingText = `👋 Hi ${userName}! I'm **Agent 365**, your AI-powered meeting assistant.

Ready to create some amazing meetings? Just tell me what you need:
• "Create a meeting about [topic] on [date] at [time]"
• "Show my recent meetings"
• "Help me plan a meeting"

What would you like to do? 🚀`;

    return MessageFactory.text(greetingText);
  }

  async handleGeneralQuery(message, userId) {
    try {
      if (!geminiAI.isAvailable()) {
        return MessageFactory.text("I can help you create and manage meetings! Try saying 'create a meeting' or 'help' for more options.");
      }

      // Use AI to provide a helpful response
      const prompt = `
        You are Agent 365, an AI meeting assistant bot for Microsoft Teams. 
        A user asked: "${message}"
        
        Provide a helpful response about meetings, scheduling, or direct them to specific features.
        Keep it conversational and helpful. If they're asking about something you can't do,
        suggest what you CAN do instead.
        
        Your main capabilities:
        - Create AI-enhanced Teams meetings
        - Monitor meeting conversations  
        - Generate meeting summaries
        - Track action items
        - Provide meeting analytics
      `;

      const result = await geminiAI.model.generateContent(prompt);
      const response = await result.response;
      const aiResponse = response.text();

      return MessageFactory.text(aiResponse);

    } catch (error) {
      return MessageFactory.text("I'm here to help with meetings! Try 'create a meeting' or 'help' to see what I can do.");
    }
  }

  // Helper methods
  extractSubject(message) {
    const aboutMatch = message.match(/about (.+?)(?:\s+(?:on|at|for|tomorrow|next|this)|\s*$)/i);
    if (aboutMatch) return aboutMatch[1].trim();
    
    const forMatch = message.match(/for (.+?)(?:\s+(?:on|at|tomorrow|next|this)|\s*$)/i);
    if (forMatch) return forMatch[1].trim();
    
    return 'New Meeting';
  }

  extractDateTime(message) {
    const now = moment();
    
    // Tomorrow
    if (message.includes('tomorrow')) {
      const timeMatch = message.match(/(\d{1,2})\s*(am|pm)/i);
      if (timeMatch) {
        const hour = parseInt(timeMatch[1]);
        const ampm = timeMatch[2].toLowerCase();
        const hour24 = ampm === 'pm' && hour !== 12 ? hour + 12 : (ampm === 'am' && hour === 12 ? 0 : hour);
        return now.add(1, 'day').hour(hour24).minute(0).second(0).toISOString();
      }
      return now.add(1, 'day').hour(14).minute(0).second(0).toISOString();
    }
    
    // Next week
    if (message.includes('next week')) {
      return now.add(1, 'week').hour(14).minute(0).second(0).toISOString();
    }
    
    // Default to tomorrow 2 PM
    return now.add(1, 'day').hour(14).minute(0).second(0).toISOString();
  }
}

module.exports = Agent365Bot;