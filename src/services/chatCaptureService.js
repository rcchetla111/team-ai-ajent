const axios = require("axios");
const authService = require("./authService");
const cosmosClient = require("../config/cosmosdb");
const geminiAI = require("./geminiAI");
const logger = require("../utils/logger");

class ChatCaptureService {
  constructor() {
    this.graphEndpoint =
      process.env.GRAPH_API_ENDPOINT || "https://graph.microsoft.com/v1.0";
    this.activeCaptures = new Map(); // Track active chat captures


    this.autoInsightsEnabled = true;
    this.insightCounters = new Map();
    this.meetingInsightTimers = new Map();

  }

  // --- MISSING FUNCTION ADDED ---
  async findChatIdWithRetries(graphEventId, maxRetries = 3) {
    if (!authService.isAvailable()) {
      throw new Error("Teams service not available for chat capture");
    }

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        logger.info(
          `🔍 Attempting to find chat ID for meeting (attempt ${attempt}/${maxRetries})`
        );

        const accessToken = await authService.getAppOnlyToken();
        const meetingEvent = await this.getMeetingEvent(
          graphEventId,
          accessToken
        );

        // Try to find associated chat/call
        if (
          meetingEvent.onlineMeeting &&
          meetingEvent.onlineMeeting.joinWebUrl
        ) {
          // Extract potential chat ID from join URL or related data
          const chatId = await this.extractChatIdFromMeeting(
            meetingEvent,
            accessToken
          );
          if (chatId) {
            logger.info(`✅ Found chat ID: ${chatId}`);
            return chatId;
          }
        }

        // If no chat found, wait and retry
        if (attempt < maxRetries) {
          logger.warn(`⚠️ Chat not found, waiting 30 seconds before retry...`);
          await new Promise((resolve) => setTimeout(resolve, 30000));
        }
      } catch (error) {
        logger.warn(`⚠️ Attempt ${attempt} failed:`, error.message);
        if (attempt === maxRetries) {
          throw new Error(
            `Failed to find chat ID after ${maxRetries} attempts: ${error.message}`
          );
        }
        await new Promise((resolve) => setTimeout(resolve, 15000));
      }
    }

    throw new Error("Could not find chat ID for meeting");
  }

  // --- MISSING FUNCTION ADDED ---
async extractChatIdFromMeeting(meetingEvent, accessToken) {
  try {
    logger.info('🔍 REAL FIX: Searching for meeting chat with multiple methods');
    
    // Method 1: Get chats associated with the calendar event
    const eventId = meetingEvent.id;
    const onlineMeetingId = meetingEvent.onlineMeeting?.id;
    
    if (onlineMeetingId) {
      // Try to get chat from online meeting
      const onlineMeetingUrl = `${this.graphEndpoint}/me/onlineMeetings/${onlineMeetingId}`;
      const onlineMeetingResponse = await axios.get(onlineMeetingUrl, {
        headers: { 'Authorization': `Bearer ${accessToken}` }
      });
      
      if (onlineMeetingResponse.data.chatInfo?.threadId) {
        logger.info('✅ REAL FIX: Found chat ID from online meeting');
        return onlineMeetingResponse.data.chatInfo.threadId;
      }
    }

    // Method 2: Search all chats for meeting-related ones
    const chatsUrl = `${this.graphEndpoint}/chats?$filter=chatType eq 'meeting'&$expand=members&$top=50`;
    const chatsResponse = await axios.get(chatsUrl, {
      headers: { 'Authorization': `Bearer ${accessToken}` }
    });
    
    if (chatsResponse.data.value && chatsResponse.data.value.length > 0) {
      // Look for chat that matches meeting time or subject
      const meetingStart = new Date(meetingEvent.start.dateTime);
      const meetingSubject = meetingEvent.subject.toLowerCase();
      
      for (const chat of chatsResponse.data.value) {
        // Check if chat was created around meeting time (within 1 hour)
        if (chat.createdDateTime) {
          const chatCreated = new Date(chat.createdDateTime);
          const timeDiff = Math.abs(meetingStart - chatCreated) / (1000 * 60); // minutes
          
          if (timeDiff <= 60) { // Within 1 hour
            logger.info('✅ REAL FIX: Found chat by time proximity', { 
              chatId: chat.id, 
              timeDiff: timeDiff + ' minutes' 
            });
            return chat.id;
          }
        }
        
        // Check if chat topic matches meeting subject
        if (chat.topic && chat.topic.toLowerCase().includes(meetingSubject.substring(0, 10))) {
          logger.info('✅ REAL FIX: Found chat by subject match', { chatId: chat.id });
          return chat.id;
        }
      }
      
      // Fallback: return the most recent meeting chat
      const sortedChats = chatsResponse.data.value.sort((a, b) => 
        new Date(b.createdDateTime) - new Date(a.createdDateTime)
      );
      
      if (sortedChats[0]) {
        logger.info('✅ REAL FIX: Using most recent meeting chat as fallback', { 
          chatId: sortedChats[0].id 
        });
        return sortedChats[0].id;
      }
    }

    // Method 3: Try direct calendar event chat
    const calendarChatUrl = `${this.graphEndpoint}/me/events/${eventId}/instances?$select=onlineMeeting`;
    try {
      const calendarResponse = await axios.get(calendarChatUrl, {
        headers: { 'Authorization': `Bearer ${accessToken}` }
      });
      
      if (calendarResponse.data.value && calendarResponse.data.value[0]?.onlineMeeting?.chatInfo?.threadId) {
        logger.info('✅ REAL FIX: Found chat from calendar event');
        return calendarResponse.data.value[0].onlineMeeting.chatInfo.threadId;
      }
    } catch (calendarError) {
      logger.warn('⚠️ Calendar event chat lookup failed:', calendarError.message);
    }
    
    logger.error('❌ REAL FIX: Could not find chat ID with any method');
    return null;
    
  } catch (error) {
    logger.error('❌ REAL FIX: Error extracting chat ID:', error.message);
    return null;
  }
}

  // --- FIXED AUTO CAPTURE FUNCTION ---
  async initiateAutomaticCapture(meeting) {
    logger.info(
      `🤖 Initiating automatic capture for meeting: ${meeting.subject}`
    );

    return [];
    try {
      if (!meeting.graphEventId) {
        logger.warn("⚠️ No graph event ID, starting simulated capture");
        await this.startSimulatedChatCapture(meeting.meetingId, meeting);
        return;
      }

      const chatId = await this.findChatIdWithRetries(meeting.graphEventId);
      await this.startChatCapture(meeting.meetingId, meeting, chatId);
    } catch (error) {
      logger.error(
        `🚨 Failed to start automatic capture for meeting ${meeting.id}, falling back to simulated:`,
        error.message
      );
      // Fallback to simulated capture
      await this.startSimulatedChatCapture(meeting.meetingId, meeting);
    }
  }

  // --- IMPROVED SIMULATED CAPTURE ---
  async startSimulatedChatCapture(meetingId, meeting) {
    try {
      logger.info("🔄 Starting simulated chat capture", { meetingId });

      const captureSession = {
        meetingId: meetingId,
        chatId: null,
        startTime: new Date().toISOString(),
        lastCaptureTime: new Date().toISOString(),
        isActive: true,
        simulated: true,
        messageCount: 0,
      };

      this.activeCaptures.set(meetingId, captureSession);

      // Start simulated monitoring
      const monitoringLoop = setInterval(async () => {
        const session = this.activeCaptures.get(meetingId);
        if (!session || !session.isActive) {
          clearInterval(monitoringLoop);
          return;
        }
        await this.simulateMessageCapture(session);
      }, 45000); // Every 45 seconds

      captureSession.monitoringLoop = monitoringLoop;
      logger.info("✅ Simulated chat capture started successfully", {
        meetingId,
      });

      return captureSession;
    } catch (error) {
      logger.error("❌ Failed to start simulated chat capture:", error);
      throw error;
    }
  }

  // --- NEW SIMULATED MESSAGE CAPTURE ---
  async simulateMessageCapture(session) {
    try {
      // Generate realistic simulated messages
      const simulatedMessages = this.generateSimulatedMessages(
        session.meetingId
      );

      if (simulatedMessages.length > 0) {
        logger.info(
          `💬 Simulated ${simulatedMessages.length} message(s) for meeting ${session.meetingId}`
        );

        for (const message of simulatedMessages) {
          await this.processMessageWithAI(session.meetingId, message);
          session.messageCount++;
        }

        session.lastCaptureTime = new Date().toISOString();
      }
    } catch (error) {
      logger.warn("⚠️ Error in simulated message capture:", error.message);
    }
  }

  // --- GENERATE REALISTIC SIMULATED MESSAGES ---
// --- GENERATE REALISTIC SIMULATED MESSAGES (DISABLED) ---
generateSimulatedMessages(meetingId) {
  // 🚫 DISABLED: Only real messages allowed
  return [];
}

  async startChatCapture(meetingId, meeting, chatId) {
    const captureSession = {
      meetingId: meetingId,
      chatId: chatId,
      startTime: new Date().toISOString(),
      lastCaptureTime: new Date().toISOString(),
      isActive: true,
      messageCount: 0,
    };
    this.activeCaptures.set(meetingId, captureSession);

    const monitoringLoop = setInterval(async () => {
      const session = this.activeCaptures.get(meetingId);
      if (!session || !session.isActive) {
        clearInterval(monitoringLoop);
        return;
      }
      await this.captureNewMessages(session);
    }, 15000);

    captureSession.monitoringLoop = monitoringLoop;
    logger.info("✅ LIVE chat capture started successfully", {
      meetingId,
      chatId,
    });
  }

  // Helper to get the meeting event to find the chatId
  async getMeetingEvent(graphEventId, accessToken) {
    const usersResponse = await axios.get(
      `${this.graphEndpoint}/users?$top=1&$select=id`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    if (!usersResponse.data.value || usersResponse.data.value.length === 0)
      throw new Error("No users found in tenant");
    const userId = usersResponse.data.value[0].id;

    const url = `${this.graphEndpoint}/users/${userId}/events/${graphEventId}`;
    const response = await axios.get(url, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    return response.data;
  }

  // Capture new messages from Teams
async captureNewMessages(session) {
  try {
    const accessToken = await authService.getAppOnlyToken();
    const lastCheckTime = session.lastCaptureTime;
    session.lastCaptureTime = new Date().toISOString();

    // REAL FIX: Better message filtering and error handling
    const url = `${this.graphEndpoint}/chats/${session.chatId}/messages?$top=50&$orderby=createdDateTime desc`;
    
    logger.info('🔍 REAL FIX: Checking for new messages', { 
      chatId: session.chatId,
      lastCheck: lastCheckTime 
    });

    const response = await axios.get(url, {
      headers: { 
        'Authorization': `Bearer ${accessToken}`,
        'ConsistencyLevel': 'eventual'
      },
    });
    
    const allMessages = response.data.value || [];
    
    // Filter for new messages since last check
    const newMessages = allMessages.filter(msg => {
      const msgTime = new Date(msg.createdDateTime);
      const lastCheck = new Date(lastCheckTime);
      return msgTime > lastCheck;
    });

    if (newMessages && newMessages.length > 0) {
      logger.info(`💬 REAL FIX: Found ${newMessages.length} NEW real messages!`, { 
        meetingId: session.meetingId 
      });
      
      for (const message of newMessages) {
        // Only process real user messages (not system/bot messages)
        if (message.from?.user && 
            !message.from.application && 
            message.body?.content && 
            message.body.content.trim() !== '') {
          
          logger.info('📝 REAL FIX: Processing real user message', { 
            sender: message.from.user.displayName,
            content: message.body.content.substring(0, 50) + '...'
          });
          
          await this.processMessageWithAI(session.meetingId, message);
          session.messageCount++;
        }
      }
    } else {
      logger.info('🔍 REAL FIX: No new messages found');
    }
  } catch (error) {
    logger.error('❌ REAL FIX: Error capturing real messages:', {
      meetingId: session.meetingId,
      chatId: session.chatId,
      error: error.message,
      status: error.response?.status
    });
    
    if (error.response?.status === 403) {
      logger.error('🚨 REAL FIX: Permission denied - check Graph API permissions for Chat.Read.All');
    }
    if (error.response?.status === 404) {
      logger.error('🚨 REAL FIX: Chat not found - meeting chat might not exist yet');
    }
  }
}

  // Process message with AI analysis, using the real message structure
  // 🔄 REPLACE YOUR EXISTING processMessageWithAI METHOD WITH THIS:
async processMessageWithAI(meetingId, message) {
  try {
    // Handle both real Teams messages and simulated messages
    let content, senderName, senderId;
    
    if (message.from && message.from.user) {
      // Real Teams message structure
      content = message.body.content.replace(/<[^>]*>?/gm, "");
      senderName = message.from.user.displayName;
      senderId = message.from.user.id;
    } else if (message.sender) {
      // Fallback structure for simulated messages from attendance service
      content = message.content;
      senderName = message.senderName || message.sender;
      senderId = message.sender;
    } else {
      // Handle other message formats
      content = message.content || message.body?.content || '';
      senderName = message.senderName || message.sender || 'Unknown';
      senderId = message.senderId || message.id || 'unknown';
    }

    if (!content || content.trim() === '') {
      logger.warn('⚠️ Empty message content, skipping processing');
      return null;
    }

    let aiAnalysis = {};

    if (geminiAI.isAvailable()) {
      aiAnalysis = await this.analyzeMessageWithGemini(content, meetingId);
    } else {
      aiAnalysis = await this.basicMessageAnalysis(content);
    }

    const enhancedMessage = {
      id: message.id,
      meetingId: meetingId,
      sender: senderName,
      senderId: senderId,
      content: content,
      timestamp: message.createdDateTime || message.timestamp || new Date().toISOString(),
      messageType: message.messageType || 'message',
      aiAnalysis: { ...aiAnalysis, processedAt: new Date().toISOString() },
      category: aiAnalysis.primaryCategory || "general",
      isActionItem: aiAnalysis.isActionItem || false,
      isQuestion: aiAnalysis.isQuestion || false,
      isDecision: aiAnalysis.isDecision || false,
      urgency: aiAnalysis.urgency || "low",
      sentiment: aiAnalysis.sentiment || "neutral",
    };

    await cosmosClient.createItem("chats", enhancedMessage);

    // 🆕 NEW: Auto-send insights
    if (this.autoInsightsEnabled) {
      await this.autoSendInsights(meetingId, enhancedMessage);
    }

    if (aiAnalysis.isActionItem || aiAnalysis.urgency === "high") {
      await this.handleUrgentMessage(meetingId, enhancedMessage);
    }

    logger.debug("✅ Message processed and stored", {
      meetingId,
      category: enhancedMessage.category,
      content: content.substring(0, 50) + '...',
      sender: senderName
    });
    return enhancedMessage;
  } catch (error) {
    logger.error("❌ Failed to process message with AI:", error);
    return null;
  }
}

  // Analyze message with Gemini AI (with added file/link detection)
  async analyzeMessageWithGemini(content, meetingId) {
    try {
      const prompt = `
        Analyze this meeting chat message and provide insights.
        Message: "${content}"
        
        Please analyze and respond in JSON format:
        {
          "primaryCategory": "question|action_item|decision|resource_sharing|general",
          "isActionItem": true/false,
          "isQuestion": true/false,
          "isDecision": true/false,
          "urgency": "low|medium|high",
          "sentiment": "positive|neutral|negative",
          "sharedResource": "Extract any URL or filename mentioned, otherwise null"
        }
      `;
      const result = await geminiAI.model.generateContent(prompt);
      const response = await result.response;
      const text = response.text();
      const analysis = JSON.parse(text.replace(/```json|```/g, "").trim());
      return analysis;
    } catch (error) {
      logger.warn("Gemini AI analysis failed, using basic analysis:", error);
      return await this.basicMessageAnalysis(content);
    }
  }

  // Basic message analysis (fallback)
  async basicMessageAnalysis(content) {
    return {
      primaryCategory: this.detectPrimaryCategory(content),
      isActionItem: this.detectActionItem(content),
      isQuestion: this.detectQuestion(content),
      isDecision: this.detectDecision(content),
      urgency: this.detectUrgency(content),
      sentiment: this.detectSentiment(content),
      keyTopics: this.extractKeyTopics(content),
      extractedActionItems: this.extractBasicActionItems(content),
      mentions: this.extractMentions(content),
      requiresFollowUp: this.detectFollowUpNeed(content),
      confidenceScore: 0.7,
    };
  }

  detectPrimaryCategory(content) {
    const contentLower = content.toLowerCase();
    if (
      contentLower.includes("?") ||
      contentLower.startsWith("what") ||
      contentLower.startsWith("how")
    )
      return "question";
    if (
      contentLower.includes("action item") ||
      contentLower.includes("will do")
    )
      return "action_item";
    if (contentLower.includes("decided") || contentLower.includes("agreed"))
      return "decision";
    if (contentLower.includes("http") || contentLower.includes("shared"))
      return "resource_sharing";
    return "general";
  }
  detectActionItem(content) {
    const keywords = ["action item", "todo", "will do", "need to", "by friday"];
    return keywords.some((k) => content.toLowerCase().includes(k));
  }
  detectQuestion(content) {
    return (
      content.includes("?") ||
      /^(what|how|when|where|why|can we)/i.test(content)
    );
  }
  detectDecision(content) {
    const keywords = ["decided", "agreed", "approved", "final decision"];
    return keywords.some((k) => content.toLowerCase().includes(k));
  }
  detectUrgency(content) {
    const lower = content.toLowerCase();
    if (["urgent", "asap", "critical"].some((w) => lower.includes(w)))
      return "high";
    if (["soon", "priority"].some((w) => lower.includes(w))) return "medium";
    return "low";
  }
  detectSentiment(content) {
    const lower = content.toLowerCase();
    const pos = ["good", "great", "excellent", "awesome"];
    const neg = ["bad", "problem", "issue", "wrong"];
    if (pos.some((w) => lower.includes(w))) return "positive";
    if (neg.some((w) => lower.includes(w))) return "negative";
    return "neutral";
  }
  extractKeyTopics(content) {
    const words = content.toLowerCase().split(/\s+/);
    return words
      .filter(
        (w) =>
          w.length > 5 &&
          ![
            "that",
            "this",
            "with",
            "from",
            "they",
            "have",
            "will",
            "were",
            "been",
            "said",
          ].includes(w)
      )
      .slice(0, 3);
  }
  extractBasicActionItems(content) {
    if (this.detectActionItem(content)) {
      return [
        {
          task: content,
          assignee: this.extractAssignee(content),
          deadline: this.extractDeadline(content),
        },
      ];
    }
    return [];
  }
  extractMentions(content) {
    const matches = content.match(/@(\w+)/g);
    return matches ? matches.map((m) => m.substring(1)) : [];
  }
  extractAssignee(content) {
    const match = content.match(/assigned to (\w+)/i);
    return match ? match[1] : null;
  }
  extractDeadline(content) {
    const match = content.match(/by (\w+day|\w+ \d{1,2})/i);
    return match ? match[1] : null;
  }
  detectFollowUpNeed(content) {
    const keywords = ["follow up", "check back", "revisit"];
    return keywords.some((k) => content.toLowerCase().includes(k));
  }

  // Handle urgent messages requiring immediate attention
  async handleUrgentMessage(meetingId, message) {
    try {
      logger.warn("🚨 Urgent message detected", {
        meetingId,
        sender: message.sender,
        urgency: message.aiAnalysis.urgency,
      });
      const urgentNotification = {
        id: `urgent-${Date.now()}`,
        meetingId,
        messageId: message.id,
        type: message.isActionItem ? "urgent_action_item" : "urgent_message",
        content: message.content,
        sender: message.sender,
        urgency: message.aiAnalysis.urgency,
        createdAt: new Date().toISOString(),
        handled: false,
      };
      await cosmosClient.createItem("notifications", urgentNotification);
      this.emitUrgentNotification(meetingId, urgentNotification);
    } catch (error) {
      logger.error("❌ Failed to handle urgent message:", error);
    }
  }

  // Emit real-time updates
  emitRealTimeUpdate(meetingId, message) {
    try {
      logger.info("📡 Real-time update emitted", {
        meetingId,
        messageCategory: message.category,
        isActionItem: message.isActionItem,
      });
    } catch (error) {
      logger.warn("Failed to emit real-time update:", error);
    }
  }

  // Emit urgent notifications
  emitUrgentNotification(meetingId, notification) {
    try {
      logger.warn("🚨 Urgent notification emitted", {
        meetingId,
        type: notification.type,
        urgency: notification.urgency,
      });
    } catch (error) {
      logger.warn("Failed to emit urgent notification:", error);
    }
  }

  // Stop chat capture for a meeting
  async stopChatCapture(meetingId) {
    try {
      const captureSession = this.activeCaptures.get(meetingId);
      if (captureSession) {
        captureSession.isActive = false;
        if (captureSession.monitoringLoop) {
          clearInterval(captureSession.monitoringLoop);
        }
        this.activeCaptures.delete(meetingId);
        logger.info("✅ Chat capture stopped", { meetingId });
      }
    } catch (error) {
      logger.error("❌ Failed to stop chat capture:", error);
    }
  }

  // Get chat capture status for all active meetings
  getActiveCaptureStatus() {
    return Array.from(this.activeCaptures.entries()).map(
      ([meetingId, session]) => ({
        meetingId,
        ...session,
        duration: new Date() - new Date(session.startTime),
        monitoringActive: !!session.monitoringLoop,
      })
    );
  }

  // Get detailed chat analysis for a meeting
  async getChatAnalysis(meetingId) {
    try {
      const messages = await cosmosClient.queryItems(
        "chats",
        "SELECT * FROM c WHERE c.meetingId = @meetingId ORDER BY c.timestamp ASC",
        [{ name: "@meetingId", value: meetingId }]
      );
      if (!messages || messages.length === 0) {
        return { meetingId, totalMessages: 0, analysis: null };
      }

      const participantAnalysis = {};
      messages.forEach((msg) => {
        const senderName = msg.sender || "Unknown";
        if (!participantAnalysis[senderName]) {
          participantAnalysis[senderName] = {
            messageCount: 0,
            questions: 0,
            actionItems: 0,
            decisions: 0,
            sentiment: { positive: 0, neutral: 0, negative: 0 },
          };
        }
        participantAnalysis[senderName].messageCount++;
        if (msg.isQuestion) participantAnalysis[senderName].questions++;
        if (msg.isActionItem) participantAnalysis[senderName].actionItems++;
        if (msg.isDecision) participantAnalysis[senderName].decisions++;
        if (msg.aiAnalysis && msg.aiAnalysis.sentiment) {
          participantAnalysis[senderName].sentiment[msg.aiAnalysis.sentiment]++;
        }
      });

      const timeline = this.createMessageTimeline(messages);

      return {
        meetingId,
        totalMessages: messages.length,
        categorizedCounts: {
          questions: messages.filter((m) => m.isQuestion).length,
          actionItems: messages.filter((m) => m.isActionItem).length,
          decisions: messages.filter((m) => m.isDecision).length,
          sharedResources: messages.filter((m) => m.aiAnalysis?.sharedResource)
            .length,
        },
        participantAnalysis: participantAnalysis,
        timeline: timeline,
        keyInsights: {
          mostActiveParticipant:
            Object.keys(participantAnalysis).length > 0
              ? Object.keys(participantAnalysis).reduce((a, b) =>
                  participantAnalysis[a].messageCount >
                  participantAnalysis[b].messageCount
                    ? a
                    : b
                )
              : "N/A",
        },
      };
    } catch (error) {
      logger.error("❌ Failed to get chat analysis:", error);
      throw error;
    }
  }

  // Create message timeline for visualization
  createMessageTimeline(messages) {
    const timeline = {};
    messages.forEach((msg) => {
      const timeKey = new Date(msg.timestamp).toISOString().slice(0, 16); // Minute precision
      if (!timeline[timeKey]) {
        timeline[timeKey] = {
          timestamp: timeKey,
          messageCount: 0,
          categories: {
            questions: 0,
            actionItems: 0,
            decisions: 0,
            general: 0,
          },
        };
      }
      timeline[timeKey].messageCount++;
      if (msg.isQuestion) timeline[timeKey].categories.questions++;
      else if (msg.isActionItem) timeline[timeKey].categories.actionItems++;
      else if (msg.isDecision) timeline[timeKey].categories.decisions++;
      else timeline[timeKey].categories.general++;
    });
    return Object.values(timeline).sort((a, b) =>
      a.timestamp.localeCompare(b.timestamp)
    );
  }


  // 🆕 ADD THESE NEW METHODS AT THE END OF ChatCaptureService CLASS:

  // Auto-send insights method
  async autoSendInsights(meetingId, message) {
    try {
      let insightMessage = null;

      // Action Item Detection
      if (message.isActionItem) {
        insightMessage = `🎯 **Action Item Detected!**\n` +
          `📋 Task: ${message.content.substring(0, 80)}...\n` +
          `👤 Assigned: ${message.sender}\n` +
          `🚨 Priority: ${message.urgency.toUpperCase()}`;
      }
      
      // Decision Detection
      else if (message.isDecision) {
        insightMessage = `✅ **Decision Made!**\n` +
          `📝 Decision: ${message.content.substring(0, 80)}...\n` +
          `👤 By: ${message.sender}\n` +
          `💡 AI will track this for follow-up`;
      }

      // Question Pattern (every 3rd question)
      else if (message.isQuestion) {
        const counter = this.getInsightCounter(meetingId);
        counter.questions++;
        
        if (counter.questions % 3 === 0) {
          insightMessage = `❓ **Question Pattern Alert**\n` +
            `📊 ${counter.questions} questions asked so far\n` +
            `💡 Consider addressing recurring topics`;
        }
      }

      // Send insight if we have one
      if (insightMessage) {
        await this.sendToMeetingChat(meetingId, insightMessage);
      }

      // Periodic summary every 20 messages
      const counter = this.getInsightCounter(meetingId);
      counter.totalMessages++;
      
      if (counter.totalMessages % 20 === 0) {
        await this.sendPeriodicUpdate(meetingId);
      }

    } catch (error) {
      logger.warn('⚠️ Auto-insight failed:', error);
    }
  }

  // Send message to meeting attendees
  // 🔄 MODIFY this method in chatCaptureService.js
// 🔄 REPLACE THIS METHOD to send to Teams chat directly
async sendToMeetingChat(meetingId, insightMessage) {
  try {
    // Try to send directly to Teams meeting chat
    const chatId = this.activeCaptures.get(meetingId)?.chatId;
    
    if (chatId && authService.isAvailable()) {
      const accessToken = await authService.getAppOnlyToken();
      
      const messagePayload = {
        body: {
          contentType: "html",
          content: `<div style="background:#e7f3ff; padding:10px; border-left:4px solid #0078d4; border-radius:5px;">
            <strong>🤖 AI Auto-Insight</strong><br/>
            ${insightMessage.replace(/\n/g, '<br/>')}
          </div>`
        }
      };

      await axios.post(
        `${this.graphEndpoint}/chats/${chatId}/messages`,
        messagePayload,
        {
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          }
        }
      );

      logger.info('✅ Insight sent directly to Teams chat', { meetingId, chatId });
      return;
    }
    
    // Fallback to email if chat not available
    logger.warn('⚠️ Could not send to Teams chat, using email fallback');
    
    // Your existing email code here...
    const teamsService = require('./teamsService');
    
    const meetings = await cosmosClient.queryItems('meetings',
      'SELECT * FROM c WHERE c.meetingId = @meetingId',
      [{ name: '@meetingId', value: meetingId }]
    );

    if (meetings && meetings.length > 0 && meetings[0].attendees) {
      const meeting = meetings[0];
      const recipient = meeting.attendees[0];
      
      const fullMessage = `🤖 **AI Live Insight**\n\n${insightMessage}\n\n` +
        `📝 Meeting: ${meeting.subject}\n⏰ ${new Date().toLocaleTimeString()}`;

      await teamsService.sendMessageToUser(recipient, fullMessage);
      logger.info('✅ Auto-insight sent via email fallback', { meetingId, recipient });
    }
  } catch (error) {
    logger.warn('⚠️ Failed to send insight:', error);
  }
}

  // Send periodic updates
  async sendPeriodicUpdate(meetingId) {
    try {
      const analysis = await this.getChatAnalysis(meetingId);
      
      const summary = `📈 **Meeting Progress Update**\n\n` +
        `💬 Messages: ${analysis.totalMessages}\n` +
        `❓ Questions: ${analysis.categorizedCounts.questions}\n` +
        `🎯 Action Items: ${analysis.categorizedCounts.actionItems}\n` +
        `✅ Decisions: ${analysis.categorizedCounts.decisions}\n` +
        `👥 Most Active: ${analysis.keyInsights.mostActiveParticipant}\n\n` +
        `💡 Meeting is progressing well! Keep up the great work! 🚀`;

      await this.sendToMeetingChat(meetingId, summary);
    } catch (error) {
      logger.warn('⚠️ Failed to send periodic update:', error);
    }
  }

  // 🆕 ADD THIS METHOD to send insights to meeting chat directly
async sendInsightToMeetingChat(meetingId, insightMessage) {
  try {
    const chatId = this.activeCaptures.get(meetingId)?.chatId;
    
    if (chatId && authService.isAvailable()) {
      // Send directly to Teams meeting chat
      const accessToken = await authService.getAppOnlyToken();
      
      const messagePayload = {
        body: {
          contentType: "html",
          content: `<div style="background:#f0f8ff; padding:10px; border-left:4px solid #0078d4;">
            <strong>🤖 AI Agent Insight</strong><br/>
            ${insightMessage.replace(/\n/g, '<br/>')}
          </div>`
        }
      };

      await axios.post(
        `${this.graphEndpoint}/chats/${chatId}/messages`,
        messagePayload,
        {
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          }
        }
      );

      logger.info('✅ Insight sent to meeting chat', { meetingId, chatId });
      return true;
    }
    
    // Fallback to email if no chat available
    return false;
  } catch (error) {
    logger.warn('⚠️ Failed to send to meeting chat, using email fallback:', error.message);
    return false;
  }
}

  // Helper to get insight counter
  getInsightCounter(meetingId) {
    if (!this.insightCounters.has(meetingId)) {
      this.insightCounters.set(meetingId, {
        totalMessages: 0,
        questions: 0,
        actionItems: 0,
        decisions: 0,
        urgentItems: 0,
        startTime: new Date()
      });
    }
    return this.insightCounters.get(meetingId);
  }

  // Start auto-insights when meeting begins
  async startAutoInsights(meetingId) {
    try {
      if (!this.autoInsightsEnabled) return;

      logger.info('🤖 Starting auto-insights', { meetingId });

      // Initialize counter
      this.getInsightCounter(meetingId);

      // Send welcome message
      const welcomeMessage = `🤖 **AI Meeting Agent Activated**\n\n` +
        `✅ Live insights enabled\n` +
        `📊 Real-time analytics active\n` +
        `🚨 Smart alerts configured\n` +
        `📈 Progress updates every 20 messages\n\n` +
        `💡 I'll automatically share insights as the meeting progresses!`;

      await this.sendToMeetingChat(meetingId, welcomeMessage);

    } catch (error) {
      logger.error('❌ Failed to start auto-insights:', error);
    }
  }

  // Stop auto-insights and send final summary
  async stopAutoInsights(meetingId) {
    try {
      logger.info('🛑 Stopping auto-insights', { meetingId });

      // Send final summary
      await this.sendFinalMeetingSummary(meetingId);

      // Cleanup
      this.insightCounters.delete(meetingId);
      
      const timer = this.meetingInsightTimers.get(meetingId);
      if (timer) {
        clearInterval(timer);
        this.meetingInsightTimers.delete(meetingId);
      }

    } catch (error) {
      logger.error('❌ Failed to stop auto-insights:', error);
    }
  }

  // Send final meeting summary
  async sendFinalMeetingSummary(meetingId) {
    try {
      const meetingSummaryService = require('./meetingSummaryService');
      const teamsService = require('./teamsService');
      
      logger.info('📋 Sending final meeting summary', { meetingId });

      // Generate summary
      const summary = await meetingSummaryService.generateMeetingSummary(meetingId, {
        includeChat: true,
        includeParticipantAnalysis: true,
        summaryType: 'comprehensive'
      });

      const finalMessage = `🏁 **AI Meeting Summary**\n\n` +
        `📝 **${summary.meeting.subject}**\n` +
        `⏰ Duration: ${summary.meeting.duration.formatted}\n\n` +
        `📊 **Key Results:**\n` +
        `• ${summary.actionItems.length} action items identified\n` +
        `• ${summary.metrics.decisionsTracked} decisions made\n` +
        `• ${summary.metrics.questionsAsked} questions discussed\n` +
        `• ${summary.metrics.totalMessages} messages exchanged\n\n` +
        `🎯 **Executive Summary:**\n${summary.executiveSummary}\n\n` +
        `✅ **Next Steps:**\n${summary.nextSteps.slice(0, 3).map(step => `• ${step}`).join('\n')}\n\n` +
        `📈 **Meeting Quality Score:** ${summary.qualityScores.overall}/10\n\n` +
        `💡 Full detailed report available via dashboard`;

      // Send to all attendees
      const meetings = await cosmosClient.queryItems('meetings',
        'SELECT * FROM c WHERE c.meetingId = @meetingId',
        [{ name: '@meetingId', value: meetingId }]
      );

      if (meetings && meetings.length > 0 && meetings[0].attendees) {
        for (const attendee of meetings[0].attendees) {
          try {
            await teamsService.sendMessageToUser(attendee, finalMessage);
            await new Promise(resolve => setTimeout(resolve, 2000)); // Rate limit
          } catch (error) {
            logger.warn(`⚠️ Failed to send summary to ${attendee}:`, error.message);
          }
        }
      }

      logger.info('✅ Final summary sent to all attendees');
    } catch (error) {
      logger.error('❌ Failed to send final summary:', error);
    }
  }






}

const chatCaptureService = new ChatCaptureService();
module.exports = chatCaptureService;
