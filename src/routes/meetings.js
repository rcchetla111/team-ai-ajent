const express = require("express");
const axios = require("axios");
const moment = require("moment");
const { v4: uuidv4 } = require("uuid");
const cosmosClient = require("../config/cosmosdb");
const geminiAI = require("../services/geminiAI");
const teamsService = require("../services/teamsService");
const logger = require("../utils/logger");
const chatCaptureService = require("../services/chatCaptureService");
const meetingAttendanceService = require("../services/meetingAttendanceService");
const meetingSummaryService = require("../services/meetingSummaryService");

const router = express.Router();

// For now, we'll simulate authentication (later we'll add real auth)
const simulateAuth = (req, res, next) => {
  // Simulate a user - replace with real auth later
  req.user = {
    userId: "demo-user-123",
    email: "demo@company.com",
    name: "Demo User",
  };
  next();
};

// Apply simulated auth to all routes (temporary)
router.use(simulateAuth);

// GET /api/meetings/status - Check overall service status
router.get("/status", (req, res) => {
  const aiStatus = geminiAI.isAvailable();
  const teamsStatus = teamsService.isAvailable();
  res.json({
    overall: {
      status: aiStatus && teamsStatus ? "fully_operational" : "partial",
      message: "Service status check complete.",
    },
    services: {
      ai: { available: aiStatus },
      teams: { available: teamsStatus },
    },
  });
});

// POST /api/meetings/create - Create a new meeting with AI assistance
router.post("/create", async (req, res) => {
  try {
    const {
      subject,
      description,
      startTime,
      endTime,
      attendees = [],
      useAI = true,
      autoJoinAgent = true,
      enableChatCapture = true,
    } = req.body;

    logger.info("🤖 Creating meeting with AI assistance", {
      subject,
      attendeesCount: attendees.length,
      useAI,
      autoJoinAgent,
    });

    // Validate required fields
    if (!subject || !startTime || !endTime) {
      return res.status(400).json({
        error: "Subject, start time, and end time are required",
      });
    }

    // Enhanced meeting object
    let enhancedMeeting = {
      subject,
      description,
      startTime,
      endTime,
      attendees,
    };

    // Create real Teams meeting using Microsoft Graph API
    logger.info("🔄 Creating real Teams meeting...");
    const teamsMeetingResult = await teamsService.createTeamsMeeting({
      subject: enhancedMeeting.subject,
      description: description,
      startTime: enhancedMeeting.startTime,
      endTime: enhancedMeeting.endTime,
      attendees: attendees,
    });

    // Create meeting in database
    const meetingData = {
      id: uuidv4(),
      meetingId: teamsMeetingResult.meetingId,
      userId: req.user.userId,
      subject: enhancedMeeting.subject,
      description: description,
      startTime: enhancedMeeting.startTime,
      endTime: enhancedMeeting.endTime,
      attendees: attendees,
      status: "scheduled",
      aiEnhanced: useAI && geminiAI.isAvailable(),
      joinUrl: teamsMeetingResult.joinUrl,
      webUrl: teamsMeetingResult.webUrl,
      graphEventId: teamsMeetingResult.graphEventId,
      isRealTeamsMeeting: teamsMeetingResult.isReal || false,
      agentAttended: false,
      agentConfig: {
        autoJoin: autoJoinAgent,
        enableChatCapture: enableChatCapture,
        generateSummary: true,
      },
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };

    // Save to Cosmos DB
    const savedMeeting = await cosmosClient.createMeeting(meetingData);

    logger.info("Meeting Event: created", {
      event: "created",
      meetingId: savedMeeting.id,
      userId: req.user.userId,
      aiEnhanced: savedMeeting.aiEnhanced,
      attendeesCount: attendees.length,
    });

    // 🚀 IMPROVED: Immediate auto-join for meetings starting soon
    if (autoJoinAgent) {
      try {
        const now = moment();
        const meetingStart = moment(enhancedMeeting.startTime);
        const minutesUntilStart = meetingStart.diff(now, 'minutes');

        logger.info(`⏰ Meeting starts in ${minutesUntilStart} minutes`);

        // If meeting starts within the next 10 minutes, join immediately
        if (minutesUntilStart <= 10 && minutesUntilStart >= -5) {
          logger.info("🤖 Meeting starting soon, joining agent immediately");
          
          try {
            // Join the meeting immediately
            await meetingAttendanceService.joinMeeting(
              savedMeeting.meetingId,
              req.user.userId
            );

            // Start chat capture if enabled
            if (enableChatCapture) {
              await chatCaptureService.initiateAutomaticCapture(savedMeeting);
            }

            // Update meeting status
            await cosmosClient.updateItem(
              "meetings",
              savedMeeting.id,
              req.user.userId,
              {
                status: "in_progress",
                agentJoinedAt: new Date().toISOString(),
                agentAttended: true,
              }
            );

            logger.info("✅ Agent immediately joined meeting", {
              meetingId: savedMeeting.id,
            });

            // Add to response
            savedMeeting.agentJoinedImmediately = true;
            savedMeeting.agentAttended = true;
            savedMeeting.status = "in_progress";

          } catch (immediateJoinError) {
            logger.error("❌ Immediate join failed:", immediateJoinError);
            savedMeeting.agentJoinError = immediateJoinError.message;
          }
        } 
        // If meeting is in the future, schedule for later
        else if (minutesUntilStart > 10) {
          logger.info("📅 Meeting scheduled for future, will join closer to start time");
          savedMeeting.agentJoinScheduled = true;
          savedMeeting.scheduledJoinTime = moment(enhancedMeeting.startTime).subtract(5, 'minutes').toISOString();
        }
        // If meeting already ended
        else if (minutesUntilStart < -60) {
          logger.warn("⚠️ Meeting appears to be in the past");
          savedMeeting.agentJoinSkipped = "Meeting in the past";
        }

      } catch (scheduleError) {
        logger.warn("⚠️ Failed to handle agent auto-join:", scheduleError);
        savedMeeting.agentJoinError = scheduleError.message;
      }
    }

    // Send response
    res.status(201).json({
      success: true,
      meeting: savedMeeting,
      message: teamsMeetingResult.isReal || teamsService.isAvailable()
        ? "🟢 Real Teams meeting created successfully!"
        : "📞 Meeting created with simulated Teams link",
      realTeamsMeeting: teamsMeetingResult.isReal || teamsService.isAvailable(),
      teamsIntegrationStatus: teamsService.getStatus(),
      agentStatus: {
        autoJoinEnabled: autoJoinAgent,
        chatCaptureEnabled: enableChatCapture,
        joinedImmediately: savedMeeting.agentJoinedImmediately || false,
        scheduledJoin: savedMeeting.agentJoinScheduled || false,
        error: savedMeeting.agentJoinError || null
      },
    });
  } catch (error) {
    logger.error("❌ Create meeting error:", error);
    res.status(500).json({
      error: "Failed to create meeting",
      details: error.message,
    });
  }
});

router.post("/create-with-real-users", async (req, res) => {
  try {
    const {
      subject,
      description,
      startTime,
      endTime,
      attendeeNames = [],      // ["John Smith", "Sarah Johnson"]
      attendeeEmails = [],     // ["john@company.com"]
      sendInviteMessages = false,
      autoJoinAgent = true,
      enableChatCapture = true,
      useAI = true
    } = req.body;

    // Validate required fields
    if (!subject || !startTime || !endTime) {
      return res.status(400).json({
        error: "Subject, start time, and end time are required"
      });
    }

    logger.info("🚀 Creating meeting with real user resolution", {
      subject,
      attendeeNames,
      attendeeEmails,
      sendInviteMessages
    });

    let meetingResult;
    
    if (teamsService.isAvailable()) {
      // Create with real user resolution
      meetingResult = await teamsService.createTeamsMeetingWithRealUsers({
        subject,
        startTime,
        endTime,
        attendeeNames,
        attendeeEmails,
        sendInviteMessages
      });
    } else {
      // Fallback to regular creation
      logger.warn("⚠️ Teams integration not available, using fallback method");
      meetingResult = await teamsService.createTeamsMeeting({
        subject,
        startTime,
        endTime,
        attendees: [...attendeeEmails] // Use emails only
      });
    }

    // Create meeting in database
    const meetingData = {
      id: uuidv4(),
      meetingId: meetingResult.meetingId,
      userId: req.user.userId,
      subject: subject,
      description: description,
      startTime: startTime,
      endTime: endTime,
      attendees: meetingResult.resolvedAttendees || [...attendeeEmails],
      attendeeNames: attendeeNames,
      status: "scheduled",
      aiEnhanced: useAI && geminiAI.isAvailable(),
      joinUrl: meetingResult.joinUrl,
      webUrl: meetingResult.webUrl,
      graphEventId: meetingResult.graphEventId,
      isRealTeamsMeeting: meetingResult.isReal || false,
      agentAttended: false,
      agentConfig: {
        autoJoin: autoJoinAgent,
        enableChatCapture: enableChatCapture,
        generateSummary: true
      },
      realUserResolution: {
        namesRequested: attendeeNames.length,
        usersResolved: meetingResult.realUsersResolved || 0,
        messagesAttempted: meetingResult.messagesAttempted || 0,
        messagesSent: meetingResult.messagesSent || 0,
        inviteResults: meetingResult.inviteResults || null
      },
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    // Save to Cosmos DB
    const savedMeeting = await cosmosClient.createMeeting(meetingData);

    logger.info("Meeting Event: created_with_real_users", {
      event: "created_with_real_users",
      meetingId: savedMeeting.id,
      userId: req.user.userId,
      realUsersResolved: meetingResult.realUsersResolved || 0,
      messagesSent: meetingResult.messagesSent || 0
    });

    // Handle auto-join logic (same as before)
    if (autoJoinAgent) {
      try {
        const now = moment();
        const meetingStart = moment(startTime);
        const minutesUntilStart = meetingStart.diff(now, 'minutes');

        if (minutesUntilStart <= 10 && minutesUntilStart >= -5) {
          logger.info("🤖 Meeting starting soon, joining agent immediately");
          
          try {
            await meetingAttendanceService.joinMeeting(
              savedMeeting.meetingId,
              req.user.userId
            );

            if (enableChatCapture) {
              await chatCaptureService.initiateAutomaticCapture(savedMeeting);
            }

            await cosmosClient.updateItem(
              "meetings",
              savedMeeting.id,
              req.user.userId,
              {
                status: "in_progress",
                agentJoinedAt: new Date().toISOString(),
                agentAttended: true
              }
            );

            savedMeeting.agentJoinedImmediately = true;
            savedMeeting.agentAttended = true;
            savedMeeting.status = "in_progress";

          } catch (immediateJoinError) {
            logger.error("❌ Immediate join failed:", immediateJoinError);
            savedMeeting.agentJoinError = immediateJoinError.message;
          }
        }
      } catch (scheduleError) {
        logger.warn("⚠️ Failed to handle agent auto-join:", scheduleError);
        savedMeeting.agentJoinError = scheduleError.message;
      }
    }

    // Send response
    res.status(201).json({
      success: true,
      meeting: savedMeeting,
      message: teamsService.isAvailable()
        ? "🚀 Real Teams meeting created with real user resolution!"
        : "📞 Meeting created with simulated Teams link",
      realUserResolution: {
        available: teamsService.isAvailable(),
        namesRequested: attendeeNames.length,
        usersResolved: meetingResult.realUsersResolved || 0,
        messagesAttempted: meetingResult.messagesAttempted || 0,
        messagesSent: meetingResult.messagesSent || 0,
        inviteResults: meetingResult.inviteResults || null
      },
      teamsIntegrationStatus: teamsService.getStatus(),
      agentStatus: {
        autoJoinEnabled: autoJoinAgent,
        chatCaptureEnabled: enableChatCapture,
        joinedImmediately: savedMeeting.agentJoinedImmediately || false,
        error: savedMeeting.agentJoinError || null
      }
    });

  } catch (error) {
    logger.error("❌ Create meeting with real users error:", error);
    res.status(500).json({
      error: "Failed to create meeting with real users",
      details: error.message
    });
  }
});

// Add this to your meetings.js routes (near the top)
router.post("/send-message-by-name", async (req, res) => {
  try {
    const { name, message } = req.body;
    
    if (!name || !message) {
      return res.status(400).json({ 
        error: "Both 'name' and 'message' are required" 
      });
    }
    
    if (!teamsService.isAvailable()) {
      return res.status(503).json({ 
        error: "Teams integration not available",
        suggestion: "Check Azure AD configuration"
      });
    }

    logger.info(`📨 Sending message to: ${name}`);
    
    const result = await teamsService.sendMessageToUser(name, message);
    
    res.json({
      success: true,
      result: result,
      message: `Message sent successfully to ${result.recipient.name}`
    });
    
  } catch (error) {
    logger.error("❌ Send message by name failed:", error);
    res.status(500).json({
      error: "Failed to send message",
      details: error.message
    });
  }
});

// GET /api/meetings - Get user's meetings
router.get("/", async (req, res) => {
  try {
    const {
      status,
      limit = 20,
      offset = 0,
      includeAIInsights = false,
    } = req.query;

    let meetings = await cosmosClient.getMeetingsByUser(req.user.userId);

    // Filter by status if specified
    if (status) {
      meetings = meetings.filter((meeting) => meeting.status === status);
    }

    // Add AI insights if requested
    if (includeAIInsights === "true" && geminiAI.isAvailable()) {
      for (let meeting of meetings.slice(0, 5)) {
        // Limit AI analysis to first 5 for performance
        try {
          if (meeting.description) {
            const insights = await geminiAI.analyzeMeetingDescription(
              meeting.description
            );
            meeting.aiInsights = insights;
          }
        } catch (aiError) {
          logger.warn("⚠️ Failed to get AI insights for meeting:", meeting.id);
        }
      }
    }

    // Apply pagination
    const paginatedMeetings = meetings.slice(
      parseInt(offset),
      parseInt(offset) + parseInt(limit)
    );

    res.json({
      meetings: paginatedMeetings,
      total: meetings.length,
      limit: parseInt(limit),
      offset: parseInt(offset),
      aiInsightsIncluded: includeAIInsights === "true",
    });
  } catch (error) {
    logger.error("❌ Get meetings error:", error);
    res.status(500).json({ error: "Failed to retrieve meetings" });
  }
});


// GET /api/meetings/teams/status - Check Teams integration status
router.get("/teams/status", (req, res) => {
  const teamsStatus = teamsService.getStatus();
  res.json({
    ...teamsStatus,
    message: teamsStatus.available
      ? "🟢 Teams integration is fully operational!"
      : "⚠️ Teams integration not available. Check Azure AD configuration.",
  });
});

// GET /api/meetings/status - Check overall service status
router.get("/status", (req, res) => {
  const aiStatus = geminiAI.isAvailable();
  const teamsStatus = teamsService.isAvailable();

  res.json({
    overall: {
      status:
        aiStatus && teamsStatus
          ? "fully_operational"
          : aiStatus || teamsStatus
          ? "partial"
          : "limited",
      message:
        aiStatus && teamsStatus
          ? "🚀 All services operational - AI + Real Teams integration!"
          : aiStatus
          ? "🤖 AI operational, Teams integration limited"
          : teamsStatus
          ? "🟢 Teams operational, AI limited"
          : "⚠️ Limited functionality - configure AI and Teams",
    },
    services: {
      ai: {
        available: aiStatus,
        model: process.env.GEMINI_MODEL || "Not configured",
      },
      teams: {
        available: teamsStatus,
        realMeetings: teamsStatus,
        fallbackMode: !teamsStatus,
      },
    },
    capabilities: {
      createRealMeetings: teamsStatus,
      aiEnhancement: aiStatus,
      smartAgendas: aiStatus,
      timeOptimization: aiStatus,
      contentAnalysis: aiStatus,
    },
  });
});

// Add these new endpoints to your src/routes/meetings.js file

// 🚀 NEW: Search for team members
router.get("/team-members/search", async (req, res) => {
  try {
    const { q } = req.query; // search query
    
    if (!q || q.trim() === '') {
      return res.status(400).json({ error: "Search query 'q' is required" });
    }
    
    if (!teamsService.isAvailable()) {
      return res.status(503).json({ 
        error: "Teams integration not available",
        suggestion: "Check Azure AD configuration"
      });
    }

    logger.info(`🔍 Searching for team members: "${q}"`);
    
    const users = await teamsService.findTeamMembers(q);
    
    res.json({
      success: true,
      query: q,
      found: users.length,
      users: users,
      message: users.length > 0 ? `Found ${users.length} team members` : 'No team members found'
    });
    
  } catch (error) {
    logger.error("❌ Team member search failed:", error);
    res.status(500).json({
      error: "Failed to search team members",
      details: error.message
    });
  }
});

// 🚀 NEW: Get all team members
router.get("/team-members", async (req, res) => {
  try {
    const { limit = 50 } = req.query;
    
    if (!teamsService.isAvailable()) {
      return res.status(503).json({ 
        error: "Teams integration not available",
        suggestion: "Check Azure AD configuration"
      });
    }

    logger.info(`📋 Getting team members (limit: ${limit})`);
    
    const users = await teamsService.getAllTeamMembers(parseInt(limit));
    
    res.json({
      success: true,
      total: users.length,
      users: users,
      message: `Retrieved ${users.length} team members`
    });
    
  } catch (error) {
    logger.error("❌ Get team members failed:", error);
    res.status(500).json({
      error: "Failed to get team members",
      details: error.message
    });
  }
});

// 🚀 NEW: Send message to team member
router.post("/send-message", async (req, res) => {
  try {
    const { email, message } = req.body;
    
    if (!email || !message) {
      return res.status(400).json({ 
        error: "Both 'email' and 'message' are required" 
      });
    }
    
    if (!teamsService.isAvailable()) {
      return res.status(503).json({ 
        error: "Teams integration not available",
        suggestion: "Check Azure AD configuration"
      });
    }

    logger.info(`📨 Sending message to: ${email}`);
    
    const result = await teamsService.sendMessageToUser(email, message);
    
    res.json({
      success: true,
      result: result,
      message: `Message sent successfully to ${result.recipient.name}`
    });
    
  } catch (error) {
    logger.error("❌ Send message failed:", error);
    res.status(500).json({
      error: "Failed to send message",
      details: error.message
    });
  }
});

// 🚀 ENHANCED: Create meeting with real user resolution and messages
router.post("/create-with-real-users", async (req, res) => {
  try {
    const {
      subject,
      description,
      startTime,
      endTime,
      attendeeNames = [],      // ["John Smith", "Sarah Johnson"]
      attendeeEmails = [],     // ["john@company.com"]
      sendInviteMessages = false,
      autoJoinAgent = true,
      enableChatCapture = true,
      useAI = true
    } = req.body;

    // Validate required fields
    if (!subject || !startTime || !endTime) {
      return res.status(400).json({
        error: "Subject, start time, and end time are required"
      });
    }

    logger.info("🚀 Creating meeting with real user resolution", {
      subject,
      attendeeNames,
      attendeeEmails,
      sendInviteMessages
    });

    let meetingResult;
    
    if (teamsService.isAvailable()) {
      // Create with real user resolution
      meetingResult = await teamsService.createTeamsMeetingWithRealUsers({
        subject,
        startTime,
        endTime,
        attendeeNames,
        attendeeEmails,
        sendInviteMessages
      });
    } else {
      // Fallback to regular creation
      logger.warn("⚠️ Teams integration not available, using fallback method");
      meetingResult = await teamsService.createTeamsMeeting({
        subject,
        startTime,
        endTime,
        attendees: [...attendeeEmails] // Use emails only
      });
    }

    // Create meeting in database
    const meetingData = {
      id: uuidv4(),
      meetingId: meetingResult.meetingId,
      userId: req.user.userId,
      subject: subject,
      description: description,
      startTime: startTime,
      endTime: endTime,
      attendees: meetingResult.resolvedAttendees || [...attendeeEmails],
      attendeeNames: attendeeNames,
      status: "scheduled",
      aiEnhanced: useAI && geminiAI.isAvailable(),
      joinUrl: meetingResult.joinUrl,
      webUrl: meetingResult.webUrl,
      graphEventId: meetingResult.graphEventId,
      isRealTeamsMeeting: meetingResult.isReal || false,
      agentAttended: false,
      agentConfig: {
        autoJoin: autoJoinAgent,
        enableChatCapture: enableChatCapture,
        generateSummary: true
      },
      realUserResolution: {
        namesRequested: attendeeNames.length,
        usersResolved: meetingResult.realUsersResolved || 0,
        messagesAttempted: meetingResult.messagesAttempted || 0,
        messagesSent: meetingResult.messagesSent || 0,
        inviteResults: meetingResult.inviteResults || null
      },
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    // Save to Cosmos DB
    const savedMeeting = await cosmosClient.createMeeting(meetingData);

    logger.info("Meeting Event: created_with_real_users", {
      event: "created_with_real_users",
      meetingId: savedMeeting.id,
      userId: req.user.userId,
      realUsersResolved: meetingResult.realUsersResolved || 0,
      messagesSent: meetingResult.messagesSent || 0
    });

    // Handle auto-join logic (same as before)
    if (autoJoinAgent) {
      try {
        const now = moment();
        const meetingStart = moment(startTime);
        const minutesUntilStart = meetingStart.diff(now, 'minutes');

        if (minutesUntilStart <= 10 && minutesUntilStart >= -5) {
          logger.info("🤖 Meeting starting soon, joining agent immediately");
          
          try {
            await meetingAttendanceService.joinMeeting(
              savedMeeting.meetingId,
              req.user.userId
            );

            if (enableChatCapture) {
              await chatCaptureService.initiateAutomaticCapture(savedMeeting);
            }

            await cosmosClient.updateItem(
              "meetings",
              savedMeeting.id,
              req.user.userId,
              {
                status: "in_progress",
                agentJoinedAt: new Date().toISOString(),
                agentAttended: true
              }
            );

            savedMeeting.agentJoinedImmediately = true;
            savedMeeting.agentAttended = true;
            savedMeeting.status = "in_progress";

          } catch (immediateJoinError) {
            logger.error("❌ Immediate join failed:", immediateJoinError);
            savedMeeting.agentJoinError = immediateJoinError.message;
          }
        }
      } catch (scheduleError) {
        logger.warn("⚠️ Failed to handle agent auto-join:", scheduleError);
        savedMeeting.agentJoinError = scheduleError.message;
      }
    }

    // Send response
    res.status(201).json({
      success: true,
      meeting: savedMeeting,
      message: teamsService.isAvailable()
        ? "🚀 Real Teams meeting created with real user resolution!"
        : "📞 Meeting created with simulated Teams link",
      realUserResolution: {
        available: teamsService.isAvailable(),
        namesRequested: attendeeNames.length,
        usersResolved: meetingResult.realUsersResolved || 0,
        messagesAttempted: meetingResult.messagesAttempted || 0,
        messagesSent: meetingResult.messagesSent || 0,
        inviteResults: meetingResult.inviteResults || null
      },
      teamsIntegrationStatus: teamsService.getStatus(),
      agentStatus: {
        autoJoinEnabled: autoJoinAgent,
        chatCaptureEnabled: enableChatCapture,
        joinedImmediately: savedMeeting.agentJoinedImmediately || false,
        error: savedMeeting.agentJoinError || null
      }
    });

  } catch (error) {
    logger.error("❌ Create meeting with real users error:", error);
    res.status(500).json({
      error: "Failed to create meeting with real users",
      details: error.message
    });
  }
});

// 🚀 NEW: Test user resolution
router.post("/test-user-resolution", async (req, res) => {
  try {
    const { names } = req.body; // ["John Smith", "Sarah Johnson"]
    
    if (!names || !Array.isArray(names)) {
      return res.status(400).json({ 
        error: "Array of 'names' is required" 
      });
    }
    
    if (!teamsService.isAvailable()) {
      return res.json({
        mode: "simulated",
        message: "Teams integration not available - any names will work in simulated mode",
        realUserLookup: false,
        input: names,
        suggestion: "Check Azure AD configuration to enable real user lookup"
      });
    }

    logger.info(`🧪 Testing user resolution for: ${names.join(', ')}`);
    
    const resolvedUsers = await teamsService.findUsersByDisplayName(names);
    
    res.json({
      mode: "real_teams",
      realUserLookup: true,
      input: names,
      foundUsers: resolvedUsers,
      summary: {
        requested: names.length,
        found: resolvedUsers.length,
        success: resolvedUsers.length > 0
      },
      message: resolvedUsers.length > 0 
        ? `✅ Found ${resolvedUsers.length}/${names.length} real Teams users!` 
        : "❌ No users found in Teams directory"
    });
    
  } catch (error) {
    logger.error("❌ User resolution test failed:", error);
    res.status(500).json({
      error: "Failed to test user resolution",
      details: error.message
    });
  }
});




// GET /api/meetings/ai/status - Check AI service status
router.get("/ai/status", (req, res) => {
  res.json({
    aiAvailable: geminiAI.isAvailable(),
    modelName: process.env.GEMINI_MODEL || "Not configured",
    features: {
      agendaGeneration: geminiAI.isAvailable(),
      timeOptimization: geminiAI.isAvailable(),
      contentAnalysis: geminiAI.isAvailable(),
      smartTitles: geminiAI.isAvailable(),
    },
    message: geminiAI.isAvailable()
      ? "🤖 AI services are fully operational!"
      : "⚠️ AI services are not available. Check Gemini API key.",
  });
});

// Add to your meetings.js routes
router.post("/test-users", async (req, res) => {
  try {
    const { names } = req.body; // e.g., ["John Smith", "Sarah Johnson"]
    
    if (!teamsService.isAvailable()) {
      return res.json({
        mode: "simulated",
        message: "Teams integration not available - any names will work",
        realUserLookup: false
      });
    }

    const resolvedUsers = await teamsService.findUsersByDisplayName(names);
    
    res.json({
      mode: "real_teams",
      input: names,
      foundUsers: resolvedUsers,
      realUserLookup: true,
      message: resolvedUsers.length > 0 ? "Found real Teams users!" : "No users found in Teams directory"
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});


router.get("/:id/status", async (req, res) => {
  try {
    let meeting = await cosmosClient.getItem(
      "meetings",
      req.params.id,
      req.user.userId
    );

    if (!meeting) {
      const meetings = await cosmosClient.queryItems(
        "meetings",
        "SELECT * FROM c WHERE c.meetingId = @meetingId",
        [{ name: "@meetingId", value: req.params.id }]
      );
      if (meetings && meetings.length > 0) {
        meeting = meetings[0];
      }
    }

    if (!meeting) {
      return res.status(404).json({ error: "Meeting not found" });
    }

    // Get agent attendance status
    let attendanceStatus = null;
    try {
      attendanceStatus = await meetingAttendanceService.getAttendanceSummary(meeting.meetingId);
    } catch (error) {
      logger.warn("Could not get attendance status:", error.message);
    }

    // Get chat capture status
    let chatCaptureStatus = null;
    try {
      const captureStatuses = chatCaptureService.getActiveCaptureStatus();
      chatCaptureStatus = captureStatuses.find(status => status.meetingId === meeting.meetingId);
    } catch (error) {
      logger.warn("Could not get chat capture status:", error.message);
    }

    const now = moment();
    const meetingStart = moment(meeting.startTime);
    const meetingEnd = moment(meeting.endTime);

    res.json({
      success: true,
      meeting: {
        id: meeting.id,
        meetingId: meeting.meetingId,
        subject: meeting.subject,
        startTime: meeting.startTime,
        endTime: meeting.endTime,
        status: meeting.status,
        agentAttended: meeting.agentAttended,
        agentJoinedAt: meeting.agentJoinedAt,
        isRealTeamsMeeting: meeting.isRealTeamsMeeting
      },
      timing: {
        minutesUntilStart: meetingStart.diff(now, 'minutes'),
        minutesSinceStart: now.diff(meetingStart, 'minutes'),
        meetingDuration: meetingEnd.diff(meetingStart, 'minutes'),
        hasStarted: now.isAfter(meetingStart),
        hasEnded: now.isAfter(meetingEnd)
      },
      agentStatus: {
        isAttending: attendanceStatus?.isActive || false,
        attendanceDetails: attendanceStatus,
        chatCapture: {
          isActive: chatCaptureStatus?.isActive || false,
          messageCount: chatCaptureStatus?.messageCount || 0,
          details: chatCaptureStatus
        }
      },
      timestamp: new Date().toISOString()
    });

  } catch (error) {
    logger.error("❌ Get meeting status error:", error);
    res.status(500).json({
      error: "Failed to get meeting status",
      details: error.message,
    });
  }
});

// NEW ROUTE FOR CHAT FEATURES
// POST /api/meetings/:id/start-monitoring
router.post("/:id/start-monitoring", async (req, res) => {
    try {
        const meeting = await cosmosClient.getItem("meetings", req.params.id, req.user.userId);
        if (!meeting) {
            return res.status(404).json({ error: "Meeting not found." });
        }
        
        await chatCaptureService.startChatCapture(meeting.meetingId, meeting);

        res.json({ success: true, message: `✅ Now monitoring chat for "${meeting.subject}".` });
    } catch (error) {
        logger.error("❌ Start monitoring error:", error);
        res.status(500).json({ error: "Failed to start monitoring", details: error.message });
    }
});

// NEW ROUTE FOR CHAT FEATURES
// GET /api/meetings/:id/chat-analysis
router.get("/:id/chat-analysis", async (req, res) => {
    try {
        const meeting = await cosmosClient.getItem("meetings", req.params.id, req.user.userId);
        if (!meeting) {
            return res.status(404).json({ error: "Meeting not found." });
        }

        const analysis = await chatCaptureService.getChatAnalysis(meeting.meetingId);
        res.json({ success: true, analysis });
    } catch (error) {
        logger.error("❌ Get chat analysis error:", error);
        res.status(500).json({ error: "Failed to get chat analysis", details: error.message });
    }
});

// GET /api/meetings/:id - Get specific meeting with AI insights
router.get("/:id", async (req, res) => {
  try {
    const meeting = await cosmosClient.getItem(
      "meetings",
      req.params.id,
      req.user.userId
    );

    if (!meeting) {
      return res.status(404).json({ error: "Meeting not found" });
    }

    // Add real-time AI insights
    if (geminiAI.isAvailable() && meeting.description) {
      try {
        const insights = await geminiAI.analyzeMeetingDescription(
          meeting.description
        );
        meeting.currentAIInsights = insights;
      } catch (aiError) {
        logger.warn("⚠️ Failed to get current AI insights");
      }
    }

    res.json(meeting);
  } catch (error) {
    logger.error("❌ Get meeting error:", error);
    res.status(500).json({ error: "Failed to retrieve meeting" });
  }
});

router.patch("/:id", async (req, res) => {
  try {
    const meetingId = req.params.id;
    const updatePayload = req.body; // e.g., { attendees: [...] }

    // 1. Find the meeting in our database to get its graphEventId
    const meeting = await cosmosClient.getItem(
      "meetings",
      meetingId,
      req.user.userId
    );
    if (!meeting || !meeting.graphEventId) {
      return res
        .status(404)
        .json({
          error:
            "Meeting not found or it is not a real Teams meeting that can be updated.",
        });
    }

    // 2. Prepare the attendees payload for the Graph API
    // The API needs a specific format for attendees
    if (updatePayload.attendees && Array.isArray(updatePayload.attendees)) {
      updatePayload.attendees = updatePayload.attendees.map((email) => ({
        emailAddress: { address: email, name: email.split("@")[0] },
        type: "required",
      }));
    }

    // 3. Call our service to update the meeting in Microsoft's system
    await teamsService.updateMeeting(meeting.graphEventId, updatePayload);

    // 4. Update our own database with the new information
    const updatedMeetingInDb = await cosmosClient.updateItem(
      "meetings",
      meetingId,
      req.user.userId,
      {
        attendees: req.body.attendees, // Save the simple email list
        updatedAt: new Date().toISOString(),
      }
    );

    logger.info("Meeting Event: updated", { meetingId: meetingId });

    res.json({
      success: true,
      message: "Meeting updated successfully!",
      meeting: updatedMeetingInDb,
    });
  } catch (error) {
    logger.error("❌ Update meeting error:", error);
    res
      .status(500)
      .json({ error: "Failed to update meeting", details: error.message });
  }
});

// POST /api/meetings/suggest-times - AI-powered time suggestions
router.post("/suggest-times", async (req, res) => {
  try {
    const {
      duration = 30,
      attendees = [],
      preferredDays = ["monday", "tuesday", "wednesday", "thursday", "friday"],
      urgency = "normal",
    } = req.body;

    logger.info("🤖 Generating AI-powered time suggestions", {
      duration,
      attendeesCount: attendees.length,
      urgency,
    });

    let suggestions = [];

    if (geminiAI.isAvailable()) {
      suggestions = await geminiAI.suggestMeetingTimes({
        duration,
        attendees,
        preferredDays,
        urgency,
        timeZone: "UTC",
      });
    } else {
      // Fallback suggestions
      suggestions = [
        {
          datetime: moment().add(1, "day").hour(10).minute(0).toISOString(),
          dayOfWeek: "Tomorrow",
          timeSlot: "Morning",
          confidence: 0.7,
          reasoning: "Standard business hours",
        },
      ];
    }

    res.json({
      success: true,
      suggestions: suggestions,
      aiPowered: geminiAI.isAvailable(),
      message: geminiAI.isAvailable()
        ? "🤖 AI-powered time suggestions generated!"
        : "📅 Basic time suggestions provided",
    });
  } catch (error) {
    logger.error("❌ Suggest times error:", error);
    res.status(500).json({ error: "Failed to suggest meeting times" });
  }
});

// POST /api/meetings/:id/enhance - Enhance existing meeting with AI
router.post("/:id/enhance", async (req, res) => {
  try {
    const meeting = await cosmosClient.getItem(
      "meetings",
      req.params.id,
      req.user.userId
    );

    if (!meeting) {
      return res.status(404).json({ error: "Meeting not found" });
    }

    if (!geminiAI.isAvailable()) {
      return res.status(503).json({ error: "AI service not available" });
    }

    const enhancements = {};

    // Generate enhanced agenda if not exists or basic
    if (!meeting.agenda || meeting.agenda.sections?.length <= 3) {
      const enhancedAgenda = await geminiAI.generateMeetingAgenda({
        subject: meeting.subject,
        attendees: meeting.attendees,
        duration: moment(meeting.endTime).diff(
          moment(meeting.startTime),
          "minutes"
        ),
        meetingType: "enhanced",
      });
      enhancements.agenda = enhancedAgenda;
    }

    // Analyze meeting for additional insights
    if (meeting.description) {
      const insights = await geminiAI.analyzeMeetingDescription(
        meeting.description
      );
      enhancements.aiInsights = insights;
    }

    // Update meeting with enhancements
    const updatedMeeting = await cosmosClient.updateItem(
      "meetings",
      req.params.id,
      req.user.userId,
      {
        ...enhancements,
        aiEnhanced: true,
        lastAIEnhancement: new Date().toISOString(),
      }
    );

    logger.info("Meeting Event: ai_enhanced", {
      event: "ai_enhanced",
      meetingId: req.params.id,
      userId: req.user.userId,
    });

    res.json({
      success: true,
      meeting: updatedMeeting,
      enhancements: enhancements,
      message: "🤖 Meeting enhanced with AI insights!",
    });
  } catch (error) {
    logger.error("❌ Enhance meeting error:", error);
    res.status(500).json({ error: "Failed to enhance meeting" });
  }
});

// --- REPLACE your old router.delete with these TWO new functions ---

// DELETE /api/meetings/:id - Cancel a SPECIFIC meeting by its ID
router.delete("/:id", async (req, res) => {
  try {
    const meeting = await cosmosClient.getItem(
      "meetings",
      req.params.id,
      req.user.userId
    );
    if (!meeting) {
      return res.status(404).json({ error: "Meeting not found" });
    }

    // In a real app, you would also call teamsService.cancelTeamsMeeting(meeting.graphEventId)
    // For now, we just update the status in our database.
    await cosmosClient.updateItem("meetings", req.params.id, req.user.userId, {
      status: "cancelled",
      cancelledAt: new Date().toISOString(),
    });

    logger.info("Meeting Event: cancelled", { meetingId: req.params.id });
    res.json({
      success: true,
      message: "Meeting cancelled successfully",
      cancelledCount: 1,
    });
  } catch (error) {
    logger.error("❌ Cancel meeting error:", error);
    res.status(500).json({ error: "Failed to cancel meeting" });
  }
});

// NEW: DELETE /api/meetings - Cancel meetings based on a date
router.delete("/", async (req, res) => {
  try {
    const { date } = req.query; // e.g., ?date=2025-08-10
    if (!date) {
      return res
        .status(400)
        .json({ error: "A date query parameter is required." });
    }

    const meetings = await cosmosClient.getMeetingsByUser(req.user.userId);

    const startOfDay = moment(date).startOf("day");
    const endOfDay = moment(date).endOf("day");

    const meetingsToCancel = meetings.filter((m) => {
      const meetingTime = moment(m.startTime);
      return (
        meetingTime.isBetween(startOfDay, endOfDay) && m.status === "scheduled"
      );
    });

    if (meetingsToCancel.length === 0) {
      return res.json({
        success: true,
        message: "No scheduled meetings found for that date to cancel.",
        cancelledCount: 0,
      });
    }

    for (const meeting of meetingsToCancel) {
      await cosmosClient.updateItem("meetings", meeting.id, req.user.userId, {
        status: "cancelled",
        cancelledAt: new Date().toISOString(),
      });
      logger.info("Meeting Event: cancelled", { meetingId: meeting.id });
    }

    res.json({
      success: true,
      message: `Successfully cancelled ${meetingsToCancel.length} meeting(s).`,
      cancelledCount: meetingsToCancel.length,
    });
  } catch (error) {
    logger.error("❌ Bulk cancel meetings error:", error);
    res.status(500).json({ error: "Failed to cancel meetings" });
  }
});




// Add these 5 missing endpoints to your existing meetings.js file:

// POST /api/meetings/:id/join-agent
// Add these endpoints to the end of your meetings.js file (before module.exports = router):

// POST /api/meetings/:id/join-agent - Join agent to meeting manually
router.post("/:id/join-agent", async (req, res) => {
  try {
    // Look up meeting by ID first, then meetingId if not found
    let meeting = await cosmosClient.getItem(
      "meetings",
      req.params.id,
      req.user.userId
    );

    if (!meeting) {
      const meetings = await cosmosClient.queryItems(
        "meetings",
        "SELECT * FROM c WHERE c.meetingId = @meetingId",
        [{ name: "@meetingId", value: req.params.id }]
      );
      if (meetings && meetings.length > 0) {
        meeting = meetings[0];
      }
    }

    if (!meeting) {
      return res.status(404).json({ error: "Meeting not found" });
    }

    logger.info("🤖 Manually joining agent to meeting", {
      meetingId: req.params.id,
      actualMeetingId: meeting.meetingId,
      subject: meeting.subject
    });

    // Join meeting using correct meetingId
    const joinResult = await meetingAttendanceService.joinMeeting(
      meeting.meetingId,
      req.user.userId
    );

    // Start chat capture if enabled
    if (meeting.agentConfig?.enableChatCapture !== false) {
      await chatCaptureService.initiateAutomaticCapture(meeting);
    }

    // Update meeting status
    await cosmosClient.updateItem("meetings", meeting.id, req.user.userId, {
      agentAttended: true,
      agentJoinedAt: new Date().toISOString(),
      status: "in_progress",
    });

    res.json({
      success: true,
      message: "Agent successfully joined meeting",
      meetingDetails: {
        id: meeting.id,
        meetingId: meeting.meetingId,
        subject: meeting.subject,
        status: "in_progress"
      },
      joinResult: joinResult,
      chatCapture: meeting.agentConfig?.enableChatCapture !== false,
    });
  } catch (error) {
    logger.error("❌ Agent join failed:", error);
    res.status(500).json({
      error: "Failed to join agent to meeting",
      details: error.message,
    });
  }
});

// 4. Similarly fix the leave endpoint:

// POST /api/meetings/:id/leave-agent - Leave agent from meeting
router.post("/:id/leave-agent", async (req, res) => {
  try {
    // FIXED: Look up meeting by both id and meetingId
    let meeting = await cosmosClient.getItem(
      "meetings",
      req.params.id,
      req.user.userId
    );

    if (!meeting) {
      // Try looking up by meetingId
      const meetings = await cosmosClient.queryItems(
        "meetings",
        "SELECT * FROM c WHERE c.meetingId = @meetingId",
        [{ name: "@meetingId", value: req.params.id }]
      );

      if (meetings && meetings.length > 0) {
        meeting = meetings[0];
      }
    }

    logger.info("🤖 Agent leaving meeting", { meetingId: req.params.id });

    // Leave meeting using correct meetingId
    const leaveResult = await meetingAttendanceService.leaveMeeting(
      meeting?.meetingId || req.params.id,
      req.user.userId
    );

    // Stop chat capture
    await chatCaptureService.stopChatCapture(
      meeting?.meetingId || req.params.id
    );

    // Update meeting status if meeting was found
    if (meeting) {
      await cosmosClient.updateItem("meetings", meeting.id, req.user.userId, {
        agentLeftAt: new Date().toISOString(),
        status: "completed",
      });
    }

    res.json({
      success: true,
      message: "Agent successfully left meeting",
      leaveResult: leaveResult,
    });
  } catch (error) {
    logger.error("❌ Agent leave failed:", error);
    res.status(500).json({
      error: "Failed to remove agent from meeting",
      details: error.message,
    });
  }
});

// 5. Fix the summary endpoint to handle meetingId correctly:

// GET /api/meetings/:id/summary - Get or generate meeting summary
router.get("/:id/summary", async (req, res) => {
  try {
    const { regenerate = false } = req.query;

    logger.info("📋 Getting meeting summary", {
      meetingId: req.params.id,
      regenerate,
    });

    // FIXED: Use the correct meetingId for summary generation
    let meetingId = req.params.id;

    // Check if this is a database ID or meetingId
    const meeting = await cosmosClient.getItem(
      "meetings",
      req.params.id,
      req.user.userId
    );
    if (meeting) {
      meetingId = meeting.meetingId; // Use the Teams meetingId for summary
    }

    let summary = null;

    // Check if summary already exists
    if (!regenerate) {
      try {
        const existingSummaries =
          await meetingSummaryService.getMeetingSummaries(meetingId);
        if (existingSummaries.length > 0) {
          summary = existingSummaries[0];
        }
      } catch (error) {
        logger.warn("Could not get existing summaries:", error);
      }
    }

    // Generate new summary if needed
    if (!summary || regenerate) {
      summary = await meetingSummaryService.generateMeetingSummary(meetingId, {
        includeChat: true,
        includeParticipantAnalysis: true,
        autoActionItems: true,
        autoFollowUp: true,
      });
    }

    res.json({
      success: true,
      summary: summary,
      generated: !summary || regenerate,
      timestamp: new Date().toISOString(),
    });
  } catch (error) {
    logger.error("❌ Summary generation failed:", error);
    res.status(500).json({
      error: "Failed to get meeting summary",
      details: error.message,
    });
  }
});

// GET /api/meetings/:id/chat-analysis - Get real-time chat analysis
router.get("/:id/chat-analysis", async (req, res) => {
  try {
    logger.info("💬 Getting chat analysis", { meetingId: req.params.id });

    // Use your existing chatCaptureService
    const chatAnalysis = await chatCaptureService.getChatAnalysis(
      req.params.id
    );

    res.json({
      success: true,
      analysis: chatAnalysis,
      timestamp: new Date().toISOString(),
    });
  } catch (error) {
    logger.error("❌ Chat analysis failed:", error);
    res.status(500).json({
      error: "Failed to get chat analysis",
      details: error.message,
    });
  }
});

// GET /api/meetings/:id/attendance-status - Get agent attendance status
router.get("/:id/attendance-status", async (req, res) => {
  try {
    const attendanceStatus =
      await meetingAttendanceService.getAttendanceSummary(req.params.id);

    res.json({
      success: true,
      attendance: attendanceStatus,
      timestamp: new Date().toISOString(),
    });
  } catch (error) {
    logger.error("❌ Attendance status failed:", error);
    res.status(500).json({
      error: "Failed to get attendance status",
      details: error.message,
    });
  }
});

module.exports = router;
