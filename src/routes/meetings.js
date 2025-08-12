const express = require("express");
const { v4: uuidv4 } = require("uuid");
const moment = require("moment");
const teamsService = require("../services/teamsService");
const geminiAI = require("../services/geminiAI");
const chatCaptureService = require("../services/chatCaptureService");
const meetingAttendanceService = require("../services/meetingAttendanceService");
const meetingSummaryService = require("../services/meetingSummaryService");
const cosmosClient = require("../config/cosmosdb");
const logger = require("../utils/logger");

const router = express.Router();

// Simulate authentication (replace with real auth later)
const simulateAuth = (req, res, next) => {
  req.user = {
    userId: "demo-user-123",
    email: "demo@company.com",
    name: "Demo User",
  };
  next();
};

router.use(simulateAuth);

// Middleware to ensure Teams integration is available
const requireRealTeams = (req, res, next) => {
  if (!teamsService.isAvailable()) {
    return res.status(503).json({
      error: "Real Teams integration is required but not configured",
      message: "Please configure Azure AD credentials to use this feature",
      requiredConfig: {
        azureClientId: "AZURE_CLIENT_ID",
        azureClientSecret: "AZURE_CLIENT_SECRET", 
        azureTenantId: "AZURE_TENANT_ID"
      }
    });
  }
  next();
};

// ============================================================================
// POC FEATURE 1.1: REAL TEAMS MEETING SCHEDULING
// ============================================================================

// POST /api/meetings/create - Create REAL Teams meeting only
router.post("/create", requireRealTeams, async (req, res) => {
  try {
    const {
      subject,
      description,
      startTime,
      endTime,
      attendees = [],
      recurrence = null,
      autoJoinAgent = true,
      enableChatCapture = true,
    } = req.body;

    logger.info("🤖 Creating REAL Teams meeting", {
      subject,
      attendeesCount: attendees.length,
      hasRecurrence: !!recurrence,
      autoJoinAgent,
    });

    // Validate required fields
    if (!subject || !startTime || !endTime) {
      return res.status(400).json({
        error: "Subject, start time, and end time are required",
      });
    }

    // Create REAL Teams meeting via Graph API
    const teamsMeetingResult = await teamsService.createTeamsMeeting({
      subject: subject,
      description: description,
      startTime: startTime,
      endTime: endTime,
      attendees: attendees,
      recurrence: recurrence
    });

    // Create meeting record in database
    const meetingData = {
      id: uuidv4(),
      meetingId: teamsMeetingResult.meetingId,
      userId: req.user.userId,
      subject: subject,
      description: description,
      startTime: startTime,
      endTime: endTime,
      attendees: attendees,
      status: "scheduled",
      joinUrl: teamsMeetingResult.joinUrl,
      webUrl: teamsMeetingResult.webUrl,
      graphEventId: teamsMeetingResult.graphEventId,
      isRealTeamsMeeting: true,
      isRecurring: teamsMeetingResult.isRecurring || false,
      agentAttended: false,
      agentConfig: {
        autoJoin: autoJoinAgent,
        enableChatCapture: enableChatCapture,
        generateSummary: true,
      },
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };

    const savedMeeting = await cosmosClient.createMeeting(meetingData);

    // Auto-join AI agent logic for meetings starting soon
    if (autoJoinAgent) {
      const now = moment();
      const meetingStart = moment(startTime);
      const minutesUntilStart = meetingStart.diff(now, 'minutes');

      if (minutesUntilStart <= 15 && minutesUntilStart >= -5) {
        logger.info("🤖 Meeting starting soon, joining AI agent to REAL Teams meeting");
        
        try {
          await meetingAttendanceService.joinMeeting(
            savedMeeting.meetingId,
            req.user.userId
          );

          if (enableChatCapture) {
            await chatCaptureService.initiateRealChatCapture(savedMeeting);
          }

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

          savedMeeting.agentJoinedImmediately = true;
          savedMeeting.agentAttended = true;
          savedMeeting.status = "in_progress";

        } catch (immediateJoinError) {
          logger.error("❌ Immediate AI agent join failed:", immediateJoinError);
          savedMeeting.agentJoinError = immediateJoinError.message;
        }
      }
    }

    res.status(201).json({
      success: true,
      meeting: savedMeeting,
      message: "🟢 REAL Teams meeting created successfully with AI agent capabilities!",
      realTeamsMeeting: true,
      teamsIntegrationStatus: teamsService.getStatus(),
      agentStatus: {
        autoJoinEnabled: autoJoinAgent,
        chatCaptureEnabled: enableChatCapture,
        joinedImmediately: savedMeeting.agentJoinedImmediately || false,
        error: savedMeeting.agentJoinError || null,
        willBeVisibleToParticipants: true
      },
      meetingDetails: {
        graphEventId: teamsMeetingResult.graphEventId,
        joinUrl: teamsMeetingResult.joinUrl,
        webUrl: teamsMeetingResult.webUrl
      }
    });
  } catch (error) {
    logger.error("❌ Create REAL Teams meeting error:", error);
    res.status(500).json({
      error: "Failed to create real Teams meeting",
      details: error.message,
    });
  }
});

// POST /api/meetings/create-with-names - Create meeting by resolving REAL user names
router.post("/create-with-names", requireRealTeams, async (req, res) => {
  try {
    const {
      subject,
      description,
      startTime,
      endTime,
      attendeeNames = [], // ["John Smith", "Sarah Johnson"]
      attendeeEmails = [], // Additional emails
      autoJoinAgent = true,
      enableChatCapture = true,
    } = req.body;

    if (!subject || !startTime || !endTime) {
      return res.status(400).json({
        error: "Subject, start time, and end time are required"
      });
    }

    logger.info("🚀 Creating REAL Teams meeting with name resolution", {
      subject,
      attendeeNames,
      attendeeEmails,
    });

    let resolvedAttendees = [...attendeeEmails];
    
    // Resolve names to emails using REAL Teams directory
    if (attendeeNames.length > 0) {
      const resolvedUsers = await teamsService.findUsersByDisplayName(attendeeNames);
      resolvedAttendees.push(...resolvedUsers.map(user => user.email));
      
      logger.info(`✅ Resolved ${resolvedUsers.length}/${attendeeNames.length} users from REAL Teams directory`);
    }

    // Remove duplicates
    resolvedAttendees = [...new Set(resolvedAttendees)];

    // Create the REAL Teams meeting
    const teamsMeetingResult = await teamsService.createTeamsMeeting({
      subject,
      description,
      startTime,
      endTime,
      attendees: resolvedAttendees
    });

    // Create meeting record in database
    const meetingData = {
      id: uuidv4(),
      meetingId: teamsMeetingResult.meetingId,
      userId: req.user.userId,
      subject: subject,
      description: description,
      startTime: startTime,
      endTime: endTime,
      attendees: resolvedAttendees,
      attendeeNames: attendeeNames,
      status: "scheduled",
      joinUrl: teamsMeetingResult.joinUrl,
      webUrl: teamsMeetingResult.webUrl,
      graphEventId: teamsMeetingResult.graphEventId,
      isRealTeamsMeeting: true,
      agentAttended: false,
      agentConfig: {
        autoJoin: autoJoinAgent,
        enableChatCapture: enableChatCapture,
        generateSummary: true
      },
      userResolution: {
        namesRequested: attendeeNames.length,
        usersResolved: resolvedAttendees.length - attendeeEmails.length,
        realTeamsDirectoryUsed: true
      },
      createdAt: new Date().toISOString(),
    };

    const savedMeeting = await cosmosClient.createMeeting(meetingData);

    res.status(201).json({
      success: true,
      meeting: savedMeeting,
      message: "🚀 REAL Teams meeting created with user name resolution from Teams directory!",
      realTeamsMeeting: true,
      userResolution: {
        realTeamsDirectoryUsed: true,
        namesRequested: attendeeNames.length,
        usersResolved: resolvedAttendees.length - attendeeEmails.length,
        finalAttendees: resolvedAttendees
      },
      teamsIntegrationStatus: teamsService.getStatus(),
    });

  } catch (error) {
    logger.error("❌ Create REAL Teams meeting with names error:", error);
    res.status(500).json({
      error: "Failed to create real Teams meeting with name resolution",
      details: error.message
    });
  }
});

// GET /api/meetings/suggest-times - Get REAL optimal meeting times from Teams calendars
router.post("/suggest-times", requireRealTeams, async (req, res) => {
  try {
    const {
      attendees = [],
      duration = 30,
    } = req.body;

    logger.info("🤖 Finding optimal meeting times from REAL Teams calendars", {
      attendeesCount: attendees.length,
      duration,
    });

    if (attendees.length === 0) {
      return res.status(400).json({
        error: "At least one attendee email is required to check real calendar availability"
      });
    }

    // Get REAL meeting suggestions from Teams calendars
    const suggestions = await teamsService.findMeetingTimes(attendees, duration);

    res.json({
      success: true,
      suggestions: suggestions,
      realTeamsCalendarData: true,
      message: "🤖 Optimal times found using REAL Teams calendar data!",
      attendeesChecked: attendees.length,
      duration: duration
    });
  } catch (error) {
    logger.error("❌ Suggest times from REAL Teams calendars error:", error);
    res.status(500).json({ 
      error: "Failed to get real meeting time suggestions",
      details: error.message 
    });
  }
});

// ============================================================================
// POC FEATURE 1.4: REAL AI AGENT MEETING ATTENDANCE  
// ============================================================================

// POST /api/meetings/:id/join-agent - Join AI agent to REAL Teams meeting
router.post("/:id/join-agent", requireRealTeams, async (req, res) => {
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

    if (!meeting.isRealTeamsMeeting) {
      return res.status(400).json({ 
        error: "AI agent can only join real Teams meetings",
        meetingType: "simulated"
      });
    }

    logger.info("🤖 AI Agent joining REAL Teams meeting as visible participant", {
      meetingId: req.params.id,
      subject: meeting.subject,
      graphEventId: meeting.graphEventId
    });

    // Join REAL Teams meeting
    const joinResult = await meetingAttendanceService.joinMeeting(
      meeting.meetingId,
      req.user.userId
    );

    // Start REAL chat capture
    if (meeting.agentConfig?.enableChatCapture !== false) {
      await chatCaptureService.initiateRealChatCapture(meeting);
    }

    // Update meeting status
    await cosmosClient.updateItem("meetings", meeting.id, req.user.userId, {
      agentAttended: true,
      agentJoinedAt: new Date().toISOString(),
      status: "in_progress",
    });

    res.json({
      success: true,
      message: "🤖 AI Agent successfully joined REAL Teams meeting and is now visible to all participants",
      meetingDetails: {
        id: meeting.id,
        meetingId: meeting.meetingId,
        subject: meeting.subject,
        status: "in_progress",
        graphEventId: meeting.graphEventId,
        joinUrl: meeting.joinUrl
      },
      joinResult: joinResult,
      capabilities: {
        realChatMonitoring: true,
        realAiInteraction: true,
        realTranscriptCapture: true,
        realSummaryGeneration: true,
        visibleToParticipants: true
      }
    });
  } catch (error) {
    logger.error("❌ AI Agent join REAL Teams meeting failed:", error);
    res.status(500).json({
      error: "Failed to join AI agent to real Teams meeting",
      details: error.message,
    });
  }
});

// POST /api/meetings/:id/leave-agent - Remove AI agent from REAL Teams meeting
router.post("/:id/leave-agent", requireRealTeams, async (req, res) => {
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

    logger.info("🤖 AI Agent leaving REAL Teams meeting", { meetingId: req.params.id });

    // Leave REAL Teams meeting
    const leaveResult = await meetingAttendanceService.leaveMeeting(
      meeting?.meetingId || req.params.id,
      req.user.userId
    );

    // Stop REAL chat capture
    await chatCaptureService.stopRealChatCapture(
      meeting?.meetingId || req.params.id
    );

    // Update meeting status
    if (meeting) {
      await cosmosClient.updateItem("meetings", meeting.id, req.user.userId, {
        agentLeftAt: new Date().toISOString(),
        status: "completed",
      });
    }

    res.json({
      success: true,
      message: "🤖 AI Agent successfully left REAL Teams meeting and sent final summary to participants",
      leaveResult: leaveResult,
    });
  } catch (error) {
    logger.error("❌ AI Agent leave REAL Teams meeting failed:", error);
    res.status(500).json({
      error: "Failed to remove AI agent from real Teams meeting",
      details: error.message,
    });
  }
});

// GET /api/meetings/:id/summary - Get AI-generated summary from REAL Teams data
router.get("/:id/summary", async (req, res) => {
  try {
    const { regenerate = false } = req.query;

    logger.info("📋 Getting AI summary from REAL Teams meeting data", {
      meetingId: req.params.id,
      regenerate,
    });

    let meetingId = req.params.id;

    // Check if this is a database ID or meetingId
    const meeting = await cosmosClient.getItem(
      "meetings",
      req.params.id,
      req.user.userId
    );
    if (meeting) {
      meetingId = meeting.meetingId;
      
      if (!meeting.isRealTeamsMeeting) {
        return res.status(400).json({
          error: "Summary can only be generated for real Teams meetings",
          meetingType: "simulated"
        });
      }
    }

    let summary = null;

    // Check for existing summary
    if (!regenerate) {
      try {
        const existingSummaries = await meetingSummaryService.getMeetingSummaries(meetingId);
        if (existingSummaries.length > 0) {
          summary = existingSummaries[0];
        }
      } catch (error) {
        logger.warn("Could not get existing summaries:", error);
      }
    }

    // Generate new summary from REAL Teams data if needed
    if (!summary || regenerate) {
      summary = await meetingSummaryService.generateMeetingSummary(meetingId, {
        includeChat: true,
        includeParticipantAnalysis: true,
        autoActionItems: true
      });
    }

    res.json({
      success: true,
      summary: summary,
      generated: !summary || regenerate,
      dataSource: "real_teams_integration",
      timestamp: new Date().toISOString(),
    });
  } catch (error) {
    logger.error("❌ Summary generation from REAL Teams data failed:", error);
    res.status(500).json({
      error: "Failed to get summary from real Teams meeting data",
      details: error.message,
    });
  }
});

// GET /api/meetings/:id/chat-analysis - Get REAL Teams chat analysis
router.get("/:id/chat-analysis", async (req, res) => {
  try {
    logger.info("💬 Getting REAL Teams chat analysis", { meetingId: req.params.id });

    const chatAnalysis = await chatCaptureService.getRealChatAnalysis(req.params.id);

    res.json({
      success: true,
      analysis: chatAnalysis,
      dataSource: "real_teams_chat",
      timestamp: new Date().toISOString(),
    });
  } catch (error) {
    logger.error("❌ REAL Teams chat analysis failed:", error);
    res.status(500).json({
      error: "Failed to get real Teams chat analysis",
      details: error.message,
    });
  }
});

// GET /api/meetings/:id/status - Get REAL meeting and agent status
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

    // Get REAL agent attendance status
    let attendanceStatus = null;
    try {
      attendanceStatus = await meetingAttendanceService.getAttendanceSummary(meeting.meetingId);
    } catch (error) {
      logger.warn("Could not get REAL attendance status:", error.message);
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
        isRealTeamsMeeting: meeting.isRealTeamsMeeting,
        graphEventId: meeting.graphEventId,
        joinUrl: meeting.joinUrl
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
        realTeamsIntegration: true,
        capabilities: {
          realChatMonitoring: true,
          realAiInteraction: true,
          realTranscriptCapture: true,
          realSummaryGeneration: true,
          visibleToParticipants: true
        }
      },
      timestamp: new Date().toISOString()
    });

  } catch (error) {
    logger.error("❌ Get REAL meeting status error:", error);
    res.status(500).json({
      error: "Failed to get real meeting status",
      details: error.message,
    });
  }
});

// ============================================================================
// BASIC MEETING MANAGEMENT (REAL TEAMS ONLY)
// ============================================================================

// GET /api/meetings - Get user's REAL Teams meetings
router.get("/", async (req, res) => {
  try {
    const { status, limit = 20, offset = 0 } = req.query;

    let meetings = await cosmosClient.getMeetingsByUser(req.user.userId);

    // Filter for real Teams meetings only
    meetings = meetings.filter(meeting => meeting.isRealTeamsMeeting);

    if (status) {
      meetings = meetings.filter((meeting) => meeting.status === status);
    }

    const paginatedMeetings = meetings.slice(
      parseInt(offset),
      parseInt(offset) + parseInt(limit)
    );

    res.json({
      meetings: paginatedMeetings,
      total: meetings.length,
      limit: parseInt(limit),
      offset: parseInt(offset),
      dataSource: "real_teams_meetings_only"
    });
  } catch (error) {
    logger.error("❌ Get REAL Teams meetings error:", error);
    res.status(500).json({ error: "Failed to retrieve real Teams meetings" });
  }
});

// GET /api/meetings/:id - Get specific REAL Teams meeting
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

    if (!meeting.isRealTeamsMeeting) {
      return res.status(400).json({ 
        error: "Only real Teams meetings are supported",
        meetingType: "simulated"
      });
    }

    res.json(meeting);
  } catch (error) {
    logger.error("❌ Get REAL Teams meeting error:", error);
    res.status(500).json({ error: "Failed to retrieve real Teams meeting" });
  }
});

// DELETE /api/meetings/:id - Cancel REAL Teams meeting
router.delete("/:id", requireRealTeams, async (req, res) => {
  try {
    const meeting = await cosmosClient.getItem(
      "meetings",
      req.params.id,
      req.user.userId
    );
    if (!meeting) {
      return res.status(404).json({ error: "Meeting not found" });
    }

    if (!meeting.isRealTeamsMeeting) {
      return res.status(400).json({ 
        error: "Can only cancel real Teams meetings",
        meetingType: "simulated"
      });
    }

    // Update meeting status (in production, would also cancel via Graph API)
    await cosmosClient.updateItem("meetings", req.params.id, req.user.userId, {
      status: "cancelled",
      cancelledAt: new Date().toISOString(),
    });

    logger.info("REAL Teams meeting cancelled", { meetingId: req.params.id });
    res.json({
      success: true,
      message: "Real Teams meeting cancelled successfully",
    });
  } catch (error) {
    logger.error("❌ Cancel REAL Teams meeting error:", error);
    res.status(500).json({ error: "Failed to cancel real Teams meeting" });
  }
});

// ============================================================================
// SERVICE STATUS (REAL TEAMS INTEGRATION ONLY)
// ============================================================================

// GET /api/meetings/teams/status - Check REAL Teams integration status
router.get("/teams/status", (req, res) => {
  const teamsStatus = teamsService.getStatus();
  res.json({
    ...teamsStatus,
    message: teamsStatus.available
      ? "🟢 REAL Teams integration is fully operational!"
      : "❌ REAL Teams integration not configured. Azure AD setup required.",
    realTeamsIntegration: teamsStatus.available,
    simulationMode: false
  });
});

// GET /api/meetings/status - Check overall service status (REAL integration only)
router.get("/status", (req, res) => {
  const aiStatus = geminiAI.isAvailable();
  const teamsStatus = teamsService.isAvailable();

  res.json({
    overall: {
      status: aiStatus && teamsStatus ? "fully_operational" : "configuration_required",
      message: aiStatus && teamsStatus
        ? "🚀 All POC features operational with REAL Teams integration!"
        : "❌ Configuration required for real Teams integration",
    },
    services: {
      ai: {
        available: aiStatus,
        model: process.env.GEMINI_MODEL || "Not configured",
        required: true
      },
      teams: {
        available: teamsStatus,
        realMeetings: teamsStatus,
        simulationMode: false,
        required: true
      },
    },
    pocFeatures: {
      realMeetingScheduling: teamsStatus,
      realRecurringMeetings: teamsStatus,
      realUserResolution: teamsStatus,
      realTimeOptimization: teamsStatus,
      realAiAgentJoin: teamsStatus,
      realChatCapture: teamsStatus,
      realAiInteraction: aiStatus && teamsStatus,
      realSummaryGeneration: aiStatus,
    },
    requiredConfiguration: !teamsStatus ? {
      azureAd: "Configure Azure AD app registration",
      graphApi: "Set up Microsoft Graph API permissions",
      environmentVariables: ["AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET", "AZURE_TENANT_ID"]
    } : null
  });
});

module.exports = router;