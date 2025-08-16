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
const { default: axios } = require("axios");

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


// Add these endpoints to your meetings.js router

// Add these endpoints to your meetings.js router - REAL TEAMS ONLY

// POST /api/meetings/validate-teams-users - Validate that attendees are real Teams users
router.post("/validate-teams-users", requireRealTeams, async (req, res) => {
  try {
    const { attendees } = req.body;

    if (!attendees || !Array.isArray(attendees) || attendees.length === 0) {
      return res.status(400).json({
        error: "An array of attendee emails is required"
      });
    }

    logger.info("🔍 Validating real Teams users", { attendees });

    const validation = await teamsService.validateTeamsUsers(attendees);

    res.json({
      success: true,
      validation: validation,
      realTeamsIntegration: true,
      message: validation.allValid 
        ? `✅ All ${validation.validUsers.length} users are valid Teams members!`
        : `❌ ${validation.invalidUsers.length} invalid users found - only real Teams members allowed`
    });

  } catch (error) {
    logger.error("❌ Validate Teams users error:", error);
    res.status(500).json({
      error: "Failed to validate Teams users",
      details: error.message
    });
  }
});

// POST /api/meetings/check-availability - Check if REAL Teams attendees are available
router.post("/check-availability", requireRealTeams, async (req, res) => {
  try {
    const { attendees, startTime, endTime } = req.body;

    if (!attendees || !Array.isArray(attendees) || attendees.length === 0) {
      return res.status(400).json({
        error: "An array of attendee emails is required"
      });
    }

    if (!startTime || !endTime) {
      return res.status(400).json({
        error: "startTime and endTime are required"
      });
    }

    logger.info("🔍 Checking REAL Teams availability for attendees", {
      attendees: attendees,
      timeSlot: `${startTime} to ${endTime}`
    });

    // First validate that all attendees are real Teams users
    const validation = await teamsService.validateTeamsUsers(attendees);
    
    if (!validation.allValid) {
      return res.status(400).json({
        error: "Some attendees are not valid Teams users",
        invalidUsers: validation.invalidUsers,
        validUsers: validation.validUsers,
        message: "Only real Teams organization members are allowed"
      });
    }

    const availability = await teamsService.checkTimeSlotAvailability(
      attendees, 
      startTime, 
      endTime
    );

    res.json({
      success: true,
      mode: "real_teams_only",
      ...availability,
      userValidation: validation,
      realCalendarData: true,
      message: availability.allAvailable 
        ? "✅ All real Teams attendees are available for this time slot!" 
        : "❌ Some real Teams attendees are not available - check conflicts"
    });

  } catch (error) {
    logger.error("❌ Check REAL Teams availability error:", error);
    res.status(500).json({
      error: "Failed to check real Teams attendee availability",
      details: error.message
    });
  }
});

// POST /api/meetings/find-available-slots - Find available time slots for REAL Teams attendees
router.post("/find-available-slots", requireRealTeams, async (req, res) => {
  try {
    const { 
      attendees, 
      duration = 30, 
      searchDays = 7
    } = req.body;

    if (!attendees || !Array.isArray(attendees) || attendees.length === 0) {
      return res.status(400).json({
        error: "An array of attendee emails is required"
      });
    }

    logger.info("🔍 Finding available time slots for REAL Teams users", {
      attendees: attendees.length,
      duration: duration,
      searchDays: searchDays
    });

    // First validate that all attendees are real Teams users
    const validation = await teamsService.validateTeamsUsers(attendees);
    
    if (!validation.allValid) {
      return res.status(400).json({
        error: "Some attendees are not valid Teams users",
        invalidUsers: validation.invalidUsers,
        validUsers: validation.validUsers,
        message: "Only real Teams organization members are allowed"
      });
    }

    const availableSlots = await teamsService.findAvailableTimeSlots(
      attendees, 
      duration, 
      searchDays
    );

    res.json({
      success: true,
      mode: "real_teams_only",
      availableSlots: availableSlots,
      searchCriteria: { 
        attendees: attendees.length, 
        duration, 
        searchDays 
      },
      userValidation: validation,
      realCalendarData: true,
      message: availableSlots.length > 0 
        ? `✅ Found ${availableSlots.length} available time slots for real Teams users!`
        : "❌ No available time slots found for all real Teams attendees"
    });

  } catch (error) {
    logger.error("❌ Find available slots for REAL Teams error:", error);
    res.status(500).json({
      error: "Failed to find available time slots for real Teams users",
      details: error.message
    });
  }
});

// GET /api/meetings/free-busy/:email - Get free/busy info for specific REAL Teams user
router.get("/free-busy/:email", requireRealTeams, async (req, res) => {
  try {
    const { email } = req.params;
    const { 
      startTime = moment().startOf('day').toISOString(),
      endTime = moment().endOf('day').toISOString()
    } = req.query;

    logger.info(`🔍 Getting free/busy info for REAL Teams user: ${email}`);

    // First validate that the user is a real Teams user
    const validation = await teamsService.validateTeamsUsers([email]);
    
    if (!validation.allValid) {
      return res.status(404).json({
        error: "User is not a valid Teams member",
        userValidation: validation,
        message: `${email} is not found in your Teams organization`
      });
    }

    const userCalendar = await teamsService.getUserCalendarEvents(email, startTime, endTime);

    res.json({
      success: true,
      mode: "real_teams_only",
      email: email,
      timeRange: { start: startTime, end: endTime },
      ...userCalendar,
      userValidation: validation,
      realCalendarData: true,
      message: `Real Teams calendar information retrieved for ${email}`
    });

  } catch (error) {
    logger.error("❌ Get free/busy for REAL Teams user error:", error);
    res.status(500).json({
      error: "Failed to get free/busy information for real Teams user",
      details: error.message
    });
  }
});

// POST /api/meetings/smart-schedule - AI-powered meeting scheduling for REAL Teams users only
router.post("/smart-schedule", requireRealTeams, async (req, res) => {
  try {
    const {
      subject,
      description,
      attendees,
      duration = 30,
      preferredDays = 7,
      autoSchedule = false
    } = req.body;

    if (!subject || !attendees || attendees.length === 0) {
      return res.status(400).json({
        error: "Subject and attendees are required"
      });
    }

    logger.info("🤖 Smart scheduling meeting for REAL Teams users", {
      subject: subject,
      attendees: attendees.length,
      duration: duration
    });

    // First validate that all attendees are real Teams users
    const validation = await teamsService.validateTeamsUsers(attendees);
    
    if (!validation.allValid) {
      return res.status(400).json({
        error: "Some attendees are not valid Teams users",
        invalidUsers: validation.invalidUsers,
        validUsers: validation.validUsers,
        message: "Only real Teams organization members can be invited to meetings"
      });
    }

    // Step 1: Find available slots
    const availabilityResponse = await teamsService.findAvailableTimeSlots(
      attendees, 
      duration, 
      preferredDays
    );

    if (availabilityResponse.length === 0) {
      return res.json({
        success: false,
        message: "❌ No available time slots found for all real Teams attendees",
        userValidation: validation,
        suggestions: [
          "Try reducing the number of attendees",
          "Increase the search duration to more days",
          "Consider a shorter meeting duration",
          "Check if all attendees have calendar permissions"
        ]
      });
    }

    // Step 2: Get the best time slot (highest confidence, earliest time)
    const bestSlot = availabilityResponse
      .filter(slot => slot.allAttendeesAvailable)
      .sort((a, b) => {
        // Sort by confidence first, then by earliest time
        if (a.confidence !== b.confidence) {
          return a.confidence === 'high' ? -1 : 1;
        }
        return new Date(a.start) - new Date(b.start);
      })[0];

    if (!bestSlot) {
      return res.json({
        success: false,
        message: "❌ No time slots where all real Teams attendees are available",
        partialOptions: availabilityResponse.slice(0, 3),
        userValidation: validation,
        suggestions: [
          "Consider making some attendees optional", 
          "Schedule multiple smaller meetings",
          "Try a different time range"
        ]
      });
    }

    const scheduleRecommendation = {
      recommendedSlot: bestSlot,
      alternativeSlots: availabilityResponse.slice(1, 4),
      meetingDetails: {
        subject: subject,
        description: description,
        startTime: bestSlot.start,
        endTime: bestSlot.end,
        attendees: attendees,
        duration: duration
      }
    };

    // Step 3: Auto-schedule if requested
    if (autoSchedule) {
      try {
        const meetingResult = await teamsService.createTeamsMeeting({
          subject: subject,
          description: description,
          startTime: bestSlot.start,
          endTime: bestSlot.end,
          attendees: attendees
        });

        return res.json({
          success: true,
          autoScheduled: true,
          meeting: meetingResult,
          schedulingDetails: scheduleRecommendation,
          userValidation: validation,
          realTeamsIntegration: true,
          message: "🚀 Meeting automatically scheduled with real Teams users at optimal time!"
        });

      } catch (scheduleError) {
        logger.error("❌ Auto-schedule failed:", scheduleError);
        // Fall through to return recommendation without auto-scheduling
      }
    }

    res.json({
      success: true,
      autoScheduled: false,
      recommendation: scheduleRecommendation,
      userValidation: validation,
      realTeamsIntegration: true,
      message: "✅ Optimal meeting time found for real Teams users! Use the recommended slot to create your meeting.",
      actions: {
        createMeeting: `/api/meetings/create`,
        checkOtherSlots: `/api/meetings/find-available-slots`
      }
    });

  } catch (error) {
    logger.error("❌ Smart schedule for REAL Teams error:", error);
    res.status(500).json({
      error: "Failed to perform smart scheduling for real Teams users",
      details: error.message
    });
  }
});

// GET /api/meetings/teams/users - Get list of real Teams users for testing
router.get("/teams/users", requireRealTeams, async (req, res) => {
  try {
    const { limit = 20, search } = req.query;
    
    logger.info(`📋 Getting real Teams users (limit: ${limit})`);

    let users;
    if (search) {
      users = await teamsService.findTeamMembers(search);
    } else {
      users = await teamsService.getAllTeamMembers(parseInt(limit));
    }
    
    res.json({
      success: true,
      mode: "real_teams_only",
      users: users,
      total: users.length,
      realTeamsDirectory: true,
      message: `Retrieved ${users.length} real Teams users from organization directory`
    });
    
  } catch (error) {
    logger.error("❌ Get real Teams users failed:", error);
    res.status(500).json({
      error: "Failed to get real Teams users",
      details: error.message
    });
  }
});// Add these endpoints to your meetings.js router

// POST /api/meetings/check-availability - Check if attendees are available for specific time
router.post("/check-availability", async (req, res) => {
  try {
    const { attendees, startTime, endTime } = req.body;

    if (!attendees || !Array.isArray(attendees) || attendees.length === 0) {
      return res.status(400).json({
        error: "An array of attendee emails is required"
      });
    }

    if (!startTime || !endTime) {
      return res.status(400).json({
        error: "startTime and endTime are required"
      });
    }

    logger.info("🔍 Checking availability for attendees", {
      attendees: attendees,
      timeSlot: `${startTime} to ${endTime}`
    });

    if (!teamsService.isAvailable()) {
      return res.json({
        success: true,
        mode: "simulated",
        timeSlot: { start: startTime, end: endTime },
        allAvailable: Math.random() > 0.3,
        attendeeStatus: attendees.map(email => ({
          email: email,
          available: Math.random() > 0.4,
          status: Math.random() > 0.4 ? 'free' : 'busy',
          conflicts: []
        })),
        message: "Teams integration not available - showing simulated availability",
        realCalendarData: false
      });
    }

    const availability = await teamsService.checkTimeSlotAvailability(
      attendees, 
      startTime, 
      endTime
    );

    res.json({
      success: true,
      mode: "real_teams",
      ...availability,
      realCalendarData: true,
      message: availability.allAvailable 
        ? "✅ All attendees are available for this time slot!" 
        : "❌ Some attendees are not available - check conflicts"
    });

  } catch (error) {
    logger.error("❌ Check availability error:", error);
    res.status(500).json({
      error: "Failed to check attendee availability",
      details: error.message
    });
  }
});

// POST /api/meetings/find-available-slots - Find available time slots for attendees
router.post("/find-available-slots", async (req, res) => {
  try {
    const { 
      attendees, 
      duration = 30, 
      searchDays = 7,
      preferredTimes = [] // Optional: ["09:00", "14:00"] 
    } = req.body;

    if (!attendees || !Array.isArray(attendees) || attendees.length === 0) {
      return res.status(400).json({
        error: "An array of attendee emails is required"
      });
    }

    logger.info("🔍 Finding available time slots", {
      attendees: attendees.length,
      duration: duration,
      searchDays: searchDays
    });

    if (!teamsService.isAvailable()) {
      // Generate simulated available slots
      const slots = [];
      const startDate = moment().add(1, 'hour');
      
      for (let i = 0; i < 5; i++) {
        const slotStart = moment(startDate).add(i * 2, 'hours');
        slots.push({
          start: slotStart.toISOString(),
          end: moment(slotStart).add(duration, 'minutes').toISOString(),
          confidence: Math.random() > 0.3 ? 'high' : 'medium',
          allAttendeesAvailable: Math.random() > 0.2
        });
      }

      return res.json({
        success: true,
        mode: "simulated",
        availableSlots: slots,
        searchCriteria: { attendees, duration, searchDays },
        realCalendarData: false,
        message: "Teams integration not available - showing simulated time slots"
      });
    }

    const availableSlots = await teamsService.findAvailableTimeSlots(
      attendees, 
      duration, 
      searchDays
    );

    res.json({
      success: true,
      mode: "real_teams",
      availableSlots: availableSlots,
      searchCriteria: { 
        attendees: attendees.length, 
        duration, 
        searchDays 
      },
      realCalendarData: true,
      message: availableSlots.length > 0 
        ? `✅ Found ${availableSlots.length} available time slots!`
        : "❌ No available time slots found for all attendees"
    });

  } catch (error) {
    logger.error("❌ Find available slots error:", error);
    res.status(500).json({
      error: "Failed to find available time slots",
      details: error.message
    });
  }
});

// GET /api/meetings/free-busy/:email - Get free/busy info for specific user
router.get("/free-busy/:email", async (req, res) => {
  try {
    const { email } = req.params;
    const { 
      startTime = moment().startOf('day').toISOString(),
      endTime = moment().endOf('day').toISOString()
    } = req.query;

    logger.info(`🔍 Getting free/busy info for ${email}`);

    if (!teamsService.isAvailable()) {
      return res.json({
        success: true,
        mode: "simulated",
        email: email,
        timeRange: { start: startTime, end: endTime },
        freeBusyStatus: Math.random() > 0.5 ? 'free' : 'busy',
        busyTimes: Math.random() > 0.5 ? [] : [{
          start: moment(startTime).add(2, 'hours').toISOString(),
          end: moment(startTime).add(3, 'hours').toISOString(),
          subject: "Existing Meeting"
        }],
        realCalendarData: false,
        message: "Teams integration not available - showing simulated data"
      });
    }

    const freeBusyInfo = await teamsService.getFreeBusyInfo([email], startTime, endTime);
    const userInfo = freeBusyInfo[0];

    res.json({
      success: true,
      mode: "real_teams",
      email: email,
      timeRange: { start: startTime, end: endTime },
      ...userInfo,
      realCalendarData: true,
      message: `Free/busy information retrieved for ${email}`
    });

  } catch (error) {
    logger.error("❌ Get free/busy error:", error);
    res.status(500).json({
      error: "Failed to get free/busy information",
      details: error.message
    });
  }
});

// POST /api/meetings/smart-schedule - AI-powered meeting scheduling
router.post("/smart-schedule", async (req, res) => {
  try {
    const {
      subject,
      description,
      attendees,
      duration = 30,
      preferredDays = 7,
      autoSchedule = false
    } = req.body;

    if (!subject || !attendees || attendees.length === 0) {
      return res.status(400).json({
        error: "Subject and attendees are required"
      });
    }

    logger.info("🤖 Smart scheduling meeting", {
      subject: subject,
      attendees: attendees.length,
      duration: duration
    });

    // Step 1: Find available slots
    const availabilityResponse = await teamsService.findAvailableTimeSlots(
      attendees, 
      duration, 
      preferredDays
    );

    if (availabilityResponse.length === 0) {
      return res.json({
        success: false,
        message: "❌ No available time slots found for all attendees",
        suggestions: [
          "Try reducing the number of attendees",
          "Increase the search duration to more days",
          "Consider a shorter meeting duration"
        ]
      });
    }

    // Step 2: Get the best time slot (highest confidence, earliest time)
    const bestSlot = availabilityResponse
      .filter(slot => slot.allAttendeesAvailable)
      .sort((a, b) => {
        // Sort by confidence first, then by earliest time
        if (a.confidence !== b.confidence) {
          return a.confidence === 'high' ? -1 : 1;
        }
        return new Date(a.start) - new Date(b.start);
      })[0];

    if (!bestSlot) {
      return res.json({
        success: false,
        message: "❌ No time slots where all attendees are available",
        partialOptions: availabilityResponse.slice(0, 3),
        suggestions: ["Consider optional attendees", "Schedule multiple smaller meetings"]
      });
    }

    const scheduleRecommendation = {
      recommendedSlot: bestSlot,
      alternativeSlots: availabilityResponse.slice(1, 4),
      meetingDetails: {
        subject: subject,
        description: description,
        startTime: bestSlot.start,
        endTime: bestSlot.end,
        attendees: attendees,
        duration: duration
      }
    };

    // Step 3: Auto-schedule if requested
    if (autoSchedule && teamsService.isAvailable()) {
      try {
        const meetingResult = await teamsService.createTeamsMeeting({
          subject: subject,
          description: description,
          startTime: bestSlot.start,
          endTime: bestSlot.end,
          attendees: attendees
        });

        return res.json({
          success: true,
          autoScheduled: true,
          meeting: meetingResult,
          schedulingDetails: scheduleRecommendation,
          message: "🚀 Meeting automatically scheduled at optimal time!"
        });

      } catch (scheduleError) {
        logger.error("❌ Auto-schedule failed:", scheduleError);
        // Fall through to return recommendation without auto-scheduling
      }
    }

    res.json({
      success: true,
      autoScheduled: false,
      recommendation: scheduleRecommendation,
      message: "✅ Optimal meeting time found! Use the recommended slot to create your meeting.",
      actions: {
        createMeeting: `/api/meetings/create`,
        checkOtherSlots: `/api/meetings/find-available-slots`
      }
    });

  } catch (error) {
    logger.error("❌ Smart schedule error:", error);
    res.status(500).json({
      error: "Failed to perform smart scheduling",
      details: error.message
    });
  }
});





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



// Add this debug endpoint to your meetings route file (meetings.js)

// GET /api/meetings/debug/user-lookup - Debug user lookup
router.get("/debug/user-lookup", async (req, res) => {
  try {
    const organizerEmail = process.env.MEETING_ORGANIZER_EMAIL || 'support@legacynote.ai';
    
    if (!teamsService.isAvailable()) {
      return res.json({
        error: "Teams service not available",
        organizerEmail: organizerEmail,
        configured: false
      });
    }

    const authService = require('../services/authService');
    const accessToken = await authService.getAppOnlyToken();

    logger.info(`🔍 Debug: Looking up user ${organizerEmail}`);

    // Try direct lookup
    let directLookup = null;
    try {
      const response = await axios.get(
        `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(organizerEmail)}`,
        { headers: { 'Authorization': `Bearer ${accessToken}` } }
      );
      directLookup = {
        success: true,
        user: {
          id: response.data.id,
          displayName: response.data.displayName,
          userPrincipalName: response.data.userPrincipalName,
          mail: response.data.mail
        }
      };
    } catch (error) {
      directLookup = {
        success: false,
        error: error.response?.data || error.message
      };
    }

    // Try search lookup
    let searchLookup = null;
    try {
      const response = await axios.get(
        `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${organizerEmail}' or mail eq '${organizerEmail}'&$select=id,displayName,userPrincipalName,mail`,
        { headers: { 'Authorization': `Bearer ${accessToken}` } }
      );
      searchLookup = {
        success: true,
        users: response.data.value
      };
    } catch (error) {
      searchLookup = {
        success: false,
        error: error.response?.data || error.message
      };
    }

    // List all users (first 10)
    let allUsers = null;
    try {
      const response = await axios.get(
        `https://graph.microsoft.com/v1.0/users?$top=10&$select=id,displayName,userPrincipalName,mail`,
        { headers: { 'Authorization': `Bearer ${accessToken}` } }
      );
      allUsers = {
        success: true,
        users: response.data.value
      };
    } catch (error) {
      allUsers = {
        success: false,
        error: error.response?.data || error.message
      };
    }

    res.json({
      debug: true,
      organizerEmail: organizerEmail,
      timestamp: new Date().toISOString(),
      lookups: {
        directLookup,
        searchLookup,
        allUsers
      },
      recommendations: [
        `Make sure ${organizerEmail} exists in your Azure AD tenant`,
        "Check if the email format is correct",
        "Verify your app has User.Read.All permission",
        "Check if the user is in the same tenant as your Azure AD app"
      ]
    });

  } catch (error) {
    logger.error("❌ Debug user lookup error:", error);
    res.status(500).json({
      error: "Debug lookup failed",
      details: error.message
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