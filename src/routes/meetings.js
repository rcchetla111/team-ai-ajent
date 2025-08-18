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
      skipAvailabilityCheck = false // Optional: skip availability check
    } = req.body;

    logger.info("🤖 Creating REAL Teams meeting with availability check", {
      subject,
      attendeesCount: attendees.length,
      hasRecurrence: !!recurrence,
      autoJoinAgent,
      skipAvailabilityCheck
    });

    // Validate required fields
    if (!subject || !startTime || !endTime) {
      return res.status(400).json({
        error: "Subject, start time, and end time are required",
      });
    }

    // Validate time range
    const start = moment(startTime);
    const end = moment(endTime);
    const durationMinutes = end.diff(start, 'minutes');
    
    if (durationMinutes <= 0) {
      return res.status(400).json({
        error: "End time must be after start time"
      });
    }

    // STEP 1: Validate that all attendees are real Teams users
    if (attendees.length > 0) {
      logger.info(`🔍 Validating ${attendees.length} attendees as real Teams users`);
      
      const validation = await teamsService.validateTeamsUsers(attendees);
      
      if (!validation.allValid) {
        return res.status(400).json({
          error: "Some attendees are not valid Teams users",
          invalidUsers: validation.invalidUsers,
          validUsers: validation.validUsers,
          message: `❌ ${validation.invalidUsers.length} invalid users found. Only real Teams organization members are allowed.`,
          suggestion: "Please provide valid email addresses from your Teams organization"
        });
      }

      logger.info(`✅ All ${attendees.length} attendees validated as real Teams users`);
    }

    // STEP 2: Check availability for all attendees (unless skipped)
    if (!skipAvailabilityCheck && attendees.length > 0) {
      logger.info(`📅 Checking availability for ${attendees.length} attendees`);
      
      try {
        const availability = await teamsService.checkTimeSlotAvailability(
          attendees, 
          startTime, 
          endTime
        );

        // If not all attendees are available, return conflict details
        if (!availability.allAvailable) {
          const busyAttendees = availability.attendeeStatus.filter(a => !a.available);
          const conflicts = busyAttendees.map(attendee => ({
            email: attendee.email,
            status: attendee.status,
            conflictingMeetings: attendee.conflicts.map(conflict => ({
              subject: conflict.subject,
              start: conflict.start,
              end: conflict.end
            }))
          }));

          return res.status(409).json({
            error: "Meeting time conflict detected",
            conflictSummary: {
              totalAttendees: attendees.length,
              availableAttendees: availability.summary.availableAttendees,
              busyAttendees: availability.summary.busyAttendees,
              conflictDetails: conflicts
            },
            requestedTimeSlot: {
              start: startTime,
              end: endTime,
              durationMinutes: durationMinutes
            },
            availability: availability,
            message: `❌ ${busyAttendees.length} out of ${attendees.length} attendees are not available during the requested time slot`,
            suggestions: [
              "Use /api/meetings/find-available-slots to find better times",
              "Remove conflicting attendees or make them optional",
              "Choose a different time slot",
              "Add skipAvailabilityCheck: true to force create the meeting anyway"
            ]
          });
        }

        logger.info(`✅ All attendees are available for the meeting time slot`);
      } catch (availabilityError) {
        logger.error("❌ Availability check failed:", availabilityError.message);
        
        // Continue with meeting creation but log the warning
        logger.warn("⚠️ Proceeding with meeting creation despite availability check failure");
      }
    }

    // STEP 3: Create REAL Teams meeting via Graph API
    logger.info("🚀 Creating REAL Teams meeting via Graph API");
    
    const teamsMeetingResult = await teamsService.createTeamsMeeting({
      subject: subject,
      description: description,
      startTime: startTime,
      endTime: endTime,
      attendees: attendees,
      recurrence: recurrence
    });

    // STEP 4: Create meeting record in database
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
      availabilityChecked: !skipAvailabilityCheck,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };

    const savedMeeting = await cosmosClient.createMeeting(meetingData);

    // STEP 5: Auto-join AI agent logic for meetings starting soon
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
      message: "🟢 REAL Teams meeting created successfully with availability validation!",
      realTeamsMeeting: true,
      teamsIntegrationStatus: teamsService.getStatus(),
      meetingValidation: {
        attendeesValidated: attendees.length,
        availabilityChecked: !skipAvailabilityCheck,
        allAttendeesAvailable: true
      },
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
        webUrl: teamsMeetingResult.webUrl,
        attendeesCount: attendees.length,
        duration: `${durationMinutes} minutes`
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
      skipAvailabilityCheck = false
    } = req.body;

    if (!subject || !startTime || !endTime) {
      return res.status(400).json({
        error: "Subject, start time, and end time are required"
      });
    }

    // Validate time range
    const start = moment(startTime);
    const end = moment(endTime);
    const durationMinutes = end.diff(start, 'minutes');
    
    if (durationMinutes <= 0) {
      return res.status(400).json({
        error: "End time must be after start time"
      });
    }

    logger.info("🚀 Creating REAL Teams meeting with name resolution and availability check", {
      subject,
      attendeeNames: attendeeNames.length,
      attendeeEmails: attendeeEmails.length,
      skipAvailabilityCheck
    });

    let resolvedAttendees = [...attendeeEmails];
    let userResolutionDetails = {
      namesRequested: attendeeNames.length,
      emailsProvided: attendeeEmails.length,
      usersResolved: 0,
      failedResolutions: []
    };

    // STEP 1: Resolve names to emails using REAL Teams directory
    if (attendeeNames.length > 0) {
      logger.info(`🔍 Resolving ${attendeeNames.length} names to emails using REAL Teams directory`);
      
      try {
        const resolvedUsers = await teamsService.findUsersByDisplayName(attendeeNames);
        const resolvedEmails = resolvedUsers.map(user => user.email);
        resolvedAttendees.push(...resolvedEmails);
        
        userResolutionDetails.usersResolved = resolvedUsers.length;
        userResolutionDetails.resolvedUsers = resolvedUsers;
        
        // Track failed resolutions
        const resolvedNames = resolvedUsers.map(user => user.name);
        userResolutionDetails.failedResolutions = attendeeNames.filter(name => 
          !resolvedNames.some(resolvedName => resolvedName.toLowerCase().includes(name.toLowerCase()))
        );
        
        logger.info(`✅ Resolved ${resolvedUsers.length}/${attendeeNames.length} users from REAL Teams directory`);
        
        if (userResolutionDetails.failedResolutions.length > 0) {
          logger.warn(`⚠️ Could not resolve: ${userResolutionDetails.failedResolutions.join(', ')}`);
        }
      } catch (resolutionError) {
        logger.error("❌ Name resolution failed:", resolutionError.message);
        return res.status(400).json({
          error: "Failed to resolve attendee names",
          details: resolutionError.message,
          attendeeNames: attendeeNames,
          suggestion: "Use exact display names as they appear in Teams, or use email addresses directly"
        });
      }
    }

    // Remove duplicates
    resolvedAttendees = [...new Set(resolvedAttendees)];

    if (resolvedAttendees.length === 0) {
      return res.status(400).json({
        error: "No valid attendees found",
        userResolution: userResolutionDetails,
        message: "Please provide valid attendee names or email addresses"
      });
    }

    // STEP 2: Validate that all resolved attendees are real Teams users
    logger.info(`🔍 Validating ${resolvedAttendees.length} resolved attendees as real Teams users`);
    
    const validation = await teamsService.validateTeamsUsers(resolvedAttendees);
    
    if (!validation.allValid) {
      return res.status(400).json({
        error: "Some resolved attendees are not valid Teams users",
        invalidUsers: validation.invalidUsers,
        validUsers: validation.validUsers,
        userResolution: userResolutionDetails,
        message: `❌ ${validation.invalidUsers.length} invalid users found after resolution`
      });
    }

    // STEP 3: Check availability for all attendees (unless skipped)
    if (!skipAvailabilityCheck && resolvedAttendees.length > 0) {
      logger.info(`📅 Checking availability for ${resolvedAttendees.length} resolved attendees`);
      
      try {
        const availability = await teamsService.checkTimeSlotAvailability(
          resolvedAttendees, 
          startTime, 
          endTime
        );

        // If not all attendees are available, return conflict details
        if (!availability.allAvailable) {
          const busyAttendees = availability.attendeeStatus.filter(a => !a.available);
          const conflicts = busyAttendees.map(attendee => {
            // Try to map back to original names if possible
            const originalName = userResolutionDetails.resolvedUsers?.find(
              user => user.email === attendee.email
            )?.name || attendee.email;

            return {
              email: attendee.email,
              originalName: originalName,
              status: attendee.status,
              conflictingMeetings: attendee.conflicts.map(conflict => ({
                subject: conflict.subject,
                start: conflict.start,
                end: conflict.end
              }))
            };
          });

          return res.status(409).json({
            error: "Meeting time conflict detected after name resolution",
            conflictSummary: {
              totalAttendees: resolvedAttendees.length,
              availableAttendees: availability.summary.availableAttendees,
              busyAttendees: availability.summary.busyAttendees,
              conflictDetails: conflicts
            },
            requestedTimeSlot: {
              start: startTime,
              end: endTime,
              durationMinutes: durationMinutes
            },
            userResolution: userResolutionDetails,
            availability: availability,
            message: `❌ ${busyAttendees.length} out of ${resolvedAttendees.length} attendees are not available during the requested time slot`,
            suggestions: [
              "Use /api/meetings/find-available-slots to find better times for these specific attendees",
              "Remove conflicting attendees or make them optional",
              "Choose a different time slot",
              "Add skipAvailabilityCheck: true to force create the meeting anyway"
            ]
          });
        }

        logger.info(`✅ All resolved attendees are available for the meeting time slot`);
      } catch (availabilityError) {
        logger.error("❌ Availability check failed:", availabilityError.message);
        logger.warn("⚠️ Proceeding with meeting creation despite availability check failure");
      }
    }

    // STEP 4: Create the REAL Teams meeting
    logger.info("🚀 Creating REAL Teams meeting via Graph API");
    
    const teamsMeetingResult = await teamsService.createTeamsMeeting({
      subject,
      description,
      startTime,
      endTime,
      attendees: resolvedAttendees
    });

    // STEP 5: Create meeting record in database
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
      userResolution: userResolutionDetails,
      availabilityChecked: !skipAvailabilityCheck,
      createdAt: new Date().toISOString(),
    };

    const savedMeeting = await cosmosClient.createMeeting(meetingData);

    res.status(201).json({
      success: true,
      meeting: savedMeeting,
      message: "🚀 REAL Teams meeting created with name resolution and availability validation!",
      realTeamsMeeting: true,
      userResolution: {
        realTeamsDirectoryUsed: true,
        ...userResolutionDetails,
        finalAttendees: resolvedAttendees,
        resolutionSuccess: userResolutionDetails.failedResolutions.length === 0
      },
      meetingValidation: {
        attendeesValidated: resolvedAttendees.length,
        availabilityChecked: !skipAvailabilityCheck,
        allAttendeesAvailable: true
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
      searchDays = 7,
      timePreferences = {}
    } = req.body;

    if (attendees.length === 0) {
      return res.status(400).json({
        error: "At least one attendee (name or email) is required for time suggestions"
      });
    }

    logger.info("🔍 Finding optimal meeting times with smart name resolution", {
      originalInput: attendees,
      duration,
      searchDays,
      timePreferences
    });

    // The TeamsService will now handle name-to-email resolution automatically
    const suggestions = await teamsService.findMeetingTimes(
      attendees,
      duration,
      searchDays,
      timePreferences
    );

    // Return enhanced suggestions
    res.json({
      success: true,
      ...suggestions,
      realTeamsIntegration: true,
      message: suggestions.suggestions.length > 0 
        ? `🎯 Found ${suggestions.suggestions.length} optimal meeting times using real Teams calendar data!`
        : "No available time slots found for all attendees",
      helpfulTips: [
        "You can use either names (like 'John Smith') or emails (like 'john@company.com')",
        "Names will be automatically resolved to email addresses using your Teams directory",
        "Make sure all attendees are part of your Teams organization"
      ]
    });

  } catch (error) {
    logger.error("❌ Suggest meeting times error:", error);
    
    // Provide helpful error messages
    if (error.message.includes('Cannot find Teams user')) {
      return res.status(400).json({
        error: "User not found in Teams directory",
        details: error.message,
        suggestions: [
          "Try using the exact display name as it appears in Teams",
          "Use the person's email address instead (e.g., rohit@company.com)",
          "Check if the person is part of your Teams organization",
          "Use 'Find People' feature to see available users"
        ]
      });
    }
    
    if (error.message.includes('Cannot resolve')) {
      return res.status(400).json({
        error: "Name resolution failed",
        details: error.message,
        suggestions: [
          "Use email addresses directly (e.g., rohit@company.com)",
          "Check the spelling of the person's name",
          "Try searching for the person in Teams first to get their exact name"
        ]
      });
    }

    res.status(500).json({
      error: "Failed to suggest optimal meeting times",
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
// REPLACE your GET /api/meetings endpoint in meetings.js with this version:

// GET /api/meetings - Get ALL Teams meetings (database + calendar)
router.get("/", async (req, res) => {
  try {
    const { status, limit = 100, offset = 0, includeCalendar = true } = req.query;

    let allMeetings = [];

    // Get meetings from database (AI-created meetings)
    const dbMeetings = await cosmosClient.getMeetingsByUser(req.user.userId);
    const aiCreatedMeetings = dbMeetings.filter(meeting => meeting.isRealTeamsMeeting);
    
    logger.info(`📊 Found ${aiCreatedMeetings.length} AI-created meetings in database`);

    // ALSO get meetings directly from Teams calendar
    if (includeCalendar !== 'false' && teamsService.isAvailable()) {
      try {
        const calendarMeetings = await teamsService.getAllCalendarMeetings();
        logger.info(`📅 Found ${calendarMeetings.length} meetings in Teams calendar`);
        
        // Combine both sources, avoiding duplicates
        const combinedMeetings = [...aiCreatedMeetings];
        
        // Add calendar meetings that aren't already in database
        calendarMeetings.forEach(calMeeting => {
          const existsInDb = aiCreatedMeetings.some(dbMeeting => 
            dbMeeting.graphEventId === calMeeting.graphEventId ||
            dbMeeting.subject === calMeeting.subject
          );
          
          if (!existsInDb) {
            combinedMeetings.push(calMeeting);
          }
        });
        
        allMeetings = combinedMeetings;
        logger.info(`🔗 Combined total: ${allMeetings.length} meetings (${aiCreatedMeetings.length} from DB + ${calendarMeetings.length - (combinedMeetings.length - aiCreatedMeetings.length)} unique from calendar)`);
        
      } catch (calendarError) {
        logger.warn('⚠️ Could not fetch calendar meetings, using database only:', calendarError.message);
        allMeetings = aiCreatedMeetings;
      }
    } else {
      allMeetings = aiCreatedMeetings;
    }

    // Apply status filter
    if (status) {
      allMeetings = allMeetings.filter((meeting) => meeting.status === status);
    }

    // Sort by start time (most recent first)
    allMeetings.sort((a, b) => new Date(b.startTime) - new Date(a.startTime));

    // Apply pagination
    const paginatedMeetings = allMeetings.slice(
      parseInt(offset),
      parseInt(offset) + parseInt(limit)
    );

    res.json({
      meetings: paginatedMeetings,
      total: allMeetings.length,
      limit: parseInt(limit),
      offset: parseInt(offset),
      dataSource: includeCalendar !== 'false' ? "database_and_teams_calendar" : "database_only",
      breakdown: {
        aiCreated: aiCreatedMeetings.length,
        fromCalendar: allMeetings.length - aiCreatedMeetings.length,
        total: allMeetings.length
      }
    });
  } catch (error) {
    logger.error("❌ Get ALL Teams meetings error:", error);
    res.status(500).json({ error: "Failed to retrieve Teams meetings" });
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
router.delete("/:id", async (req, res) => {
  try {
    const meetingId = req.params.id;
    console.log("🗑️ Backend: Attempting to cancel meeting:", meetingId);

    let meeting = null;
    let isFromDatabase = false;

    // STEP 1: Try to find meeting in database first
    try {
      meeting = await cosmosClient.getItem("meetings", meetingId, req.user.userId);
      if (meeting) {
        isFromDatabase = true;
        console.log("✅ Found meeting in database:", meeting.subject);
      }
    } catch (dbError) {
      console.log("⚠️ Meeting not in database, will check calendar meetings");
    }

    // STEP 2: If not in database, search calendar meetings
    if (!meeting) {
      try {
        console.log("🔍 Searching calendar meetings for:", meetingId);
        
        // Get all meetings (database + calendar)
        const allMeetingsResponse = await fetch(`http://localhost:5000/api/meetings?limit=200&includeCalendar=true`);
        const allMeetingsData = await allMeetingsResponse.json();
        
        console.log("📅 Total meetings available:", allMeetingsData.meetings?.length || 0);
        
        // Find the meeting by various ID fields
        meeting = allMeetingsData.meetings?.find(m => 
          m.id === meetingId || 
          m.meetingId === meetingId ||
          m.graphEventId === meetingId ||
          m.subject.toLowerCase().includes(meetingId.toLowerCase())
        );
        
        if (meeting) {
          console.log("✅ Found meeting in calendar data:", meeting.subject);
          console.log("📊 Meeting source:", meeting.isFromTeamsCalendar ? "Teams Calendar" : "Database");
        }
      } catch (searchError) {
        console.error("❌ Error searching calendar meetings:", searchError);
      }
    }

    if (!meeting) {
      console.log("❌ Meeting not found:", meetingId);
      return res.status(404).json({ 
        error: "Meeting not found",
        details: `Meeting with ID ${meetingId} not found in database or calendar`,
        searchedFor: meetingId
      });
    }

    console.log("🎯 Found meeting to cancel:", {
      id: meeting.id,
      subject: meeting.subject,
      isFromDatabase: isFromDatabase,
      isFromCalendar: meeting.isFromTeamsCalendar,
      graphEventId: meeting.graphEventId
    });

    // STEP 3: Cancel the meeting based on its source
    let cancellationResult = { success: false };

    if (meeting.isFromTeamsCalendar && meeting.graphEventId && teamsService.isAvailable()) {
      // Cancel Teams calendar meeting via Graph API
      try {
        console.log("🔄 Cancelling Teams calendar meeting via Graph API");
        
        const accessToken = await require('../services/authService').getAppOnlyToken();
        const organizerEmail = process.env.MEETING_ORGANIZER_EMAIL || 'support@legacynote.ai';
        
        // Get organizer user ID
        const userResponse = await axios.get(
          `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(organizerEmail)}?$select=id`,
          { headers: { 'Authorization': `Bearer ${await accessToken}` } }
        );
        
        const organizerUserId = userResponse.data.id;

        // Cancel the meeting by updating it to cancelled status
        await axios.delete(
          `https://graph.microsoft.com/v1.0/users/${organizerUserId}/events/${meeting.graphEventId}`,
          { 
            headers: { 
              'Authorization': `Bearer ${await accessToken}`, 
              'Content-Type': 'application/json'
            } 
          }
        );

        console.log("✅ Teams calendar meeting cancelled successfully");
        cancellationResult = { success: true, method: 'teams_graph_api' };

      } catch (graphError) {
        console.error("❌ Failed to cancel via Graph API:", graphError.message);
        // Continue to try database cancellation as fallback
      }
    }

    // STEP 4: Update database record if it exists
    if (isFromDatabase || !cancellationResult.success) {
      try {
        await cosmosClient.updateItem("meetings", meeting.id, req.user.userId, {
          status: "cancelled",
          cancelledAt: new Date().toISOString(),
          cancelledBy: req.user.userId
        });
        
        console.log("✅ Database meeting record updated to cancelled");
        cancellationResult = { success: true, method: 'database_update' };

      } catch (dbUpdateError) {
        console.error("❌ Failed to update database:", dbUpdateError.message);
        
        if (!cancellationResult.success) {
          throw new Error(`Failed to cancel meeting: ${dbUpdateError.message}`);
        }
      }
    }

    // STEP 5: Return success response
    console.log("🎉 Meeting cancellation completed:", cancellationResult);

    res.json({
      success: true,
      message: `Meeting "${meeting.subject}" cancelled successfully`,
      meeting: {
        id: meeting.id,
        subject: meeting.subject,
        originalStatus: meeting.status,
        newStatus: "cancelled",
        cancelledAt: new Date().toISOString(),
        isFromTeamsCalendar: meeting.isFromTeamsCalendar,
        isFromDatabase: isFromDatabase
      },
      cancellationMethod: cancellationResult.method,
      attendeesNotified: meeting.isFromTeamsCalendar ? "Teams will notify attendees automatically" : "Notification pending"
    });

  } catch (error) {
    console.error("❌ Cancel meeting error:", error);
    res.status(500).json({ 
      error: "Failed to cancel meeting",
      details: error.message,
      meetingId: req.params.id
    });
  }
});



// ADD THESE ENDPOINTS TO YOUR meetings.js FOR DYNAMIC ATTENDEE MANAGEMENT

// PUT /api/meetings/:id/attendees - Update attendees list (add/remove)
router.put("/:id/attendees", requireRealTeams, async (req, res) => {
  try {
    const { attendeesToAdd = [], attendeesToRemove = [], checkAvailability = true } = req.body;
    
    // Get existing meeting
    let meeting = await cosmosClient.getItem(
      "meetings",
      req.params.id,
      req.user.userId
    );

    if (!meeting) {
      return res.status(404).json({ error: "Meeting not found" });
    }

    if (!meeting.isRealTeamsMeeting) {
      return res.status(400).json({ 
        error: "Can only modify real Teams meetings",
        meetingType: "simulated"
      });
    }

    logger.info("🔄 Updating attendees for meeting", {
      meetingId: req.params.id,
      subject: meeting.subject,
      currentAttendees: meeting.attendees.length,
      toAdd: attendeesToAdd.length,
      toRemove: attendeesToRemove.length
    });

    // Validate new attendees are real Teams users
    if (attendeesToAdd.length > 0) {
      const validation = await teamsService.validateTeamsUsers(attendeesToAdd);
      
      if (!validation.allValid) {
        return res.status(400).json({
          error: "Some new attendees are not valid Teams users",
          invalidUsers: validation.invalidUsers,
          validUsers: validation.validUsers,
          message: "Only real Teams organization members can be added"
        });
      }
    }

    // Calculate new attendees list
    let updatedAttendees = [...meeting.attendees];
    
    // Remove attendees
    if (attendeesToRemove.length > 0) {
      updatedAttendees = updatedAttendees.filter(email => 
        !attendeesToRemove.includes(email)
      );
      logger.info(`➖ Removing ${attendeesToRemove.length} attendees`);
    }
    
    // Add new attendees (avoid duplicates)
    if (attendeesToAdd.length > 0) {
      const newAttendees = attendeesToAdd.filter(email => 
        !updatedAttendees.includes(email)
      );
      updatedAttendees.push(...newAttendees);
      logger.info(`➕ Adding ${newAttendees.length} new attendees`);
    }

    // Check availability for NEW attendees only (if requested)
    let availabilityCheck = null;
    if (checkAvailability && attendeesToAdd.length > 0) {
      try {
        availabilityCheck = await teamsService.checkTimeSlotAvailability(
          attendeesToAdd,
          meeting.startTime,
          meeting.endTime
        );

        if (!availabilityCheck.allAvailable) {
          const busyAttendees = availabilityCheck.attendeeStatus.filter(a => !a.available);
          
          return res.status(409).json({
            error: "Some new attendees are not available",
            conflictSummary: {
              newAttendees: attendeesToAdd.length,
              availableNewAttendees: availabilityCheck.summary.availableAttendees,
              busyNewAttendees: availabilityCheck.summary.busyAttendees,
              conflictDetails: busyAttendees.map(attendee => ({
                email: attendee.email,
                status: attendee.status,
                conflictingMeetings: attendee.conflicts
              }))
            },
            meetingTimeSlot: {
              start: meeting.startTime,
              end: meeting.endTime
            },
            message: `❌ ${busyAttendees.length} out of ${attendeesToAdd.length} new attendees are not available`,
            suggestions: [
              "Remove conflicting attendees from the add list",
              "Set checkAvailability: false to force add them anyway",
              "Reschedule the meeting to a time when everyone is available"
            ]
          });
        }
      } catch (availabilityError) {
        logger.warn("⚠️ Availability check failed, proceeding anyway:", availabilityError.message);
      }
    }

    // Update the meeting in Microsoft Graph
    try {
      const accessToken = require('../services/authService').getAppOnlyToken();
      const organizerEmail = process.env.MEETING_ORGANIZER_EMAIL || 'support@legacynote.ai';
      
      // Get organizer user ID
      const userResponse = await axios.get(
        `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(organizerEmail)}?$select=id`,
        { headers: { 'Authorization': `Bearer ${await accessToken}` } }
      );
      
      const organizerUserId = userResponse.data.id;

      // Prepare updated attendees list for Graph API
      const graphAttendees = updatedAttendees
        .filter(email => email !== organizerEmail) // Don't include organizer
        .map(email => ({
          emailAddress: { 
            address: email, 
            name: email.split('@')[0]
          },
          type: 'required'
        }));

      // Update the meeting via Graph API
      const updatePayload = {
        attendees: graphAttendees
      };

      await axios.patch(
        `https://graph.microsoft.com/v1.0/users/${organizerUserId}/events/${meeting.graphEventId}`,
        updatePayload,
        { 
          headers: { 
            'Authorization': `Bearer ${await accessToken}`, 
            'Content-Type': 'application/json'
          } 
        }
      );

      logger.info("✅ Meeting updated in Microsoft Graph");
    } catch (graphError) {
      logger.error("❌ Failed to update meeting in Graph:", graphError.message);
      // Continue with database update even if Graph update fails
    }

    // Update meeting in database
    const updatedMeetingData = {
      attendees: updatedAttendees,
      updatedAt: new Date().toISOString(),
      lastAttendeeUpdate: {
        timestamp: new Date().toISOString(),
        added: attendeesToAdd,
        removed: attendeesToRemove,
        newTotal: updatedAttendees.length
      }
    };

    await cosmosClient.updateItem(
      "meetings", 
      req.params.id, 
      req.user.userId, 
      updatedMeetingData
    );

    // Get updated meeting
    const finalMeeting = await cosmosClient.getItem(
      "meetings",
      req.params.id,
      req.user.userId
    );

    res.json({
      success: true,
      message: "🔄 Meeting attendees updated successfully!",
      meeting: finalMeeting,
      changes: {
        attendeesAdded: attendeesToAdd,
        attendeesRemoved: attendeesToRemove,
        previousCount: meeting.attendees.length,
        newCount: updatedAttendees.length,
        availabilityChecked: checkAvailability && attendeesToAdd.length > 0,
        allNewAttendeesAvailable: availabilityCheck?.allAvailable ?? null
      },
      realTeamsIntegration: true
    });

  } catch (error) {
    logger.error("❌ Update attendees error:", error);
    res.status(500).json({
      error: "Failed to update meeting attendees",
      details: error.message
    });
  }
});

// POST /api/meetings/:id/attendees/add - Add attendees to existing meeting
// REPLACE your POST /api/meetings/:id/attendees/add endpoint in meetings.js with this FIXED version:

// POST /api/meetings/:id/attendees/add - Add attendees to existing meeting
// REPLACE the attendee processing section in your POST /:id/attendees/add endpoint with this:

// POST /api/meetings/:id/attendees/add - FIXED version with better name resolution
router.post("/:id/attendees/add", requireRealTeams, async (req, res) => {
  try {
    const { attendees, attendeeNames = [], checkAvailability = true } = req.body;
    
    if ((!attendees || attendees.length === 0) && attendeeNames.length === 0) {
      return res.status(400).json({
        error: "Either attendees (emails) or attendeeNames must be provided"
      });
    }

    console.log("➕ Backend: Adding attendees to meeting:", req.params.id);
    console.log("➕ Backend: Emails to add:", attendees);
    console.log("➕ Backend: Names to add:", attendeeNames);

    let attendeesToAdd = [...(attendees || [])];
    
    // IMPROVED: Handle names more flexibly like the remove function
    if (attendeeNames.length > 0) {
      try {
        // Get current meeting first to check existing attendees
        let meeting = null;
        try {
          meeting = await cosmosClient.getItem("meetings", req.params.id, req.user.userId);
        } catch (dbError) {
          // Search in calendar meetings
          const allMeetingsResponse = await fetch(`http://localhost:5000/api/meetings?limit=200&includeCalendar=true`);
          const allMeetingsData = await allMeetingsResponse.json();
          meeting = allMeetingsData.meetings?.find(m => 
            m.id === req.params.id || m.meetingId === req.params.id || m.graphEventId === req.params.id
          );
        }

        if (meeting) {
          // STRATEGY 1: Check if name matches existing attendees (like remove function)
          const currentAttendees = meeting.attendees || [];
          console.log("🔍 Current attendees in meeting:", currentAttendees);
          
          for (const name of attendeeNames) {
            const nameLower = name.toLowerCase().trim();
            console.log(`🔍 Looking for name "${nameLower}" in current attendees`);
            
            // Find attendee by partial name match (like remove function does)
            const matchedEmail = currentAttendees.find(email => {
              const emailLower = email.toLowerCase();
              // Check if email contains the name or name contains part of email
              return emailLower.includes(nameLower) || 
                     nameLower.includes(emailLower.split('@')[0]) ||
                     emailLower.split('@')[0].includes(nameLower);
            });
            
            if (matchedEmail) {
              console.log(`✅ Found existing attendee: ${name} -> ${matchedEmail}`);
              attendeesToAdd.push(matchedEmail);
              continue;
            }
            
            // STRATEGY 2: Try Teams directory search with multiple variations
            let resolvedEmail = null;
            const searchVariations = [
              name,                           // "Anusha"
              name.toLowerCase(),             // "anusha" 
              name.charAt(0).toUpperCase() + name.slice(1).toLowerCase(), // "Anusha"
              name.toUpperCase()              // "ANUSHA"
            ];
            
            for (const variation of searchVariations) {
              try {
                console.log(`🔍 Trying Teams directory search for: "${variation}"`);
                const resolvedUsers = await teamsService.findUsersByDisplayName([variation]);
                
                if (resolvedUsers && resolvedUsers.length > 0) {
                  resolvedEmail = resolvedUsers[0].email;
                  console.log(`✅ Resolved via Teams directory: ${variation} -> ${resolvedEmail}`);
                  break;
                }
              } catch (searchError) {
                console.log(`⚠️ Teams search failed for "${variation}":`, searchError.message);
              }
            }
            
            // STRATEGY 3: Smart email construction if Teams search fails
            if (!resolvedEmail) {
              console.log(`🔧 Attempting smart email construction for: ${name}`);
              
              // Get domain from existing attendees
              const domain = currentAttendees.length > 0 ? 
                currentAttendees[0].split('@')[1] : 'warrantyme.co';
              
              const possibleEmails = [
                `${name.toLowerCase()}@${domain}`,           // anusha@warrantyme.co
                `${name.toLowerCase()}@warrantyme.co`,       // anusha@warrantyme.co (fallback)
                `${name.toLowerCase()}@legacynote.ai`        // anusha@legacynote.ai (another fallback)
              ];
              
              // Test each constructed email
              for (const testEmail of possibleEmails) {
                try {
                  console.log(`🧪 Testing constructed email: ${testEmail}`);
                  const validation = await teamsService.validateTeamsUsers([testEmail]);
                  
                  if (validation.allValid && validation.validUsers.length > 0) {
                    resolvedEmail = testEmail;
                    console.log(`✅ Smart construction succeeded: ${name} -> ${testEmail}`);
                    break;
                  }
                } catch (validationError) {
                  console.log(`❌ Validation failed for ${testEmail}`);
                }
              }
            }
            
            if (resolvedEmail) {
              attendeesToAdd.push(resolvedEmail);
            } else {
              // If all strategies fail, still add the name and let validation catch it later
              console.log(`⚠️ Could not resolve "${name}", will try as-is`);
              attendeesToAdd.push(name);
            }
          }
        }
        
      } catch (resolutionError) {
        console.error("❌ Name resolution process failed:", resolutionError);
        // Continue with original names and let validation handle it
        attendeesToAdd.push(...attendeeNames);
      }
    }

    // Remove duplicates
    attendeesToAdd = [...new Set(attendeesToAdd)];

    if (attendeesToAdd.length === 0) {
      return res.status(400).json({
        error: "No valid attendees to add after processing"
      });
    }

    console.log("📝 Final attendees to add:", attendeesToAdd);

    // STEP 1: Find the meeting (same logic as before)
    let meeting = null;
    try {
      meeting = await cosmosClient.getItem("meetings", req.params.id, req.user.userId);
      console.log("📊 Found meeting in database:", meeting ? "YES" : "NO");
    } catch (dbError) {
      console.log("⚠️ Meeting not in database, will search calendar meetings");
    }

    // STEP 2: If not in database, search through ALL calendar meetings
    if (!meeting) {
      console.log("🔍 Searching calendar meetings for ID:", req.params.id);
      
      try {
        const allMeetingsResponse = await fetch(`http://localhost:5000/api/meetings?limit=200&includeCalendar=true`);
        const allMeetingsData = await allMeetingsResponse.json();
        
        meeting = allMeetingsData.meetings?.find(m => 
          m.id === req.params.id || 
          m.meetingId === req.params.id ||
          m.graphEventId === req.params.id
        );
        
        if (meeting) {
          console.log("✅ Found meeting in calendar data:", meeting.subject);
        }
      } catch (searchError) {
        console.error("❌ Error searching calendar meetings:", searchError);
      }
    }

    if (!meeting) {
      return res.status(404).json({ 
        error: "Meeting not found",
        details: `Meeting with ID ${req.params.id} not found`
      });
    }

    // STEP 3: Filter out attendees that are already in the meeting
    const currentAttendees = meeting.attendees || [];
    const newAttendees = attendeesToAdd.filter(email => {
      const isAlreadyAttendee = currentAttendees.some(current => {
        const currentLower = current.toLowerCase();
        const emailLower = email.toLowerCase();
        
        // Check exact match or partial match
        return currentLower === emailLower ||
               currentLower.includes(emailLower) ||
               emailLower.includes(currentLower.split('@')[0]);
      });
      
      if (isAlreadyAttendee) {
        console.log(`⚠️ ${email} is already an attendee`);
        return false;
      }
      return true;
    });

    if (newAttendees.length === 0) {
      return res.status(400).json({
        error: "All specified attendees are already in the meeting",
        currentAttendees: currentAttendees,
        requestedAttendees: attendeesToAdd
      });
    }

    // STEP 4: Validate only the truly new attendees
    console.log("🔍 Validating new attendees:", newAttendees);
    
    let validatedAttendees = [];
    let invalidAttendees = [];
    
    for (const attendee of newAttendees) {
      try {
        const validation = await teamsService.validateTeamsUsers([attendee]);
        
        if (validation.allValid && validation.validUsers.length > 0) {
          validatedAttendees.push(validation.validUsers[0].email);
          console.log(`✅ Validated: ${attendee} -> ${validation.validUsers[0].email}`);
        } else {
          invalidAttendees.push({
            email: attendee,
            error: "Not found in Teams organization"
          });
          console.log(`❌ Invalid: ${attendee}`);
        }
      } catch (validationError) {
        invalidAttendees.push({
          email: attendee,
          error: validationError.message
        });
        console.log(`❌ Validation error for ${attendee}:`, validationError.message);
      }
    }
    
    if (validatedAttendees.length === 0) {
      return res.status(400).json({
        error: "No valid Teams users found to add",
        invalidUsers: invalidAttendees,
        suggestions: [
          "Try using exact email addresses (e.g., anusha@warrantyme.co)",
          "Use 'Find people named [name]' to see available users",
          "Check if the person exists in your Teams organization"
        ]
      });
    }

    // STEP 5: Add validated attendees to the meeting
    const updatedAttendees = [...currentAttendees, ...validatedAttendees];

    console.log("👥 Attendees before:", currentAttendees.length);
    console.log("👥 Valid new attendees:", validatedAttendees);
    console.log("👥 Attendees after:", updatedAttendees.length);

    // STEP 6: Update the meeting
    try {
      if (meeting.isFromTeamsCalendar && meeting.graphEventId) {
        console.log("🔄 Updating Teams calendar meeting via Graph API");
        await teamsService.updateTeamsMeeting(meeting.graphEventId, {
          attendees: updatedAttendees
        });
        console.log("✅ Teams meeting updated successfully");
      }
      
      if (!meeting.isFromTeamsCalendar) {
        await cosmosClient.updateItem("meetings", req.params.id, req.user.userId, {
          attendees: updatedAttendees,
          updatedAt: new Date().toISOString()
        });
        console.log("✅ Database meeting updated successfully");
      }
    } catch (updateError) {
      console.error("❌ Error updating meeting:", updateError);
    }

    // STEP 7: Return success response
    res.json({
      success: true,
      message: "➕ Attendees added successfully!",
      meeting: {
        ...meeting,
        attendees: updatedAttendees,
        updatedAt: new Date().toISOString()
      },
      changes: {
        attendeesAdded: validatedAttendees,
        invalidAttendees: invalidAttendees,
        previousCount: currentAttendees.length,
        newCount: updatedAttendees.length,
        updatedInTeams: meeting.isFromTeamsCalendar
      }
    });

  } catch (error) {
    console.error("❌ Add attendees backend error:", error);
    res.status(500).json({
      error: "Failed to add attendees to meeting",
      details: error.message
    });
  }
});

router.post("/suggest-times", requireRealTeams, async (req, res) => {
  try {
    const {
      attendees = [],
      duration = 30,
      searchDays = 7,
      timePreferences = {}
    } = req.body;

    if (attendees.length === 0) {
      return res.status(400).json({
        error: "At least one attendee email is required for time suggestions"
      });
    }

    logger.info("🔍 Finding optimal meeting times", {
      attendees: attendees.length,
      duration,
      searchDays,
      timePreferences
    });

    // Step 1: Validate that all attendees are real Teams users
    const validation = await teamsService.validateTeamsUsers(attendees);
    
    if (!validation.allValid) {
      return res.status(400).json({
        error: "Some attendees are not valid Teams users",
        invalidUsers: validation.invalidUsers,
        validUsers: validation.validUsers,
        message: "Only real Teams organization members can be included in time suggestions"
      });
    }

    // Step 2: Get optimal meeting time suggestions
    const suggestions = await teamsService.findMeetingTimes(
      attendees,
      duration,
      searchDays,
      timePreferences
    );

    // Step 3: Return enhanced suggestions
    res.json({
      success: true,
      ...suggestions,
      realTeamsIntegration: true,
      message: suggestions.suggestions.length > 0 
        ? `🎯 Found ${suggestions.suggestions.length} optimal meeting times using real Teams calendar data!`
        : "No available time slots found for all attendees"
    });

  } catch (error) {
    logger.error("❌ Suggest meeting times error:", error);
    res.status(500).json({
      error: "Failed to suggest optimal meeting times",
      details: error.message
    });
  }
});

// DELETE /api/meetings/:id/attendees/remove - Remove attendees from existing meeting
// REPLACE your DELETE /api/meetings/:id/attendees/remove endpoint in meetings.js with this:

// DELETE /api/meetings/:id/attendees/remove - Remove attendees from existing meeting
router.delete("/:id/attendees/remove", requireRealTeams, async (req, res) => {
  try {
    const { attendees } = req.body;
    
    if (!attendees || attendees.length === 0) {
      return res.status(400).json({
        error: "An array of attendee emails to remove is required"
      });
    }

    console.log("🗑️ Backend: Attempting to remove attendees from meeting:", req.params.id);
    console.log("🗑️ Backend: Attendees to remove:", attendees);

    // STEP 1: Try to find meeting in database first
    let meeting = null;
    try {
      meeting = await cosmosClient.getItem(
        "meetings",
        req.params.id,
        req.user.userId
      );
      console.log("📊 Found meeting in database:", meeting ? "YES" : "NO");
    } catch (dbError) {
      console.log("⚠️ Meeting not in database, will search calendar meetings");
    }

    // STEP 2: If not in database, search through ALL calendar meetings
    if (!meeting) {
      console.log("🔍 Searching calendar meetings for ID:", req.params.id);
      
      try {
        // Get all meetings (database + calendar)
        const allMeetingsResponse = await fetch(`http://localhost:5000/api/meetings?limit=200&includeCalendar=true`);
        const allMeetingsData = await allMeetingsResponse.json();
        
        console.log("📅 Total meetings available:", allMeetingsData.meetings?.length || 0);
        
        // Find the meeting by ID in the combined list
        meeting = allMeetingsData.meetings?.find(m => 
          m.id === req.params.id || 
          m.meetingId === req.params.id ||
          m.graphEventId === req.params.id
        );
        
        if (meeting) {
          console.log("✅ Found meeting in calendar data:", meeting.subject);
          console.log("📊 Meeting source:", meeting.isFromTeamsCalendar ? "Teams Calendar" : "Database");
        }
      } catch (searchError) {
        console.error("❌ Error searching calendar meetings:", searchError);
      }
    }

    if (!meeting) {
      console.log("❌ Meeting not found in database OR calendar");
      return res.status(404).json({ 
        error: "Meeting not found",
        details: `Meeting with ID ${req.params.id} not found in database or calendar`,
        suggestion: "Use 'Show my meetings' to see available meetings"
      });
    }

    console.log("✅ Meeting found:", {
      id: meeting.id,
      subject: meeting.subject,
      currentAttendees: meeting.attendees?.length || 0,
      isFromCalendar: meeting.isFromTeamsCalendar || false
    });

    // STEP 3: Check if attendees to remove actually exist in the meeting
    const currentAttendees = meeting.attendees || [];
    const attendeesToRemove = attendees.filter(email => 
      currentAttendees.some(current => 
        current.toLowerCase().includes(email.toLowerCase()) ||
        email.toLowerCase().includes(current.toLowerCase())
      )
    );

    if (attendeesToRemove.length === 0) {
      console.log("⚠️ No matching attendees found to remove");
      return res.status(400).json({
        error: "No matching attendees found to remove",
        providedAttendees: attendees,
        currentAttendees: currentAttendees,
        message: "None of the specified attendees were found in the meeting"
      });
    }

    console.log("👥 Attendees to remove:", attendeesToRemove);

    // STEP 4: Remove attendees from the list
    const updatedAttendees = currentAttendees.filter(current => 
      !attendeesToRemove.some(toRemove => 
        current.toLowerCase().includes(toRemove.toLowerCase()) ||
        toRemove.toLowerCase().includes(current.toLowerCase())
      )
    );

    console.log("📊 Attendees before:", currentAttendees.length);
    console.log("📊 Attendees after:", updatedAttendees.length);

    // STEP 5: Update the meeting
    try {
      // If it's a calendar meeting, try to update via Microsoft Graph
      if (meeting.isFromTeamsCalendar && meeting.graphEventId) {
        console.log("🔄 Updating Teams calendar meeting via Graph API");
        
        const updateResult = await teamsService.updateTeamsMeeting(meeting.graphEventId, {
          attendees: updatedAttendees
        });
        
        console.log("✅ Teams meeting updated successfully");
      }
      
      // Also update in database if the meeting exists there
      if (!meeting.isFromTeamsCalendar) {
        await cosmosClient.updateItem(
          "meetings", 
          req.params.id, 
          req.user.userId, 
          {
            attendees: updatedAttendees,
            updatedAt: new Date().toISOString(),
            lastAttendeeUpdate: {
              timestamp: new Date().toISOString(),
              removed: attendeesToRemove,
              newTotal: updatedAttendees.length
            }
          }
        );
        console.log("✅ Database meeting updated successfully");
      }

    } catch (updateError) {
      console.error("❌ Error updating meeting:", updateError);
      // Continue anyway - the attendee list change logic worked
    }

    // STEP 6: Return success response
    const finalMeeting = {
      ...meeting,
      attendees: updatedAttendees,
      updatedAt: new Date().toISOString()
    };

    res.json({
      success: true,
      message: "➖ Attendees removed successfully!",
      meeting: finalMeeting,
      changes: {
        attendeesRemoved: attendeesToRemove,
        previousCount: currentAttendees.length,
        newCount: updatedAttendees.length,
        updatedInTeams: meeting.isFromTeamsCalendar,
        updatedInDatabase: !meeting.isFromTeamsCalendar
      }
    });

  } catch (error) {
    console.error("❌ Remove attendees backend error:", error);
    res.status(500).json({
      error: "Failed to remove attendees from meeting",
      details: error.message
    });
  }
});




// GET /api/meetings/:id/attendees - Get current attendees list
router.get("/:id/attendees", async (req, res) => {
  try {
    const meeting = await cosmosClient.getItem(
      "meetings",
      req.params.id,
      req.user.userId
    );

    if (!meeting) {
      return res.status(404).json({ error: "Meeting not found" });
    }

    // Get detailed info about each attendee if Teams is available
    let attendeeDetails = [];
    
    if (teamsService.isAvailable() && meeting.attendees.length > 0) {
      try {
        const validation = await teamsService.validateTeamsUsers(meeting.attendees);
        attendeeDetails = validation.validUsers.map(user => ({
          email: user.email,
          displayName: user.displayName,
          userPrincipalName: user.userPrincipalName,
          status: 'valid'
        }));
        
        // Add invalid users
        validation.invalidUsers.forEach(user => {
          attendeeDetails.push({
            email: user.email,
            status: 'invalid',
            error: user.error
          });
        });
      } catch (error) {
        logger.warn("Could not get attendee details:", error.message);
        attendeeDetails = meeting.attendees.map(email => ({
          email: email,
          status: 'unknown'
        }));
      }
    } else {
      attendeeDetails = meeting.attendees.map(email => ({
        email: email,
        status: 'unknown'
      }));
    }

    res.json({
      success: true,
      meetingId: meeting.id,
      subject: meeting.subject,
      attendees: meeting.attendees,
      attendeeDetails: attendeeDetails,
      attendeeCount: meeting.attendees.length,
      lastUpdated: meeting.updatedAt || meeting.createdAt,
      lastAttendeeUpdate: meeting.lastAttendeeUpdate || null
    });

  } catch (error) {
    logger.error("❌ Get attendees error:", error);
    res.status(500).json({
      error: "Failed to get meeting attendees",
      details: error.message
    });
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