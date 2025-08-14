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

// FIX: Add missing router declaration
const router = express.Router();

// FIX: Add missing authentication middleware
const simulateAuth = (req, res, next) => {
  req.user = {
    userId: "demo-user-123",
    email: "demo@company.com",
    name: "Demo User",
  };
  next();
};

router.use(simulateAuth);

// FIX: Add missing requireRealTeams middleware
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
// CHANNEL MANAGEMENT ENDPOINTS
// ============================================================================

// POST /api/channels/create - Create Teams Channel
router.post("/create", requireRealTeams, async (req, res) => {
  try {
    const {
      teamId,
      displayName,
      description,
      membershipType = 'standard' // 'standard' or 'private'
    } = req.body;

    if (!teamId || !displayName) {
      return res.status(400).json({
        error: "Team ID and channel name are required"
      });
    }

    logger.info("🆕 Creating Teams channel", {
      teamId,
      displayName,
      membershipType
    });

    const channelResult = await teamsService.createTeamsChannel({
      teamId,
      displayName,
      description,
      membershipType
    });

    res.status(201).json({
      success: true,
      channel: channelResult,
      message: `✅ Channel "${displayName}" created successfully!`,
      teamId: teamId
    });

  } catch (error) {
    logger.error("❌ Create Teams channel error:", error);
    res.status(500).json({
      error: "Failed to create Teams channel",
      details: error.message
    });
  }
});

// GET /api/channels/teams - Get available teams
router.get("/teams", requireRealTeams, async (req, res) => {
  try {
    const teams = await teamsService.getAvailableTeams();
    
    res.json({
      success: true,
      teams: teams,
      message: `Found ${teams.length} teams where you can create channels`
    });

  } catch (error) {
    logger.error("❌ Get teams error:", error);
    res.status(500).json({
      error: "Failed to get available teams",
      details: error.message
    });
  }
});

// GET /api/channels/:teamId - Get channels in a team
router.get("/:teamId", requireRealTeams, async (req, res) => {
  try {
    const { teamId } = req.params;
    
    // DEBUG: Add logging to see what teamId you're getting
    console.log("🔍 Getting channels for teamId:", teamId);
    console.log("🔍 Full request params:", req.params);
    
    // ISSUE: Make sure you're passing the teamId correctly
    const channels = await teamsService.getTeamChannels(teamId);
    
    res.json({
      success: true,
      teamId: teamId,
      channels: channels,
      message: `Found ${channels.length} channels in this team`
    });

  } catch (error) {
    console.error("❌ Channel route error:", error);
    logger.error("❌ Get channels error:", error);
    res.status(500).json({
      error: "Failed to get team channels",
      details: error.message
    });
  }
});


router.get("/channels/debug", async (req, res) => {
  try {
    // First get available teams
    const teams = await teamsService.getAvailableTeams();
    console.log("🔍 Available teams:", teams);
    
    if (teams.length > 0) {
      // Try to get channels for the first team
      const firstTeam = teams[0];
      console.log("🔍 Trying to get channels for first team:", firstTeam);
      
      const channels = await teamsService.getTeamChannels(firstTeam.id);
      
      res.json({
        success: true,
        debugInfo: {
          teamsFound: teams.length,
          firstTeamId: firstTeam.id,
          firstTeamName: firstTeam.displayName,
          channelsInFirstTeam: channels.length
        },
        teams: teams,
        sampleChannels: channels
      });
    } else {
      res.json({
        success: true,
        message: "No teams found",
        teams: []
      });
    }

  } catch (error) {
    console.error("❌ Debug route error:", error);
    res.status(500).json({
      error: "Debug failed",
      details: error.message
    });
  }
});


router.get("/debug/teams-only", async (req, res) => {
  try {
    console.log("🔍 Testing teams-only endpoint");
    
    if (!teamsService.isAvailable()) {
      return res.json({
        success: false,
        message: "Teams service not available",
        isAvailable: false
      });
    }

    // Just test getting teams, not channels
    const teams = await teamsService.getAvailableTeams();
    console.log("✅ Teams retrieved:", teams.length);
    
    res.json({
      success: true,
      message: `Found ${teams.length} teams`,
      teams: teams,
      teamsServiceAvailable: true
    });

  } catch (error) {
    console.error("❌ Teams-only test failed:", error);
    res.status(500).json({
      error: "Failed to get teams",
      details: error.message,
      stack: error.stack
    });
  }
});

// ============================================================================
// ALTERNATIVE: Routes without requireRealTeams (for testing/simulation)
// ============================================================================

// If you want to test without Teams integration, use these routes instead:

// POST /api/channels/create-simulation - Create simulated channel
router.post("/create-simulation", async (req, res) => {
  try {
    const {
      teamId,
      displayName,
      description,
      membershipType = 'standard'
    } = req.body;

    if (!displayName) {
      return res.status(400).json({
        error: "Channel name is required"
      });
    }

    logger.info("🆕 Creating simulated Teams channel", {
      teamId: teamId || "simulated-team",
      displayName,
      membershipType
    });

    // Simulated response
    const simulatedChannel = {
      success: true,
      channelId: uuidv4(),
      displayName: displayName,
      description: description || `${displayName} - Created via AI Agent`,
      membershipType: membershipType,
      teamId: teamId || "simulated-team",
      webUrl: `https://teams.microsoft.com/l/channel/simulated-${displayName.replace(/\s+/g, '-').toLowerCase()}`,
      createdDateTime: new Date().toISOString()
    };

    res.status(201).json({
      success: true,
      channel: simulatedChannel,
      message: `✅ Simulated channel "${displayName}" created successfully!`,
      teamId: teamId || "simulated-team",
      mode: "simulation"
    });

  } catch (error) {
    logger.error("❌ Create simulated channel error:", error);
    res.status(500).json({
      error: "Failed to create simulated channel",
      details: error.message
    });
  }
});

// GET /api/channels/teams-simulation - Get simulated teams
router.get("/teams-simulation", async (req, res) => {
  try {
    const simulatedTeams = [
      {
        id: "team-engineering",
        displayName: "Engineering Team",
        description: "Software development and technical discussions"
      },
      {
        id: "team-product", 
        displayName: "Product Team",
        description: "Product management and strategy"
      },
      {
        id: "team-marketing",
        displayName: "Marketing Team", 
        description: "Marketing campaigns and communications"
      },
      {
        id: "team-general",
        displayName: "General",
        description: "Company-wide discussions and announcements"
      }
    ];
    
    res.json({
      success: true,
      teams: simulatedTeams,
      message: `Found ${simulatedTeams.length} simulated teams where you can create channels`,
      mode: "simulation"
    });

  } catch (error) {
    logger.error("❌ Get simulated teams error:", error);
    res.status(500).json({
      error: "Failed to get simulated teams",
      details: error.message
    });
  }
});

// FIX: Add missing module.exports
module.exports = router;