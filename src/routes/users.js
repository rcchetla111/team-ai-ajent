const express = require("express");
const { v4: uuidv4 } = require("uuid");
const teamsService = require("../services/teamsService");
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

// ============================================================================
// POC FEATURE 1.1: USER RESOLUTION FOR MEETING SCHEDULING
// ============================================================================

// POST /api/users/resolve - Find user emails from display names
router.post("/resolve", async (req, res) => {
  try {
    const { names } = req.body;

    if (!names || !Array.isArray(names) || names.length === 0) {
      return res.status(400).json({ 
        error: "An array of 'names' is required in the request body." 
      });
    }

    logger.info("🔍 Resolving user names to emails", { names });

    if (!teamsService.isAvailable()) {
      return res.json({
        success: true,
        mode: "simulated",
        resolvedUsers: names.map(name => ({
          name: name,
          email: `${name.toLowerCase().replace(/\s+/g, '.')}@company.com`
        })),
        message: "Teams integration not available - using simulated user resolution",
        realUserLookup: false,
        suggestion: "Configure Azure AD to enable real user lookup"
      });
    }

    const resolvedUsers = await teamsService.findUsersByDisplayName(names);

    res.json({
      success: true,
      mode: "real_teams",
      resolvedUsers: resolvedUsers,
      summary: {
        requested: names.length,
        found: resolvedUsers.length,
        successRate: `${Math.round((resolvedUsers.length / names.length) * 100)}%`
      },
      realUserLookup: true,
      message: resolvedUsers.length > 0 
        ? `✅ Successfully resolved ${resolvedUsers.length}/${names.length} users from Teams directory!` 
        : "❌ No users found in Teams directory"
    });

  } catch (error) {
    logger.error("❌ Resolve users error:", error);
    res.status(500).json({ 
      error: "Failed to resolve user names", 
      details: error.message 
    });
  }
});

// GET /api/users/search - Search for team members
router.get("/search", async (req, res) => {
  try {
    const { q, limit = 10 } = req.query;
    
    if (!q || q.trim() === '') {
      return res.status(400).json({ 
        error: "Search query 'q' is required" 
      });
    }
    
    logger.info(`🔍 Searching for team members: "${q}"`);

    if (!teamsService.isAvailable()) {
      return res.json({
        success: true,
        mode: "simulated",
        query: q,
        users: [
          {
            id: uuidv4(),
            name: `${q} Smith`,
            email: `${q.toLowerCase()}@company.com`,
            jobTitle: "Software Engineer",
            department: "Engineering"
          },
          {
            id: uuidv4(),
            name: `${q} Johnson`,
            email: `${q.toLowerCase()}.johnson@company.com`,
            jobTitle: "Product Manager",
            department: "Product"
          }
        ],
        found: 2,
        realUserLookup: false,
        message: "Teams integration not available - showing simulated results"
      });
    }

    const users = await teamsService.findTeamMembers(q);
    
    res.json({
      success: true,
      mode: "real_teams",
      query: q,
      users: users.slice(0, parseInt(limit)),
      found: users.length,
      realUserLookup: true,
      message: users.length > 0 
        ? `Found ${users.length} team members matching "${q}"` 
        : `No team members found matching "${q}"`
    });
    
  } catch (error) {
    logger.error("❌ Team member search failed:", error);
    res.status(500).json({
      error: "Failed to search team members",
      details: error.message
    });
  }
});

// GET /api/users - Get all team members
router.get("/", async (req, res) => {
  try {
    const { limit = 50 } = req.query;
    
    logger.info(`📋 Getting team members (limit: ${limit})`);

    if (!teamsService.isAvailable()) {
      return res.json({
        success: true,
        mode: "simulated",
        users: [
          {
            id: uuidv4(),
            name: "John Smith",
            email: "john.smith@company.com",
            jobTitle: "Software Engineer",
            department: "Engineering"
          },
          {
            id: uuidv4(),
            name: "Sarah Johnson",
            email: "sarah.johnson@company.com",
            jobTitle: "Product Manager",
            department: "Product"
          },
          {
            id: uuidv4(),
            name: "Mike Wilson",
            email: "mike.wilson@company.com",
            jobTitle: "Design Lead",
            department: "Design"
          },
          {
            id: uuidv4(),
            name: "Lisa Chen",
            email: "lisa.chen@company.com",
            jobTitle: "Data Scientist",
            department: "Analytics"
          },
          {
            id: uuidv4(),
            name: "David Brown",
            email: "david.brown@company.com",
            jobTitle: "DevOps Engineer",
            department: "Engineering"
          }
        ],
        total: 5,
        realUserLookup: false,
        message: "Teams integration not available - showing simulated team directory"
      });
    }

    const users = await teamsService.getAllTeamMembers(parseInt(limit));
    
    res.json({
      success: true,
      mode: "real_teams",
      users: users,
      total: users.length,
      realUserLookup: true,
      message: `Retrieved ${users.length} team members from Teams directory`
    });
    
  } catch (error) {
    logger.error("❌ Get team members failed:", error);
    res.status(500).json({
      error: "Failed to get team members",
      details: error.message
    });
  }
});

// POST /api/users/test-resolution - Test user resolution functionality
router.post("/test-resolution", async (req, res) => {
  try {
    const { names } = req.body;
    
    if (!names || !Array.isArray(names)) {
      return res.status(400).json({ 
        error: "Array of 'names' is required" 
      });
    }
    
    logger.info(`🧪 Testing user resolution for: ${names.join(', ')}`);

    if (!teamsService.isAvailable()) {
      return res.json({
        success: true,
        mode: "simulated",
        message: "Teams integration not available - any names will work in simulated mode",
        realUserLookup: false,
        input: names,
        testResults: {
          simulatedUsers: names.map(name => ({
            inputName: name,
            resolvedEmail: `${name.toLowerCase().replace(/\s+/g, '.')}@company.com`,
            status: "simulated"
          })),
          allNamesWork: true
        },
        suggestion: "Configure Azure AD to test real user resolution"
      });
    }

    const resolvedUsers = await teamsService.findUsersByDisplayName(names);
    
    res.json({
      success: true,
      mode: "real_teams",
      realUserLookup: true,
      input: names,
      foundUsers: resolvedUsers,
      testResults: {
        requested: names.length,
        found: resolvedUsers.length,
        successRate: `${Math.round((resolvedUsers.length / names.length) * 100)}%`,
        details: names.map(name => {
          const found = resolvedUsers.find(user => 
            user.name.toLowerCase().includes(name.toLowerCase())
          );
          return {
            inputName: name,
            found: !!found,
            resolvedUser: found || null
          };
        })
      },
      message: resolvedUsers.length > 0 
        ? `✅ Successfully found ${resolvedUsers.length}/${names.length} real Teams users!` 
        : "❌ No users found in Teams directory - try different names"
    });
    
  } catch (error) {
    logger.error("❌ User resolution test failed:", error);
    res.status(500).json({
      error: "Failed to test user resolution",
      details: error.message
    });
  }
});

// GET /api/users/status - Get user resolution service status
router.get("/status", (req, res) => {
  const teamsStatus = teamsService.isAvailable();
  
  res.json({
    success: true,
    userResolutionAvailable: teamsStatus,
    teamsIntegration: teamsStatus,
    features: {
      nameToEmailResolution: teamsStatus,
      teamMemberSearch: teamsStatus,
      realTimeDirectoryLookup: teamsStatus,
      simulatedMode: !teamsStatus
    },
    message: teamsStatus
      ? "🟢 User resolution fully operational with real Teams directory!"
      : "⚠️ User resolution in simulated mode - configure Azure AD for real functionality",
    configuration: {
      azureAdRequired: !teamsStatus,
      graphApiAccess: teamsStatus,
      userReadPermissions: teamsStatus
    }
  });
});

module.exports = router;