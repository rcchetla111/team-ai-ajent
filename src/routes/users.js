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

// For now, we'll simulate authentication (replace with real auth later)
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
// 1.1 MEETING SCHEDULING CAPABILITIES
// ============================================================================

// POST /api/users/resolve - Find user emails from display names (for scheduling)
router.post("/resolve", async (req, res) => {
    try {
        const { names } = req.body;

        if (!names || !Array.isArray(names) || names.length === 0) {
            return res.status(400).json({ error: "An array of 'names' is required in the request body." });
        }

        const resolvedUsers = await teamsService.findUsersByDisplayName(names);

        res.json({
            success: true,
            resolvedUsers: resolvedUsers,
            message: `Resolved ${resolvedUsers.length}/${names.length} users for meeting scheduling`
        });

    } catch (error) {
        logger.error("❌ Resolve users error:", error);
        res.status(500).json({ error: "Failed to resolve user names", details: error.message });
    }
});

module.exports = router;