const express = require('express');
const router = express.Router();

// Teams bot processing endpoint
router.post('/agent/process', async (req, res) => {
    try {
        const { message, user, context } = req.body;
        
        console.log('📨 Teams message received:', {
            message,
            user: user.name,
            conversationId: context.conversationId
        });

        // Use your existing Agent 365 logic
        const response = await processAgentMessage(message, user, context);
        
        res.json({
            success: true,
            message: response,
            timestamp: new Date().toISOString()
        });
    } catch (error) {
        console.error('Teams agent processing error:', error);
        res.status(500).json({
            success: false,
            error: 'Failed to process message',
            message: 'I encountered an error processing your request. Please try again.'
        });
    }
});

async function processAgentMessage(message, user, context) {
    // Import your existing Agent 365 logic here
    const { getAgentAction, handleUserMessage } = require('../services/agent365Service');
    
    try {
        // Process using your existing Agent 365 system
        const action = await getAgentAction(message);
        const response = await executeAgentAction(action, user, context);
        
        return response;
    } catch (error) {
        console.error('Agent processing error:', error);
        return 'I apologize, but I encountered an error processing your request. Please try again.';
    }
}

async function executeAgentAction(action, user, context) {
    // Route to your existing tool implementations
    switch (action.tool_name) {
        case 'create_meeting':
            return await handleCreateMeetingFromTeams(action.parameters, user);
        case 'get_meetings':
            return await handleGetMeetingsFromTeams(action.parameters, user);
        case 'find_people':
            return await handleFindPeopleFromTeams(action.parameters, user);
        default:
            return action.parameters?.responseText || 'Command processed successfully!';
    }
}

module.exports = router;