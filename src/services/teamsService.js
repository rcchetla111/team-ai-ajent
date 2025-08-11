const axios = require('axios');
const moment = require('moment');
const authService = require('./authService');
const logger = require('../utils/logger');

class TeamsService {
  constructor() {
    this.graphEndpoint = process.env.GRAPH_API_ENDPOINT || 'https://graph.microsoft.com/v1.0';
  }

  // Check if service is available
  isAvailable() {
    return authService.isAvailable();
  }

    // This function builds the recurrence object for the Graph API
  buildRecurrencePattern(recurrence) {
    if (!recurrence || !recurrence.frequency) {
        return null;
    }

    const pattern = {
        type: recurrence.frequency, // "daily", "weekly", "monthly"
        interval: recurrence.interval || 1, // e.g., every 1 day, every 2 weeks
    };

    if (recurrence.frequency === 'weekly' && recurrence.daysOfWeek) {
        pattern.daysOfWeek = recurrence.daysOfWeek; // e.g., ["monday", "wednesday"]
    }

    const range = {
        type: recurrence.rangeType || 'endDate', // "endDate", "noEnd", "numbered"
        startDate: moment(recurrence.startDate).format('YYYY-MM-DD'),
    };

    if (range.type === 'endDate' && recurrence.endDate) {
        range.endDate = moment(recurrence.endDate).format('YYYY-MM-DD');
    } else if (range.type === 'numbered' && recurrence.occurrences) {
        range.numberOfOccurrences = recurrence.occurrences;
    } else {
        // Default to ending after 1 year if no end date is specified for safety
        range.type = 'endDate';
        range.endDate = moment(recurrence.startDate).add(1, 'year').format('YYYY-MM-DD');
    }

    return { pattern, range };
  }

  // Create real Teams meeting using the EXACT working approach
 async createTeamsMeeting(meetingData) {
        if (!this.isAvailable()) {
            throw new Error('Teams service not available - check Azure AD configuration');
        }

        try {
            const accessToken = await authService.getAppOnlyToken();

            const {
                subject,
                startTime,
                endTime,
                attendees = [],
                recurrence // NEW: Accept a recurrence object
            } = meetingData;

            const eventDetails = {
                subject: subject,
                start: {
                    dateTime: startTime,
                    timeZone: "UTC"
                },
                end: {
                    dateTime: endTime,
                    timeZone: "UTC"
                },
                isOnlineMeeting: true,
                onlineMeetingProvider: "teamsForBusiness"
            };

            // NEW: Add recurrence pattern if it exists
            const recurrencePattern = this.buildRecurrencePattern(recurrence);
            if (recurrencePattern) {
                eventDetails.recurrence = recurrencePattern;
                logger.info('🔄 Creating a recurring meeting with pattern:', { recurrencePattern });
            }

            if (attendees.length > 0) {
                eventDetails.attendees = attendees.map(email => ({
                    emailAddress: { address: email, name: email.split('@')[0] },
                    type: 'required'
                }));
            }

            logger.info('🔄 Creating calendar event with Teams meeting');

            // Using a specific user ID is the correct approach for app-only permissions
            const usersResponse = await axios.get(
                `${this.graphEndpoint}/users?$top=1&$select=id`, // More efficient query
                { headers: { 'Authorization': `Bearer ${accessToken}` } }
            );

            if (!usersResponse.data.value || usersResponse.data.value.length === 0) {
                throw new Error('No users found in tenant to create a meeting on behalf of.');
            }
            const userId = usersResponse.data.value[0].id;
            logger.info(`🔄 Using user ${userId} to create calendar event`);

            const response = await axios.post(
                `${this.graphEndpoint}/users/${userId}/events`,
                eventDetails,
                { headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' } }
            );

            const eventData = response.data;

            logger.info('✅ Teams meeting created successfully!', {
                eventId: eventData.id,
                subject: eventData.subject,
                isRecurring: !!eventData.recurrence
            });

            return {
                success: true,
                meetingId: eventData.id,
                subject: eventData.subject,
                startTime: eventData.start.dateTime,
                endTime: eventData.end.dateTime,
                joinUrl: eventData.onlineMeeting?.joinUrl,
                webUrl: eventData.webLink,
                graphEventId: eventData.id,
                isReal: true,
                isRecurring: !!eventData.recurrence
            };

        } catch (error) {
            logger.error('❌ Failed to create Teams meeting:', {
                error: error.message,
                status: error.response?.status,
                data: error.response?.data
            });
            throw new Error(`Teams meeting creation failed: ${error.message}`);
        }
    }

  // Update existing Teams meeting
  async updateTeamsMeeting(graphEventId, updateData) {
    if (!this.isAvailable()) {
      throw new Error('Teams service not available');
    }

    try {
      const accessToken = await authService.getAppOnlyToken();
      
      const response = await axios.patch(
        `${this.graphEndpoint}/beta/communications/onlineMeetings/${graphEventId}`,
        updateData,
        {
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          }
        }
      );

      logger.info('✅ Teams meeting updated successfully');
      return { success: true, data: response.data };

    } catch (error) {
      logger.error('❌ Failed to update Teams meeting:', error);
      throw new Error(`Teams meeting update failed: ${error.message}`);
    }
  }

   async updateMeeting(graphEventId, updatePayload) {
    if (!this.isAvailable()) {
      throw new Error('Teams service not available');
    }
    if (!graphEventId || !updatePayload) {
        throw new Error('graphEventId and updatePayload are required for updating a meeting.');
    }

    try {
        const accessToken = await authService.getAppOnlyToken();

        // We need a user's context to update their calendar event
        const usersResponse = await axios.get(
            `${this.graphEndpoint}/users?$top=1&$select=id`,
            { headers: { 'Authorization': `Bearer ${accessToken}` } }
        );
        if (!usersResponse.data.value || usersResponse.data.value.length === 0) {
            throw new Error('No users found in tenant to update a meeting on behalf of.');
        }
        const userId = usersResponse.data.value[0].id;

        logger.info(`🔄 Updating meeting event ${graphEventId} for user ${userId}`, { payload: updatePayload });

        // Use a PATCH request to update only the specified fields (e.g., attendees)
        const response = await axios.patch(
            `${this.graphEndpoint}/users/${userId}/events/${graphEventId}`,
            updatePayload,
            {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            }
        );

        logger.info('✅ Teams meeting updated successfully!', { eventId: response.data.id });
        return { success: true, data: response.data };

    } catch (error) {
        logger.error('❌ Failed to update Teams meeting:', {
            error: error.message,
            status: error.response?.status,
            data: error.response?.data
        });
        throw new Error(`Teams meeting update failed: ${error.message}`);
    }
  }

  // Cancel Teams meeting
  async cancelTeamsMeeting(graphEventId) {
    if (!this.isAvailable()) {
      throw new Error('Teams service not available');
    }

    try {
      const accessToken = await authService.getAppOnlyToken();
      
      await axios.delete(
        `${this.graphEndpoint}/beta/communications/onlineMeetings/${graphEventId}`,
        {
          headers: {
            'Authorization': `Bearer ${accessToken}`
          }
        }
      );

      logger.info('✅ Teams meeting cancelled successfully');
      return { success: true };

    } catch (error) {
      logger.error('❌ Failed to cancel Teams meeting:', error);
      throw new Error(`Teams meeting cancellation failed: ${error.message}`);
    }
  }

  // Create simulated meeting (fallback)
  createSimulatedMeeting(meetingData) {
    const timestamp = Date.now();
    const meetingId = `meeting-${timestamp}-${Math.random().toString(36).substr(2, 9)}`;
    
    return {
      success: true,
      meetingId: meetingId,
      subject: meetingData.subject,
      startTime: meetingData.startTime,
      endTime: meetingData.endTime,
      joinUrl: `https://teams.microsoft.com/join/meeting-${timestamp}`,
      webUrl: `https://teams.microsoft.com/meeting/${timestamp}`,
      graphEventId: null,
      isReal: false,
      note: 'This is a simulated meeting. Configure Azure AD for real Teams integration.'
    };
  }

  // Get available meeting times (uses Graph findMeetingTimes API)
  async findMeetingTimes(attendees, duration = 30) {
    if (!this.isAvailable()) {
      throw new Error('Teams service not available for finding meeting times');
    }

    try {
      const accessToken = await authService.getAppOnlyToken();
      
      const findMeetingTimesRequest = {
        attendees: attendees.map(email => ({
          emailAddress: {
            address: email,
            name: email.split('@')[0]
          }
        })),
        meetingDuration: `PT${duration}M`,
        maxCandidates: 10,
        timeConstraint: {
          timeslots: [{
            start: {
              dateTime: moment().add(1, 'hour').toISOString(),
              timeZone: 'UTC'
            },
            end: {
              dateTime: moment().add(7, 'days').toISOString(),
              timeZone: 'UTC'
            }
          }]
        }
      };

      const response = await axios.post(
        `${this.graphEndpoint}/me/calendar/getSchedule`,
        findMeetingTimesRequest,
        {
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          }
        }
      );

      logger.info('✅ Retrieved meeting time suggestions from Graph API');
      return response.data.meetingTimeSuggestions || [];

    } catch (error) {
      logger.error('❌ Failed to get meeting times from Graph API:', error);
      throw new Error(`Finding meeting times failed: ${error.message}`);
    }
  }

  async findUsersByDisplayName(displayNames) {
    if (!this.isAvailable()) {
      throw new Error('Teams service not available for user lookup.');
    }
    if (!Array.isArray(displayNames) || displayNames.length === 0) {
        return [];
    }

    try {
        const accessToken = await authService.getAppOnlyToken();
        const resolvedUsers = [];

        // We will search for each name individually
        for (const name of displayNames) {
            // The Graph API filter to find a user whose display name starts with the given name
            const filter = `$filter=startswith(displayName, '${name}')`;
            // We only need the displayName and userPrincipalName (email)
            const select = `$select=displayName,userPrincipalName`;
            
            const url = `${this.graphEndpoint}/users?${filter}&${select}`;

            logger.info(`🔍 Searching for user: ${name}`);
            
            const response = await axios.get(url, {
                headers: { 'Authorization': `Bearer ${accessToken}` }
            });

            if (response.data.value && response.data.value.length > 0) {
                // For this POC, we'll take the first match.
                const foundUser = response.data.value[0];
                resolvedUsers.push({
                    name: foundUser.displayName,
                    email: foundUser.userPrincipalName
                });
                logger.info(`✅ Found user: ${foundUser.displayName} (${foundUser.userPrincipalName})`);
            } else {
                logger.warn(`⚠️ Could not find a user matching: ${name}`);
            }
        }

        return resolvedUsers;

    } catch (error) {
        logger.error('❌ Failed to find users by display name:', {
            error: error.message,
            status: error.response?.status,
            data: error.response?.data
        });
        throw new Error(`User lookup failed: ${error.message}`);
    }
  }

  
async findTeamMembers(searchTerm) {
    if (!this.isAvailable()) {
      throw new Error('Teams service not available');
    }

    try {
      const accessToken = await authService.getAppOnlyToken();
      
      // Search for users in the organization
      const searchUrl = `${this.graphEndpoint}/users?$filter=startswith(displayName,'${searchTerm}') or startswith(givenName,'${searchTerm}') or startswith(surname,'${searchTerm}')&$select=id,displayName,mail,userPrincipalName,jobTitle,department&$top=10`;
      
      const response = await axios.get(searchUrl, {
        headers: { 'Authorization': `Bearer ${accessToken}` }
      });

      const users = response.data.value.map(user => ({
        id: user.id,
        name: user.displayName,
        email: user.mail || user.userPrincipalName,
        jobTitle: user.jobTitle,
        department: user.department
      }));

      logger.info(`✅ Found ${users.length} team members matching "${searchTerm}"`);
      return users;

    } catch (error) {
      logger.error('❌ Failed to find team members:', error);
      throw error;
    }
  }


  async getAllTeamMembers(limit = 50) {
    if (!this.isAvailable()) {
      throw new Error('Teams service not available');
    }

    try {
      const accessToken = await authService.getAppOnlyToken();
      
      const usersUrl = `${this.graphEndpoint}/users?$select=id,displayName,mail,userPrincipalName,jobTitle,department&$top=${limit}&$filter=accountEnabled eq true`;
      
      const response = await axios.get(usersUrl, {
        headers: { 'Authorization': `Bearer ${accessToken}` }
      });

      const users = response.data.value.map(user => ({
        id: user.id,
        name: user.displayName,
        email: user.mail || user.userPrincipalName,
        jobTitle: user.jobTitle,
        department: user.department
      }));

      logger.info(`✅ Retrieved ${users.length} team members`);
      return users;

    } catch (error) {
      logger.error('❌ Failed to get team members:', error);
      throw error;
    }
  }

  // 🚀 NEW: Send message to team member
 // In src/services/teamsService.js
// Replace the sendMessageToUser function with this fixed version:

// Replace your sendMessageToUser function in teamsService.js with this:

// Replace sendMessageToUser with this email-based function in teamsService.js

async sendMessageToUser(userIdentifier, message) {
  if (!this.isAvailable()) {
    throw new Error('Teams service not available');
  }

  try {
    const accessToken = await authService.getAppOnlyToken();
    
    let targetUser;
    
    // Resolve user (email or name)
    if (userIdentifier.includes('@')) {
      const userResponse = await axios.get(
        `${this.graphEndpoint}/users/${userIdentifier}`,
        { headers: { 'Authorization': `Bearer ${accessToken}` } }
      );
      targetUser = userResponse.data;
    } else {
      const users = await this.findTeamMembers(userIdentifier);
      if (users.length === 0) {
        throw new Error(`No user found with name: ${userIdentifier}`);
      }
      const userResponse = await axios.get(
        `${this.graphEndpoint}/users/${users[0].email}`,
        { headers: { 'Authorization': `Bearer ${accessToken}` } }
      );
      targetUser = userResponse.data;
    }

    // Get sender user (first user in tenant)
    const sendersResponse = await axios.get(
      `${this.graphEndpoint}/users?$top=1&$select=id,displayName,userPrincipalName`,
      { headers: { 'Authorization': `Bearer ${accessToken}` } }
    );
    
    if (!sendersResponse.data.value || sendersResponse.data.value.length === 0) {
      throw new Error('No users found in tenant to send message from');
    }
    
    const senderUser = sendersResponse.data.value[0];
    logger.info(`📤 Sending email from: ${senderUser.displayName} to: ${targetUser.displayName}`);

    // Send email message
    const emailPayload = {
      message: {
        subject: `🤖 Message from AI Meeting Agent`,
        body: {
          contentType: 'HTML',
          content: `
            <div style="font-family: Arial, sans-serif; max-width: 600px;">
              <h2 style="color: #0078d4;">🤖 AI Meeting Agent</h2>
              <div style="background: #f5f5f5; padding: 20px; border-radius: 8px; margin: 20px 0;">
                <p style="font-size: 16px; line-height: 1.5; margin: 0;">
                  ${message.replace(/\n/g, '<br>')}
                </p>
              </div>
              <p style="color: #666; font-size: 12px; margin-top: 20px;">
                Sent by AI Meeting Agent on behalf of ${senderUser.displayName}
              </p>
            </div>
          `
        },
        toRecipients: [
          {
            emailAddress: {
              address: targetUser.userPrincipalName,
              name: targetUser.displayName
            }
          }
        ]
      }
    };

    await axios.post(
      `${this.graphEndpoint}/users/${senderUser.id}/sendMail`,
      emailPayload,
      { headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' } }
    );

    logger.info(`✅ Email sent successfully to ${targetUser.displayName}`);
    
    return {
      success: true,
      method: 'email',
      recipient: {
        name: targetUser.displayName,
        email: targetUser.userPrincipalName
      },
      sender: {
        name: senderUser.displayName,
        email: senderUser.userPrincipalName
      },
      message: 'Email sent successfully via AI Agent'
    };

  } catch (error) {
    logger.error(`❌ Failed to send email to ${userIdentifier}:`, {
      message: error.message,
      status: error.response?.status,
      statusText: error.response?.statusText,
      details: error.response?.data
    });
    
    const cleanError = new Error(`Failed to send email: ${error.message}`);
    cleanError.status = error.response?.status;
    throw cleanError;
  }
}

  // 🚀 NEW: Send meeting invite message to attendees
  async sendMeetingInviteMessages(meetingDetails, attendeeEmails) {
    if (!this.isAvailable()) {
      throw new Error('Teams service not available');
    }

    const results = [];
    
    for (const email of attendeeEmails) {
      try {
        const message = `🎯 **Meeting Invitation: ${meetingDetails.subject}**

📅 **When:** ${moment(meetingDetails.startTime).format('MMMM Do YYYY, h:mm A')}
⏱️ **Duration:** ${moment(meetingDetails.endTime).diff(moment(meetingDetails.startTime), 'minutes')} minutes

🔗 **Join Link:** ${meetingDetails.joinUrl}

Hope to see you there! 🤖`;

        const result = await this.sendMessageToUser(email, message);
        results.push({ email, status: 'sent', ...result });
        
        // Add delay to avoid rate limiting
        await new Promise(resolve => setTimeout(resolve, 1000));
        
      } catch (error) {
        logger.warn(`⚠️ Failed to send invite to ${email}:`, error.message);
        results.push({ 
          email, 
          status: 'failed', 
          error: error.message 
        });
      }
    }

    logger.info(`📨 Sent ${results.filter(r => r.status === 'sent').length}/${results.length} meeting invites`);
    return results;
  }

  // 🚀 ENHANCED: Create meeting with real user resolution
  async createTeamsMeetingWithRealUsers(meetingData) {
    if (!this.isAvailable()) {
      throw new Error('Teams service not available - check Azure AD configuration');
    }

    try {
      const {
        subject,
        startTime,
        endTime,
        attendeeNames = [], // Array of names like ["John Smith", "Sarah Johnson"]
        attendeeEmails = [], // Array of emails
        sendInviteMessages = false
      } = meetingData;

      // Step 1: Resolve names to emails
      let resolvedAttendees = [...attendeeEmails];
      
      if (attendeeNames.length > 0) {
        logger.info(`🔍 Resolving ${attendeeNames.length} attendee names to emails`);
        
        for (const name of attendeeNames) {
          try {
            const users = await this.findTeamMembers(name);
            if (users.length > 0) {
              resolvedAttendees.push(users[0].email); // Take first match
              logger.info(`✅ Resolved "${name}" → ${users[0].email}`);
            } else {
              logger.warn(`⚠️ Could not find user: ${name}`);
            }
          } catch (error) {
            logger.warn(`⚠️ Error resolving ${name}:`, error.message);
          }
        }
      }

      // Remove duplicates
      resolvedAttendees = [...new Set(resolvedAttendees)];
      
      // Step 2: Create the Teams meeting
      const meetingResult = await this.createTeamsMeeting({
        subject,
        startTime,
        endTime,
        attendees: resolvedAttendees
      });

      // Step 3: Send invite messages if requested
      let inviteResults = null;
      if (sendInviteMessages && resolvedAttendees.length > 0) {
        logger.info(`📨 Sending invite messages to ${resolvedAttendees.length} attendees`);
        
        inviteResults = await this.sendMeetingInviteMessages({
          subject,
          startTime,
          endTime,
          joinUrl: meetingResult.joinUrl
        }, resolvedAttendees);
      }

      return {
        ...meetingResult,
        resolvedAttendees,
        inviteResults,
        realUsersResolved: attendeeNames.length,
        messagesAttempted: sendInviteMessages ? resolvedAttendees.length : 0,
        messagesSent: inviteResults ? inviteResults.filter(r => r.status === 'sent').length : 0
      };

    } catch (error) {
      logger.error('❌ Failed to create meeting with real users:', error);
      throw error;
    }
  }







  // Get service status
  getStatus() {
    return {
      available: this.isAvailable(),
      authConfigured: authService.isAvailable(),
      features: {
        createMeetings: this.isAvailable(),
        updateMeetings: this.isAvailable(),
        cancelMeetings: this.isAvailable(),
        findTimes: this.isAvailable()
      }
    };
  }
}

// Create singleton instance
const teamsService = new TeamsService();

module.exports = teamsService;