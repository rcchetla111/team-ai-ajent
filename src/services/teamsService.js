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

  // Add these methods to the TeamsService class in teamsService.js

// Find team members by search term (POC Feature 1.1)
async findTeamMembers(searchTerm) {
  if (!this.isAvailable()) {
    // Return simulated data when Teams is not available
    return [
      {
        id: `user-${Date.now()}-1`,
        name: `${searchTerm} Smith`,
        email: `${searchTerm.toLowerCase()}@company.com`,
        jobTitle: "Software Engineer",
        department: "Engineering"
      },
      {
        id: `user-${Date.now()}-2`, 
        name: `${searchTerm} Johnson`,
        email: `${searchTerm.toLowerCase()}.johnson@company.com`,
        jobTitle: "Product Manager",
        department: "Product"
      }
    ];
  }

  try {
    const accessToken = await authService.getAppOnlyToken();
    const filter = `$filter=startswith(displayName, '${searchTerm}') or startswith(givenName, '${searchTerm}') or startswith(surname, '${searchTerm}')`;
    const select = `$select=id,displayName,userPrincipalName,jobTitle,department`;
    const url = `${this.graphEndpoint}/users?${filter}&${select}&$top=20`;

    logger.info(`🔍 Searching for team members: ${searchTerm}`);
    
    const response = await axios.get(url, {
      headers: { 'Authorization': `Bearer ${accessToken}` }
    });

    const users = response.data.value || [];
    return users.map(user => ({
      id: user.id,
      name: user.displayName,
      email: user.userPrincipalName,
      jobTitle: user.jobTitle || 'N/A',
      department: user.department || 'N/A'
    }));

  } catch (error) {
    logger.error('❌ Failed to search team members:', error);
    throw new Error(`Team member search failed: ${error.message}`);
  }
}

// Get all team members (POC Feature 1.1)
async getAllTeamMembers(limit = 50) {
  if (!this.isAvailable()) {
    // Return simulated data when Teams is not available
    return [
      {
        id: "user-1",
        name: "John Smith",
        email: "john.smith@company.com",
        jobTitle: "Software Engineer",
        department: "Engineering"
      },
      {
        id: "user-2",
        name: "Sarah Johnson", 
        email: "sarah.johnson@company.com",
        jobTitle: "Product Manager",
        department: "Product"
      },
      {
        id: "user-3",
        name: "Mike Wilson",
        email: "mike.wilson@company.com",
        jobTitle: "Design Lead",
        department: "Design"
      },
      {
        id: "user-4",
        name: "Lisa Chen",
        email: "lisa.chen@company.com",
        jobTitle: "Data Scientist", 
        department: "Analytics"
      },
      {
        id: "user-5",
        name: "David Brown",
        email: "david.brown@company.com",
        jobTitle: "DevOps Engineer",
        department: "Engineering"
      }
    ];
  }

  try {
    const accessToken = await authService.getAppOnlyToken();
    const select = `$select=id,displayName,userPrincipalName,jobTitle,department`;
    const url = `${this.graphEndpoint}/users?${select}&$top=${limit}`;

    logger.info(`📋 Getting team members (limit: ${limit})`);
    
    const response = await axios.get(url, {
      headers: { 'Authorization': `Bearer ${accessToken}` }
    });

    const users = response.data.value || [];
    return users.map(user => ({
      id: user.id,
      name: user.displayName,
      email: user.userPrincipalName,
      jobTitle: user.jobTitle || 'N/A',
      department: user.department || 'N/A'
    }));

  } catch (error) {
    logger.error('❌ Failed to get team members:', error);
    throw new Error(`Get team members failed: ${error.message}`);
  }
}

  // Build recurrence pattern for recurring meetings
  buildRecurrencePattern(recurrence) {
    if (!recurrence || !recurrence.frequency) {
      return null;
    }

    const pattern = {
      type: recurrence.frequency, // "daily", "weekly", "monthly"
      interval: recurrence.interval || 1,
    };

    if (recurrence.frequency === 'weekly' && recurrence.daysOfWeek) {
      pattern.daysOfWeek = recurrence.daysOfWeek;
    }

    const range = {
      type: recurrence.rangeType || 'endDate',
      startDate: moment(recurrence.startDate).format('YYYY-MM-DD'),
    };

    if (range.type === 'endDate' && recurrence.endDate) {
      range.endDate = moment(recurrence.endDate).format('YYYY-MM-DD');
    } else if (range.type === 'numbered' && recurrence.occurrences) {
      range.numberOfOccurrences = recurrence.occurrences;
    } else {
      range.type = 'endDate';
      range.endDate = moment(recurrence.startDate).add(1, 'year').format('YYYY-MM-DD');
    }

    return { pattern, range };
  }

  // Create Teams meeting (POC Core Feature 1.1)


// Replace your createTeamsMeeting method with this version that has better error handling:

// Enhanced version that adds organizer info to meeting subject and body

async createTeamsMeeting(meetingData) {
  if (!this.isAvailable()) {
    throw new Error('Teams service not available - check Azure AD configuration');
  }

  try {
    const accessToken = await authService.getAppOnlyToken();
    const { subject, startTime, endTime, attendees = [], recurrence, description } = meetingData;

    // Use the specific organizer email from environment variable
    const organizerEmail = process.env.MEETING_ORGANIZER_EMAIL || 'support@legacynote.ai';
    
    // Get the specific user ID for the organizer
    logger.info(`🔍 Creating meeting for organizer: ${organizerEmail}`);
    
    const userResponse = await axios.get(
      `${this.graphEndpoint}/users/${encodeURIComponent(organizerEmail)}?$select=id,displayName,userPrincipalName`,
      { headers: { 'Authorization': `Bearer ${accessToken}` } }
    );

    if (!userResponse.data || !userResponse.data.id) {
      throw new Error(`Organizer ${organizerEmail} not found in tenant`);
    }

    const organizerUserId = userResponse.data.id;
    const organizerName = userResponse.data.displayName;
    
    logger.info(` Found organizer: ${organizerName} (${organizerEmail}) - ID: ${organizerUserId}`);

    // Enhanced event details with clearer organizer info
    const eventDetails = {
      subject: `${subject}`, // Keep subject clean
      body: {
        contentType: "html",
        content: `
          <div>
            <p><strong>Meeting organized by:</strong> ${organizerName} (${organizerEmail})</p>
            ${description ? `<p><strong>Description:</strong> ${description}</p>` : ''}
            <p><em>This meeting was created via AI Agent system</em></p>
          </div>
        `
      },
      start: {
        dateTime: startTime,
        timeZone: "UTC"
      },
      end: {
        dateTime: endTime,
        timeZone: "UTC"
      },
      isOnlineMeeting: true,
      onlineMeetingProvider: "teamsForBusiness",
      // Explicitly set importance to ensure visibility
      importance: "normal",
      sensitivity: "normal"
    };

    // Add recurrence pattern if provided
    const recurrencePattern = this.buildRecurrencePattern(recurrence);
    if (recurrencePattern) {
      eventDetails.recurrence = recurrencePattern;
      logger.info('🔄 Creating recurring meeting');
    }

    // Prepare attendees list (don't include organizer - they're automatic)
    const inviteeAttendees = attendees.filter(email => email !== organizerEmail);

    if (inviteeAttendees.length > 0) {
      eventDetails.attendees = inviteeAttendees.map(email => ({
        emailAddress: { 
          address: email, 
          name: email.split('@')[0]
        },
        type: 'required'
      }));
    }

    // Create the meeting using the specific organizer's calendar
    logger.info(`🔄 Creating Teams meeting on ${organizerName}'s calendar`);
    
    const response = await axios.post(
      `${this.graphEndpoint}/users/${organizerUserId}/events`,
      eventDetails,
      { 
        headers: { 
          'Authorization': `Bearer ${accessToken}`, 
          'Content-Type': 'application/json',
          'Prefer': 'return=representation'  // Get full response back
        } 
      }
    );

    const eventData = response.data;
    
    logger.info('✅ Teams meeting created successfully', {
      meetingId: eventData.id,
      subject: eventData.subject,
      organizer: `${organizerName} (${organizerEmail})`,
      organizerInResponse: eventData.organizer?.emailAddress?.address,
      attendeesCount: inviteeAttendees.length
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
      isRecurring: !!eventData.recurrence,
      organizer: {
        name: organizerName,
        email: organizerEmail,
        userId: organizerUserId,
        confirmedInResponse: eventData.organizer?.emailAddress?.address === organizerEmail
      }
    };

  } catch (error) {
    // Better error logging that avoids circular references
    const errorInfo = {
      message: error.message,
      status: error.response?.status,
      statusText: error.response?.statusText,
      data: error.response?.data,
      url: error.config?.url,
      method: error.config?.method
    };
    
    logger.error('❌ Failed to create Teams meeting:', errorInfo);
    
    // More specific error messages
    if (error.response?.status === 403) {
      throw new Error(`Permission denied: Check if your app has Calendars.ReadWrite permission for ${organizerEmail}`);
    } else if (error.response?.status === 404) {
      throw new Error(`User not found: ${organizerEmail} does not exist or is not accessible`);
    } else if (error.response?.status === 400) {
      throw new Error(`Bad request: ${error.response?.data?.error?.message || 'Invalid meeting data'}`);
    } else {
      throw new Error(`Teams meeting creation failed: ${error.message}`);
    }
  }
}

  // Find users by display names (POC Feature 1.1)
  async findUsersByDisplayName(displayNames) {
    if (!this.isAvailable()) {
      throw new Error('Teams service not available for user lookup');
    }

    if (!Array.isArray(displayNames) || displayNames.length === 0) {
      return [];
    }

    try {
      const accessToken = await authService.getAppOnlyToken();
      const resolvedUsers = [];

      for (const name of displayNames) {
        const filter = `$filter=startswith(displayName, '${name}')`;
        const select = `$select=displayName,userPrincipalName`;
        const url = `${this.graphEndpoint}/users?${filter}&${select}`;

        logger.info(`🔍 Searching for user: ${name}`);
        
        const response = await axios.get(url, {
          headers: { 'Authorization': `Bearer ${accessToken}` }
        });

        if (response.data.value && response.data.value.length > 0) {
          const foundUser = response.data.value[0];
          resolvedUsers.push({
            name: foundUser.displayName,
            email: foundUser.userPrincipalName
          });
          logger.info(`✅ Found user: ${foundUser.displayName}`);
        } else {
          logger.warn(`⚠️ Could not find user: ${name}`);
        }
      }

      return resolvedUsers;

    } catch (error) {
      logger.error('❌ Failed to find users:', error);
      throw new Error(`User lookup failed: ${error.message}`);
    }
  }

  // Get available meeting times (POC Feature 1.1)
  async findMeetingTimes(attendees, duration = 30) {
    if (!this.isAvailable()) {
      throw new Error('Teams service not available for finding meeting times');
    }

    try {
      const accessToken = await authService.getAppOnlyToken();
      
      const findMeetingTimesRequest = {
        attendees: attendees.map(email => ({
          emailAddress: { address: email, name: email.split('@')[0] }
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

      logger.info('✅ Retrieved meeting time suggestions');
      return response.data.meetingTimeSuggestions || [];

    } catch (error) {
      logger.error('❌ Failed to get meeting times:', error);
      throw new Error(`Finding meeting times failed: ${error.message}`);
    }
  }


  // Create Teams Channel (New Feature)
async createTeamsChannel(channelData) {
  if (!this.isAvailable()) {
    throw new Error('Teams service not available - check Azure AD configuration');
  }

  try {
    const accessToken = await authService.getAppOnlyToken();
    const { teamId, displayName, description, membershipType = 'standard' } = channelData;

    // Channel creation payload
    const channelDetails = {
      displayName: displayName,
      description: description || `${displayName} - Created via AI Agent`,
      membershipType: membershipType, // 'standard' or 'private'
      '@odata.type': '#Microsoft.Graph.channel'
    };

    logger.info(`🆕 Creating Teams channel: ${displayName} in team: ${teamId}`);
    
    const response = await axios.post(
      `${this.graphEndpoint}/teams/${teamId}/channels`,
      channelDetails,
      { 
        headers: { 
          'Authorization': `Bearer ${accessToken}`, 
          'Content-Type': 'application/json'
        } 
      }
    );

    const channelDataResult = response.data;
    
    logger.info('✅ Teams channel created successfully', {
      channelId: channelDataResult.id,
      displayName: channelDataResult.displayName,
      teamId: teamId
    });

    return {
      success: true,
      channelId: channelDataResult.id,
      displayName: channelDataResult.displayName,
      description: channelDataResult.description,
      webUrl: channelDataResult.webUrl,
      membershipType: channelDataResult.membershipType,
      teamId: teamId,
      createdDateTime: channelDataResult.createdDateTime
    };

  } catch (error) {
    logger.error('❌ Failed to create Teams channel:', error);
    
    if (error.response?.status === 403) {
      throw new Error(`Permission denied: Check if your app has Channel.Create permission for team ${teamId}`);
    } else if (error.response?.status === 404) {
      throw new Error(`Team not found: ${teamId} does not exist or is not accessible`);
    } else if (error.response?.status === 409) {
      throw new Error(`Channel already exists: A channel with name "${displayName}" already exists in this team`);
    } else {
      throw new Error(`Teams channel creation failed: ${error.message}`);
    }
  }
}

// Get Teams/Groups that user can create channels in
async getAvailableTeams() {
  if (!this.isAvailable()) {
    console.log("⚠️ Teams service not available, returning mock data");
    return [
      {
        id: "team-1",
        displayName: "Engineering Team",
        description: "Software development team"
      },
      {
        id: "team-2", 
        displayName: "Product Team",
        description: "Product management team"
      }
    ];
  }

  try {
    const accessToken = await authService.getAppOnlyToken();
    console.log("✅ Got access token, length:", accessToken.length);
    
    logger.info('📋 Getting available teams for channel creation');
    
    // Try the simpler groups endpoint first
    const url = `${this.graphEndpoint}/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName,description&$top=10`;
    console.log("🔍 Calling URL:", url);
    
    const response = await axios.get(url, {
      headers: { 'Authorization': `Bearer ${accessToken}` }
    });

    console.log("✅ Graph API response status:", response.status);
    console.log("✅ Groups found:", response.data.value?.length || 0);

    const teams = response.data.value || [];
    const result = teams.map(team => ({
      id: team.id,
      displayName: team.displayName,
      description: team.description || 'No description'
    }));

    console.log("✅ Processed teams:", result.length);
    return result;

  } catch (error) {
    console.error('❌ getAvailableTeams error details:', {
      message: error.message,
      status: error.response?.status,
      statusText: error.response?.statusText,
      errorData: error.response?.data
    });
    
    logger.error('❌ Failed to get available teams:', error);
    throw new Error(`Get teams failed: ${error.message}`);
  }
}

// List channels in a team
async getTeamChannels(teamId) {
  if (!this.isAvailable()) {
    return [
      {
        id: "channel-1",
        displayName: "General",
        description: "General discussion"
      }
    ];
  }

  try {
    const accessToken = await authService.getAppOnlyToken();
    
    // FIX: The URL was wrong - it had /teams/teams/ instead of /teams/{teamId}/
    const response = await axios.get(
      `${this.graphEndpoint}/teams/${teamId}/channels`, // CORRECTED URL
      { headers: { 'Authorization': `Bearer ${accessToken}` } }
    );

    const channels = response.data.value || [];
    return channels.map(channel => ({
      id: channel.id,
      displayName: channel.displayName,
      description: channel.description,
      membershipType: channel.membershipType,
      webUrl: channel.webUrl
    }));

  } catch (error) {
    logger.error('❌ Failed to get team channels:', error);
    throw new Error(`Get channels failed: ${error.message}`);
  }
}




  // Get service status
 // Update the existing getStatus method to include the new features
getStatus() {
  return {
    available: this.isAvailable(),
    authConfigured: authService.isAvailable(),
    features: {
      createMeetings: this.isAvailable(),
      recurringMeetings: this.isAvailable(),
      userResolution: this.isAvailable(),
      timeOptimization: this.isAvailable(),
      teamMemberSearch: this.isAvailable(),
      getAllTeamMembers: this.isAvailable(),
      createChannels: this.isAvailable(),        // ADD
      listTeams: this.isAvailable(),             // ADD  
      listChannels: this.isAvailable()           // ADD
    }
  };
}
}

// Create singleton instance
const teamsService = new TeamsService();
module.exports = teamsService;