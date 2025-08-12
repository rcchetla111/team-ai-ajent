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
  async createTeamsMeeting(meetingData) {
    if (!this.isAvailable()) {
      throw new Error('Teams service not available - check Azure AD configuration');
    }

    try {
      const accessToken = await authService.getAppOnlyToken();
      const { subject, startTime, endTime, attendees = [], recurrence } = meetingData;

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

      // Add recurrence pattern if provided
      const recurrencePattern = this.buildRecurrencePattern(recurrence);
      if (recurrencePattern) {
        eventDetails.recurrence = recurrencePattern;
        logger.info('🔄 Creating recurring meeting');
      }

      // Add attendees if provided
      if (attendees.length > 0) {
        eventDetails.attendees = attendees.map(email => ({
          emailAddress: { address: email, name: email.split('@')[0] },
          type: 'required'
        }));
      }

      // Get a user to create meeting on behalf of
      const usersResponse = await axios.get(
        `${this.graphEndpoint}/users?$top=1&$select=id`,
        { headers: { 'Authorization': `Bearer ${accessToken}` } }
      );

      if (!usersResponse.data.value || usersResponse.data.value.length === 0) {
        throw new Error('No users found in tenant to create meeting');
      }

      const userId = usersResponse.data.value[0].id;
      logger.info(`🔄 Creating meeting for user ${userId}`);

      const response = await axios.post(
        `${this.graphEndpoint}/users/${userId}/events`,
        eventDetails,
        { headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' } }
      );

      const eventData = response.data;
      logger.info('✅ Teams meeting created successfully');

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
      logger.error('❌ Failed to create Teams meeting:', error);
      throw new Error(`Teams meeting creation failed: ${error.message}`);
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
      teamMemberSearch: this.isAvailable(),  // Add this line
      getAllTeamMembers: this.isAvailable()  // Add this line
    }
  };
}
}

// Create singleton instance
const teamsService = new TeamsService();
module.exports = teamsService;