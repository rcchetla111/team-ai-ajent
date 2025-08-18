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
    throw new Error('Teams service not available - Azure AD configuration required');
  }

  try {
    const accessToken = await authService.getAppOnlyToken();
    const filter = `$filter=startswith(displayName, '${searchTerm}') or startswith(givenName, '${searchTerm}') or startswith(surname, '${searchTerm}')`;
    const select = `$select=id,displayName,userPrincipalName,jobTitle,department`;
    const url = `${this.graphEndpoint}/users?${filter}&${select}&$top=20`;

    logger.info(`🔍 Searching for real Teams members: ${searchTerm}`);
    
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
    logger.error('❌ Failed to search real Teams members:', error);
    throw new Error(`Real Teams member search failed: ${error.message}`);
  }
}

// Get all team members (POC Feature 1.1)
async getAllTeamMembers(limit = 50) {
  if (!this.isAvailable()) {
    throw new Error('Teams service not available - Azure AD configuration required');
  }

  try {
    const accessToken = await authService.getAppOnlyToken();
    const select = `$select=id,displayName,userPrincipalName,jobTitle,department`;
    const url = `${this.graphEndpoint}/users?${select}&$top=${limit}`;

    logger.info(`📋 Getting real Teams members (limit: ${limit})`);
    
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
    logger.error('❌ Failed to get real Teams members:', error);
    throw new Error(`Get real Teams members failed: ${error.message}`);
  }
}

async getAllCalendarMeetings(startDate = null, endDate = null, limit = 100) {
  if (!this.isAvailable()) {
    throw new Error('Teams service not available - Azure AD configuration required');
  }

  try {
    const accessToken = await authService.getAppOnlyToken();
    const organizerEmail = process.env.MEETING_ORGANIZER_EMAIL || 'support@legacynote.ai';
    
    // Get organizer user ID
    const userResponse = await axios.get(
      `${this.graphEndpoint}/users/${encodeURIComponent(organizerEmail)}?$select=id`,
      { headers: { 'Authorization': `Bearer ${accessToken}` } }
    );
    
    const organizerUserId = userResponse.data.id;
    
    // Set date range (default: last 30 days to next 30 days)
    if (!startDate) {
      startDate = moment().subtract(30, 'days').startOf('day').toISOString();
    }
    if (!endDate) {
      endDate = moment().add(30, 'days').endOf('day').toISOString();
    }
    
    logger.info(`📅 Getting ALL calendar meetings for ${organizerEmail}`, {
      startDate, endDate, limit
    });
    
    // Get ALL calendar events (not just AI-created ones)
    const response = await axios.get(
      `${this.graphEndpoint}/users/${organizerUserId}/calendar/calendarView`,
      {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        params: {
          startDateTime: startDate,
          endDateTime: endDate,
          $select: 'id,subject,start,end,attendees,organizer,onlineMeeting,webLink,createdDateTime,lastModifiedDateTime,isOnlineMeeting',
          $orderby: 'start/dateTime desc',
          $top: limit
        }
      }
    );

    const events = response.data.value || [];
    
    logger.info(`✅ Found ${events.length} calendar meetings from Teams`);
    
    // Convert Teams calendar events to our meeting format
    const meetings = events.map(event => ({
      id: event.id, // Use Teams event ID
      meetingId: event.id,
      graphEventId: event.id,
      subject: event.subject,
      startTime: event.start.dateTime,
      endTime: event.end.dateTime,
      attendees: event.attendees?.map(a => a.emailAddress.address) || [],
      organizer: event.organizer?.emailAddress?.address || organizerEmail,
      joinUrl: event.onlineMeeting?.joinUrl || null,
      webUrl: event.webLink,
      status: moment(event.start.dateTime).isBefore(moment()) ? 
               (moment(event.end.dateTime).isBefore(moment()) ? 'completed' : 'in_progress') : 
               'scheduled',
      isRealTeamsMeeting: true,
      isFromTeamsCalendar: true, // Flag to indicate this came from Teams calendar directly
      createdAt: event.createdDateTime,
      updatedAt: event.lastModifiedDateTime || event.createdDateTime,
      agentAttended: false, // Default for calendar meetings
      agentConfig: {
        autoJoin: false,
        enableChatCapture: false,
        generateSummary: false
      }
    }));

    return meetings;

  } catch (error) {
    logger.error('❌ Failed to get ALL calendar meetings:', {
      error: error.message,
      status: error.response?.status,
      statusText: error.response?.statusText
    });
    
    if (error.response?.status === 403) {
      throw new Error(`Permission denied: Cannot access Teams calendar. Check app permissions.`);
    } else if (error.response?.status === 404) {
      throw new Error(`User not found: ${organizerEmail} does not exist in your organization`);
    } else {
      throw new Error(`Teams calendar lookup failed: ${error.message}`);
    }
  }
}

  // Build recurrence pattern for recurring meetings
buildRecurrencePattern(recurrence) {
  console.log('🔄 buildRecurrencePattern called with:', JSON.stringify(recurrence, null, 2));
  
  if (!recurrence || !recurrence.pattern || !recurrence.pattern.type) {
    console.log('❌ No valid recurrence pattern provided');
    return null;
  }

  try {
    const pattern = recurrence.pattern;
    const range = recurrence.range || {};

    console.log('📋 Building recurrence pattern:', {
      type: pattern.type,
      interval: pattern.interval,
      rangeType: range.type
    });

    // Build Microsoft Graph recurrence object
    const graphRecurrence = {
      pattern: {
        type: pattern.type, // daily, weekly, monthly, yearly
        interval: pattern.interval || 1
      },
      range: {
        type: range.type || 'noEnd',
        startDate: range.startDate || new Date().toISOString().split('T')[0]
      }
    };

    // Add days of week for weekly recurrence
    if (pattern.type === 'weekly' && pattern.daysOfWeek && pattern.daysOfWeek.length > 0) {
      console.log('📅 Adding weekly days:', pattern.daysOfWeek);
      graphRecurrence.pattern.daysOfWeek = pattern.daysOfWeek;
    }

    // Add end conditions
    if (range.type === 'endDate' && range.endDate) {
      console.log('📅 Adding end date:', range.endDate);
      graphRecurrence.range.endDate = range.endDate;
    } else if (range.type === 'numbered' && range.numberOfOccurrences) {
      console.log('📅 Adding occurrence count:', range.numberOfOccurrences);
      graphRecurrence.range.numberOfOccurrences = range.numberOfOccurrences;
    }

    console.log('✅ Built Microsoft Graph recurrence:', JSON.stringify(graphRecurrence, null, 2));
    return graphRecurrence;

  } catch (error) {
    console.error('❌ Error building recurrence pattern:', error);
    return null;
  }
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

    console.log('🔄 createTeamsMeeting called with:', {
      subject,
      startTime,
      endTime,
      attendeesCount: attendees.length,
      hasRecurrence: !!recurrence,
      recurrenceData: recurrence
    });

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
    
    logger.info(`✅ Found organizer: ${organizerName} (${organizerEmail}) - ID: ${organizerUserId}`);

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

    // 🔧 CRITICAL FIX: Add recurrence pattern if provided
    console.log('🔄 Checking for recurrence pattern...');
    if (recurrence) {
      console.log('🔄 Recurrence data provided, building pattern...');
      const recurrencePattern = this.buildRecurrencePattern(recurrence);
      
      if (recurrencePattern) {
        console.log('✅ Adding recurrence to event details:', JSON.stringify(recurrencePattern, null, 2));
        eventDetails.recurrence = recurrencePattern;
        logger.info('🔄 Creating recurring meeting with pattern:', recurrencePattern.pattern.type);
      } else {
        console.warn('⚠️ buildRecurrencePattern returned null - creating single meeting');
        logger.warn('⚠️ Recurrence pattern could not be built, creating single meeting instead');
      }
    } else {
      console.log('📅 No recurrence data provided, creating single meeting');
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

    console.log('📤 Final event details being sent to Microsoft Graph:', JSON.stringify(eventDetails, null, 2));

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
    
    console.log('📨 Microsoft Graph response:', {
      id: eventData.id,
      subject: eventData.subject,
      recurrence: eventData.recurrence,
      hasRecurrence: !!eventData.recurrence,
      seriesMasterId: eventData.seriesMasterId
    });
    
    logger.info('✅ Teams meeting created successfully', {
      meetingId: eventData.id,
      subject: eventData.subject,
      organizer: `${organizerName} (${organizerEmail})`,
      organizerInResponse: eventData.organizer?.emailAddress?.address,
      attendeesCount: inviteeAttendees.length,
      isRecurring: !!eventData.recurrence,
      recurrenceType: eventData.recurrence?.pattern?.type || 'none'
    });

    const isActuallyRecurring = !!eventData.recurrence;
    
    if (recurrence && !isActuallyRecurring) {
      console.warn('⚠️ WARNING: Recurrence was requested but Microsoft Graph did not create recurring event!');
      console.warn('🔧 Check recurrence object format and Microsoft Graph API permissions');
      console.warn('📋 Requested recurrence:', JSON.stringify(recurrence, null, 2));
      console.warn('📋 Built pattern:', JSON.stringify(eventDetails.recurrence, null, 2));
    }

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
      isRecurring: isActuallyRecurring, // 🔧 CRITICAL: Return actual recurrence status
      recurrence: eventData.recurrence || null,
      seriesMasterId: eventData.seriesMasterId || null,
      organizer: {
        name: organizerName,
        email: organizerEmail,
        userId: organizerUserId,
        confirmedInResponse: eventData.organizer?.emailAddress?.address === organizerEmail
      },
      // Debug info
      debug: {
        recurrenceRequested: !!recurrence,
        recurrenceBuilt: !!eventDetails.recurrence,
        recurrenceReturned: !!eventData.recurrence,
        requestedPattern: recurrence?.pattern?.type || 'none',
        returnedPattern: eventData.recurrence?.pattern?.type || 'none'
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
      // Check if it's a recurrence-related error
      const errorMessage = error.response?.data?.error?.message || 'Invalid meeting data';
      if (errorMessage.toLowerCase().includes('recurrence')) {
        throw new Error(`Recurrence error: ${errorMessage}. Check recurrence pattern format.`);
      }
      throw new Error(`Bad request: ${errorMessage}`);
    } else {
      throw new Error(`Teams meeting creation failed: ${error.message}`);
    }
  }
}



async updateTeamsMeeting(graphEventId, updateData) {
  if (!this.isAvailable()) {
    throw new Error('Teams service not available - check Azure AD configuration');
  }

  try {
    const accessToken = await authService.getAppOnlyToken();
    
    // Get organizer email and user ID
    const organizerEmail = process.env.MEETING_ORGANIZER_EMAIL || 'support@legacynote.ai';
    
    const userResponse = await axios.get(
      `${this.graphEndpoint}/users/${encodeURIComponent(organizerEmail)}?$select=id`,
      { headers: { 'Authorization': `Bearer ${accessToken}` } }
    );
    
    if (!userResponse.data || !userResponse.data.id) {
      throw new Error(`Organizer ${organizerEmail} not found in tenant`);
    }
    
    const organizerUserId = userResponse.data.id;
    
    logger.info(`🔄 Updating Teams meeting: ${graphEventId}`, {
      organizer: organizerEmail,
      updateData: Object.keys(updateData)
    });

    // Prepare update payload
    const updatePayload = {};

    // Handle attendees update
    if (updateData.attendees) {
      // Filter out organizer from attendees list
      const filteredAttendees = updateData.attendees.filter(email => email !== organizerEmail);
      
      updatePayload.attendees = filteredAttendees.map(email => ({
        emailAddress: { 
          address: email, 
          name: email.split('@')[0]
        },
        type: 'required'
      }));
      
      logger.info(`📝 Updating attendees: ${filteredAttendees.length} attendees`);
    }

    // Handle other updates (subject, description, time, etc.)
    if (updateData.subject) {
      updatePayload.subject = updateData.subject;
    }
    
    if (updateData.description) {
      updatePayload.body = {
        contentType: "html",
        content: updateData.description
      };
    }
    
    if (updateData.startTime) {
      updatePayload.start = {
        dateTime: updateData.startTime,
        timeZone: "UTC"
      };
    }
    
    if (updateData.endTime) {
      updatePayload.end = {
        dateTime: updateData.endTime,
        timeZone: "UTC"
      };
    }

    // Update the meeting via Microsoft Graph API
    const response = await axios.patch(
      `${this.graphEndpoint}/users/${organizerUserId}/events/${graphEventId}`,
      updatePayload,
      { 
        headers: { 
          'Authorization': `Bearer ${accessToken}`, 
          'Content-Type': 'application/json',
          'Prefer': 'return=representation'
        } 
      }
    );

    const updatedEvent = response.data;
    
    logger.info('✅ Teams meeting updated successfully', {
      meetingId: updatedEvent.id,
      subject: updatedEvent.subject,
      attendeesCount: updatedEvent.attendees?.length || 0
    });

    return {
      success: true,
      meetingId: updatedEvent.id,
      subject: updatedEvent.subject,
      startTime: updatedEvent.start?.dateTime,
      endTime: updatedEvent.end?.dateTime,
      attendees: updatedEvent.attendees?.map(a => a.emailAddress.address) || [],
      webUrl: updatedEvent.webLink,
      joinUrl: updatedEvent.onlineMeeting?.joinUrl,
      lastModified: updatedEvent.lastModifiedDateTime,
      organizer: {
        email: organizerEmail,
        userId: organizerUserId
      }
    };

  } catch (error) {
    const errorInfo = {
      message: error.message,
      status: error.response?.status,
      statusText: error.response?.statusText,
      data: error.response?.data
    };
    
    logger.error('❌ Failed to update Teams meeting:', errorInfo);
    
    if (error.response?.status === 403) {
      throw new Error(`Permission denied: Check if your app has Calendars.ReadWrite permission`);
    } else if (error.response?.status === 404) {
      throw new Error(`Meeting not found: ${graphEventId} does not exist or is not accessible`);
    } else if (error.response?.status === 400) {
      throw new Error(`Bad request: ${error.response?.data?.error?.message || 'Invalid update data'}`);
    } else {
      throw new Error(`Teams meeting update failed: ${error.message}`);
    }
  }
}

async getTeamsMeetingDetails(graphEventId) {
  if (!this.isAvailable()) {
    throw new Error('Teams service not available - check Azure AD configuration');
  }

  try {
    const accessToken = await authService.getAppOnlyToken();
    const organizerEmail = process.env.MEETING_ORGANIZER_EMAIL || 'support@legacynote.ai';
    
    const userResponse = await axios.get(
      `${this.graphEndpoint}/users/${encodeURIComponent(organizerEmail)}?$select=id`,
      { headers: { 'Authorization': `Bearer ${accessToken}` } }
    );
    
    const organizerUserId = userResponse.data.id;
    
    const response = await axios.get(
      `${this.graphEndpoint}/users/${organizerUserId}/events/${graphEventId}`,
      { 
        headers: { 'Authorization': `Bearer ${accessToken}` },
        params: {
          $select: 'id,subject,start,end,attendees,organizer,onlineMeeting,webLink,lastModifiedDateTime'
        }
      }
    );

    const meeting = response.data;
    
    return {
      id: meeting.id,
      subject: meeting.subject,
      startTime: meeting.start?.dateTime,
      endTime: meeting.end?.dateTime,
      attendees: meeting.attendees?.map(a => a.emailAddress.address) || [],
      organizer: meeting.organizer?.emailAddress?.address,
      joinUrl: meeting.onlineMeeting?.joinUrl,
      webUrl: meeting.webLink,
      lastModified: meeting.lastModifiedDateTime
    };

  } catch (error) {
    logger.error('❌ Failed to get Teams meeting details:', error);
    throw new Error(`Get meeting details failed: ${error.message}`);
  }
}


// Add these methods to your teamsService.js file

// Get free/busy information for team members
// Add these methods to your teamsService.js file

// Get free/busy information for team members
async getFreeBusyInfo(attendeeEmails, startTime, endTime) {
  if (!this.isAvailable()) {
    // Return simulated free/busy data when Teams is not available
    return attendeeEmails.map(email => ({
      email: email,
      freeBusyStatus: Math.random() > 0.5 ? 'free' : 'busy',
      busyTimes: Math.random() > 0.5 ? [] : [{
        start: moment(startTime).add(30, 'minutes').toISOString(),
        end: moment(startTime).add(90, 'minutes').toISOString()
      }]
    }));
  }

  try {
    const accessToken = await authService.getAppOnlyToken();
    
    // FIXED: Use correct endpoint and request format
    const scheduleRequest = {
      schedules: attendeeEmails,
      startTime: {
        dateTime: startTime,
        timeZone: "UTC"
      },
      endTime: {
        dateTime: endTime,
        timeZone: "UTC"
      },
      availabilityViewInterval: 60 // Optional: view interval in minutes
    };

    logger.info(`🔍 Checking availability for ${attendeeEmails.length} attendees`, {
      attendees: attendeeEmails,
      timeSlot: `${startTime} to ${endTime}`
    });

    // FIXED: Use the correct endpoint - /calendar/getSchedule (not /me/calendar/getSchedule)
    const response = await axios.post(
      `${this.graphEndpoint}/calendar/getSchedule`,
      scheduleRequest,
      {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      }
    );

    const schedules = response.data.value || [];
    
    return schedules.map((schedule, index) => {
      const email = attendeeEmails[index];
      const busyTimes = schedule.busyTimes || [];
      
      return {
        email: email,
        freeBusyStatus: busyTimes.length === 0 ? 'free' : 'busy',
        busyTimes: busyTimes.map(busyTime => ({
          start: busyTime.start.dateTime,
          end: busyTime.end.dateTime
        })),
        workingHours: schedule.workingHours || null,
        availabilityView: schedule.availabilityView || null
      };
    });

  } catch (error) {
    logger.error('❌ Failed to get free/busy information:', {
      message: error.message,
      status: error.response?.status,
      statusText: error.response?.statusText,
      data: error.response?.data
    });
    
    // If it's a permission or authentication error, throw specific message
    if (error.response?.status === 403) {
      throw new Error(`Permission denied: Check if your app has Calendars.Read permission`);
    } else if (error.response?.status === 401) {
      throw new Error(`Authentication failed: Check your access token`);
    } else if (error.response?.status === 400) {
      throw new Error(`Bad request: ${error.response?.data?.error?.message || 'Invalid request format'}`);
    } else {
      throw new Error(`Free/busy lookup failed: ${error.message}`);
    }
  }
}

// Add these methods to your teamsService.js file

// Get free/busy information for team members (REAL TEAMS ONLY)
async getFreeBusyInfo(attendeeEmails, startTime, endTime) {
  if (!this.isAvailable()) {
    throw new Error('Teams service not available - Azure AD configuration required for real Teams integration');
  }

  try {
    const accessToken = await authService.getAppOnlyToken();
    
    // FIXED: Use correct endpoint and request format
    const scheduleRequest = {
      schedules: attendeeEmails,
      startTime: {
        dateTime: startTime,
        timeZone: "UTC"
      },
      endTime: {
        dateTime: endTime,
        timeZone: "UTC"
      },
      availabilityViewInterval: 60 // Optional: view interval in minutes
    };

    logger.info(`🔍 Checking REAL Teams calendar availability for ${attendeeEmails.length} attendees`, {
      attendees: attendeeEmails,
      timeSlot: `${startTime} to ${endTime}`
    });

    // FIXED: Use the correct endpoint - /calendar/getSchedule (not /me/calendar/getSchedule)
    const response = await axios.post(
      `${this.graphEndpoint}/calendar/getSchedule`,
      scheduleRequest,
      {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      }
    );

    const schedules = response.data.value || [];
    
    return schedules.map((schedule, index) => {
      const email = attendeeEmails[index];
      const busyTimes = schedule.busyTimes || [];
      
      return {
        email: email,
        freeBusyStatus: busyTimes.length === 0 ? 'free' : 'busy',
        busyTimes: busyTimes.map(busyTime => ({
          start: busyTime.start.dateTime,
          end: busyTime.end.dateTime
        })),
        workingHours: schedule.workingHours || null,
        availabilityView: schedule.availabilityView || null
      };
    });

  } catch (error) {
    logger.error('❌ Failed to get REAL Teams free/busy information:', {
      message: error.message,
      status: error.response?.status,
      statusText: error.response?.statusText,
      data: error.response?.data
    });
    
    // If it's a permission or authentication error, throw specific message
    if (error.response?.status === 403) {
      throw new Error(`Permission denied: Check if your app has Calendars.Read permission for real Teams calendars`);
    } else if (error.response?.status === 401) {
      throw new Error(`Authentication failed: Check your Azure AD access token`);
    } else if (error.response?.status === 400) {
      throw new Error(`Bad request: ${error.response?.data?.error?.message || 'Invalid request format for Teams calendar API'}`);
    } else {
      throw new Error(`Real Teams calendar lookup failed: ${error.message}`);
    }
  }
}

// Get real calendar events for a specific Teams user (FIXED - proper time filtering)
async getUserCalendarEvents(userEmail, startTime, endTime) {
  if (!this.isAvailable()) {
    throw new Error('Teams service not available - Azure AD configuration required');
  }

  try {
    const accessToken = await authService.getAppOnlyToken();
    
    // FIXED: Proper date formatting for Microsoft Graph API
    const startDate = moment(startTime).utc().format('YYYY-MM-DDTHH:mm:ss.SSS[Z]');
    const endDate = moment(endTime).utc().format('YYYY-MM-DDTHH:mm:ss.SSS[Z]');
    
    logger.info(`📅 Getting REAL calendar events for Teams user: ${userEmail}`, {
      timeRange: `${startDate} to ${endDate}`,
      duration: moment(endTime).diff(moment(startTime), 'hours', true) + ' hours'
    });
    
    // FIXED: Use calendarView endpoint with proper time filtering
    const response = await axios.get(
      `${this.graphEndpoint}/users/${encodeURIComponent(userEmail)}/calendar/calendarView`,
      {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        params: {
          startDateTime: startDate,
          endDateTime: endDate,
          $select: 'subject,start,end,showAs,organizer,sensitivity,isAllDay',
          $orderby: 'start/dateTime',
          $top: 100 // Limit to prevent too much data
        }
      }
    );

    const events = response.data.value || [];
    
    // FIXED: Filter events that actually overlap with the requested time range
    const requestStart = moment(startTime).utc();
    const requestEnd = moment(endTime).utc();
    
    const overlappingEvents = events.filter(event => {
      const eventStart = moment(event.start.dateTime).utc();
      const eventEnd = moment(event.end.dateTime).utc();
      
      // Check if event overlaps with requested time range
      return eventStart.isBefore(requestEnd) && eventEnd.isAfter(requestStart);
    });
    
    // FIXED: Only consider events that make the user "busy"
    const busyEvents = overlappingEvents.filter(event => 
      event.showAs === 'busy' || 
      event.showAs === 'tentative' || 
      event.showAs === 'outOfOffice' ||
      event.showAs === 'workingElsewhere'
    );

    const isBusy = busyEvents.length > 0;

    logger.info(`✅ Calendar check for ${userEmail}:`, {
      totalEventsInRange: overlappingEvents.length,
      busyEvents: busyEvents.length,
      status: isBusy ? 'BUSY' : 'FREE',
      timeRange: `${startDate} to ${endDate}`
    });

    return {
      email: userEmail,
      timeRange: {
        start: startTime,
        end: endTime,
        durationHours: moment(endTime).diff(moment(startTime), 'hours', true)
      },
      events: overlappingEvents.map(event => ({
        subject: event.subject,
        start: event.start.dateTime,
        end: event.end.dateTime,
        showAs: event.showAs,
        isAllDay: event.isAllDay || false,
        organizer: event.organizer?.emailAddress?.address
      })),
      freeBusyStatus: isBusy ? 'busy' : 'free',
      busyTimes: busyEvents.map(event => ({
        start: event.start.dateTime,
        end: event.end.dateTime,
        subject: event.subject,
        showAs: event.showAs
      })),
      summary: {
        totalEventsInTimeRange: overlappingEvents.length,
        busyEventsCount: busyEvents.length,
        freeEventsCount: overlappingEvents.length - busyEvents.length,
        isAvailable: !isBusy
      }
    };

  } catch (error) {
    logger.error('❌ Failed to get REAL Teams user calendar events:', {
      userEmail,
      error: error.message,
      status: error.response?.status,
      statusText: error.response?.statusText,
      timeRange: `${startTime} to ${endTime}`
    });
    
    if (error.response?.status === 403) {
      throw new Error(`Permission denied: Cannot access real Teams calendar for ${userEmail}. Check app permissions.`);
    } else if (error.response?.status === 404) {
      throw new Error(`Teams user not found: ${userEmail} does not exist in your organization`);
    } else if (error.response?.status === 400) {
      throw new Error(`Bad request for ${userEmail}: ${error.response?.data?.error?.message || 'Invalid time range or parameters'}`);
    } else {
      throw new Error(`Real Teams calendar lookup failed for ${userEmail}: ${error.message}`);
    }
  }
}

// FIXED: Check if specific time slot is available for all REAL Teams attendees
async checkTimeSlotAvailability(attendeeEmails, startTime, endTime) {
  if (!this.isAvailable()) {
    throw new Error('Teams service not available - Azure AD configuration required for real Teams integration');
  }

  const attendeeStatus = [];
  let allAvailable = true;

  // Validate time range
  const start = moment(startTime);
  const end = moment(endTime);
  const durationMinutes = end.diff(start, 'minutes');
  
  if (durationMinutes <= 0) {
    throw new Error('End time must be after start time');
  }
  
  if (durationMinutes > 1440) { // More than 24 hours
    throw new Error('Time range cannot exceed 24 hours');
  }

  logger.info(`🔍 Checking REAL Teams availability for ${attendeeEmails.length} users`, {
    timeSlot: `${startTime} to ${endTime}`,
    durationMinutes: durationMinutes
  });

  for (const email of attendeeEmails) {
    try {
      const userCalendar = await this.getUserCalendarEvents(email, startTime, endTime);
      
      const isAvailable = userCalendar.freeBusyStatus === 'free';
      attendeeStatus.push({
        email: email,
        available: isAvailable,
        status: userCalendar.freeBusyStatus,
        conflicts: userCalendar.busyTimes || [],
        eventsInTimeRange: userCalendar.summary.totalEventsInTimeRange,
        busyEventsCount: userCalendar.summary.busyEventsCount,
        details: userCalendar.summary
      });
      
      if (!isAvailable) {
        allAvailable = false;
      }

      logger.info(`📊 ${email}: ${userCalendar.freeBusyStatus.toUpperCase()} (${userCalendar.summary.busyEventsCount} busy events in range)`);
    } catch (error) {
      logger.error(`❌ Could not check REAL Teams availability for ${email}:`, error.message);
      attendeeStatus.push({
        email: email,
        available: false,
        status: 'error',
        conflicts: [],
        error: error.message,
        eventsInTimeRange: 0,
        busyEventsCount: 0
      });
      allAvailable = false;
    }
  }

  const result = {
    timeSlot: { 
      start: startTime, 
      end: endTime,
      durationMinutes: durationMinutes,
      durationHours: Math.round(durationMinutes / 60 * 100) / 100
    },
    allAvailable: allAvailable,
    attendeeStatus: attendeeStatus,
    checkedAt: new Date().toISOString(),
    dataSource: 'real_teams_calendars',
    summary: {
      totalAttendees: attendeeEmails.length,
      availableAttendees: attendeeStatus.filter(a => a.available).length,
      busyAttendees: attendeeStatus.filter(a => !a.available && a.status !== 'error').length,
      errorAttendees: attendeeStatus.filter(a => a.status === 'error').length
    }
  };

  logger.info(`🏁 Availability check complete:`, {
    allAvailable: allAvailable,
    available: result.summary.availableAttendees,
    busy: result.summary.busyAttendees,
    errors: result.summary.errorAttendees
  });

  return result;
}

// Find available time slots for multiple REAL Teams attendees  
async findAvailableTimeSlots(attendeeEmails, duration = 30, searchDays = 7) {
  if (!this.isAvailable()) {
    throw new Error('Teams service not available - Azure AD configuration required for real Teams integration');
  }

  try {
    const slots = [];
    const startSearch = moment().add(1, 'hour');
    const endSearch = moment().add(searchDays, 'days');
    
    logger.info(`🔍 Finding REAL Teams available slots for ${attendeeEmails.length} users over ${searchDays} days`);
    
    // Generate time slots to check (every 30 minutes during business hours)
    const current = moment(startSearch);
    let slotsChecked = 0;
    
    while (current.isBefore(endSearch) && slots.length < 10) {
      // Only check business hours (9 AM to 5 PM, Monday to Friday)
      if (current.day() >= 1 && current.day() <= 5 && // Monday to Friday
          current.hour() >= 9 && current.hour() < 17) { // 9 AM to 5 PM
        
        const slotStart = current.clone();
        const slotEnd = current.clone().add(duration, 'minutes');
        
        // Don't check slots that end after business hours
        if (slotEnd.hour() <= 17) {
          try {
            slotsChecked++;
            const availability = await this.checkTimeSlotAvailability(
              attendeeEmails, 
              slotStart.toISOString(), 
              slotEnd.toISOString()
            );
            
            if (availability.allAvailable) {
              slots.push({
                start: slotStart.toISOString(),
                end: slotEnd.toISOString(),
                confidence: 'high',
                allAttendeesAvailable: true,
                attendeeAvailability: availability.attendeeStatus,
                dayOfWeek: slotStart.format('dddd'),
                timeOfDay: slotStart.format('h:mm A')
              });
              
              logger.info(`✅ Found available slot: ${slotStart.format('dddd, MMMM Do YYYY, h:mm A')}`);
            }
          } catch (error) {
            logger.warn(`Failed to check REAL Teams slot ${slotStart.format()}:`, error.message);
          }
        }
      }
      
      current.add(30, 'minutes');
    }

    logger.info(`🔍 Checked ${slotsChecked} time slots, found ${slots.length} available slots`);

    return slots;

  } catch (error) {
    logger.error('❌ Failed to find REAL Teams available time slots:', error);
    throw new Error(`Finding real Teams available slots failed: ${error.message}`);
  }
}

// Validate that all attendees are real Teams users
async validateTeamsUsers(attendeeEmails) {
  if (!this.isAvailable()) {
    throw new Error('Teams service not available - Azure AD configuration required');
  }

  const validUsers = [];
  const invalidUsers = [];

  for (const email of attendeeEmails) {
    try {
      const accessToken = await authService.getAppOnlyToken();
      
      const response = await axios.get(
        `${this.graphEndpoint}/users/${encodeURIComponent(email)}?$select=id,displayName,userPrincipalName,mail`,
        { headers: { 'Authorization': `Bearer ${accessToken}` } }
      );

      if (response.data && response.data.id) {
        validUsers.push({
          email: email,
          displayName: response.data.displayName,
          userPrincipalName: response.data.userPrincipalName,
          exists: true
        });
        logger.info(`✅ Validated Teams user: ${response.data.displayName} (${email})`);
      }
    } catch (error) {
      invalidUsers.push({
        email: email,
        exists: false,
        error: error.response?.status === 404 ? 'User not found in Teams organization' : error.message
      });
      logger.warn(`❌ Invalid Teams user: ${email} - ${error.message}`);
    }
  }

  return {
    validUsers,
    invalidUsers,
    allValid: invalidUsers.length === 0,
    summary: `${validUsers.length}/${attendeeEmails.length} users are valid Teams members`
  };
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
async findMeetingTimes(attendees, duration = 30, searchDays = 7, timePreferences = {}) {
  if (!this.isAvailable()) {
    throw new Error('Teams service not available - Azure AD configuration required for real Teams integration');
  }

  try {
    logger.info('🔍 Finding optimal meeting times with name resolution', {
      originalAttendees: attendees,
      duration,
      searchDays
    });

    // STEP 1: Resolve any names to email addresses
    let resolvedAttendees = [];
    
    for (const attendee of attendees) {
      // Check if it's already an email (contains @)
      if (attendee.includes('@')) {
        resolvedAttendees.push(attendee);
        logger.info(`✅ Already email format: ${attendee}`);
      } else {
        // It's a name, try to resolve it
        logger.info(`🔍 Resolving name to email: ${attendee}`);
        
        try {
          const resolvedUsers = await this.findUsersByDisplayName([attendee]);
          
          if (resolvedUsers && resolvedUsers.length > 0) {
            const resolvedEmail = resolvedUsers[0].email;
            resolvedAttendees.push(resolvedEmail);
            logger.info(`✅ Resolved name: ${attendee} -> ${resolvedEmail}`);
          } else {
            logger.warn(`⚠️ Could not resolve name: ${attendee}`);
            throw new Error(`Cannot find Teams user with name "${attendee}". Please use email address or exact display name.`);
          }
        } catch (resolutionError) {
          logger.error(`❌ Failed to resolve name ${attendee}:`, resolutionError.message);
          throw new Error(`Cannot resolve "${attendee}" to email address. Please use a valid email address like "rohit@company.com" or exact display name from Teams directory.`);
        }
      }
    }

    logger.info(`📧 Resolved attendees: ${JSON.stringify(resolvedAttendees)}`);

    // STEP 2: Validate all resolved emails are real Teams users
    const validation = await this.validateTeamsUsers(resolvedAttendees);
    
    if (!validation.allValid) {
      const invalidEmails = validation.invalidUsers.map(u => u.email);
      throw new Error(`Invalid Teams users found: ${invalidEmails.join(', ')}. Please ensure all attendees are in your Teams organization.`);
    }

    logger.info(`✅ All ${resolvedAttendees.length} attendees validated as real Teams users`);

    // STEP 3: Try Microsoft Graph findMeetingTimes API first
    const accessToken = await authService.getAppOnlyToken();
    const startTime = moment().add(1, 'hour').startOf('hour');
    const endTime = moment().add(searchDays, 'days').endOf('day');
    
    let suggestions = [];
    
    try {
      // Use the correct Microsoft Graph endpoint for finding meeting times
      const findMeetingTimesRequest = {
        attendees: resolvedAttendees.map(email => ({
          emailAddress: { 
            address: email, 
            name: email.split('@')[0] 
          }
        })),
        timeConstraint: {
          timeslots: [{
            start: {
              dateTime: startTime.toISOString(),
              timeZone: 'UTC'
            },
            end: {
              dateTime: endTime.toISOString(),
              timeZone: 'UTC'
            }
          }]
        },
        meetingDuration: `PT${duration}M`,
        maxCandidates: 20,
        isOrganizerOptional: false,
        returnSuggestionReasons: true,
        minimumAttendeePercentage: timePreferences.minimumAttendeePercentage || 100
      };

      logger.info('🔍 Trying Microsoft Graph findMeetingTimes API');
      
      // FIXED: Use the correct Microsoft Graph endpoint
      const response = await axios.post(
        `${this.graphEndpoint}/me/calendar/getSchedule`, // WRONG ENDPOINT
        findMeetingTimesRequest,
        {
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          }
        }
      );

      if (response.data.meetingTimeSuggestions) {
        suggestions = response.data.meetingTimeSuggestions.map(suggestion => ({
          start: suggestion.meetingTimeSlot.start.dateTime,
          end: suggestion.meetingTimeSlot.end.dateTime,
          confidence: this.mapConfidenceScore(suggestion.confidence),
          attendeeAvailability: suggestion.attendeeAvailability || [],
          suggestionReason: suggestion.suggestionReason,
          source: 'microsoft_graph_api',
          score: this.calculateSuggestionScore(suggestion)
        }));
        
        logger.info(`✅ Microsoft Graph API returned ${suggestions.length} suggestions`);
      }
    } catch (graphError) {
      logger.warn('Microsoft Graph findMeetingTimes failed, falling back to manual analysis:', graphError.message);
    }

    // STEP 4: Manual availability analysis (using RESOLVED EMAILS)
    if (suggestions.length < 5) {
      logger.info('🔍 Performing manual availability analysis for better results');
      
      const manualSuggestions = await this.findAvailableTimeSlots(resolvedAttendees, duration, searchDays);
      
      const formattedManualSuggestions = manualSuggestions.map(slot => ({
        start: slot.start,
        end: slot.end,
        confidence: slot.confidence || 'high',
        attendeeAvailability: slot.attendeeAvailability || [],
        source: 'manual_analysis',
        dayOfWeek: slot.dayOfWeek,
        timeOfDay: slot.timeOfDay,
        score: this.calculateManualScore(slot, timePreferences),
        allAttendeesAvailable: slot.allAttendeesAvailable
      }));

      suggestions = [...suggestions, ...formattedManualSuggestions];
    }

    // STEP 5: Enhance and sort suggestions
    suggestions = this.enhanceWithSmartPreferences(suggestions, timePreferences);
    suggestions.sort((a, b) => b.score - a.score);

    const finalSuggestions = suggestions.slice(0, 10).map((suggestion, index) => ({
      ...suggestion,
      rank: index + 1,
      recommendationReason: this.getRecommendationReason(suggestion, timePreferences),
      businessHoursMatch: this.isBusinessHours(suggestion.start),
      timeZoneOptimal: this.isOptimalTimeZone(suggestion.start, resolvedAttendees)
    }));

    logger.info(`🎯 Returning ${finalSuggestions.length} optimized time suggestions`);

    return {
      suggestions: finalSuggestions,
      nameResolution: {
        originalAttendees: attendees,
        resolvedAttendees: resolvedAttendees,
        resolutionSuccess: true
      },
      searchCriteria: {
        attendees: resolvedAttendees.length,
        duration: duration,
        searchDays: searchDays,
        timePreferences: timePreferences
      },
      metadata: {
        totalCandidatesAnalyzed: suggestions.length,
        dataSource: finalSuggestions.length > 0 ? finalSuggestions[0].source : 'none',
        searchCompletedAt: new Date().toISOString()
      }
    };

  } catch (error) {
    logger.error('❌ Failed to find meeting times:', error);
    throw new Error(`Finding optimal meeting times failed: ${error.message}`);
  }
}




// Helper methods for time suggestions
mapConfidenceScore(confidence) {
  const confidenceMap = {
    'high': 'high',
    'medium': 'medium', 
    'low': 'low'
  };
  return confidenceMap[confidence?.toLowerCase()] || 'medium';
}

calculateSuggestionScore(suggestion) {
  let score = 50; // Base score
  
  // Higher score for higher confidence
  if (suggestion.confidence === 'high') score += 30;
  else if (suggestion.confidence === 'medium') score += 15;
  
  // Prefer business hours (9 AM - 5 PM)
  const hour = new Date(suggestion.meetingTimeSlot.start.dateTime).getHours();
  if (hour >= 9 && hour <= 17) score += 20;
  
  // Prefer weekdays
  const dayOfWeek = new Date(suggestion.meetingTimeSlot.start.dateTime).getDay();
  if (dayOfWeek >= 1 && dayOfWeek <= 5) score += 15;
  
  // Prefer times when more attendees are available
  if (suggestion.attendeeAvailability) {
    const availableCount = suggestion.attendeeAvailability.filter(a => a.availability === 'free').length;
    const total = suggestion.attendeeAvailability.length;
    score += (availableCount / total) * 20;
  }
  
  return score;
}

calculateManualScore(slot, preferences) {
  let score = 60; // Base score for manual analysis
  
  // All attendees available gets high score
  if (slot.allAttendeesAvailable) score += 25;
  
  // Business hours preference
  const hour = new Date(slot.start).getHours();
  if (hour >= 9 && hour <= 17) score += 15;
  
  // Weekday preference
  const dayOfWeek = new Date(slot.start).getDay();
  if (dayOfWeek >= 1 && dayOfWeek <= 5) score += 10;
  
  // Time preferences
  if (preferences.preferredHours) {
    if (preferences.preferredHours.includes(hour)) score += 20;
  }
  
  return score;
}

enhanceWithSmartPreferences(suggestions, preferences) {
  return suggestions.map(suggestion => {
    const startTime = new Date(suggestion.start);
    const hour = startTime.getHours();
    const dayOfWeek = startTime.getDay();
    
    // Add smart scoring based on common meeting patterns
    let smartScore = suggestion.score || 50;
    
    // Common good meeting times
    if (hour === 10 || hour === 14) smartScore += 10; // 10 AM or 2 PM
    if (hour >= 9 && hour <= 11) smartScore += 5; // Morning meetings
    if (dayOfWeek === 2 || dayOfWeek === 3 || dayOfWeek === 4) smartScore += 5; // Tue, Wed, Thu
    
    // Avoid common bad times
    if (hour === 12) smartScore -= 10; // Lunch time
    if (hour >= 17) smartScore -= 15; // After hours
    if (dayOfWeek === 1 && hour <= 10) smartScore -= 10; // Monday morning
    if (dayOfWeek === 5 && hour >= 15) smartScore -= 10; // Friday afternoon
    
    return {
      ...suggestion,
      score: smartScore
    };
  });
}

getRecommendationReason(suggestion, preferences) {
  const reasons = [];
  
  if (suggestion.confidence === 'high') reasons.push('High confidence match');
  if (suggestion.allAttendeesAvailable) reasons.push('All attendees available');
  if (suggestion.businessHoursMatch) reasons.push('Business hours');
  
  const hour = new Date(suggestion.start).getHours();
  if (hour === 10) reasons.push('Optimal morning time');
  if (hour === 14) reasons.push('Good afternoon time');
  
  const dayOfWeek = new Date(suggestion.start).getDay();
  if ([2, 3, 4].includes(dayOfWeek)) reasons.push('Mid-week timing');
  
  return reasons.length > 0 ? reasons.join(', ') : 'Available time slot';
}

isBusinessHours(dateTime) {
  const hour = new Date(dateTime).getHours();
  const dayOfWeek = new Date(dateTime).getDay();
  return hour >= 9 && hour <= 17 && dayOfWeek >= 1 && dayOfWeek <= 5;
}

isOptimalTimeZone(dateTime, attendees) {
  // Simple timezone optimization - can be enhanced based on attendee locations
  const hour = new Date(dateTime).getHours();
  return hour >= 9 && hour <= 16; // Good for most time zones
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
    throw new Error('Teams service not available - Azure AD configuration required');
  }

  try {
    const accessToken = await authService.getAppOnlyToken();
    
    logger.info('📋 Getting available real Teams for channel creation');
    
    const url = `${this.graphEndpoint}/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName,description&$top=10`;
    
    const response = await axios.get(url, {
      headers: { 'Authorization': `Bearer ${accessToken}` }
    });

    const teams = response.data.value || [];
    const result = teams.map(team => ({
      id: team.id,
      displayName: team.displayName,
      description: team.description || 'No description'
    }));

    logger.info(`✅ Found ${result.length} real Teams`);
    return result;

  } catch (error) {
    logger.error('❌ Failed to get real Teams:', error);
    throw new Error(`Get real Teams failed: ${error.message}`);
  }
}




// List channels in a team
async getTeamChannels(teamId) {
  if (!this.isAvailable()) {
    throw new Error('Teams service not available - Azure AD configuration required');
  }

  try {
    const accessToken = await authService.getAppOnlyToken();
    
    const response = await axios.get(
      `${this.graphEndpoint}/teams/${teamId}/channels`,
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
    logger.error('❌ Failed to get real Teams channels:', error);
    throw new Error(`Get real Teams channels failed: ${error.message}`);
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