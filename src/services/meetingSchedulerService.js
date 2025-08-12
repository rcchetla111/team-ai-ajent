const cron = require('node-cron');
const moment = require('moment');
const cosmosClient = require('../config/cosmosdb');
const meetingAttendanceService = require('./meetingAttendanceService');
const chatCaptureService = require('./chatCaptureService');
const logger = require('../utils/logger');

class MeetingSchedulerService {
  constructor() {
    this.scheduledJobs = new Map(); // Track scheduled join jobs
    this.activeMonitors = new Map(); // Track active meeting monitors
    this.isInitialized = false;
  }

  // Initialize the scheduler service
  initialize() {
    if (this.isInitialized) {
      return;
    }

    logger.info('🤖 Initializing Automatic Meeting Join Scheduler');

    // Run every minute to check for meetings to join
    cron.schedule('* * * * *', async () => {
      await this.checkForMeetingsToJoin();
    });

    // Run every 5 minutes to clean up completed meetings
    cron.schedule('*/5 * * * *', async () => {
      await this.cleanupCompletedMeetings();
    });

    // Load existing scheduled meetings on startup
    this.loadScheduledMeetings();

    this.isInitialized = true;
    logger.info('✅ Meeting Scheduler Service initialized');
  }

  // Schedule automatic join for a meeting
  async scheduleMeetingJoin(meeting) {
    try {
      if (!meeting.agentConfig?.autoJoin) {
        logger.debug('Auto-join disabled for meeting:', meeting.meetingId);
        return;
      }

      const meetingStart = moment(meeting.startTime);
      const now = moment();
      const minutesUntilStart = meetingStart.diff(now, 'minutes');

      logger.info('📅 Scheduling automatic join for meeting', {
        meetingId: meeting.meetingId,
        subject: meeting.subject,
        startTime: meeting.startTime,
        minutesUntilStart
      });

      // Store in database for persistence
      const scheduleRecord = {
        id: `schedule_${meeting.meetingId}`,
        meetingId: meeting.meetingId,
        meetingDbId: meeting.id,
        userId: meeting.userId,
        scheduledJoinTime: meetingStart.toISOString(),
        status: 'scheduled',
        createdAt: new Date().toISOString(),
        autoJoinEnabled: true
      };

      await cosmosClient.createItem('schedules', scheduleRecord);

      // If meeting starts within 2 minutes, join immediately
      if (minutesUntilStart <= 2 && minutesUntilStart >= -1) {
        logger.info('🚀 Meeting starting very soon, joining immediately');
        await this.executeAutomaticJoin(meeting);
      }

      logger.info('✅ Meeting join scheduled successfully');

    } catch (error) {
      logger.error('❌ Failed to schedule meeting join:', error);
      throw error;
    }
  }

  // Check for meetings that need to be joined
  async checkForMeetingsToJoin() {
    try {
      const now = moment();
      const checkWindowStart = now.subtract(1, 'minute').toISOString();
      const checkWindowEnd = now.add(3, 'minutes').toISOString();

      // Get scheduled meetings within the join window
      const scheduledMeetings = await cosmosClient.queryItems('schedules',
        'SELECT * FROM c WHERE c.status = "scheduled" AND c.scheduledJoinTime >= @start AND c.scheduledJoinTime <= @end',
        [
          { name: '@start', value: checkWindowStart },
          { name: '@end', value: checkWindowEnd }
        ]
      );

      for (const schedule of scheduledMeetings) {
        await this.processScheduledMeeting(schedule);
      }

    } catch (error) {
      logger.error('❌ Error checking for meetings to join:', error);
    }
  }

  // Process a scheduled meeting for auto-join
  async processScheduledMeeting(schedule) {
    try {
      // Get the full meeting details
      const meeting = await cosmosClient.getItem('meetings', schedule.meetingDbId, schedule.userId);
      
      if (!meeting) {
        logger.warn('⚠️ Meeting not found for schedule:', schedule.meetingId);
        await this.markScheduleCompleted(schedule.id, 'meeting_not_found');
        return;
      }

      if (meeting.agentAttended) {
        logger.info('✅ Agent already joined meeting:', meeting.meetingId);
        await this.markScheduleCompleted(schedule.id, 'already_joined');
        return;
      }

      const meetingStart = moment(meeting.startTime);
      const now = moment();
      const minutesSinceStart = now.diff(meetingStart, 'minutes');

      // Join if meeting has started (within 5 minutes of start time)
      if (minutesSinceStart >= 0 && minutesSinceStart <= 5) {
        logger.info('🤖 Meeting started, executing automatic join');
        await this.executeAutomaticJoin(meeting);
        await this.markScheduleCompleted(schedule.id, 'joined');
      }
      // If meeting started more than 5 minutes ago, mark as missed
      else if (minutesSinceStart > 5) {
        logger.warn('⚠️ Meeting start window missed:', meeting.meetingId);
        await this.markScheduleCompleted(schedule.id, 'missed');
      }

    } catch (error) {
      logger.error('❌ Error processing scheduled meeting:', error);
      await this.markScheduleCompleted(schedule.id, 'error');
    }
  }

  // Execute automatic join for a meeting
  async executeAutomaticJoin(meeting) {
    try {
      logger.info('🚀 Executing automatic agent join for REAL Teams meeting', {
        meetingId: meeting.meetingId,
        subject: meeting.subject
      });

      // Join the real Teams meeting
      const joinResult = await meetingAttendanceService.joinMeeting(
        meeting.meetingId,
        meeting.userId
      );

      // Start real chat capture if enabled
      if (meeting.agentConfig?.enableChatCapture !== false) {
        await chatCaptureService.initiateRealChatCapture(meeting);
      }

      // Update meeting status
      await cosmosClient.updateItem('meetings', meeting.id, meeting.userId, {
        agentAttended: true,
        agentJoinedAt: new Date().toISOString(),
        status: 'in_progress',
        autoJoinExecuted: true
      });

      logger.info('✅ Automatic agent join completed successfully', {
        meetingId: meeting.meetingId,
        joinResult: joinResult?.success
      });

      // Start monitoring for meeting end
      this.startMeetingEndMonitor(meeting);

    } catch (error) {
      logger.error('❌ Automatic agent join failed:', error);
      
      // Update meeting with error info
      await cosmosClient.updateItem('meetings', meeting.id, meeting.userId, {
        autoJoinError: error.message,
        autoJoinAttemptedAt: new Date().toISOString()
      });

      throw error;
    }
  }

  // Start monitoring for meeting end to auto-leave
  startMeetingEndMonitor(meeting) {
    const meetingEnd = moment(meeting.endTime);
    const now = moment();
    const minutesUntilEnd = meetingEnd.diff(now, 'minutes');

    if (minutesUntilEnd > 0) {
      logger.info('📊 Starting meeting end monitor', {
        meetingId: meeting.meetingId,
        minutesUntilEnd
      });

      const endMonitor = setTimeout(async () => {
        await this.executeAutomaticLeave(meeting);
      }, minutesUntilEnd * 60 * 1000);

      this.activeMonitors.set(meeting.meetingId, endMonitor);
    }
  }

  // Execute automatic leave when meeting ends
  async executeAutomaticLeave(meeting) {
    try {
      logger.info('🚪 Executing automatic agent leave', {
        meetingId: meeting.meetingId
      });

      // Leave the meeting
      await meetingAttendanceService.leaveMeeting(meeting.meetingId, meeting.userId);

      // Stop chat capture
      await chatCaptureService.stopRealChatCapture(meeting.meetingId);

      // Update meeting status
      await cosmosClient.updateItem('meetings', meeting.id, meeting.userId, {
        agentLeftAt: new Date().toISOString(),
        status: 'completed',
        autoLeaveExecuted: true
      });

      // Clean up monitor
      this.activeMonitors.delete(meeting.meetingId);

      logger.info('✅ Automatic agent leave completed');

    } catch (error) {
      logger.error('❌ Automatic agent leave failed:', error);
    }
  }

  // Mark a schedule as completed
  async markScheduleCompleted(scheduleId, status) {
    try {
      const schedules = await cosmosClient.queryItems('schedules',
        'SELECT * FROM c WHERE c.id = @scheduleId',
        [{ name: '@scheduleId', value: scheduleId }]
      );

      if (schedules.length > 0) {
        const schedule = schedules[0];
        await cosmosClient.updateItem('schedules', schedule.id, schedule.userId || 'system', {
          status: 'completed',
          completionReason: status,
          completedAt: new Date().toISOString()
        });
      }
    } catch (error) {
      logger.error('❌ Failed to mark schedule completed:', error);
    }
  }

  // Load existing scheduled meetings on startup
  async loadScheduledMeetings() {
    try {
      const now = moment();
      const futureThreshold = now.add(24, 'hours').toISOString();

      const scheduledMeetings = await cosmosClient.queryItems('schedules',
        'SELECT * FROM c WHERE c.status = "scheduled" AND c.scheduledJoinTime <= @threshold',
        [{ name: '@threshold', value: futureThreshold }]
      );

      logger.info(`📋 Loaded ${scheduledMeetings.length} scheduled meetings on startup`);

    } catch (error) {
      logger.error('❌ Failed to load scheduled meetings:', error);
    }
  }

  // Clean up completed meetings
  async cleanupCompletedMeetings() {
    try {
      const twentyFourHoursAgo = moment().subtract(24, 'hours').toISOString();

      const oldSchedules = await cosmosClient.queryItems('schedules',
        'SELECT * FROM c WHERE c.status = "completed" AND c.completedAt < @threshold',
        [{ name: '@threshold', value: twentyFourHoursAgo }]
      );

      for (const schedule of oldSchedules) {
        await cosmosClient.deleteItem('schedules', schedule.id, schedule.userId || 'system');
      }

      if (oldSchedules.length > 0) {
        logger.info(`🧹 Cleaned up ${oldSchedules.length} old schedule records`);
      }

    } catch (error) {
      logger.error('❌ Failed to cleanup completed meetings:', error);
    }
  }

  // Cancel scheduled join for a meeting
  async cancelScheduledJoin(meetingId) {
    try {
      const schedules = await cosmosClient.queryItems('schedules',
        'SELECT * FROM c WHERE c.meetingId = @meetingId AND c.status = "scheduled"',
        [{ name: '@meetingId', value: meetingId }]
      );

      for (const schedule of schedules) {
        await cosmosClient.updateItem('schedules', schedule.id, schedule.userId || 'system', {
          status: 'cancelled',
          cancelledAt: new Date().toISOString()
        });
      }

      // Cancel active monitor if exists
      if (this.activeMonitors.has(meetingId)) {
        clearTimeout(this.activeMonitors.get(meetingId));
        this.activeMonitors.delete(meetingId);
      }

      logger.info('✅ Scheduled join cancelled for meeting:', meetingId);

    } catch (error) {
      logger.error('❌ Failed to cancel scheduled join:', error);
    }
  }

  // Get scheduler status
  getSchedulerStatus() {
    return {
      initialized: this.isInitialized,
      activeSchedules: this.scheduledJobs.size,
      activeMonitors: this.activeMonitors.size,
      features: {
        automaticJoin: true,
        automaticLeave: true,
        realTimeMonitoring: true,
        persistentScheduling: true
      }
    };
  }

  // Force check for meetings (manual trigger)
  async forceCheckMeetings() {
    logger.info('🔄 Manually triggering meeting check');
    await this.checkForMeetingsToJoin();
  }
}

// Create singleton instance
const meetingSchedulerService = new MeetingSchedulerService();

module.exports = meetingSchedulerService;