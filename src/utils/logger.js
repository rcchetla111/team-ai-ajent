const winston = require('winston');
const path = require('path');
const fs = require('fs');

// Create logs directory if it doesn't exist
const logDir = 'logs';
if (!fs.existsSync(logDir)) {
  fs.mkdirSync(logDir);
}

// Console format for development
const consoleFormat = winston.format.combine(
  winston.format.colorize(),
  winston.format.timestamp({
    format: 'HH:mm:ss'
  }),
  winston.format.printf(({ timestamp, level, message, ...meta }) => {
    let log = `${timestamp} [${level}]: ${message}`;
    
    // Add metadata if present
    if (Object.keys(meta).length > 0) {
      log += `\n${JSON.stringify(meta, null, 2)}`;
    }
    
    return log;
  })
);

// Create logger instance
const logger = winston.createLogger({
  level: process.env.LOG_LEVEL || 'debug',
  format: winston.format.combine(
    winston.format.timestamp(),
    winston.format.errors({ stack: true }),
    winston.format.json()
  ),
  transports: [
    // Console transport for development
    new winston.transports.Console({
      format: consoleFormat,
      level: 'debug'
    }),
    
    // File transport for all logs
    new winston.transports.File({
      filename: path.join(logDir, 'app.log'),
      maxsize: 5242880, // 5MB
      maxFiles: 2,
    })
  ]
});

// Helper functions for specific log types
logger.logMeetingEvent = function(event, meetingId, userId, metadata = {}) {
  this.info('Meeting Event', {
    event,
    meetingId,
    userId,
    ...metadata,
    type: 'meeting_event'
  });
};

logger.logError = function(error, context = {}) {
  this.error('Application Error', {
    message: error.message,
    stack: error.stack,
    ...context,
    type: 'application_error'
  });
};

module.exports = logger;