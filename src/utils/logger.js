// Update your logger.js to handle circular references

const winston = require('winston');

// Safe JSON stringify that handles circular references
const safeStringify = (obj) => {
  try {
    return JSON.stringify(obj, (key, value) => {
      // Skip circular references and functions
      if (typeof value === 'object' && value !== null) {
        if (value.constructor?.name === 'ClientRequest' || 
            value.constructor?.name === 'IncomingMessage' ||
            value.constructor?.name === 'Socket') {
          return '[Circular Reference]';
        }
      }
      if (typeof value === 'function') {
        return '[Function]';
      }
      return value;
    }, 2);
  } catch (error) {
    return '[Unable to stringify object]';
  }
};

const logger = winston.createLogger({
  level: process.env.LOG_LEVEL || 'info',
  format: winston.format.combine(
    winston.format.timestamp({
      format: 'HH:mm:ss'
    }),
    winston.format.errors({ stack: true }),
    winston.format.printf(({ level, message, timestamp, ...meta }) => {
      let output = `${timestamp} [${level}]: ${message}`;
      
      // Safely handle additional metadata
      if (Object.keys(meta).length > 0) {
        output += '\n' + safeStringify(meta);
      }
      
      return output;
    })
  ),
  transports: [
    new winston.transports.Console({
      format: winston.format.combine(
        winston.format.colorize(),
        winston.format.simple()
      )
    })
  ]
});

// Add file transport for production
if (process.env.NODE_ENV === 'production') {
  logger.add(new winston.transports.File({ 
    filename: 'logs/error.log', 
    level: 'error' 
  }));
  logger.add(new winston.transports.File({ 
    filename: 'logs/combined.log' 
  }));
}

module.exports = logger;