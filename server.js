const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const compression = require('compression');
const bodyParser = require('body-parser');
const path = require('path');

// Load environment variables
require('dotenv').config();

// Import custom modules
const logger = require('./src/utils/logger');
const cosmosClient = require('./src/config/cosmosdb');
const userRoutes = require('./src/routes/users'); // Add this line
const meetingRoutes = require('./src/routes/meetings');
const teamsService = require('./src/services/teamsService');


// Initialize Express app
const app = express();

// Middleware
app.use(helmet({
  contentSecurityPolicy: false // Disable for local development
}));

app.use(compression());
// This is the NEW, corrected code
// Add this before your routes
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization');
  if (req.method === 'OPTIONS') {
    res.sendStatus(200);
  } else {
    next();
  }
});

// Body parsing middleware
app.use(bodyParser.json({ limit: '10mb' }));
app.use(bodyParser.urlencoded({ extended: true, limit: '10mb' }));

// Static files
app.use(express.static(path.join(__dirname, 'public')));

// Health check endpoint
// Add this route to your main router or app
app.get('/health', (req, res) => {
  res.json({ status: 'OK', timestamp: new Date().toISOString() });
});

// Test endpoint
app.get('/api/test', (req, res) => {
  res.json({
    message: 'Agent 365 API is working!',
    timestamp: new Date().toISOString()
  });
});

// API Routes
app.use('/api/meetings', meetingRoutes);
app.use('/api/users', userRoutes); // Add this line

// --- NEW CODE TO SERVE THE CHAT UI ---

// This tells the server to serve the chat.html file
app.get('/chat', (req, res) => {
  res.sendFile(path.join(__dirname, 'chat.html'));
});

// This redirects the main URL to your new chat page
app.get('/', (req, res) => {
    res.redirect('/chat');
});

// --- END OF NEW CODE ---

// Serve a simple HTML page for testing
// app.get('/', (req, res) => {
//   res.send(`
//     <!DOCTYPE html>
//     <html>
//     <head>
//         <title>Agent 365 - Digital Worker</title>
//         <style>
//             body { font-family: Arial, sans-serif; margin: 40px; background: #f5f5f5; }
//             .container { max-width: 800px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
//             h1 { color: #0078d4; }
//             .status { padding: 10px; margin: 10px 0; border-radius: 5px; }
//             .success { background: #d4edda; border: 1px solid #c3e6cb; color: #155724; }
//             .info { background: #d1ecf1; border: 1px solid #bee5eb; color: #0c5460; }
//         </style>
//     </head>
//     <body>
//         <div class="container">
//             <h1>🟢 Agent 365 - Teams Integration</h1>
//             <div class="status success">✅ Server is running successfully!</div>
//             <div class="status info">📊 Database: Cosmos DB Emulator</div>
//             <div class="status info">🔧 Environment: ${process.env.NODE_ENV || 'development'}</div>
            
//             <h3>Service Status:</h3>
//             <div class="status info">
//                 🟢 Teams Integration: ${process.env.AZURE_CLIENT_ID ? 'Configured' : 'Not Configured'}
//             </div>
            
//             <h3>Test Endpoints:</h3>
//             <ul>
//                 <li><a href="/health">Health Check</a></li>
//                 <li><a href="/api/test">API Test</a></li>
//                 <li><a href="/api/meetings/teams/status">Teams Status</a></li>
//                 <li><a href="/api/meetings">View Meetings</a></li>
//             </ul>
            
//             <h3>Next Steps:</h3>
//             <ol>
//                 <li>✅ Server started</li>
//                 <li>✅ Database connection ready</li>
//                 <li>🔄 Teams integration ready...</li>
//                 <li>🎯 Focus: REAL Teams meetings only</li>
//             </ol>
//         </div>
//     </body>
//     </html>
//   `);
// });

// Error handling middleware
app.use((err, req, res, next) => {
  logger.error('Unhandled error:', {
    error: err.message,
    stack: err.stack,
    url: req.url,
    method: req.method
  });
  
  res.status(err.status || 500).json({
    error: process.env.NODE_ENV === 'production' 
      ? 'Internal server error' 
      : err.message
  });
});

// Initialize Cosmos DB
async function initializeDatabase() {
  try {
    logger.info('🔄 Initializing Cosmos DB...');
    await cosmosClient.createDatabaseIfNotExists();
    await cosmosClient.createContainersIfNotExists();
    logger.info('✅ Cosmos DB initialized successfully');
  } catch (error) {
    logger.error('❌ Failed to initialize Cosmos DB:', error);
    logger.error('💡 Make sure Cosmos DB Emulator is running on https://localhost:8081');
    process.exit(1);
  }
}

// Start server
const PORT = process.env.PORT || 5000;

async function startServer() {
  try {
    await initializeDatabase();
    
    app.listen(PORT, () => {
      logger.info(`🚀 Agent 365 server running on port ${PORT}`);
      logger.info(`🌐 Environment: ${process.env.NODE_ENV || 'development'}`);
      logger.info(`📊 Dashboard: http://localhost:${PORT}`);
      logger.info(`🔍 Health Check: http://localhost:${PORT}/health`);
      logger.info(`🔧 Cosmos DB: ${process.env.COSMOS_ENDPOINT}`);
    });
  } catch (error) {
    logger.error('❌ Failed to start server:', error);
    process.exit(1);
  }
}

// Graceful shutdown
process.on('SIGTERM', () => {
  logger.info('SIGTERM received, shutting down gracefully');
  process.exit(0);
});

process.on('SIGINT', () => {
  logger.info('SIGINT received, shutting down gracefully');
  process.exit(0);
});

startServer();