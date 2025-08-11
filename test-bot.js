// Create this as test-bot.js in your root directory
const axios = require('axios');

async function testBotIntegration() {
  console.log('🤖 Testing Agent 365 Bot Integration');
  console.log('====================================\n');

  try {
    // Test 1: Check if main server is running
    console.log('1. Testing main Agent 365 server...');
    try {
      const mainResponse = await axios.get('http://localhost:5000/health');
      console.log('   ✅ Main server is running');
      console.log(`   📊 Status: ${mainResponse.data.status}\n`);
    } catch (error) {
      console.log('   ❌ Main server not running - start with: npm start\n');
      return;
    }

    // Test 2: Check if bot server is running
    console.log('2. Testing bot server...');
    try {
      const botResponse = await axios.get('http://localhost:3978/bot/health');
      console.log('   ✅ Bot server is running');
      console.log(`   🤖 Bot: ${botResponse.data.bot}\n`);
    } catch (error) {
      console.log('   ❌ Bot server not running - start with: npm run bot\n');
      return;
    }

    // Test 3: Check bot info
    console.log('3. Testing bot information...');
    try {
      const infoResponse = await axios.get('http://localhost:3978/bot/info');
      const info = infoResponse.data;
      console.log('   ✅ Bot info retrieved');
      console.log(`   📱 Name: ${info.name}`);
      console.log(`   📝 Description: ${info.description}`);
      console.log(`   🔧 Capabilities: ${info.capabilities.length} features\n`);
    } catch (error) {
      console.log('   ⚠️  Could not get bot info\n');
    }

    // Test 4: Check API integration
    console.log('4. Testing API integration...');
    try {
      const statusResponse = await axios.get('http://localhost:5000/api/meetings/status');
      const status = statusResponse.data;
      console.log('   ✅ API integration working');
      console.log(`   🤖 AI Available: ${status.services.ai.available ? 'Yes' : 'No'}`);
      console.log(`   🟢 Teams Available: ${status.services.teams.available ? 'Yes' : 'No'}\n`);
    } catch (error) {
      console.log('   ❌ API integration issue\n');
    }

    // Test 5: Environment check
    console.log('5. Checking bot configuration...');
    const hasAppId = !!process.env.MICROSOFT_APP_ID;
    const hasAppPassword = !!process.env.MICROSOFT_APP_PASSWORD;
    
    console.log(`   🆔 App ID configured: ${hasAppId ? 'Yes' : 'No'}`);
    console.log(`   🔑 App Password configured: ${hasAppPassword ? 'Yes' : 'No'}`);
    
    if (!hasAppId || !hasAppPassword) {
      console.log('   ⚠️  Bot credentials missing - add to .env file\n');
    } else {
      console.log('   ✅ Bot credentials configured\n');
    }

    // Summary
    console.log('📋 INTEGRATION TEST SUMMARY');
    console.log('===========================');
    
    const mainServerOk = true; // We got here, so it's running
    const botServerOk = true; // We got here, so it's running
    const credsConfigured = hasAppId && hasAppPassword;
    
    if (mainServerOk && botServerOk && credsConfigured) {
      console.log('🎉 ALL SYSTEMS GO! Your bot is ready for Teams deployment!');
      console.log('\n📱 Next Steps:');
      console.log('   1. Register bot in Azure Portal');
      console.log('   2. Configure messaging endpoint');
      console.log('   3. Create Teams app manifest');
      console.log('   4. Upload to Teams and test!');
      console.log('\n🚀 Once deployed, you can:');
      console.log('   • Chat with Agent 365 in Teams');
      console.log('   • Ask it to create AI-enhanced meetings');
      console.log('   • Join meetings and get real-time insights');
      console.log('   • Receive automatic summaries and action items');
      
    } else {
      console.log('⚠️  Some components need attention:');
      if (!mainServerOk) console.log('   ❌ Start main server: npm start');
      if (!botServerOk) console.log('   ❌ Start bot server: npm run bot');
      if (!credsConfigured) console.log('   ❌ Configure bot credentials in .env');
    }

    console.log('\n🔍 Current Status:');
    console.log(`   Main Server: http://localhost:5000`);
    console.log(`   Bot Server: http://localhost:3978`);
    console.log(`   Bot Endpoint: http://localhost:3978/api/messages`);
    
  } catch (error) {
    console.error('❌ Integration test failed:', error.message);
    console.log('\n🔧 Troubleshooting:');
    console.log('   1. Make sure both servers are running');
    console.log('   2. Check your .env file configuration');
    console.log('   3. Verify network connectivity');
  }
}

// Load environment variables
require('dotenv').config();

// Run the test
testBotIntegration();