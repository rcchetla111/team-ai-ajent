// Create this file as quick-test.js
const axios = require('axios');

async function quickTest() {
  try {
    console.log('🔍 Quick Service Test');
    console.log('===================');
    
    // Test service status
    const response = await axios.get('http://localhost:5000/api/meetings/status');
    const status = response.data;
    
    console.log('🤖 AI Service:', status.services.ai.available ? '✅ Working' : '❌ Not Available');
    console.log('🟢 Teams Service:', status.services.teams.available ? '✅ Working' : '❌ Not Available');
    console.log('📊 Overall Status:', status.overall.status);
    console.log('💬 Message:', status.overall.message);
    
    // Test creating a simple meeting
    const meetingResponse = await axios.post('http://localhost:5000/api/meetings/create', {
      subject: 'Quick Test Meeting',
      description: 'Testing AI and Teams integration',
      startTime: new Date(Date.now() + 5 * 60 * 1000).toISOString(),
      endTime: new Date(Date.now() + 35 * 60 * 1000).toISOString(),
      attendees: ['test@company.com'],
      useAI: true,
      autoJoinAgent: false, // Skip auto-join for quick test
      enableChatCapture: false // Skip chat capture for quick test
    });
    
    const meeting = meetingResponse.data.meeting;
    console.log('\n📅 Meeting Creation Test:');
    console.log('✅ Meeting Created:', meeting.subject);
    console.log('🤖 AI Enhanced:', meeting.aiEnhanced);
    console.log('🟢 Real Teams:', meeting.isRealTeamsMeeting);
    
    if (meeting.aiEnhanced && meeting.isRealTeamsMeeting) {
      console.log('\n🎉 SUCCESS: Both AI and Teams are working perfectly!');
    } else {
      console.log('\n⚠️  Some services may need configuration');
    }
    
  } catch (error) {
    console.error('❌ Test failed:', error.response?.data || error.message);
  }
}

quickTest();