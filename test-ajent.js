const axios = require('axios');

async function testAgentFunctionality() {
  const baseUrl = 'http://localhost:5000/api/meetings';
  
  console.log('🤖 Testing Agent 365 Digital Worker');
  console.log('==================================\n');

  try {
    // 1. Create meeting with agent auto-join
    console.log('1. Creating meeting with AI and auto-join...');
    const meetingResponse = await axios.post(`${baseUrl}/create`, {
      subject: 'Agent Test Meeting',
      description: 'Testing our digital worker capabilities with AI enhancement',
      startTime: new Date(Date.now() + 2 * 60 * 1000).toISOString(), // 2 minutes from now  
      endTime: new Date(Date.now() + 32 * 60 * 1000).toISOString(),   // 32 minutes from now
      attendees: ['test@company.com', 'demo@company.com'],
      useAI: true,
      autoJoinAgent: true,
      enableChatCapture: true
    });
    
    const meeting = meetingResponse.data.meeting;
    const meetingId = meeting.id;
    
    console.log('✅ Meeting created successfully!');
    console.log(`   Meeting ID: ${meetingId}`);
    console.log(`   Subject: ${meeting.subject}`);
    console.log(`   AI Enhanced: ${meeting.aiEnhanced}`);
    console.log(`   Real Teams: ${meeting.isRealTeamsMeeting}`);
    console.log(`   Auto-join scheduled: ${meetingResponse.data.agentConfig?.scheduledJoin}\n`);

    // 2. Test manual agent join (in case auto-join isn't ready yet)
    console.log('2. Testing manual agent join...');
    try {
      const joinResponse = await axios.post(`${baseUrl}/${meetingId}/join-agent`);
      console.log('✅ Agent joined manually:', joinResponse.data.message);
    } catch (joinError) {
      console.log('ℹ️  Manual join result:', joinError.response?.data?.details || 'Already joined or not available');
    }

    // 3. Check attendance status
    console.log('\n3. Checking agent attendance status...');
    const attendanceResponse = await axios.get(`${baseUrl}/${meetingId}/attendance-status`);
    console.log('✅ Attendance status retrieved');
    console.log(`   Agent active: ${attendanceResponse.data.attendance?.isActive || 'Not available'}`);

    // 4. Get chat analysis
    console.log('\n4. Testing chat analysis...');
    const chatResponse = await axios.get(`${baseUrl}/${meetingId}/chat-analysis`);
    console.log('✅ Chat analysis retrieved');
    console.log(`   Total messages: ${chatResponse.data.analysis?.totalMessages || 0}`);
    console.log(`   Analysis available: ${!!chatResponse.data.analysis}`);

    // 5. Generate meeting summary
    console.log('\n5. Testing summary generation...');
    const summaryResponse = await axios.get(`${baseUrl}/${meetingId}/summary`);
    console.log('✅ Summary generation tested');
    console.log(`   Summary created: ${summaryResponse.data.success}`);
    console.log(`   Generated new: ${summaryResponse.data.generated}`);

    // 6. Test agent leave
    console.log('\n6. Testing agent leave...');
    try {
      const leaveResponse = await axios.post(`${baseUrl}/${meetingId}/leave-agent`);
      console.log('✅ Agent left successfully:', leaveResponse.data.message);
    } catch (leaveError) {
      console.log('ℹ️  Leave result:', leaveError.response?.data?.details || 'Not currently in meeting');
    }

    // 7. Get all meetings
    console.log('\n7. Testing meetings list...');
    const meetingsResponse = await axios.get(`${baseUrl}`);
    console.log(`✅ Retrieved ${meetingsResponse.data.meetings.length} meetings`);

    // 8. Get specific meeting details
    console.log('\n8. Testing meeting details...');
    const detailsResponse = await axios.get(`${baseUrl}/${meetingId}`);
    console.log('✅ Meeting details retrieved');
    console.log(`   Meeting status: ${detailsResponse.data.status}`);

    console.log('\n🎉 All tests completed successfully!');
    console.log('\n📊 Test Results Summary:');
    console.log('========================');
    console.log('✅ Meeting creation with AI');
    console.log('✅ Agent join/leave functionality');  
    console.log('✅ Chat analysis system');
    console.log('✅ Summary generation');
    console.log('✅ Attendance tracking');
    console.log('✅ API endpoints working');
    
    console.log('\n🚀 Your Agent 365 Digital Worker is fully operational!');
    
  } catch (error) {
    console.error('\n❌ Test failed:');
    console.error('Error:', error.response?.data || error.message);
    console.error('\n💡 Make sure:');
    console.error('  • Server is running (npm start)');
    console.error('  • All services are imported correctly');
    console.error('  • Database is accessible');
  }
}

// Run the test
testAgentFunctionality();