const { GoogleGenerativeAI } = require('@google/generative-ai');
const logger = require('../utils/logger');

class GeminiAIService {
  constructor() {
    this.apiKey = process.env.GEMINI_API_KEY;
    this.modelName = process.env.GEMINI_MODEL || 'gemini-2.0-flash-exp';
    
    if (!this.apiKey) {
      logger.error('❌ Gemini API key not provided. AI features will not be available.');
      logger.error('   Please set GEMINI_API_KEY environment variable');
      this.genAI = null;
      this.model = null;
      return;
    }

    try {
      this.genAI = new GoogleGenerativeAI(this.apiKey);
      this.model = this.genAI.getGenerativeModel({ model: this.modelName });
      logger.info('✅ Gemini AI initialized successfully');
      logger.info(`🤖 Using model: ${this.modelName}`);
    } catch (error) {
      logger.error('❌ Failed to initialize Gemini AI:', error);
      this.genAI = null;
      this.model = null;
    }
  }

  // Check if AI is available
  isAvailable() {
    return this.model !== null;
  }

  // Ensure AI is available
  ensureAIAvailable() {
    if (!this.isAvailable()) {
      throw new Error('AI service is required but not configured. Please set up Gemini API key.');
    }
  }

  // Test AI connection
  async testConnection() {
    if (!this.isAvailable()) {
      return {
        success: false,
        error: 'AI not configured',
        details: 'Gemini API key missing'
      };
    }

    try {
      const testPrompt = "Respond with 'AI connection successful' if you can understand this message.";
      const result = await this.model.generateContent(testPrompt);
      const response = await result.response;
      const text = response.text();

      logger.info('✅ Gemini AI connection test successful');
      
      return {
        success: true,
        message: 'AI service ready for real Teams meeting analysis',
        model: this.modelName,
        response: text.substring(0, 100)
      };

    } catch (error) {
      logger.error('❌ Gemini AI connection test failed:', error);
      
      return {
        success: false,
        error: 'AI connection failed',
        details: error.message
      };
    }
  }

  // Generate intelligent meeting agenda for real Teams meetings
  async generateMeetingAgenda(meetingInfo) {
    this.ensureAIAvailable();

    try {
      const { subject, attendees = [], duration = 30, meetingType = 'general' } = meetingInfo;
      
      const prompt = `
        Generate a professional meeting agenda for a REAL Microsoft Teams meeting:
        
        Subject: ${subject}
        Duration: ${duration} minutes
        Number of Attendees: ${attendees.length}
        Meeting Type: ${meetingType}
        
        This is for a real business meeting with actual participants. Create a practical, actionable agenda.
        
        Provide a structured agenda in JSON format:
        {
          "title": "Professional meeting title",
          "estimatedDuration": ${duration},
          "sections": [
            {
              "title": "Section name",
              "duration": "X minutes",
              "description": "What will be covered",
              "objectives": ["specific objective 1", "specific objective 2"]
            }
          ],
          "objectives": ["main objective 1", "main objective 2"],
          "suggestedPreparation": ["prep item 1", "prep item 2"],
          "expectedOutcomes": ["expected outcome 1", "expected outcome 2"]
        }
      `;

      const result = await this.model.generateContent(prompt);
      const response = await result.response;
      const text = response.text();
      
      try {
        const agenda = JSON.parse(text.replace(/```json|```/g, '').trim());
        logger.info('✅ AI-generated meeting agenda created for real Teams meeting');
        return agenda;
      } catch (parseError) {
        logger.warn('⚠️ Failed to parse AI agenda response, using structured fallback');
        return this.getFallbackAgenda(meetingInfo);
      }

    } catch (error) {
      logger.error('❌ Error generating meeting agenda with AI:', error);
      return this.getFallbackAgenda(meetingInfo);
    }
  }

  // Analyze real Teams meeting description
  async analyzeMeetingDescription(description) {
    this.ensureAIAvailable();

    try {
      const prompt = `
        Analyze this real Teams meeting description and extract key insights:
        
        "${description}"
        
        This is for an actual business meeting. Provide practical analysis.
        
        Provide analysis in JSON format:
        {
          "urgency": "low|medium|high",
          "estimatedDuration": "suggested duration in minutes",
          "topics": ["topic1", "topic2", "topic3"],
          "suggestedAttendees": ["role1", "role2", "role3"],
          "meetingType": "planning|review|decision|update|training|brainstorming",
          "preparationNeeded": true/false,
          "keyQuestions": ["important question 1", "important question 2"],
          "expectedOutcomes": ["concrete outcome 1", "concrete outcome 2"],
          "recommendedTimeSlot": "morning|afternoon|any",
          "complexity": "low|medium|high"
        }
      `;

      const result = await this.model.generateContent(prompt);
      const response = await result.response;
      const text = response.text();
      
      try {
        const analysis = JSON.parse(text.replace(/```json|```/g, '').trim());
        logger.info('✅ AI meeting description analysis completed for real Teams meeting');
        return analysis;
      } catch (parseError) {
        return this.getFallbackAnalysis(description);
      }

    } catch (error) {
      logger.error('❌ Error analyzing meeting description with AI:', error);
      return this.getFallbackAnalysis(description);
    }
  }

  // Generate smart meeting title for real Teams meetings
  async generateMeetingTitle(description, attendees = []) {
    this.ensureAIAvailable();

    try {
      const prompt = `
        Generate a concise, professional meeting title for a real Teams meeting based on this description:
        
        "${description}"
        
        Number of attendees: ${attendees.length}
        
        Requirements:
        - Maximum 60 characters
        - Clear and descriptive for business context
        - Professional tone appropriate for Teams
        - Include key topic/purpose
        - Should work well as a Teams meeting title
        
        Respond with just the title, no additional text.
      `;

      const result = await this.model.generateContent(prompt);
      const response = await result.response;
      const title = response.text().trim().replace(/['"]/g, '');
      
      logger.info('✅ AI-generated meeting title created for real Teams meeting');
      return title.substring(0, 60); // Ensure max length

    } catch (error) {
      logger.error('❌ Error generating meeting title with AI:', error);
      return `${description.substring(0, 40)} - Meeting`;
    }
  }

  // Validate meeting content for real business appropriateness
  async validateMeetingContent(content) {
    this.ensureAIAvailable();

    try {
      const prompt = `
        Analyze this meeting content for business appropriateness in a real Teams environment:
        
        "${content}"
        
        Check for:
        - Professional language suitable for Teams
        - Clear business objectives
        - Appropriate tone for corporate environment
        - Sensitive information considerations
        - Compliance with business communication standards
        
        Respond in JSON format:
        {
          "isAppropriate": true/false,
          "confidence": 0.0-1.0,
          "issues": ["specific issue 1", "specific issue 2"],
          "suggestions": ["specific suggestion 1", "specific suggestion 2"],
          "sensitivityLevel": "low|medium|high",
          "businessContext": "appropriate|needs_revision|inappropriate"
        }
      `;

      const result = await this.model.generateContent(prompt);
      const response = await result.response;
      const text = response.text();
      
      try {
        const validation = JSON.parse(text.replace(/```json|```/g, '').trim());
        logger.info('✅ AI content validation completed for real Teams meeting');
        return validation;
      } catch (parseError) {
        return { 
          isAppropriate: true, 
          confidence: 0.7, 
          suggestions: [],
          businessContext: 'appropriate' 
        };
      }

    } catch (error) {
      logger.error('❌ Error validating meeting content with AI:', error);
      return { 
        isAppropriate: true, 
        confidence: 0.5, 
        suggestions: [],
        businessContext: 'appropriate'
      };
    }
  }

  // Analyze real Teams chat message for meeting insights
  async analyzeChatMessage(content, context = {}) {
    this.ensureAIAvailable();

    try {
      const prompt = `
        Analyze this message from a real Teams meeting chat:
        
        Message: "${content}"
        Context: Meeting about ${context.meetingSubject || 'business topics'}
        
        Provide analysis in JSON format:
        {
          "category": "question|action_item|decision|resource_sharing|general|concern",
          "urgency": "low|medium|high", 
          "sentiment": "positive|neutral|negative",
          "actionRequired": true/false,
          "keyTopics": ["topic1", "topic2"],
          "mentions": ["person1", "person2"],
          "followUpNeeded": true/false,
          "businessImpact": "low|medium|high"
        }
      `;

      const result = await this.model.generateContent(prompt);
      const response = await result.response;
      const text = response.text();
      
      try {
        const analysis = JSON.parse(text.replace(/```json|```/g, '').trim());
        return analysis;
      } catch (parseError) {
        return this.getBasicMessageAnalysis(content);
      }

    } catch (error) {
      logger.warn('⚠️ AI chat message analysis failed, using basic analysis:', error);
      return this.getBasicMessageAnalysis(content);
    }
  }

  // Generate response for AI agent in real Teams meeting
  async generateMeetingResponse(userMessage, meetingContext = {}) {
    this.ensureAIAvailable();

    try {
      const prompt = `
        You are an AI Meeting Assistant participating in a live Microsoft Teams meeting.
        
        Participant message: "${userMessage}"
        Meeting context: ${meetingContext.subject || 'Business meeting'}
        
        Generate a helpful, professional response as if you're a meeting participant. 
        Keep it concise (2-3 sentences max) and actionable.
        
        You can:
        - Answer questions about the meeting
        - Summarize discussion points
        - Track action items and decisions
        - Provide meeting facilitation help
        - Offer relevant business insights
        
        Respond naturally as a professional meeting participant would.
      `;

      const result = await this.model.generateContent(prompt);
      const response = await result.response;
      const text = response.text().trim();
      
      logger.info('✅ AI meeting response generated for real Teams interaction');
      return text;

    } catch (error) {
      logger.error('❌ Error generating AI meeting response:', error);
      return "I'm here to help with the meeting. Could you please repeat your question?";
    }
  }

  // Get AI service status
  getStatus() {
    const available = this.isAvailable();
    
    return {
      available: available,
      model: this.modelName || 'not_configured',
      capabilities: {
        meetingAnalysis: available,
        agendaGeneration: available,
        chatAnalysis: available,
        summaryGeneration: available,
        realTimeInteraction: available
      },
      configuration: {
        apiKey: this.apiKey ? 'configured' : 'missing',
        modelName: this.modelName
      },
      message: available 
        ? 'AI service ready for real Teams meeting intelligence'
        : 'AI service not configured - set GEMINI_API_KEY'
    };
  }

  // Fallback methods when AI is not available
  getFallbackAgenda(meetingInfo) {
    const { subject, duration = 30 } = meetingInfo;
    return {
      title: subject,
      estimatedDuration: duration,
      sections: [
        {
          title: 'Welcome & Introductions',
          duration: '5 minutes',
          description: 'Brief introductions and meeting overview',
          objectives: ['Set meeting tone', 'Confirm attendees']
        },
        {
          title: 'Main Discussion',
          duration: `${Math.max(duration - 15, 15)} minutes`,
          description: `Focused discussion on: ${subject}`,
          objectives: ['Address main topics', 'Gather input from all participants']
        },
        {
          title: 'Action Items & Next Steps',
          duration: '10 minutes',
          description: 'Review decisions and assign next steps',
          objectives: ['Define clear action items', 'Set follow-up timeline']
        }
      ],
      objectives: ['Discuss main topic', 'Make informed decisions', 'Assign clear action items'],
      suggestedPreparation: ['Review agenda', 'Prepare relevant materials', 'Think of questions'],
      expectedOutcomes: ['Clear action items', 'Defined next steps', 'Aligned understanding']
    };
  }

  getFallbackAnalysis(description) {
    return {
      urgency: 'medium',
      estimatedDuration: '30',
      topics: ['Discussion topics from description'],
      suggestedAttendees: ['Relevant team members'],
      meetingType: 'general',
      preparationNeeded: true,
      keyQuestions: ['What are our main objectives?', 'What decisions need to be made?'],
      expectedOutcomes: ['Clear action items', 'Aligned understanding'],
      recommendedTimeSlot: 'any',
      complexity: 'medium'
    };
  }

  getBasicMessageAnalysis(content) {
    const contentLower = content.toLowerCase();
    
    return {
      category: contentLower.includes('?') ? 'question' : 
                contentLower.includes('action') ? 'action_item' :
                contentLower.includes('decided') ? 'decision' : 'general',
      urgency: contentLower.includes('urgent') ? 'high' : 'low',
      sentiment: 'neutral',
      actionRequired: contentLower.includes('action') || contentLower.includes('todo'),
      keyTopics: [],
      mentions: [],
      followUpNeeded: false,
      businessImpact: 'medium'
    };
  }
}

// Create singleton instance
const geminiAI = new GeminiAIService();

module.exports = geminiAI;