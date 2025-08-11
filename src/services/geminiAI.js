const { GoogleGenerativeAI } = require('@google/generative-ai');
const logger = require('../utils/logger');

class GeminiAIService {
  constructor() {
    this.apiKey = process.env.GEMINI_API_KEY;
    this.modelName = process.env.GEMINI_MODEL || 'gemini-1.5-flash';
    
    if (!this.apiKey) {
      logger.warn('⚠️ Gemini API key not provided. AI features will be limited.');
      this.genAI = null;
      this.model = null;
      return;
    }

    try {
      this.genAI = new GoogleGenerativeAI(this.apiKey);
      this.model = this.genAI.getGenerativeModel({ model: this.modelName });
      logger.info('✅ Gemini AI initialized successfully');
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

  // Generate intelligent meeting agenda
  async generateMeetingAgenda(meetingInfo) {
    if (!this.isAvailable()) {
      return this.getFallbackAgenda(meetingInfo);
    }

    try {
      const { subject, attendees = [], duration = 30, meetingType = 'general' } = meetingInfo;
      
      const prompt = `
        Generate a professional meeting agenda for the following meeting:
        
        Subject: ${subject}
        Duration: ${duration} minutes
        Number of Attendees: ${attendees.length}
        Meeting Type: ${meetingType}
        
        Please provide:
        1. A clear agenda with time allocations
        2. Suggested discussion points
        3. Recommended meeting structure
        4. Expected outcomes
        
        Format the response as JSON with this structure:
        {
          "title": "Meeting title",
          "estimatedDuration": ${duration},
          "sections": [
            {
              "title": "Section name",
              "duration": "minutes",
              "description": "What will be covered"
            }
          ],
          "objectives": ["objective1", "objective2"],
          "suggestedPreparation": ["prep item 1", "prep item 2"]
        }
      `;

      const result = await this.model.generateContent(prompt);
      const response = await result.response;
      const text = response.text();
      
      // Try to parse JSON response
      try {
        const agenda = JSON.parse(text.replace(/```json|```/g, '').trim());
        logger.info('✅ AI-generated meeting agenda created');
        return agenda;
      } catch (parseError) {
        logger.warn('⚠️ Failed to parse AI response, using fallback');
        return this.getFallbackAgenda(meetingInfo);
      }

    } catch (error) {
      logger.error('❌ Error generating meeting agenda:', error);
      return this.getFallbackAgenda(meetingInfo);
    }
  }

  // Suggest optimal meeting times using AI
  async suggestMeetingTimes(requirements) {
    if (!this.isAvailable()) {
      return this.getFallbackTimeSlots(requirements);
    }

    try {
      const { 
        duration = 30, 
        attendees = [], 
        preferredDays = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'],
        timeZone = 'UTC',
        urgency = 'normal'
      } = requirements;

      const prompt = `
        Suggest optimal meeting times based on these requirements:
        
        Duration: ${duration} minutes
        Number of Attendees: ${attendees.length}
        Preferred Days: ${preferredDays.join(', ')}
        Urgency: ${urgency}
        Time Zone: ${timeZone}
        
        Consider:
        - Best times for productivity and engagement
        - Common working hours across time zones
        - Meeting fatigue (avoid back-to-back meetings)
        - Energy levels throughout the day
        
        Provide 5 suggestions for the next 7 days in JSON format:
        {
          "suggestions": [
            {
              "datetime": "2024-MM-DDTHH:mm:ss.000Z",
              "dayOfWeek": "Monday",
              "timeSlot": "Morning/Afternoon/Evening",
              "confidence": 0.9,
              "reasoning": "Why this time is optimal"
            }
          ]
        }
      `;

      const result = await this.model.generateContent(prompt);
      const response = await result.response;
      const text = response.text();
      
      try {
        const suggestions = JSON.parse(text.replace(/```json|```/g, '').trim());
        logger.info('✅ AI-generated meeting time suggestions created');
        return suggestions.suggestions || [];
      } catch (parseError) {
        return this.getFallbackTimeSlots(requirements);
      }

    } catch (error) {
      logger.error('❌ Error suggesting meeting times:', error);
      return this.getFallbackTimeSlots(requirements);
    }
  }

  // Analyze meeting description and extract insights
  async analyzeMeetingDescription(description) {
    if (!this.isAvailable()) {
      return this.getFallbackAnalysis(description);
    }

    try {
      const prompt = `
        Analyze this meeting description and extract key insights:
        
        "${description}"
        
        Provide analysis in JSON format:
        {
          "urgency": "low/medium/high",
          "estimatedDuration": "minutes",
          "topics": ["topic1", "topic2"],
          "suggestedAttendees": ["role1", "role2"],
          "meetingType": "planning/review/decision/update/training",
          "preparationNeeded": true/false,
          "keyQuestions": ["question1", "question2"],
          "expectedOutcomes": ["outcome1", "outcome2"]
        }
      `;

      const result = await this.model.generateContent(prompt);
      const response = await result.response;
      const text = response.text();
      
      try {
        const analysis = JSON.parse(text.replace(/```json|```/g, '').trim());
        logger.info('✅ AI meeting description analysis completed');
        return analysis;
      } catch (parseError) {
        return this.getFallbackAnalysis(description);
      }

    } catch (error) {
      logger.error('❌ Error analyzing meeting description:', error);
      return this.getFallbackAnalysis(description);
    }
  }

  // Generate smart meeting title based on description
  async generateMeetingTitle(description, attendees = []) {
    if (!this.isAvailable()) {
      return `Meeting - ${new Date().toLocaleDateString()}`;
    }

    try {
      const prompt = `
        Generate a concise, professional meeting title based on this description:
        
        "${description}"
        
        Attendees: ${attendees.length} people
        
        Requirements:
        - Maximum 60 characters
        - Clear and descriptive
        - Professional tone
        - Include key topic/purpose
        
        Respond with just the title, no additional text.
      `;

      const result = await this.model.generateContent(prompt);
      const response = await result.response;
      const title = response.text().trim().replace(/['"]/g, '');
      
      logger.info('✅ AI-generated meeting title created');
      return title;

    } catch (error) {
      logger.error('❌ Error generating meeting title:', error);
      return `Meeting - ${new Date().toLocaleDateString()}`;
    }
  }

  // Validate meeting content for appropriateness
  async validateMeetingContent(content) {
    if (!this.isAvailable()) {
      return { isAppropriate: true, confidence: 0.5, suggestions: [] };
    }

    try {
      const prompt = `
        Analyze this meeting content for business appropriateness and provide feedback:
        
        "${content}"
        
        Check for:
        - Professional language
        - Clear objectives
        - Appropriate tone
        - Sensitive information handling
        - Compliance considerations
        
        Respond in JSON format:
        {
          "isAppropriate": true/false,
          "confidence": 0.0-1.0,
          "issues": ["issue1", "issue2"],
          "suggestions": ["suggestion1", "suggestion2"],
          "sensitivityLevel": "low/medium/high"
        }
      `;

      const result = await this.model.generateContent(prompt);
      const response = await result.response;
      const text = response.text();
      
      try {
        const validation = JSON.parse(text.replace(/```json|```/g, '').trim());
        logger.info('✅ AI content validation completed');
        return validation;
      } catch (parseError) {
        return { isAppropriate: true, confidence: 0.5, suggestions: [] };
      }

    } catch (error) {
      logger.error('❌ Error validating meeting content:', error);
      return { isAppropriate: true, confidence: 0.5, suggestions: [] };
    }
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
          description: 'Brief introductions and meeting overview'
        },
        {
          title: 'Main Discussion',
          duration: `${duration - 15} minutes`,
          description: `Focused discussion on: ${subject}`
        },
        {
          title: 'Action Items & Next Steps',
          duration: '10 minutes',
          description: 'Review decisions and assign next steps'
        }
      ],
      objectives: ['Discuss main topic', 'Make decisions', 'Assign action items'],
      suggestedPreparation: ['Review agenda', 'Prepare questions']
    };
  }

  getFallbackTimeSlots(requirements) {
    const { duration = 30 } = requirements;
    const suggestions = [];
    const now = new Date();
    
    // Generate 5 basic time suggestions
    for (let i = 1; i <= 5; i++) {
      const futureDate = new Date(now);
      futureDate.setDate(futureDate.getDate() + i);
      futureDate.setHours(10, 0, 0, 0); // 10 AM
      
      suggestions.push({
        datetime: futureDate.toISOString(),
        dayOfWeek: futureDate.toLocaleDateString('en-US', { weekday: 'long' }),
        timeSlot: 'Morning',
        confidence: 0.7,
        reasoning: 'Standard business hours, good for productivity'
      });
    }
    
    return suggestions;
  }

  getFallbackAnalysis(description) {
    return {
      urgency: 'medium',
      estimatedDuration: '30',
      topics: ['General discussion'],
      suggestedAttendees: ['Team members'],
      meetingType: 'general',
      preparationNeeded: true,
      keyQuestions: ['What are our objectives?'],
      expectedOutcomes: ['Clear next steps']
    };
  }
}

// Create singleton instance
const geminiAI = new GeminiAIService();

module.exports = geminiAI;