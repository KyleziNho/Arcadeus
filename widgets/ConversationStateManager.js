/**
 * Conversation State Manager - Manages agent conversation context and memory
 * Implements OpenAI's best practices for conversation management
 */

class ConversationStateManager {
    constructor() {
        this.conversations = new Map();
        this.activeConversationId = null;
        this.maxContextLength = 10; // Keep last 10 interactions
        this.maxConversations = 50; // Keep last 50 conversations
        this.persistenceKey = 'arcadeus_conversation_state';
        
        this.loadFromStorage();
    }

    /**
     * Start a new conversation
     */
    startNewConversation(userId = 'default', context = {}) {
        const conversationId = this.generateConversationId();
        
        const conversation = {
            id: conversationId,
            userId,
            startTime: new Date().toISOString(),
            lastActivity: new Date().toISOString(),
            context: { ...context },
            messages: [],
            agentState: {
                currentWorkflow: null,
                pendingActions: [],
                lastIntent: null,
                confidence: 0
            },
            metadata: {
                totalMessages: 0,
                avgResponseTime: 0,
                errors: 0,
                successfulActions: 0
            }
        };

        this.conversations.set(conversationId, conversation);
        this.activeConversationId = conversationId;
        
        this.saveToStorage();
        console.log(`💬 Started new conversation: ${conversationId}`);
        
        return conversationId;
    }

    /**
     * Get active conversation or create new one
     */
    getActiveConversation() {
        if (!this.activeConversationId || !this.conversations.has(this.activeConversationId)) {
            return this.conversations.get(this.startNewConversation());
        }
        return this.conversations.get(this.activeConversationId);
    }

    /**
     * Add message to conversation history
     */
    addMessage(message, role = 'user', metadata = {}) {
        const conversation = this.getActiveConversation();
        if (!conversation) return;

        const messageEntry = {
            id: this.generateMessageId(),
            role, // 'user', 'assistant', 'system'
            content: message,
            timestamp: new Date().toISOString(),
            metadata: {
                responseTime: metadata.responseTime || null,
                intent: metadata.intent || null,
                confidence: metadata.confidence || null,
                actionsTaken: metadata.actionsTaken || [],
                ...metadata
            }
        };

        conversation.messages.push(messageEntry);
        conversation.lastActivity = new Date().toISOString();
        conversation.metadata.totalMessages++;

        // Update response time average
        if (metadata.responseTime && role === 'assistant') {
            const totalTime = conversation.metadata.avgResponseTime * (conversation.metadata.totalMessages - 1);
            conversation.metadata.avgResponseTime = (totalTime + metadata.responseTime) / conversation.metadata.totalMessages;
        }

        // Trim conversation to max length
        if (conversation.messages.length > this.maxContextLength * 2) {
            // Keep system messages and trim user/assistant pairs
            const systemMessages = conversation.messages.filter(m => m.role === 'system');
            const recentMessages = conversation.messages.slice(-this.maxContextLength * 2);
            conversation.messages = [...systemMessages, ...recentMessages];
        }

        this.saveToStorage();
        return messageEntry.id;
    }

    /**
     * Update agent state in current conversation
     */
    updateAgentState(updates) {
        const conversation = this.getActiveConversation();
        if (!conversation) return;

        conversation.agentState = {
            ...conversation.agentState,
            ...updates
        };

        this.saveToStorage();
    }

    /**
     * Get conversation context for agent processing
     */
    getConversationContext(conversationId = null) {
        const conversation = conversationId ? 
            this.conversations.get(conversationId) : 
            this.getActiveConversation();

        if (!conversation) return null;

        // Build context from recent messages
        const recentMessages = conversation.messages.slice(-this.maxContextLength);
        
        return {
            conversationId: conversation.id,
            userId: conversation.userId,
            context: conversation.context,
            recentMessages: recentMessages.map(m => ({
                role: m.role,
                content: m.content,
                intent: m.metadata.intent,
                timestamp: m.timestamp
            })),
            agentState: conversation.agentState,
            summary: this.generateConversationSummary(conversation)
        };
    }

    /**
     * Generate conversation summary for context
     */
    generateConversationSummary(conversation) {
        const messages = conversation.messages;
        if (messages.length === 0) return 'No previous conversation';

        const userMessages = messages.filter(m => m.role === 'user').length;
        const assistantMessages = messages.filter(m => m.role === 'assistant').length;
        const intents = [...new Set(messages
            .map(m => m.metadata.intent)
            .filter(intent => intent))];

        const recentTopics = messages.slice(-6)
            .filter(m => m.role === 'user')
            .map(m => this.extractTopic(m.content))
            .filter(topic => topic);

        return {
            messageCount: { user: userMessages, assistant: assistantMessages },
            commonIntents: intents.slice(0, 3),
            recentTopics: [...new Set(recentTopics)].slice(0, 3),
            avgResponseTime: conversation.metadata.avgResponseTime,
            successRate: conversation.metadata.successfulActions / Math.max(assistantMessages, 1)
        };
    }

    /**
     * Extract topic from user message (simple keyword extraction)
     */
    extractTopic(content) {
        const topicKeywords = [
            'formatting', 'analysis', 'calculation', 'chart', 'formula',
            'color', 'bold', 'italic', 'sum', 'average', 'irr', 'moic',
            'revenue', 'expenses', 'debt', 'equity', 'assumptions'
        ];

        const words = content.toLowerCase().split(/\s+/);
        for (const keyword of topicKeywords) {
            if (words.some(word => word.includes(keyword))) {
                return keyword;
            }
        }
        return null;
    }

    /**
     * Search conversation history
     */
    searchConversations(query, options = {}) {
        const results = [];
        const searchTerms = query.toLowerCase().split(/\s+/);

        for (const [id, conversation] of this.conversations) {
            let relevanceScore = 0;
            const matchingMessages = [];

            for (const message of conversation.messages) {
                const content = message.content.toLowerCase();
                const matches = searchTerms.filter(term => content.includes(term));
                
                if (matches.length > 0) {
                    const score = matches.length / searchTerms.length;
                    relevanceScore = Math.max(relevanceScore, score);
                    
                    if (score > 0.5) {
                        matchingMessages.push({
                            ...message,
                            relevanceScore: score
                        });
                    }
                }
            }

            if (relevanceScore > 0.3) {
                results.push({
                    conversationId: id,
                    relevanceScore,
                    matchingMessages: matchingMessages.slice(0, 3), // Top 3 matches
                    summary: this.generateConversationSummary(conversation),
                    lastActivity: conversation.lastActivity
                });
            }
        }

        // Sort by relevance and recency
        results.sort((a, b) => {
            const relevanceDiff = b.relevanceScore - a.relevanceScore;
            if (Math.abs(relevanceDiff) > 0.1) return relevanceDiff;
            return new Date(b.lastActivity) - new Date(a.lastActivity);
        });

        return results.slice(0, options.limit || 10);
    }

    /**
     * Get conversation analytics
     */
    getAnalytics() {
        const conversations = Array.from(this.conversations.values());
        
        const analytics = {
            totalConversations: conversations.length,
            totalMessages: conversations.reduce((sum, c) => sum + c.messages.length, 0),
            avgMessagesPerConversation: 0,
            avgResponseTime: 0,
            commonIntents: {},
            userSatisfaction: 0,
            timeDistribution: {},
            errorRate: 0
        };

        if (conversations.length > 0) {
            analytics.avgMessagesPerConversation = analytics.totalMessages / conversations.length;
            
            // Calculate average response time
            let totalResponseTime = 0;
            let responseCount = 0;
            
            conversations.forEach(conv => {
                if (conv.metadata.avgResponseTime > 0) {
                    totalResponseTime += conv.metadata.avgResponseTime;
                    responseCount++;
                }
            });
            
            analytics.avgResponseTime = responseCount > 0 ? totalResponseTime / responseCount : 0;

            // Count intents
            conversations.forEach(conv => {
                conv.messages.forEach(msg => {
                    if (msg.metadata.intent) {
                        analytics.commonIntents[msg.metadata.intent] = 
                            (analytics.commonIntents[msg.metadata.intent] || 0) + 1;
                    }
                });
            });

            // Time distribution (by hour of day)
            conversations.forEach(conv => {
                conv.messages.forEach(msg => {
                    const hour = new Date(msg.timestamp).getHours();
                    analytics.timeDistribution[hour] = (analytics.timeDistribution[hour] || 0) + 1;
                });
            });

            // Error rate
            const totalErrors = conversations.reduce((sum, c) => sum + c.metadata.errors, 0);
            analytics.errorRate = totalErrors / analytics.totalMessages;
        }

        return analytics;
    }

    /**
     * Cleanup old conversations
     */
    cleanup() {
        if (this.conversations.size <= this.maxConversations) return;

        // Sort by last activity
        const sorted = Array.from(this.conversations.entries())
            .sort(([,a], [,b]) => new Date(b.lastActivity) - new Date(a.lastActivity));

        // Keep only the most recent conversations
        const toKeep = sorted.slice(0, this.maxConversations);
        const toRemove = sorted.slice(this.maxConversations);

        // Remove old conversations
        toRemove.forEach(([id]) => {
            this.conversations.delete(id);
        });

        console.log(`🧹 Cleaned up ${toRemove.length} old conversations`);
        this.saveToStorage();
    }

    /**
     * Persistence methods
     */
    saveToStorage() {
        try {
            const data = {
                conversations: Array.from(this.conversations.entries()),
                activeConversationId: this.activeConversationId,
                lastSaved: new Date().toISOString()
            };
            
            localStorage.setItem(this.persistenceKey, JSON.stringify(data));
        } catch (error) {
            console.error('Failed to save conversation state:', error);
        }
    }

    loadFromStorage() {
        try {
            const data = localStorage.getItem(this.persistenceKey);
            if (data) {
                const parsed = JSON.parse(data);
                this.conversations = new Map(parsed.conversations || []);
                this.activeConversationId = parsed.activeConversationId;
                
                console.log(`💾 Loaded ${this.conversations.size} conversations from storage`);
            }
        } catch (error) {
            console.error('Failed to load conversation state:', error);
            this.conversations = new Map();
        }
    }

    /**
     * Export conversation data
     */
    exportConversations(format = 'json') {
        const data = {
            conversations: Array.from(this.conversations.values()),
            exportDate: new Date().toISOString(),
            analytics: this.getAnalytics()
        };

        if (format === 'json') {
            return JSON.stringify(data, null, 2);
        } else if (format === 'csv') {
            // Simple CSV export of messages
            let csv = 'ConversationID,Role,Content,Timestamp,Intent,ResponseTime\n';
            data.conversations.forEach(conv => {
                conv.messages.forEach(msg => {
                    csv += `"${conv.id}","${msg.role}","${msg.content.replace(/"/g, '""')}","${msg.timestamp}","${msg.metadata.intent || ''}","${msg.metadata.responseTime || ''}"\n`;
                });
            });
            return csv;
        }

        return data;
    }

    /**
     * Utility methods
     */
    generateConversationId() {
        return 'conv_' + Date.now() + '_' + Math.random().toString(36).substr(2, 8);
    }

    generateMessageId() {
        return 'msg_' + Date.now() + '_' + Math.random().toString(36).substr(2, 6);
    }

    /**
     * Reset all conversation data (for testing/development)
     */
    reset() {
        this.conversations.clear();
        this.activeConversationId = null;
        localStorage.removeItem(this.persistenceKey);
        console.log('🔄 Reset all conversation data');
    }
}

// Export conversation state manager
window.ConversationStateManager = ConversationStateManager;

console.log('💬 Conversation State Manager loaded');