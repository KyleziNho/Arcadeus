/**
 * Enhanced Excel Agent - Built following OpenAI's Agent Best Practices
 * Implements proper foundations, orchestration, guardrails, and human intervention
 */

class EnhancedExcelAgent {
    constructor() {
        this.isInitialized = false;
        this.conversationId = this.generateConversationId();
        this.actionHistory = [];
        this.guardrails = new AgentGuardrails();
        this.tools = new ExcelToolkit();
        this.orchestrator = new AgentOrchestrator(this);
        
        this.setupAgent();
    }

    async setupAgent() {
        console.log('🚀 Initializing Enhanced Excel Agent...');
        
        try {
            // Wait for Office.js to be ready
            await this.waitForOffice();
            
            // Initialize components
            await this.tools.initialize();
            await this.guardrails.initialize();
            
            this.isInitialized = true;
            console.log('✅ Enhanced Excel Agent ready');
            
            // Log agent readiness
            this.logEvent('agent_initialized', { conversationId: this.conversationId });
            
        } catch (error) {
            console.error('❌ Failed to initialize Enhanced Excel Agent:', error);
            throw error;
        }
    }

    /**
     * Main agent entry point - processes user requests with full workflow
     */
    async processRequest(userMessage, context = {}) {
        if (!this.isInitialized) {
            throw new Error('Enhanced Excel Agent not initialized');
        }

        const requestId = this.generateRequestId();
        console.log(`🎯 Processing request ${requestId}:`, userMessage);

        try {
            // Step 1: Apply guardrails
            const guardrailsResult = await this.guardrails.validateRequest(userMessage, context);
            if (!guardrailsResult.approved) {
                return this.createErrorResponse(guardrailsResult.reason, 'guardrails_blocked');
            }

            // Step 2: Analyze and classify request
            const analysis = await this.analyzeRequest(userMessage, context);
            
            // Step 3: Route to appropriate orchestrator
            const response = await this.orchestrator.handleRequest(analysis, context);
            
            // Step 4: Apply output guardrails
            const finalResponse = await this.guardrails.validateResponse(response);
            
            // Step 5: Log the interaction
            this.logInteraction(requestId, userMessage, finalResponse, analysis);
            
            return finalResponse;

        } catch (error) {
            console.error('Error processing request:', error);
            return this.createErrorResponse(error.message, 'processing_error');
        }
    }

    /**
     * Analyze user request using structured approach
     */
    async analyzeRequest(message, context) {
        const analysis = {
            requestId: this.generateRequestId(),
            timestamp: new Date().toISOString(),
            message: message,
            context: context,
            intent: null,
            confidence: 0,
            requiredTools: [],
            riskLevel: 'low',
            requiresConfirmation: false,
            extractedEntities: {}
        };

        // Intent classification using structured patterns
        const intents = [
            {
                name: 'format_cells',
                patterns: [/change.*color/i, /format.*cell/i, /make.*bold/i, /style.*range/i],
                tools: ['cellFormatter'],
                riskLevel: 'low'
            },
            {
                name: 'modify_data',
                patterns: [/update.*value/i, /change.*number/i, /set.*cell/i, /replace.*data/i],
                tools: ['dataModifier'],
                riskLevel: 'medium',
                requiresConfirmation: true
            },
            {
                name: 'add_formula',
                patterns: [/add.*formula/i, /calculate.*sum/i, /create.*function/i],
                tools: ['formulaBuilder'],
                riskLevel: 'medium'
            },
            {
                name: 'analyze_data',
                patterns: [/analyze.*data/i, /what.*value/i, /tell.*me.*about/i, /explain/i],
                tools: ['dataAnalyzer'],
                riskLevel: 'low'
            },
            {
                name: 'delete_data',
                patterns: [/delete.*row/i, /remove.*column/i, /clear.*range/i],
                tools: ['dataModifier'],
                riskLevel: 'high',
                requiresConfirmation: true
            }
        ];

        // Find matching intent
        let bestMatch = null;
        let highestConfidence = 0;

        for (const intent of intents) {
            for (const pattern of intent.patterns) {
                if (pattern.test(message)) {
                    const confidence = this.calculateConfidence(message, pattern);
                    if (confidence > highestConfidence) {
                        highestConfidence = confidence;
                        bestMatch = intent;
                    }
                }
            }
        }

        if (bestMatch) {
            analysis.intent = bestMatch.name;
            analysis.confidence = highestConfidence;
            analysis.requiredTools = bestMatch.tools;
            analysis.riskLevel = bestMatch.riskLevel;
            analysis.requiresConfirmation = bestMatch.requiresConfirmation;
        }

        // Extract entities (colors, ranges, values, etc.)
        analysis.extractedEntities = this.extractEntities(message);

        return analysis;
    }

    /**
     * Extract entities from user message
     */
    extractEntities(message) {
        const entities = {};

        // Extract colors
        const colors = ['red', 'blue', 'green', 'yellow', 'orange', 'purple', 'pink', 'gray', 'black', 'white'];
        const foundColors = colors.filter(color => 
            new RegExp(`\\b${color}\\b`, 'i').test(message)
        );
        if (foundColors.length > 0) {
            entities.colors = foundColors;
        }

        // Extract cell ranges (A1, B2:D4, etc.)
        const rangePattern = /\b[A-Z]+\d+(?::[A-Z]+\d+)?\b/g;
        const ranges = message.match(rangePattern);
        if (ranges) {
            entities.ranges = ranges;
        }

        // Extract numbers
        const numberPattern = /\b\d+(?:\.\d+)?\b/g;
        const numbers = message.match(numberPattern);
        if (numbers) {
            entities.numbers = numbers.map(n => parseFloat(n));
        }

        return entities;
    }

    /**
     * Calculate confidence score for pattern matching
     */
    calculateConfidence(message, pattern) {
        const match = message.match(pattern);
        if (!match) return 0;
        
        // Base confidence from match
        let confidence = 0.7;
        
        // Boost for exact keyword matches
        const keywords = ['change', 'format', 'update', 'add', 'delete', 'analyze'];
        for (const keyword of keywords) {
            if (new RegExp(`\\b${keyword}\\b`, 'i').test(message)) {
                confidence += 0.1;
            }
        }
        
        return Math.min(confidence, 1.0);
    }

    /**
     * Create standardized error response
     */
    createErrorResponse(message, type) {
        return {
            success: false,
            type: type,
            message: message,
            timestamp: new Date().toISOString(),
            conversationId: this.conversationId
        };
    }

    /**
     * Log agent interactions for evaluation
     */
    logInteraction(requestId, userMessage, response, analysis) {
        const logEntry = {
            requestId,
            timestamp: new Date().toISOString(),
            conversationId: this.conversationId,
            userMessage,
            analysis,
            response: {
                success: response.success,
                type: response.type,
                actionsTaken: response.actionsTaken || []
            }
        };

        this.actionHistory.push(logEntry);
        
        // Keep only last 50 interactions
        if (this.actionHistory.length > 50) {
            this.actionHistory = this.actionHistory.slice(-50);
        }

        console.log('📊 Logged interaction:', logEntry);
    }

    /**
     * Log significant events
     */
    logEvent(eventType, data) {
        console.log(`📅 Event: ${eventType}`, data);
    }

    /**
     * Wait for Office.js to be ready
     */
    async waitForOffice() {
        return new Promise((resolve, reject) => {
            const checkOffice = () => {
                if (typeof Office !== 'undefined' && Office.context) {
                    resolve();
                } else {
                    setTimeout(checkOffice, 100);
                }
            };
            checkOffice();
            
            // Timeout after 10 seconds
            setTimeout(() => reject(new Error('Office.js not available')), 10000);
        });
    }

    /**
     * Generate unique IDs
     */
    generateConversationId() {
        return 'conv_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
    }

    generateRequestId() {
        return 'req_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
    }

    /**
     * Get agent performance metrics
     */
    getPerformanceMetrics() {
        const totalRequests = this.actionHistory.length;
        const successfulRequests = this.actionHistory.filter(h => h.response.success).length;
        const successRate = totalRequests > 0 ? (successfulRequests / totalRequests) : 0;

        return {
            totalRequests,
            successfulRequests,
            successRate,
            averageConfidence: this.actionHistory.reduce((sum, h) => sum + (h.analysis.confidence || 0), 0) / totalRequests || 0,
            conversationId: this.conversationId,
            agentUptime: Date.now() - this.startTime
        };
    }
}

/**
 * Agent Guardrails - Safety and validation system
 */
class AgentGuardrails {
    constructor() {
        this.safeguards = {
            maxCellsPerOperation: 1000,
            maxFormulaComplexity: 50,
            prohibitedPatterns: [
                /delete.*all/i,
                /clear.*everything/i,
                /format.*entire.*sheet/i
            ],
            requireConfirmation: [
                /delete/i,
                /remove/i,
                /clear/i
            ]
        };
    }

    async initialize() {
        console.log('🛡️ Initializing guardrails...');
    }

    /**
     * Validate user request before processing
     */
    async validateRequest(message, context) {
        // Check for prohibited patterns
        for (const pattern of this.safeguards.prohibitedPatterns) {
            if (pattern.test(message)) {
                return {
                    approved: false,
                    reason: 'Request contains prohibited operation',
                    severity: 'high'
                };
            }
        }

        // Check message length and complexity
        if (message.length > 1000) {
            return {
                approved: false,
                reason: 'Request too long',
                severity: 'medium'
            };
        }

        // All checks passed
        return { approved: true };
    }

    /**
     * Validate response before sending to user
     */
    async validateResponse(response) {
        // Ensure response has required fields
        if (!response.success !== undefined || !response.message) {
            response.success = false;
            response.message = 'Invalid response format';
        }

        return response;
    }
}

/**
 * Excel Toolkit - Standardized tools for Excel operations
 */
class ExcelToolkit {
    constructor() {
        this.tools = new Map();
    }

    async initialize() {
        console.log('🔧 Initializing Excel toolkit...');
        
        // Register available tools
        this.registerTool('cellFormatter', new CellFormatterTool());
        this.registerTool('dataModifier', new DataModifierTool());
        this.registerTool('formulaBuilder', new FormulaBuilderTool());
        this.registerTool('dataAnalyzer', new DataAnalyzerTool());
    }

    registerTool(name, toolInstance) {
        this.tools.set(name, toolInstance);
    }

    getTool(name) {
        return this.tools.get(name);
    }

    getAvailableTools() {
        return Array.from(this.tools.keys());
    }
}

/**
 * Agent Orchestrator - Manages workflow and tool coordination
 */
class AgentOrchestrator {
    constructor(agent) {
        this.agent = agent;
        this.workflows = new Map();
        this.setupWorkflows();
    }

    setupWorkflows() {
        // Define workflows for different request types
        this.workflows.set('format_cells', new FormatCellsWorkflow());
        this.workflows.set('modify_data', new ModifyDataWorkflow());
        this.workflows.set('add_formula', new AddFormulaWorkflow());
        this.workflows.set('analyze_data', new AnalyzeDataWorkflow());
        this.workflows.set('delete_data', new DeleteDataWorkflow());
    }

    /**
     * Handle request using appropriate workflow
     */
    async handleRequest(analysis, context) {
        const workflow = this.workflows.get(analysis.intent);
        
        if (!workflow) {
            return {
                success: false,
                type: 'unknown_intent',
                message: 'I don\'t understand what you want me to do. Please try rephrasing your request.'
            };
        }

        // Execute workflow
        return await workflow.execute(analysis, context, this.agent.tools);
    }
}

// Export for global use
window.EnhancedExcelAgent = EnhancedExcelAgent;
console.log('🚀 Enhanced Excel Agent loaded');