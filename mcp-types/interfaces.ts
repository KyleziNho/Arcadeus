/**
 * MCP Type Definitions and Interfaces
 * Core types for the Model Context Protocol implementation
 */

// ===== Core MCP Types =====

export interface MCPCapabilities {
  tools?: {
    listChanged?: boolean;
  };
  resources?: {
    listChanged?: boolean;
    subscribe?: boolean;
  };
  prompts?: {
    listChanged?: boolean;
  };
  logging?: {};
  notifications?: boolean;
}

export interface MCPServerInfo {
  name: string;
  version: string;
  description?: string;
  capabilities: MCPCapabilities;
}

export interface MCPClientInfo {
  name: string;
  version: string;
  capabilities: MCPCapabilities;
}

// ===== Tool Definitions =====

export interface MCPTool {
  name: string;
  title?: string;
  description: string;
  inputSchema: {
    type: string;
    properties?: Record<string, any>;
    required?: string[];
  };
}

export interface MCPToolCall {
  name: string;
  arguments: Record<string, any>;
}

export interface MCPToolResult {
  content: MCPContent[];
  isError?: boolean;
  errorMessage?: string;
}

// ===== Resource Definitions =====

export interface MCPResource {
  uri: string;
  name: string;
  description?: string;
  mimeType?: string;
}

export interface MCPResourceContent {
  uri: string;
  content: MCPContent[];
}

// ===== Content Types =====

export interface MCPContent {
  type: 'text' | 'image' | 'resource' | 'error';
  text?: string;
  data?: string; // Base64 for images
  mimeType?: string;
  uri?: string; // For resource references
}

// ===== Prompt Definitions =====

export interface MCPPrompt {
  name: string;
  description?: string;
  arguments?: MCPPromptArgument[];
}

export interface MCPPromptArgument {
  name: string;
  description?: string;
  required?: boolean;
}

// ===== Notification Types =====

export interface MCPNotification {
  method: string;
  params?: any;
}

// ===== Excel-Specific Types =====

export interface ExcelRange {
  worksheet: string;
  address: string;
  values?: any[][];
  formulas?: string[][];
  format?: ExcelFormat;
}

export interface ExcelFormat {
  backgroundColor?: string;
  fontColor?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  numberFormat?: string;
  borders?: ExcelBorder;
}

export interface ExcelBorder {
  top?: BorderStyle;
  bottom?: BorderStyle;
  left?: BorderStyle;
  right?: BorderStyle;
}

export interface BorderStyle {
  style: 'thin' | 'medium' | 'thick' | 'double';
  color: string;
}

export interface ExcelConditionalFormat {
  type: 'cellValue' | 'colorScale' | 'dataBar' | 'iconSet';
  priority: number;
  condition?: {
    operator: string;
    formula1: string;
    formula2?: string;
  };
  format?: ExcelFormat;
}

// ===== Conversation Context Types =====

export interface ConversationContext {
  sessionId: string;
  history: ConversationMessage[];
  currentExcelState: ExcelState;
  lastOperation?: Operation;
  mentionedRanges: ExcelRange[];
  userPreferences: UserPreferences;
}

export interface ConversationMessage {
  id: string;
  role: 'user' | 'assistant' | 'system';
  content: string;
  timestamp: Date;
  toolCalls?: MCPToolCall[];
  toolResults?: MCPToolResult[];
  metadata?: MessageMetadata;
}

export interface MessageMetadata {
  excelContext?: ExcelRange[];
  operations?: Operation[];
  confidence?: number;
  intentType?: string;
}

export interface ExcelState {
  activeWorksheet: string;
  selectedRange: string;
  worksheets: WorksheetInfo[];
  namedRanges: NamedRange[];
  recentChanges: ChangeRecord[];
}

export interface WorksheetInfo {
  name: string;
  index: number;
  visible: boolean;
  protected: boolean;
}

export interface NamedRange {
  name: string;
  range: string;
  scope: 'workbook' | 'worksheet';
}

export interface ChangeRecord {
  timestamp: Date;
  range: ExcelRange;
  previousValue: any;
  newValue: any;
  operation: string;
}

// ===== Operation Types for Undo/Redo =====

export interface Operation {
  id: string;
  type: OperationType;
  timestamp: Date;
  description: string;
  params: Record<string, any>;
  inverse?: Operation; // For undo
  status: 'pending' | 'completed' | 'failed' | 'undone';
  error?: string;
}

export type OperationType = 
  | 'write-value'
  | 'write-formula'
  | 'apply-format'
  | 'insert-rows'
  | 'insert-columns'
  | 'delete-range'
  | 'create-chart'
  | 'apply-conditional-format'
  | 'merge-cells'
  | 'sort-range'
  | 'filter-range';

export interface OperationHistory {
  operations: Operation[];
  currentIndex: number;
  maxSize: number;
}

// ===== User Preferences =====

export interface UserPreferences {
  defaultNumberFormat: string;
  defaultDateFormat: string;
  preferredColorScheme: ColorScheme;
  autoCalculate: boolean;
  showFormulas: boolean;
  language: string;
}

export interface ColorScheme {
  positive: string;
  negative: string;
  neutral: string;
  highlight: string;
}

// ===== AI Integration Types =====

export interface AIIntent {
  type: IntentType;
  confidence: number;
  entities: Entity[];
  suggestedTools: string[];
  clarificationNeeded?: string;
}

export type IntentType = 
  | 'read-data'
  | 'write-data'
  | 'format-cells'
  | 'calculate'
  | 'create-chart'
  | 'analyze-data'
  | 'validate-model'
  | 'explain-formula'
  | 'find-errors'
  | 'optimize-model';

export interface Entity {
  type: EntityType;
  value: string;
  confidence: number;
  normalizedValue?: any;
}

export type EntityType = 
  | 'range'
  | 'worksheet'
  | 'value'
  | 'formula'
  | 'color'
  | 'format-type'
  | 'calculation-type'
  | 'chart-type'
  | 'metric-name';

// ===== Financial Modeling Types =====

export interface FinancialMetric {
  name: string;
  value: number;
  formula?: string;
  range?: ExcelRange;
  dependencies?: string[];
  confidence?: number;
}

export interface FinancialModel {
  metrics: {
    irr?: FinancialMetric;
    npv?: FinancialMetric;
    moic?: FinancialMetric;
    paybackPeriod?: FinancialMetric;
    dcf?: FinancialMetric;
  };
  assumptions: Record<string, any>;
  scenarios: Scenario[];
}

export interface Scenario {
  name: string;
  description: string;
  assumptions: Record<string, any>;
  results: Record<string, number>;
}

// ===== Transport Types =====

export interface MCPTransport {
  type: 'stdio' | 'http';
  send(message: any): Promise<void>;
  receive(): AsyncIterator<any>;
  close(): Promise<void>;
}

export interface MCPSession {
  id: string;
  client: MCPClientInfo;
  server: MCPServerInfo;
  transport: MCPTransport;
  context: ConversationContext;
  isConnected: boolean;
  createdAt: Date;
  lastActivity: Date;
}

// ===== Error Types =====

export class MCPError extends Error {
  constructor(
    message: string,
    public code: MCPErrorCode,
    public details?: any
  ) {
    super(message);
    this.name = 'MCPError';
  }
}

export enum MCPErrorCode {
  // Protocol errors
  INVALID_REQUEST = 'INVALID_REQUEST',
  METHOD_NOT_FOUND = 'METHOD_NOT_FOUND',
  INVALID_PARAMS = 'INVALID_PARAMS',
  INTERNAL_ERROR = 'INTERNAL_ERROR',
  
  // Connection errors
  CONNECTION_FAILED = 'CONNECTION_FAILED',
  CONNECTION_LOST = 'CONNECTION_LOST',
  TIMEOUT = 'TIMEOUT',
  
  // Excel errors
  EXCEL_NOT_READY = 'EXCEL_NOT_READY',
  INVALID_RANGE = 'INVALID_RANGE',
  WORKSHEET_NOT_FOUND = 'WORKSHEET_NOT_FOUND',
  PROTECTED_RANGE = 'PROTECTED_RANGE',
  
  // Operation errors
  OPERATION_FAILED = 'OPERATION_FAILED',
  UNDO_FAILED = 'UNDO_FAILED',
  REDO_FAILED = 'REDO_FAILED',
  
  // AI errors
  AI_SERVICE_ERROR = 'AI_SERVICE_ERROR',
  INTENT_UNCLEAR = 'INTENT_UNCLEAR',
  INSUFFICIENT_CONTEXT = 'INSUFFICIENT_CONTEXT'
}

// ===== Event Types =====

export interface MCPEvent {
  type: MCPEventType;
  timestamp: Date;
  data: any;
}

export type MCPEventType = 
  | 'connection-established'
  | 'connection-lost'
  | 'tool-executed'
  | 'resource-accessed'
  | 'notification-received'
  | 'error-occurred'
  | 'context-updated'
  | 'excel-changed';

// ===== Configuration Types =====

export interface MCPConfig {
  servers: MCPServerConfig[];
  client: MCPClientConfig;
  transport: MCPTransportConfig;
  logging: LoggingConfig;
}

export interface MCPServerConfig {
  name: string;
  type: 'excel' | 'ai' | 'file';
  transport: 'stdio' | 'http';
  endpoint?: string;
  capabilities: MCPCapabilities;
  autoConnect: boolean;
}

export interface MCPClientConfig {
  name: string;
  version: string;
  capabilities: MCPCapabilities;
  maxSessions: number;
  sessionTimeout: number;
}

export interface MCPTransportConfig {
  type: 'stdio' | 'http';
  timeout: number;
  retryAttempts: number;
  retryDelay: number;
}

export interface LoggingConfig {
  level: 'debug' | 'info' | 'warn' | 'error';
  logToConsole: boolean;
  logToFile: boolean;
  maxLogSize: number;
}