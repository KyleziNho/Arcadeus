/**
 * MCP JSON-RPC Schemas
 * Defines the structure of MCP protocol messages
 */

// ===== Base JSON-RPC Types =====

export interface JSONRPCRequest {
  jsonrpc: '2.0';
  id?: string | number;
  method: string;
  params?: any;
}

export interface JSONRPCResponse {
  jsonrpc: '2.0';
  id: string | number;
  result?: any;
  error?: JSONRPCError;
}

export interface JSONRPCNotification {
  jsonrpc: '2.0';
  method: string;
  params?: any;
}

export interface JSONRPCError {
  code: number;
  message: string;
  data?: any;
}

// ===== MCP Protocol Messages =====

// Initialize
export interface InitializeRequest extends JSONRPCRequest {
  method: 'initialize';
  params: {
    protocolVersion: string;
    capabilities: any;
    clientInfo: {
      name: string;
      version: string;
    };
  };
}

export interface InitializeResponse extends JSONRPCResponse {
  result: {
    protocolVersion: string;
    capabilities: any;
    serverInfo: {
      name: string;
      version: string;
    };
  };
}

// Tools
export interface ListToolsRequest extends JSONRPCRequest {
  method: 'tools/list';
}

export interface ListToolsResponse extends JSONRPCResponse {
  result: {
    tools: Array<{
      name: string;
      title?: string;
      description: string;
      inputSchema: any;
    }>;
  };
}

export interface CallToolRequest extends JSONRPCRequest {
  method: 'tools/call';
  params: {
    name: string;
    arguments: Record<string, any>;
  };
}

export interface CallToolResponse extends JSONRPCResponse {
  result: {
    content: Array<{
      type: string;
      text?: string;
      data?: string;
      mimeType?: string;
    }>;
  };
}

// Resources
export interface ListResourcesRequest extends JSONRPCRequest {
  method: 'resources/list';
}

export interface ListResourcesResponse extends JSONRPCResponse {
  result: {
    resources: Array<{
      uri: string;
      name: string;
      description?: string;
      mimeType?: string;
    }>;
  };
}

export interface ReadResourceRequest extends JSONRPCRequest {
  method: 'resources/read';
  params: {
    uri: string;
  };
}

export interface ReadResourceResponse extends JSONRPCResponse {
  result: {
    content: Array<{
      type: string;
      text?: string;
      data?: string;
      mimeType?: string;
    }>;
  };
}

// Prompts
export interface ListPromptsRequest extends JSONRPCRequest {
  method: 'prompts/list';
}

export interface ListPromptsResponse extends JSONRPCResponse {
  result: {
    prompts: Array<{
      name: string;
      description?: string;
      arguments?: Array<{
        name: string;
        description?: string;
        required?: boolean;
      }>;
    }>;
  };
}

export interface GetPromptRequest extends JSONRPCRequest {
  method: 'prompts/get';
  params: {
    name: string;
    arguments?: Record<string, any>;
  };
}

export interface GetPromptResponse extends JSONRPCResponse {
  result: {
    messages: Array<{
      role: 'user' | 'assistant' | 'system';
      content: {
        type: string;
        text?: string;
      };
    }>;
  };
}

// Notifications
export interface ToolsListChangedNotification extends JSONRPCNotification {
  method: 'notifications/tools/list_changed';
}

export interface ResourcesListChangedNotification extends JSONRPCNotification {
  method: 'notifications/resources/list_changed';
}

export interface PromptsListChangedNotification extends JSONRPCNotification {
  method: 'notifications/prompts/list_changed';
}

// Progress
export interface ProgressNotification extends JSONRPCNotification {
  method: 'notifications/progress';
  params: {
    progressToken: string;
    progress: number;
    total?: number;
  };
}

// Logging
export interface LoggingMessageNotification extends JSONRPCNotification {
  method: 'notifications/message';
  params: {
    level: 'debug' | 'info' | 'warning' | 'error';
    logger?: string;
    data: any;
  };
}

// ===== Excel-Specific Protocol Extensions =====

export interface ExcelSelectionChangedNotification extends JSONRPCNotification {
  method: 'notifications/excel/selection_changed';
  params: {
    worksheet: string;
    range: string;
    values?: any[][];
  };
}

export interface ExcelDataChangedNotification extends JSONRPCNotification {
  method: 'notifications/excel/data_changed';
  params: {
    worksheet: string;
    range: string;
    changeType: 'value' | 'formula' | 'format';
    previousValues?: any[][];
    newValues?: any[][];
  };
}

export interface ExcelWorksheetChangedNotification extends JSONRPCNotification {
  method: 'notifications/excel/worksheet_changed';
  params: {
    changeType: 'added' | 'deleted' | 'renamed' | 'moved';
    worksheet: string;
    newName?: string;
    newIndex?: number;
  };
}

// ===== Validation Schemas =====

export const PROTOCOL_VERSION = '2025-06-18';

export function validateRequest(request: any): boolean {
  return (
    request &&
    request.jsonrpc === '2.0' &&
    typeof request.method === 'string'
  );
}

export function validateResponse(response: any): boolean {
  return (
    response &&
    response.jsonrpc === '2.0' &&
    (response.id !== undefined) &&
    (response.result !== undefined || response.error !== undefined)
  );
}

export function validateNotification(notification: any): boolean {
  return (
    notification &&
    notification.jsonrpc === '2.0' &&
    typeof notification.method === 'string' &&
    notification.id === undefined
  );
}

// ===== Message Factory Functions =====

export function createRequest(
  method: string,
  params?: any,
  id?: string | number
): JSONRPCRequest {
  return {
    jsonrpc: '2.0',
    id: id || generateId(),
    method,
    params
  };
}

export function createResponse(
  id: string | number,
  result?: any,
  error?: JSONRPCError
): JSONRPCResponse {
  const response: JSONRPCResponse = {
    jsonrpc: '2.0',
    id
  };
  
  if (error) {
    response.error = error;
  } else {
    response.result = result;
  }
  
  return response;
}

export function createNotification(
  method: string,
  params?: any
): JSONRPCNotification {
  return {
    jsonrpc: '2.0',
    method,
    params
  };
}

export function createError(
  code: number,
  message: string,
  data?: any
): JSONRPCError {
  return {
    code,
    message,
    data
  };
}

// ===== Error Codes =====

export const ErrorCodes = {
  // JSON-RPC 2.0 standard error codes
  PARSE_ERROR: -32700,
  INVALID_REQUEST: -32600,
  METHOD_NOT_FOUND: -32601,
  INVALID_PARAMS: -32602,
  INTERNAL_ERROR: -32603,
  
  // MCP-specific error codes
  PROTOCOL_ERROR: -32000,
  NOT_INITIALIZED: -32001,
  CAPABILITY_NOT_SUPPORTED: -32002,
  RESOURCE_NOT_FOUND: -32003,
  TOOL_NOT_FOUND: -32004,
  PROMPT_NOT_FOUND: -32005,
  
  // Excel-specific error codes
  EXCEL_ERROR: -32100,
  RANGE_ERROR: -32101,
  WORKSHEET_ERROR: -32102,
  FORMULA_ERROR: -32103,
  PROTECTED_ERROR: -32104
};

// ===== Helper Functions =====

let requestIdCounter = 0;

export function generateId(): string {
  return `req_${Date.now()}_${++requestIdCounter}`;
}

export function isRequest(message: any): message is JSONRPCRequest {
  return validateRequest(message) && message.id !== undefined;
}

export function isNotification(message: any): message is JSONRPCNotification {
  return validateNotification(message);
}

export function isResponse(message: any): message is JSONRPCResponse {
  return validateResponse(message);
}