/**
 * @deprecated Replaced by llm-client.ts.
 * GroqClient has been replaced by LlmClient which routes through our backend proxy.
 * See llm-client.ts for the new implementation.
 */
export { LlmClient as GroqClient, MODEL_CHAIN } from './llm-client';
export type {
  LlmConfig as GroqConfig,
  ChatMessage,
  ToolCall,
  ToolDefinition,
  LlmResponse as GroqResponse,
} from './llm-client';
