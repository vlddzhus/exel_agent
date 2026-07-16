import { create } from "zustand";

export interface Message {
  id: string;
  role: "user" | "assistant" | "system" | "error";
  content: string;
  timestamp: number;
}

interface ChatState {
  messages: Message[];
  isProcessing: boolean;
  addUserMessage: (content: string) => void;
  addAssistantMessage: (content: string) => void;
  addSystemMessage: (content: string) => void;
  addErrorMessage: (content: string) => void;
  setProcessing: (v: boolean) => void;
  loadMessages: (msgs: { role: string; content: string }[]) => void;
  clear: () => void;
}

let idCounter = 0;

export const useChatStore = create<ChatState>((set) => ({
  messages: [],
  isProcessing: false,
  addUserMessage: (content) =>
    set((s) => ({
      messages: [
        ...s.messages,
        {
          id: `m_${++idCounter}`,
          role: "user",
          content,
          timestamp: Date.now(),
        },
      ],
    })),
  addAssistantMessage: (content) =>
    set((s) => ({
      messages: [
        ...s.messages,
        {
          id: `m_${++idCounter}`,
          role: "assistant",
          content,
          timestamp: Date.now(),
        },
      ],
    })),
  addSystemMessage: (content) =>
    set((s) => ({
      messages: [
        ...s.messages,
        {
          id: `m_${++idCounter}`,
          role: "system",
          content,
          timestamp: Date.now(),
        },
      ],
    })),
  addErrorMessage: (content) =>
    set((s) => ({
      messages: [
        ...s.messages,
        {
          id: `m_${++idCounter}`,
          role: "error",
          content,
          timestamp: Date.now(),
        },
      ],
    })),
  setProcessing: (v) => set({ isProcessing: v }),
  loadMessages: (msgs) =>
    set({
      messages: msgs.map((m) => ({
        id: `m_${++idCounter}`,
        role: m.role as Message["role"],
        content: m.content,
        timestamp: Date.now(),
      })),
    }),
  clear: () => set({ messages: [], isProcessing: false }),
}));
