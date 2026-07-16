import { create } from "zustand";
import {
  ChatSession,
  getSessionsAsync,
  deleteSessionAsync,
  loadSessionAsync,
} from "../utils/session-store";

interface SessionState {
  sessions: ChatSession[];
  loading: boolean;
  load: () => Promise<void>;
  remove: (id: string) => Promise<void>;
  getById: (id: string) => ChatSession | undefined;
}

export const useSessionStore = create<SessionState>((set, get) => ({
  sessions: [],
  loading: false,
  load: async () => {
    set({ loading: true });
    const sessions = await getSessionsAsync();
    set({ sessions, loading: false });
  },
  remove: async (id: string) => {
    await deleteSessionAsync(id);
    set((s) => ({ sessions: s.sessions.filter((x) => x.id !== id) }));
  },
  getById: (id: string) => get().sessions.find((s) => s.id === id),
}));
