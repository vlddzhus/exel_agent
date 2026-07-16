import { create } from "zustand";

export interface PlanStep {
  id: string;
  label: string;
  status: "pending" | "current" | "done";
  durationMs?: number;
}

export interface ThoughtLine {
  text: string;
  timestamp: number;
}

export interface LiveStats {
  tokensIn: number;
  tokensOut: number;
  elapsedMs: number;
  provider: string;
}

type LiveStatus =
  "idle" | "thinking" | "executing" | "done" | "error" | "cancelled";

let timerInterval: ReturnType<typeof setInterval> | null = null;

interface LiveActivityState {
  status: LiveStatus;
  plan: PlanStep[];
  thoughts: ThoughtLine[];
  progress: { done: number; total: number; currentLabel: string };
  stats: LiveStats;
  error: { code: string; message: string } | null;
  hasChanges: boolean;
  startTime: number;
  start: () => void;
  setChanges: () => void;
  setPlan: (steps: PlanStep[]) => void;
  addThought: (text: string) => void;
  updateProgress: (done: number, total: number, label: string) => void;
  markStepDone: (id: string, durationMs: number) => void;
  setStats: (stats: Partial<LiveStats>) => void;
  finish: () => void;
  fail: (code: string, message: string) => void;
  cancel: () => void;
  reset: () => void;
}

export const useLiveActivityStore = create<LiveActivityState>((set, get) => ({
  status: "idle",
  plan: [],
  thoughts: [],
  progress: { done: 0, total: 0, currentLabel: "" },
  stats: { tokensIn: 0, tokensOut: 0, elapsedMs: 0, provider: "" },
  error: null,
  hasChanges: false,
  startTime: 0,
  start: () => {
    const t0 = Date.now();
    if (timerInterval) clearInterval(timerInterval);
    timerInterval = setInterval(() => {
      set((s) => ({
        stats: { ...s.stats, elapsedMs: Date.now() - t0 },
      }));
    }, 1000);
    set({
      status: "thinking",
      thoughts: [],
      error: null,
      startTime: t0,
      stats: { tokensIn: 0, tokensOut: 0, elapsedMs: 0, provider: "" },
      progress: { done: 0, total: 0, currentLabel: "" },
    });
  },
  setPlan: (steps) =>
    set({
      plan: steps,
      progress: { done: 0, total: steps.length, currentLabel: "" },
    }),
  addThought: (text) =>
    set((s) => ({
      thoughts: [...s.thoughts, { text, timestamp: Date.now() }].slice(-50),
    })),
  setChanges: () => set({ hasChanges: true }),
  updateProgress: (done, total, label) =>
    set({ progress: { done, total, currentLabel: label } }),
  markStepDone: (id, durationMs) =>
    set((s) => ({
      plan: s.plan.map((step) =>
        step.id === id
          ? { ...step, status: "done" as const, durationMs }
          : step,
      ),
      progress: { ...s.progress, done: s.progress.done + 1 },
    })),
  setStats: (partial) => set((s) => ({ stats: { ...s.stats, ...partial } })),
  finish: () => {
    if (timerInterval) clearInterval(timerInterval);
    timerInterval = null;
    set({ status: "done" });
  },
  fail: (code, message) => {
    if (timerInterval) clearInterval(timerInterval);
    timerInterval = null;
    set({ status: "error", error: { code, message } });
  },
  cancel: () => {
    if (timerInterval) clearInterval(timerInterval);
    timerInterval = null;
    set({
      status: "cancelled",
      hasChanges: false,
      thoughts: [],
      plan: [],
      progress: { done: 0, total: 0, currentLabel: "" },
      error: null,
    });
  },
  reset: () => {
    if (timerInterval) clearInterval(timerInterval);
    timerInterval = null;
    set({
      status: "idle",
      plan: [],
      thoughts: [],
      progress: { done: 0, total: 0, currentLabel: "" },
      stats: { tokensIn: 0, tokensOut: 0, elapsedMs: 0, provider: "" },
      error: null,
      hasChanges: false,
      startTime: 0,
    });
  },
}));
