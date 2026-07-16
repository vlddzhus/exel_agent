import { create } from "zustand";

interface UserState {
  email: string;
  tier: "free" | "pro" | "team";
  usage: number;
  limit: number;
  setUser: (u: {
    email: string;
    tier: "free" | "pro" | "team";
    usage: number;
    limit: number;
  }) => void;
  setTier: (tier: "free" | "pro" | "team") => void;
  setUsage: (usage: number) => void;
}

export const useUserStore = create<UserState>((set) => ({
  userId: "",
  email: "",
  tier: "free",
  usage: 0,
  limit: 10,
  setUser: (u) => set(u),
  setTier: (tier) => set({ tier }),
  setUsage: (usage) => set({ usage }),
}));
