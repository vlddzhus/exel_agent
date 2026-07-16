import { getProviderEntries, ProviderEntry } from "./provider-factory";

export const TaskType = {
  SIMPLE: "simple" as const,
  COMPLEX: "complex" as const,
};
export type TaskType = (typeof TaskType)[keyof typeof TaskType];
export type UserTier = "free" | "pro" | "team";

const MODEL_MAP: Record<UserTier, Record<TaskType, string>> = {
  free: {
    simple: "meta-llama/llama-4-scout-17b-16e-instruct",
    complex: "meta-llama/llama-4-scout-17b-16e-instruct",
  },
  pro: {
    simple: "meta-llama/llama-4-scout-17b-16e-instruct",
    complex: "meta-llama/llama-4-scout-17b-16e-instruct",
  },
  team: {
    simple: "meta-llama/llama-4-scout-17b-16e-instruct",
    complex: "meta-llama/llama-4-scout-17b-16e-instruct",
  },
};

const PRIORITY_MAP: Record<string, number> = {
  groq: 0,
  openmodel: 10,
  openai: 20,
  anthropic: 30,
};

export function pickModelName(tier: UserTier, taskType: TaskType): string {
  return MODEL_MAP[tier]?.[taskType] ?? MODEL_MAP.free.simple;
}

export function getFilteredProviderChain(
  tier: UserTier,
  taskType: TaskType,
): ProviderEntry[] {
  const targetModel = pickModelName(tier, taskType);
  const allProviders = getProviderEntries();
  const matched = allProviders.filter((p) => p.defaultModel === targetModel);

  if (matched.length > 0) return matched;

  return allProviders.sort(
    (a, b) => (PRIORITY_MAP[a.name] ?? 99) - (PRIORITY_MAP[b.name] ?? 99),
  );
}
