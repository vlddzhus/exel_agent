import { getProviderEntries, ProviderEntry } from "./provider-factory";

export const TaskType = {
  SIMPLE: "simple" as const,
  COMPLEX: "complex" as const,
};
export type TaskType = (typeof TaskType)[keyof typeof TaskType];
export type UserTier = "free" | "pro" | "team";

const MODEL_MAP: Record<UserTier, Record<TaskType, string>> = {
  free: { simple: "claude-haiku-4-5", complex: "claude-haiku-4-5" },
  pro: { simple: "claude-haiku-4-5", complex: "gpt-5.2" },
  team: { simple: "claude-haiku-4-5", complex: "gpt-5.2" },
};

const PRIORITY_MAP: Record<string, number> = {
  openmodel: 0,
  groq: 10,
  gigaChat: 20,
  openai: 30,
  anthropic: 40,
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
