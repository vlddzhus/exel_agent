import { v4 as uuid } from "uuid";
import { getDb } from "../db/kysely";
import { getRedis } from "./redis";

export const TIER_LIMITS: Record<string, number> = {
  free: 10,
  pro: 500,
  team: 2000,
};

export interface DailyUsage {
  count: number;
  date: string;
  inputTokens: number;
  outputTokens: number;
  totalTokens: number;
  cost: string;
  resetDate: string;
}

export interface UsageEventInput {
  provider?: string;
  model?: string;
  inputTokens?: number;
  outputTokens?: number;
  cost?: string;
}

export interface UsageSummary extends DailyUsage {
  recentRequests: Array<{
    timestamp: string;
    userId: string;
    model: string;
    provider: string;
    inputTokens: number;
    outputTokens: number;
    totalTokens: number;
    cost: string;
  }>;
}

const MSK_OFFSET_MS = 3 * 60 * 60 * 1000;
const DAY_MS = 24 * 60 * 60 * 1000;

function getMskWindow(now = new Date()) {
  const shifted = new Date(now.getTime() + MSK_OFFSET_MS);
  const date = shifted.toISOString().slice(0, 10);
  const [year, month, day] = date.split("-").map(Number);
  const startMs = Date.UTC(year, month - 1, day) - MSK_OFFSET_MS;
  const endMs = startMs + DAY_MS;
  return {
    date,
    start: new Date(startMs),
    end: new Date(endMs),
    resetDate: new Date(endMs).toISOString(),
    ttlSeconds: Math.max(1, Math.ceil((endMs - now.getTime()) / 1000)),
  };
}

function redisKey(userId: string, date: string): string {
  return `usage:${date}:${userId}`;
}

function toNumber(value: unknown): number {
  if (typeof value === "number") return value;
  if (typeof value === "bigint") return Number(value);
  if (typeof value === "string") return Number(value) || 0;
  return 0;
}

async function getUsageFromDb(userId: string): Promise<DailyUsage> {
  const window = getMskWindow();
  const row = await getDb()
    .selectFrom("usage_events")
    .select((eb) => [
      eb.fn.countAll().as("count"),
      eb.fn.sum<number>("input_tokens").as("input_tokens"),
      eb.fn.sum<number>("output_tokens").as("output_tokens"),
    ])
    .where("user_id", "=", userId)
    .where("created_at", ">=", window.start)
    .where("created_at", "<", window.end)
    .executeTakeFirst();

  const costRows = await getDb()
    .selectFrom("usage_events")
    .select(["cost"])
    .where("user_id", "=", userId)
    .where("created_at", ">=", window.start)
    .where("created_at", "<", window.end)
    .execute();

  const inputTokens = toNumber(row?.input_tokens);
  const outputTokens = toNumber(row?.output_tokens);
  const cost = costRows.reduce((sum, r) => sum + parseFloat(r.cost || "0"), 0);
  return {
    count: toNumber(row?.count),
    date: window.date,
    inputTokens,
    outputTokens,
    totalTokens: inputTokens + outputTokens,
    cost: cost.toFixed(6),
    resetDate: window.resetDate,
  };
}

export async function getUsage(userId: string): Promise<DailyUsage> {
  const window = getMskWindow();
  const redis = getRedis();

  if (redis) {
    try {
      const count = await redis.get(redisKey(userId, window.date));
      if (count !== null) {
        return {
          count: Number(count) || 0,
          date: window.date,
          inputTokens: 0,
          outputTokens: 0,
          totalTokens: 0,
          cost: "0",
          resetDate: window.resetDate,
        };
      }

      const usage = await getUsageFromDb(userId);
      await redis.set(
        redisKey(userId, window.date),
        usage.count,
        "EX",
        window.ttlSeconds,
      );
      return usage;
    } catch {
      // ponytail: Redis is an acceleration path; DB remains the source of truth.
    }
  }

  try {
    return await getUsageFromDb(userId);
  } catch {
    // Dev-resilience: если Postgres не поднят, возвращаем пустой usage,
    // чтобы LLM-запросы работали без БД (см. server.ts dev-старт).
    return {
      count: 0,
      date: window.date,
      inputTokens: 0,
      outputTokens: 0,
      totalTokens: 0,
      cost: "0",
      resetDate: window.resetDate,
    };
  }
}

export async function incrementUsage(
  userId: string,
  event: UsageEventInput = {},
): Promise<void> {
  const window = getMskWindow();
  const inputTokens = event.inputTokens ?? 0;
  const outputTokens = event.outputTokens ?? 0;

  try {
    await getDb()
      .insertInto("usage_events")
      .values({
        id: uuid(),
        user_id: userId,
        provider: event.provider ?? "unknown",
        model: event.model ?? "unknown",
        input_tokens: inputTokens,
        output_tokens: outputTokens,
        cost: event.cost ?? "0",
      })
      .execute();
  } catch {
    // Dev-resilience: без Postgres учёт usage пропускаем.
  }

  const redis = getRedis();
  if (!redis) return;

  try {
    const key = redisKey(userId, window.date);
    const count = await redis.incr(key);
    if (count === 1) await redis.expire(key, window.ttlSeconds);
  } catch {
    // DB write already succeeded; Redis can warm again on next read.
  }
}

export async function getUsageSummary(userId: string): Promise<UsageSummary> {
  const window = getMskWindow();
  try {
    const usage = await getUsageFromDb(userId);
    const recentRows = await getDb()
      .selectFrom("usage_events")
      .select([
        "user_id",
        "model",
        "provider",
        "input_tokens",
        "output_tokens",
        "cost",
        "created_at",
      ])
      .where("user_id", "=", userId)
      .where("created_at", ">=", window.start)
      .where("created_at", "<", window.end)
      .orderBy("created_at", "desc")
      .limit(50)
      .execute();

    return {
      ...usage,
      recentRequests: recentRows.map((row) => ({
        timestamp: row.created_at.toISOString(),
        userId: row.user_id,
        model: row.model,
        provider: row.provider,
        inputTokens: row.input_tokens,
        outputTokens: row.output_tokens,
        totalTokens: row.input_tokens + row.output_tokens,
        cost: row.cost,
      })),
    };
  } catch {
    // Dev-resilience: без Postgres возвращаем пустую сводку.
    return {
      count: 0,
      date: window.date,
      inputTokens: 0,
      outputTokens: 0,
      totalTokens: 0,
      cost: "0",
      resetDate: window.resetDate,
      recentRequests: [],
    };
  }
}
