import rateLimit from "express-rate-limit";
import { getRedis } from "../utils/redis";

const WINDOW_MS = parseInt(process.env.RATE_LIMIT_WINDOW_MS || "60000", 10);
const MAX = parseInt(process.env.RATE_LIMIT_MAX || "30", 10);
const MESSAGE = {
  error: "Too many requests. Please try again later.",
  code: "RATE_LIMIT_EXCEEDED",
};

// In-memory fallback store
const hits = new Map<string, { count: number; resetAt: number }>();

function createStore() {
  return {
    async increment(key: string) {
      const redis = getRedis();
      if (redis) {
        const now = Date.now();
        const windowKey = `ratelimit:${key}:${Math.floor(now / WINDOW_MS)}`;
        const count = await redis.incr(windowKey);
        if (count === 1) await redis.pexpire(windowKey, WINDOW_MS);
        return { totalHits: count, resetTime: new Date(now + WINDOW_MS) };
      }
      // ponytail: in-memory fallback, replace with Redis-only when infra is ready
      const now = Date.now();
      const entry = hits.get(key);
      if (!entry || now > entry.resetAt) {
        hits.set(key, { count: 1, resetAt: now + WINDOW_MS });
        return { totalHits: 1, resetTime: new Date(now + WINDOW_MS) };
      }
      entry.count++;
      return {
        totalHits: entry.count,
        resetTime: new Date(entry.resetAt),
      };
    },
    async decrement(key: string) {
      const redis = getRedis();
      if (redis) {
        const windowKey = `ratelimit:${key}:${Math.floor(Date.now() / WINDOW_MS)}`;
        await redis.decr(windowKey);
        return;
      }
      const entry = hits.get(key);
      if (entry) entry.count = Math.max(0, entry.count - 1);
    },
    async resetKey(key: string) {
      const redis = getRedis();
      if (redis) {
        const windowKey = `ratelimit:${key}:${Math.floor(Date.now() / WINDOW_MS)}`;
        await redis.del(windowKey);
        return;
      }
      hits.delete(key);
    },
    async resetAll() {
      const redis = getRedis();
      if (redis) {
        const keys = await redis.keys("ratelimit:*");
        if (keys.length > 0) await redis.del(...keys);
        return;
      }
      hits.clear();
    },
  };
}

export const publicRateLimiter = rateLimit({
  windowMs: WINDOW_MS,
  max: MAX,
  standardHeaders: true,
  legacyHeaders: false,
  message: MESSAGE,
  keyGenerator: (req) => `ip:${req.ip || "unknown"}`,
  store: createStore(),
});

export const userRateLimiter = rateLimit({
  windowMs: WINDOW_MS,
  max: MAX,
  standardHeaders: true,
  legacyHeaders: false,
  message: MESSAGE,
  keyGenerator: (req) => {
    if (!req.userId) throw new Error("userRateLimiter requires auth first");
    return `user:${req.userId}`;
  },
  store: createStore(),
});
