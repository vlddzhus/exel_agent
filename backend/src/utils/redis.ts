import Redis from "ioredis";

let client: Redis | null = null;
let enabled = false;

export function getRedis(): Redis | null {
  if (!enabled) {
    const url = process.env.REDIS_URL;
    if (url) {
      client = new Redis(url, {
        maxRetriesPerRequest: 1,
        retryStrategy() {
          return null;
        },
        lazyConnect: true,
      });
      enabled = true;
    }
  }
  return client;
}

export async function closeRedis(): Promise<void> {
  if (client) {
    await client.quit();
    client = null;
    enabled = false;
  }
}
