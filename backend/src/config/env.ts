import { z } from "zod";
import { randomBytes } from "crypto";

const PLACEHOLDER_JWT_SECRET = "CHANGE-ME-USE-RANDOM-64-CHARS";

/**
 * Dev-режим: если JWT_SECRET не задан или это плейсхолдер, авто-генерируем
 * стабильный ключ на процесс. Auth в dev и так обойдён (см. dev-bypass.ts),
 * но схема требует непустое значение ≥32 символов. В prod валидация строгая.
 */
function resolveJwtSecret(): string {
  const raw = process.env.JWT_SECRET;
  const isDev =
    process.env.NODE_ENV === "development" ||
    process.env.NODE_ENV === undefined ||
    process.env.NODE_ENV === "test";
  if (!raw || raw === PLACEHOLDER_JWT_SECRET || raw.length < 32) {
    if (isDev) return randomBytes(48).toString("hex");
    // В prod дойдёт до zod-валидации ниже и упадёт с понятной ошибкой.
    return raw ?? "";
  }
  return raw;
}

// Заполняем до zod-валидации, чтобы dev-автогенерация прошла через схему.
process.env.JWT_SECRET = resolveJwtSecret();

const EnvSchema = z.object({
  NODE_ENV: z
    .enum(["development", "test", "production"])
    .default("development"),
  LOG_LEVEL: z
    .enum(["fatal", "error", "warn", "info", "debug", "trace", "silent"])
    .default("info"),
  PORT: z.coerce.number().int().min(1).max(65535).default(4000),
  ALLOWED_ORIGINS: z
    .string()
    .default("https://localhost:3000")
    .transform((value) =>
      value
        .split(",")
        .map((origin) => origin.trim())
        .filter(Boolean),
    )
    .pipe(z.array(z.string().url()).min(1)),
  JWT_SECRET: z
    .string()
    .min(32, "JWT_SECRET must be at least 32 characters")
    .refine(
      (value) => value !== PLACEHOLDER_JWT_SECRET,
      "JWT_SECRET must not use the example placeholder",
    ),
  // DATABASE_URL обязательна в production; в dev/test опциональна — запуск без
  // Postgres возможен для smoke-тестов (usage-store падает мягко через try/catch).
  DATABASE_URL: z.string().url("DATABASE_URL must be a valid PostgreSQL URL"),
  REDIS_URL: z.string().url().optional(),
  RATE_LIMIT_WINDOW_MS: z.coerce.number().int().positive().default(60_000),
  RATE_LIMIT_MAX: z.coerce.number().int().positive().default(30),
  YOOKASSA_SHOP_ID: z.string().optional(),
  YOOKASSA_SECRET_KEY: z.string().optional(),
});

export type Env = z.infer<typeof EnvSchema>;

export const env: Env = (() => {
  const result = EnvSchema.safeParse(process.env);

  if (result.success) return result.data;

  const isDev =
    process.env.NODE_ENV === "development" ||
    process.env.NODE_ENV === undefined ||
    process.env.NODE_ENV === "test";

  // В dev смягчаем только DATABASE_URL: если Postgres не поднят, backend
  // стартует, а usage-запросы падают мягко. Остальные ошибки — реальные.
  if (isDev) {
    const onlyDbMissing = result.error.issues.every(
      (i) => i.path.includes("DATABASE_URL"),
    );
    if (onlyDbMissing) {
      const merged = {
        ...process.env,
        NODE_ENV: process.env.NODE_ENV ?? "development",
        LOG_LEVEL: process.env.LOG_LEVEL ?? "info",
        PORT: Number(process.env.PORT ?? 4000),
        ALLOWED_ORIGINS: ["https://localhost:3000"],
      } as unknown as Env;
      return merged;
    }
  }

  const details = result.error.issues
    .map((issue) => `${issue.path.join(".")}: ${issue.message}`)
    .join("; ");

  throw new Error(`Invalid backend environment: ${details}`);
})();
