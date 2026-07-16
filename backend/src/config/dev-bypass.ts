/**
 * dev-bypass.ts — единая точка правды для обхода платных проверок в локальном режиме.
 *
 * Принцип (AGENTS.md / docs): боевой код биллинга НЕ удаляется. В режимах
 * `development` / `test` все гейты (тариф, дневной лимит, подписка, auth)
 * прозрачно обходятся, чтобы владелец проекта мог вручную протестировать все
 * 24 инструмента без платежа. В `production` обход полностью выключен.
 *
 * См. docs/07-PAYMENTS-SPEC.md — боевые проверки остаются нетронутыми.
 */

const BYPASSED_ENVS = new Set(["development", "test"]);

/** Активен ли обход тарифных/лимитных гейтов. Только для dev/test, никогда в production. */
export function isDevBypassEnabled(): boolean {
  return BYPASSED_ENVS.has(process.env.NODE_ENV ?? "development");
}

/**
 * Активен ли обход JWT-авторизации. ТОЛЬКО `development` — в `test`
 * security-boundary тесты (401 без токена) обязаны проходить и проверять
 * реальный auth. В `production` auth работает по-боевому.
 */
export function isAuthBypassed(): boolean {
  return process.env.NODE_ENV === "development";
}

/** Локальный пользователь по умолчанию, когда JWT отсутствует. */
export const DEV_USER_ID = "dev-local-user";

/**
 * Самый высокий тариф в системе (docs/01-PRODUCT-SPEC.md: Team 799₽/мес,
 * 2000 задач/день). В dev каждый запрос получает максимум возможностей.
 */
export const DEV_TIER = "team" as const;

/** Лимит запросов в день, недостижимый при ручном тестировании. */
export const DEV_DAILY_LIMIT = Number.MAX_SAFE_INTEGER;

/**
 * Возвращает эффективный тариф пользователя.
 * В dev всегда `team`; в prod — реальный тариф из JWT/БД.
 */
export function resolveTier(tier: string | undefined): string {
  if (isDevBypassEnabled()) return DEV_TIER;
  return tier || "free";
}

/**
 * Возвращает эффективный userId.
 * В dev анонимные запросы маппятся на `dev-local-user`, чтобы у каждого
 * запроса был стабильный идентификатор (для usage/логов). Реальный
 * JWT-userId в dev сохраняется, если токен присутствует.
 */
export function resolveUserId(
  userId: string | undefined,
  fallback = DEV_USER_ID,
): string {
  if (isDevBypassEnabled()) {
    if (userId && userId !== "anonymous") return userId;
    return fallback;
  }
  return userId || "anonymous";
}

/** Должен ли дневной лимит считаться исчерпанным. В dev — никогда. */
export function isUsageBlocked(used: number, limit: number): boolean {
  if (isDevBypassEnabled()) return false;
  return used >= limit;
}
