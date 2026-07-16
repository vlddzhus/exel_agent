// Ambient-расширение Express.Request полями userId/tier, проставляемыми
// middleware/auth.ts. Используется во всех route-хендлерах.
// Формат `declare global` гарантирует применение и через `tsc`, и через `ts-node`.
export {};

declare global {
  namespace Express {
    interface Request {
      userId?: string;
      tier?: string;
    }
  }
}
