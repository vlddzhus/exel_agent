import { Request, Response, NextFunction } from "express";
import { verifyToken } from "../utils/auth";
import {
  isAuthBypassed,
  resolveTier,
  resolveUserId,
  DEV_TIER,
} from "../config/dev-bypass";

export function validateApiKey(
  req: Request,
  res: Response,
  next: NextFunction,
): void {
  // ── Dev bypass: фронтенд надстройки работает без JWT в `development`,
  // чтобы владелец мог тестировать все 24 инструмента без регистрации.
  // В `test`/`production` боевой auth ниже работает по-настоящему.
  if (isAuthBypassed()) {
    req.userId = resolveUserId(undefined);
    req.tier = DEV_TIER;
    next();
    return;
  }

  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    res.status(401).json({
      error: "Missing or malformed Authorization header",
      code: "AUTH_MISSING_TOKEN",
    });
    return;
  }

  const token = authHeader.slice(7);
  try {
    const payload = verifyToken(token);
    req.userId = payload.userId;
    req.tier = resolveTier(payload.tier);
    next();
  } catch {
    res
      .status(401)
      .json({ error: "Invalid or expired token", code: "AUTH_INVALID_TOKEN" });
  }
}

export function optionalAuth(
  req: Request,
  _res: Response,
  next: NextFunction,
): void {
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    req.userId = "anonymous";
    next();
    return;
  }

  const token = authHeader.slice(7);
  try {
    const payload = verifyToken(token);
    req.userId = payload.userId;
  } catch {
    req.userId = "anonymous";
  }
  next();
}
