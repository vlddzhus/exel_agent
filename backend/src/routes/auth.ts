import { Router, Request, Response } from "express";
import { z } from "zod";
import {
  signAccessToken,
  signRefreshToken,
  verifyToken,
  comparePassword,
} from "../utils/auth";
import * as userRepo from "../db/repositories/user-repo";
import { v4 as uuid } from "uuid";

const router = Router();

const registerSchema = z.object({
  email: z.string().email(),
  password: z.string().min(8).max(128),
});

const loginSchema = z.object({
  email: z.string().email(),
  password: z.string(),
});

const refreshSchema = z.object({
  refreshToken: z.string().min(1),
});

router.post("/register", async (req: Request, res: Response) => {
  const parsed = registerSchema.safeParse(req.body);
  if (!parsed.success) {
    res
      .status(400)
      .json({ error: parsed.error.flatten(), code: "VALIDATION_ERROR" });
    return;
  }

  const { email, password } = parsed.data;

  const existing = await userRepo.findByEmail(email);
  if (existing) {
    res
      .status(409)
      .json({ error: "Email already registered", code: "EMAIL_EXISTS" });
    return;
  }

  const user = await userRepo.createUser(email, password);

  const payload = { userId: user.id, email: user.email, tier: user.tier };
  const accessToken = signAccessToken(payload);
  const refreshToken = signRefreshToken(payload);

  res.status(201).json({
    user: { id: user.id, email: user.email, tier: user.tier },
    accessToken,
    refreshToken,
  });
});

router.post("/login", async (req: Request, res: Response) => {
  const parsed = loginSchema.safeParse(req.body);
  if (!parsed.success) {
    res
      .status(400)
      .json({ error: parsed.error.flatten(), code: "VALIDATION_ERROR" });
    return;
  }

  const { email, password } = parsed.data;
  const user = await userRepo.findByEmail(email);

  if (!user || !(await comparePassword(password, user.password))) {
    res
      .status(401)
      .json({
        error: "Invalid email or password",
        code: "INVALID_CREDENTIALS",
      });
    return;
  }

  const payload = { userId: user.id, email: user.email, tier: user.tier };
  const accessToken = signAccessToken(payload);
  const refreshToken = signRefreshToken(payload);

  res.json({
    user: { id: user.id, email: user.email, tier: user.tier },
    accessToken,
    refreshToken,
  });
});

router.post("/refresh", async (req: Request, res: Response) => {
  const parsed = refreshSchema.safeParse(req.body);
  if (!parsed.success) {
    res
      .status(400)
      .json({ error: parsed.error.flatten(), code: "VALIDATION_ERROR" });
    return;
  }

  try {
    const payload = verifyToken(parsed.data.refreshToken);
    const user = await userRepo.findById(payload.userId);
    if (!user) {
      res.status(401).json({ error: "User not found", code: "USER_NOT_FOUND" });
      return;
    }

    const newPayload = { userId: user.id, email: user.email, tier: user.tier };
    const accessToken = signAccessToken(newPayload);

    res.json({ accessToken });
  } catch {
    res.status(401).json({
      error: "Invalid or expired refresh token",
      code: "INVALID_REFRESH_TOKEN",
    });
  }
});

export default router;
