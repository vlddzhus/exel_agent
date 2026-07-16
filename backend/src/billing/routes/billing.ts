import { Router, Request, Response } from "express";
import { z } from "zod";
import { validateApiKey } from "../../middleware/auth";
import { getDb } from "../../db/kysely";
import { YooKassaAdapter } from "../yookassa.adapter";
import { SubscriptionService } from "../subscription.service";
import { env } from "../../config/env";
import pino from "pino";

const router = Router();
const logger = pino({ level: env.LOG_LEVEL });

function getSubService(): SubscriptionService {
  if (!env.YOOKASSA_SHOP_ID || !env.YOOKASSA_SECRET_KEY) {
    throw Object.assign(new Error("YooKassa credentials are not configured"), {
      statusCode: 503,
    });
  }

  const gateway = new YooKassaAdapter(
    env.YOOKASSA_SHOP_ID,
    env.YOOKASSA_SECRET_KEY,
  );
  return new SubscriptionService(gateway, getDb());
}

const subscribeSchema = z.object({
  planId: z.enum(["pro", "team"]),
  period: z.enum(["month", "year"]),
});

router.get("/plans", (_req: Request, res: Response) => {
  const service = getSubService();
  res.json({ plans: service.getPlanList() });
});

router.post(
  "/subscribe",
  validateApiKey,
  async (req: Request, res: Response) => {
    const parsed = subscribeSchema.safeParse(req.body);
    if (!parsed.success) {
      res
        .status(400)
        .json({ error: parsed.error.flatten(), code: "VALIDATION_ERROR" });
      return;
    }

    try {
      const service = getSubService();
      const result = await service.subscribe(
        req.userId!,
        parsed.data.planId,
        parsed.data.period,
      );
      res.json(result);
    } catch (err: unknown) {
      const status = (err as { statusCode?: number }).statusCode ?? 500;
      const message = err instanceof Error ? err.message : "Internal error";
      res.status(status).json({
        error: message,
        code: status === 404 ? "NOT_FOUND" : "PAYMENT_ERROR",
      });
    }
  },
);

router.post("/webhook", async (req: Request, res: Response) => {
  // ponytail: respond fast (200), process async — YooKassa expects quick ack
  res.status(200).json({ status: "ok" });

  try {
    if (!env.YOOKASSA_SHOP_ID || !env.YOOKASSA_SECRET_KEY) {
      throw new Error("YooKassa credentials are not configured");
    }

    const gateway = new YooKassaAdapter(
      env.YOOKASSA_SHOP_ID,
      env.YOOKASSA_SECRET_KEY,
    );
    const event = await gateway.parseWebhook({
      ip: req.ip,
      body: req.body,
    });
    const service = getSubService();
    await service.handleWebhook(event);
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    logger.error({ error: msg }, "webhook processing error");
  }
});

router.post("/cancel", validateApiKey, async (req: Request, res: Response) => {
  try {
    const service = getSubService();
    const result = await service.cancelSubscription(req.userId!);
    res.json(result);
  } catch (err: unknown) {
    const status = (err as { statusCode?: number }).statusCode ?? 500;
    const message = err instanceof Error ? err.message : "Internal error";
    res.status(status).json({
      error: message,
      code: status === 404 ? "NOT_FOUND" : "CANCEL_ERROR",
    });
  }
});

export default router;
