import dotenv from "dotenv";
import path from "path";
import os from "os";

dotenv.config({ path: path.resolve(__dirname, "../.env") });

import express from "express";
import cors from "cors";
import helmet from "helmet";
import { env } from "./config/env";
import { publicRateLimiter, userRateLimiter } from "./middleware/rate-limit";
import { validateApiKey } from "./middleware/auth";
import { sessionStore } from "./session-store";
import { getModelChain } from "./utils/provider-factory";
import authRoutes from "./routes/auth";
import agentRoutes from "./routes/agent";
import { getDb, closeDb } from "./db/kysely";
import { runMigrations } from "./db/migrate";
import { closeRedis } from "./utils/redis";
import billingRoutes from "./billing/routes/billing";
import { SubscriptionService } from "./billing/subscription.service";
import { YooKassaAdapter } from "./billing/yookassa.adapter";
import { getUsageSummary } from "./utils/usage-store";
import pino from "pino";
import fs from "fs";

const logger = pino({ level: env.LOG_LEVEL });

const app = express();
const PORT = env.PORT;

app.set("trust proxy", 1);
app.use(helmet());
app.use(
  cors({
    origin: env.ALLOWED_ORIGINS,
    methods: ["GET", "POST", "PATCH", "DELETE", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
  }),
);
app.use(express.json({ limit: "1mb" }));

app.use("/api/auth", publicRateLimiter, authRoutes);
app.use("/api/billing", billingRoutes);
app.use("/api/agent", validateApiKey, userRateLimiter, agentRoutes);

// ── Feedback store ──

interface FeedbackEntry {
  id: string;
  timestamp: string;
  userId: string;
  messageId?: string;
  rating: "helpful" | "unhelpful";
  comment?: string;
}
const feedbackStore: FeedbackEntry[] = [];

// ── Routes ──

app.get("/api/health", (_req, res) => {
  res.json({
    status: "ok",
    timestamp: new Date().toISOString(),
    providers: getModelChain(),
  });
});

app.get("/api/providers", (_req, res) => {
  res.json({ providers: getModelChain() });
});

// ── Feedback ──

app.post("/api/feedback", validateApiKey, async (req, res) => {
  const { messageId, rating, comment } = req.body;
  const userId = req.userId || "anonymous";

  if (!rating || !["helpful", "unhelpful"].includes(rating)) {
    res.status(400).json({
      error: 'rating must be "helpful" or "unhelpful"',
      code: "INVALID_RATING",
    });
    return;
  }

  const entry: FeedbackEntry = {
    id: `fb_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
    timestamp: new Date().toISOString(),
    userId,
    messageId: messageId || undefined,
    rating,
    comment: comment || undefined,
  };

  feedbackStore.push(entry);

  try {
    const feedbackPath = path.join(__dirname, "../../data/feedback.json");
    const dir = path.dirname(feedbackPath);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

    let existing: FeedbackEntry[] = [];
    try {
      existing = JSON.parse(fs.readFileSync(feedbackPath, "utf-8"));
    } catch {}
    existing.push(entry);
    fs.writeFileSync(feedbackPath, JSON.stringify(existing, null, 2));
  } catch {}

  logger.info(
    { userId, rating, messageId: messageId || undefined },
    "feedback",
  );
  res.json({ status: "ok", id: entry.id });
});

app.get("/api/feedback", validateApiKey, async (req, res) => {
  res.json({
    count: feedbackStore.length,
    feedbacks: feedbackStore.slice(-100),
  });
});

// ── Usage ──

app.get("/api/usage/:userId", validateApiKey, async (req, res) => {
  const requestedId = req.params.userId;
  const userId = req.userId || "anonymous";
  if (requestedId !== userId) {
    res.status(403).json({ error: "Forbidden", code: "FORBIDDEN" });
    return;
  }

  const usage = await getUsageSummary(userId);
  res.json({
    userId,
    totalRequests: usage.count,
    totalTokens: usage.totalTokens,
    totalCost: usage.cost,
    recentRequests: usage.recentRequests,
  });
});

// ── Session Store ──

app.get("/api/sessions", validateApiKey, (_req, res) => {
  res.json({ sessions: sessionStore.getAllMeta() });
});

app.get("/api/sessions/:id", validateApiKey, (req, res) => {
  const id = req.params.id as string;
  const session = sessionStore.getById(id);
  if (!session) {
    res.status(404).json({ error: "Session not found" });
    return;
  }
  res.json({ session });
});

app.post("/api/sessions", validateApiKey, (req, res) => {
  const session = req.body.session;
  if (!session || !session.id) {
    res.status(400).json({ error: "Session with id is required" });
    return;
  }
  sessionStore.save(session);
  res.json({ status: "ok" });
});

app.delete("/api/sessions/:id", validateApiKey, (req, res) => {
  const id = req.params.id as string;
  const deleted = sessionStore.delete(id);
  res.json({ status: deleted ? "deleted" : "not_found" });
});

app.patch("/api/sessions/:id/rename", validateApiKey, (req, res) => {
  const id = req.params.id as string;
  const { title } = req.body;
  if (!title) {
    res.status(400).json({ error: "title is required" });
    return;
  }
  const ok = sessionStore.rename(id, title);
  res.json({ status: ok ? "renamed" : "not_found" });
});

// ── Recurring billing cron ──
function startBillingCron(): void {
  const intervalMs = 3600000;
  setInterval(async () => {
    try {
      const shopId = env.YOOKASSA_SHOP_ID || "";
      const secretKey = env.YOOKASSA_SECRET_KEY || "";
      if (!shopId || !secretKey) return;
      const gateway = new YooKassaAdapter(shopId, secretKey);
      const service = new SubscriptionService(gateway, getDb());
      const renewed = await service.renewDueSubscriptions();
      if (renewed > 0) {
        logger.info({ count: renewed }, "recurring billing cron completed");
      }
    } catch (err) {
      logger.error({ error: err }, "billing cron error");
    }
  }, intervalMs);
  logger.info({ intervalMs }, "billing cron started");
}

// ── Start (only when run directly, not in tests) ──
if (!process.env.JEST_WORKER_ID) {
  const isDev = env.NODE_ENV !== "production";

  function startServer() {
    const isSecure = process.env.USE_HTTPS === "true" || isDev;
    if (isSecure) {
      const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
      const keyPath = path.join(certDir, "localhost.key");
      const certPath = path.join(certDir, "localhost.crt");
      if (fs.existsSync(keyPath) && fs.existsSync(certPath)) {
        const https = require("https");
        https
          .createServer(
            { key: fs.readFileSync(keyPath), cert: fs.readFileSync(certPath) },
            app,
          )
          .listen(PORT, () => {
            logger.info(
              { port: PORT, env: env.NODE_ENV, https: true },
              "server started",
            );
          });
        return;
      }
    }
    app.listen(PORT, () => {
      logger.info(
        { port: PORT, env: env.NODE_ENV, https: false },
        "server started",
      );
    });
  }

  runMigrations(getDb())
    .then(() => {
      startServer();
      startBillingCron();
    })
    .catch((err) => {
      // В dev стартуем даже без Postgres: владелец может тестировать LLM-агента
      // без `docker compose up`. Usage-запросы упадут мягко (try/catch в route).
      // В production недоступная БД — критично, падаем.
      if (isDev) {
        logger.warn(
          { error: err.message },
          "DB migration failed — starting anyway in dev mode (run `docker compose up -d` for full functionality)",
        );
        startServer();
      } else {
        logger.error(err, "migration failed");
        process.exit(1);
      }
    });
}

export { app, closeDb, closeRedis };
export type { FeedbackEntry };
