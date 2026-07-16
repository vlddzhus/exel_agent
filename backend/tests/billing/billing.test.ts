import request from "supertest";
import jwt from "jsonwebtoken";
import {
  PaymentGateway,
  SubscriptionPlan,
} from "../../src/billing/gateway.interface";

const testToken = jwt.sign(
  { userId: "test-user", email: "test@example.com", tier: "free" },
  process.env.JWT_SECRET ||
    "test-secret-64-chars-long-at-least-for-hs256-validation",
  { expiresIn: "15m" },
);
const authHeader = { Authorization: `Bearer ${testToken}` };

process.env.NODE_ENV = "test";
process.env.JWT_SECRET =
  "test-secret-64-chars-long-at-least-for-hs256-validation";
process.env.YOOKASSA_SHOP_ID = "test_shop";
process.env.YOOKASSA_SECRET_KEY = "test_secret";

// ── Mocks for modules that server.ts pulls in (ESM deps) ──

jest.mock("ai", () => ({
  streamText: jest.fn(),
  generateText: jest.fn(),
  tool: jest.fn((def: unknown) => def),
  jsonSchema: jest.fn((s: unknown) => s),
}));

jest.mock("../../src/utils/provider-factory", () => ({
  getProviderEntries: jest.fn(() => []),
  getModelChain: jest.fn(() => []),
  convertTools: jest.fn(() => ({})),
}));

jest.mock("../../src/db/migrate", () => ({
  runMigrations: jest.fn(() => Promise.resolve()),
}));

jest.mock("../../src/utils/auth", () => {
  const actual = jest.requireActual("../../src/utils/auth") as Record<
    string,
    unknown
  >;
  return {
    signAccessToken: actual.signAccessToken,
    signRefreshToken: actual.signRefreshToken,
    verifyToken: actual.verifyToken,
    hashPassword: jest.fn((pw: string) => Promise.resolve(pw)),
    comparePassword: jest.fn((pw: string, hash: string) =>
      Promise.resolve(pw === hash),
    ),
  };
});

// shallow db mock so subscription.service and routes compile
jest.mock("../../src/db/kysely", () => ({
  getDb: jest.fn(() => ({
    insertInto: jest.fn(() => ({
      values: jest.fn(() => ({ execute: jest.fn(() => Promise.resolve()) })),
    })),
    selectFrom: jest.fn(() => ({
      selectAll: jest.fn(() => {
        const chainable: Record<string, jest.Mock> = {
          where: jest.fn(() => chainable),
          executeTakeFirst: jest.fn(() => Promise.resolve(undefined)),
        };
        return chainable;
      }),
      select: jest.fn(() => ({
        where: jest.fn(() => ({
          executeTakeFirst: jest.fn(() => Promise.resolve(undefined)),
        })),
      })),
    })),
    updateTable: jest.fn(() => ({
      set: jest.fn(() => ({
        where: jest.fn(() => ({ execute: jest.fn(() => Promise.resolve()) })),
      })),
    })),
  })),
  closeDb: jest.fn(),
}));

jest.mock("../../src/db/repositories/user-repo", () => ({
  findById: jest.fn(() =>
    Promise.resolve({
      id: "test-user",
      email: "test@example.com",
      tier: "free",
    }),
  ),
}));

jest.mock("../../src/db/repositories/session-repo", () => ({
  findByUserId: jest.fn(() => Promise.resolve([])),
  findById: jest.fn(() => Promise.resolve(null)),
  create: jest.fn(() => Promise.resolve()),
  updateTitle: jest.fn(() => Promise.resolve()),
  remove: jest.fn(() => Promise.resolve()),
}));

jest.mock("../../src/billing/yookassa-client", () => ({
  YooKassaClient: jest.fn().mockImplementation(() => ({
    createPayment: jest.fn(async () => ({
      id: "pm_test",
      status: "pending",
      confirmation: {
        confirmation_url:
          "https://yoomoney.ru/checkout/payments/v2/confirmation?orderId=pm_test",
      },
    })),
    getPayment: jest.fn(async () => ({ id: "pm_123", status: "active" })),
    deactivateSavedPayment: jest.fn(async () => ({})),
  })),
}));

const { app } = require("../../src/server");

// ── Gateway interface contract ──

describe("PaymentGateway interface contract", () => {
  it("defines createSubscriptionPayment method signature", () => {
    const gateway: Partial<PaymentGateway> = {
      name: "yookassa",
      createSubscriptionPayment: jest.fn(),
    };
    expect(gateway.name).toBe("yookassa");
    expect(typeof gateway.createSubscriptionPayment).toBe("function");
  });

  it("defines cancelSubscription method signature", () => {
    const gateway: Partial<PaymentGateway> = {
      cancelSubscription: jest.fn(),
    };
    expect(typeof gateway.cancelSubscription).toBe("function");
  });

  it("defines getSubscriptionStatus method signature", () => {
    const gateway: Partial<PaymentGateway> = {
      getSubscriptionStatus: jest.fn(),
    };
    expect(typeof gateway.getSubscriptionStatus).toBe("function");
  });

  it("defines parseWebhook method signature", () => {
    const gateway: Partial<PaymentGateway> = {
      parseWebhook: jest.fn(),
    };
    expect(typeof gateway.parseWebhook).toBe("function");
  });

  it("accepts valid subscription plan types", () => {
    const plans: SubscriptionPlan[] = ["free", "pro", "team"];
    expect(plans).toContain("free");
    expect(plans).toContain("pro");
    expect(plans).toContain("team");
  });
});

// ── YooKassa adapter (mocked) ──

describe("YooKassaAdapter", () => {
  let YooKassaAdapter: typeof import("../../src/billing/yookassa.adapter").YooKassaAdapter;

  beforeAll(async () => {
    YooKassaAdapter = (await import("../../src/billing/yookassa.adapter"))
      .YooKassaAdapter;
  });

  it("has name 'yookassa'", () => {
    const adapter = new YooKassaAdapter("shop", "key");
    expect(adapter.name).toBe("yookassa");
  });

  it("createSubscriptionPayment returns pending payment with confirmationUrl", async () => {
    const adapter = new YooKassaAdapter("shop", "key");
    const result = await adapter.createSubscriptionPayment({
      userId: "u1",
      planId: "pro",
      period: "month",
      amountRub: 299,
      description: "Pro (month)",
      returnUrl: "http://localhost:3000/settings",
      email: "user@example.com",
    });

    expect(result).toHaveProperty("paymentId");
    expect(result).toHaveProperty("confirmationUrl");
    expect(result.confirmationUrl).toContain("yoomoney.ru");
    expect(result.status).toBe("pending");
  });

  it("cancelSubscription returns canceledAt", async () => {
    const adapter = new YooKassaAdapter("shop", "key");
    const result = await adapter.cancelSubscription("sub_123");
    expect(result).toHaveProperty("canceledAt");
    expect(result.canceledAt).toBeInstanceOf(Date);
  });

  it("getSubscriptionStatus returns active status", async () => {
    const adapter = new YooKassaAdapter("shop", "key");
    const result = await adapter.getSubscriptionStatus("sub_123");
    expect(result).toHaveProperty("status", "active");
  });

  it("parseWebhook throws for invalid IP", async () => {
    const adapter = new YooKassaAdapter("shop", "key");
    await expect(
      adapter.parseWebhook({
        ip: "1.2.3.4",
        body: { event: "payment.succeeded", object: {} },
      }),
    ).rejects.toThrow("Forbidden");
  });
});

// ── Subscription service (mocked gateway) ──

describe("SubscriptionService", () => {
  let SubscriptionService: typeof import("../../src/billing/subscription.service").SubscriptionService;
  let mockGateway: jest.Mocked<PaymentGateway>;

  beforeAll(async () => {
    SubscriptionService = (
      await import("../../src/billing/subscription.service")
    ).SubscriptionService;
  });

  beforeEach(() => {
    mockGateway = {
      name: "yookassa",
      createSubscriptionPayment: jest.fn().mockResolvedValue({
        paymentId: "pm_test",
        confirmationUrl: "https://yoomoney.ru/checkout/confirm",
        status: "pending" as const,
      }),
      cancelSubscription: jest.fn().mockResolvedValue({
        canceledAt: new Date(),
      }),
      getSubscriptionStatus: jest.fn().mockResolvedValue({
        status: "active" as const,
      }),
      parseWebhook: jest.fn(),
      chargeSavedPayment: jest.fn().mockResolvedValue({
        paymentId: "pm_recurring",
        status: "succeeded" as const,
      }),
    };
  });

  it("getPlanList returns all plans with prices", () => {
    const db = require("../../src/db/kysely").getDb();
    const service = new SubscriptionService(mockGateway, db);
    const plans = service.getPlanList();
    expect(plans).toHaveLength(3);
    const pro = plans.find((p: any) => p.id === "pro");
    expect(pro).toBeDefined();
    expect(pro!.priceMonth).toBeGreaterThan(0);
  });
});

// ── Billing routes via supertest ──

describe("Billing routes", () => {
  describe("GET /api/billing/plans", () => {
    it("returns plan list without auth", async () => {
      const res = await request(app).get("/api/billing/plans");
      expect(res.status).toBe(200);
      expect(res.body).toHaveProperty("plans");
      expect(Array.isArray(res.body.plans)).toBe(true);
    });
  });

  describe("POST /api/billing/subscribe", () => {
    it("returns 401 without auth", async () => {
      const res = await request(app)
        .post("/api/billing/subscribe")
        .send({ planId: "pro", period: "month" });
      expect(res.status).toBe(401);
    });

    it("returns 400 for invalid plan", async () => {
      const res = await request(app)
        .post("/api/billing/subscribe")
        .set(authHeader)
        .send({ planId: "invalid", period: "month" });
      expect(res.status).toBe(400);
    });

    it("returns confirmationUrl for valid request", async () => {
      const res = await request(app)
        .post("/api/billing/subscribe")
        .set(authHeader)
        .send({ planId: "pro", period: "month" });
      expect(res.status).toBe(200);
      expect(res.body).toHaveProperty("confirmationUrl");
    });
  });

  describe("POST /api/billing/cancel", () => {
    it("returns 401 without auth", async () => {
      const res = await request(app).post("/api/billing/cancel");
      expect(res.status).toBe(401);
    });

    it("returns 404 when no active subscription", async () => {
      const res = await request(app)
        .post("/api/billing/cancel")
        .set(authHeader);
      expect(res.status).toBe(404);
    });
  });
});
