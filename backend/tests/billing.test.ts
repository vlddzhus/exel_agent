import { YooKassaAdapter } from "../src/billing/yookassa.adapter";
import { SubscriptionService } from "../src/billing/subscription.service";
import { createMockDb } from "./mocks/db";

jest.mock("../src/db/kysely", () => ({
  getDb: jest.fn(() => createMockDb()),
  closeDb: jest.fn(),
}));

jest.mock("../src/db/repositories/user-repo", () => ({
  findById: jest.fn((id: string) => {
    if (id === "existing-user") {
      return Promise.resolve({ id, email: "user@test.com", tier: "free" });
    }
    return Promise.resolve(null);
  }),
}));

jest.mock("../src/billing/yookassa-client", () => ({
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

const shopId = "test_shop";
const secretKey = "test_secret";

describe("YooKassaAdapter", () => {
  let adapter: YooKassaAdapter;

  beforeEach(() => {
    adapter = new YooKassaAdapter(shopId, secretKey);
  });

  describe("createSubscriptionPayment", () => {
    it("creates payment with receipt and save_payment_method", async () => {
      const result = await adapter.createSubscriptionPayment({
        userId: "user-1",
        planId: "pro",
        period: "month",
        amountRub: 299,
        description: "Подписка Pro (month)",
        returnUrl: "http://localhost:3000/settings",
        email: "user@test.com",
      });

      expect(result.paymentId).toMatch(/^pm_/);
      expect(result.confirmationUrl).toContain("yoomoney.ru");
      expect(result.status).toBe("pending");
    });

    it("creates yearly team payment", async () => {
      const result = await adapter.createSubscriptionPayment({
        userId: "user-2",
        planId: "team",
        period: "year",
        amountRub: 7990,
        description: "Подписка Team (year)",
        returnUrl: "http://localhost:3000/settings",
        email: "team@test.com",
      });

      expect(result.paymentId).toBeDefined();
      expect(result.status).toBe("pending");
    });
  });

  describe("parseWebhook", () => {
    it("parses payment.succeeded webhook", async () => {
      const event = await adapter.parseWebhook({
        ip: "185.71.76.10",
        body: {
          event: "payment.succeeded",
          object: {
            id: "pm_123",
            status: "succeeded",
            amount: { value: "299.00", currency: "RUB" },
            payment_method: { id: "pmt_456", type: "bank_card", saved: true },
            created_at: "2026-07-08T12:00:00.000Z",
          },
        },
      });

      expect(event.id).toBe("payment.succeeded:pm_123");
      expect(event.type).toBe("payment.succeeded");
      expect(event.providerSubId).toBe("pmt_456");
      expect(event.amountRub).toBe(299);
    });

    it("parses payment.canceled as payment.failed", async () => {
      const event = await adapter.parseWebhook({
        ip: "185.71.76.10",
        body: {
          event: "payment.canceled",
          object: { id: "pm_789", created_at: "2026-07-08T12:00:00.000Z" },
        },
      });

      expect(event.type).toBe("payment.failed");
    });

    it("parses refund.succeeded", async () => {
      const event = await adapter.parseWebhook({
        ip: "185.71.76.10",
        body: {
          event: "refund.succeeded",
          object: {
            id: "rf_123",
            created_at: "2026-07-08T12:00:00.000Z",
          },
        },
      });

      expect(event.type).toBe("refund.succeeded");
    });

    it("rejects unknown IP", async () => {
      await expect(
        adapter.parseWebhook({
          ip: "1.2.3.4",
          body: { event: "payment.succeeded", object: { id: "pm_1" } },
        }),
      ).rejects.toThrow("Forbidden");
    });

    it("rejects YooKassa IP outside range", async () => {
      await expect(
        adapter.parseWebhook({
          ip: "185.71.79.1",
          body: { event: "payment.succeeded", object: { id: "pm_1" } },
        }),
      ).rejects.toThrow("Forbidden");
    });

    it("allows IPv6 YooKassa range", async () => {
      const event = await adapter.parseWebhook({
        ip: "2a02:5180::1",
        body: {
          event: "payment.succeeded",
          object: {
            id: "pm_v6",
            created_at: "2026-07-08T12:00:00.000Z",
          },
        },
      });

      expect(event.id).toBe("payment.succeeded:pm_v6");
    });
  });

  describe("cancelSubscription", () => {
    it("returns canceledAt", async () => {
      const result = await adapter.cancelSubscription("pmt_456");
      expect(result.canceledAt).toBeInstanceOf(Date);
    });
  });

  describe("getSubscriptionStatus", () => {
    it("returns active for existing payment", async () => {
      const status = await adapter.getSubscriptionStatus("pm_123");
      expect(status.status).toBe("active");
    });
  });
});

describe("SubscriptionService", () => {
  let service: SubscriptionService;

  beforeEach(() => {
    const gateway = new YooKassaAdapter(shopId, secretKey);
    const { getDb } = require("../src/db/kysely");
    service = new SubscriptionService(gateway, getDb());
  });

  describe("getPlanList", () => {
    it("returns all plans with prices in kopecks", () => {
      const plans = service.getPlanList();
      expect(plans).toHaveLength(3);
      const pro = plans.find((p) => p.id === "pro");
      expect(pro).toBeDefined();
      expect(pro!.priceMonth).toBe(29900);
      expect(pro!.priceYear).toBe(299000);
      expect(pro!.dailyLimit).toBe(500);
    });

    it("returns free plan with zero prices", () => {
      const plans = service.getPlanList();
      const free = plans.find((p) => p.id === "free");
      expect(free).toBeDefined();
      expect(free!.priceMonth).toBe(0);
      expect(free!.priceYear).toBe(0);
      expect(free!.dailyLimit).toBe(10);
    });
  });

  describe("subscribe", () => {
    it("throws 404 for unknown user", async () => {
      await expect(
        service.subscribe("nonexistent", "pro", "month"),
      ).rejects.toMatchObject({ statusCode: 404 });
    });

    it("creates subscription for existing user", async () => {
      const { findById } = require("../src/db/repositories/user-repo");
      findById.mockResolvedValueOnce({
        id: "existing-user",
        email: "user@test.com",
        tier: "free",
      });

      const result = await service.subscribe("existing-user", "pro", "month");
      expect(result).toHaveProperty("confirmationUrl");
      expect(result.confirmationUrl).toContain("yoomoney.ru");
    });
  });
});
