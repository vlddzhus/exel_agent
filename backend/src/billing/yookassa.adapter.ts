import { randomBytes } from "crypto";
import { z } from "zod";
import { YooKassaClient } from "./yookassa-client";
import {
  PaymentGateway,
  CreatePaymentSchema,
  SubscriptionStatus,
  WebhookEvent,
} from "./gateway.interface";
import pino from "pino";

const logger = pino({ level: process.env.LOG_LEVEL || "info" });

const YOOKASSA_IPS = [
  "185.71.76.0/27",
  "185.71.77.0/27",
  "77.75.153.0/27",
  "77.75.154.128/25",
  "2a02:5180::/32",
];

const VAT_CODE = 1; // 20% НДС, уточнить у бухгалтерии

const PLANS: Record<string, { name: string }> = {
  pro: { name: "Pro" },
  team: { name: "Team" },
};

function ipInCidr(ip: string, cidr: string): boolean {
  // ponytail: IPv6 check — simplified prefix match for YooKassa range
  if (ip.includes(":")) {
    if (!cidr.includes(":")) return false;
    const [range] = cidr.split("/");
    return ip.toLowerCase().startsWith(range.toLowerCase().slice(0, 6));
  }

  const [range, bitsStr] = cidr.split("/");
  const bits = parseInt(bitsStr, 10);
  if (cidr.includes(":")) return false;

  const ipLong = ipv4ToLong(ip);
  if (ipLong === null) return false;
  const rangeLong = ipv4ToLong(range);
  if (rangeLong === null) return false;

  const mask = ~(2 ** (32 - bits) - 1);
  return (ipLong & mask) === (rangeLong & mask);
}

function ipv4ToLong(ip: string): number | null {
  const parts = ip.split(".").map(Number);
  if (parts.length !== 4 || parts.some((p) => isNaN(p) || p < 0 || p > 255))
    return null;
  return (
    ((parts[0] << 24) | (parts[1] << 16) | (parts[2] << 8) | parts[3]) >>> 0
  );
}

function assertYookassaIp(ip: string): void {
  const isAllowed = YOOKASSA_IPS.some((cidr) => ipInCidr(ip, cidr));
  if (!isAllowed) {
    throw new Error(`Forbidden: IP ${ip} is not a YooKassa IP`);
  }
}

function generateIdempotencyKey(): string {
  return randomBytes(16).toString("hex");
}

function mapYooKassaEvent(event: string): WebhookEvent["type"] {
  switch (event) {
    case "payment.waiting_for_capture":
    case "payment.succeeded":
      return "payment.succeeded";
    case "payment.canceled":
      return "payment.failed";
    case "refund.succeeded":
      return "refund.succeeded";
    default:
      return "payment.failed";
  }
}

export class YooKassaAdapter implements PaymentGateway {
  readonly name = "yookassa" as const;

  private client: YooKassaClient;

  constructor(shopId: string, secretKey: string) {
    this.client = new YooKassaClient(shopId, secretKey);
  }

  async createSubscriptionPayment(
    params: z.infer<typeof CreatePaymentSchema>,
  ): Promise<{
    paymentId: string;
    confirmationUrl: string;
    status: "pending" | "succeeded";
  }> {
    const idempotencyKey = generateIdempotencyKey();
    const body = {
      amount: { value: params.amountRub.toFixed(2), currency: "RUB" },
      capture: true,
      confirmation: { type: "redirect", return_url: params.returnUrl },
      description: params.description,
      save_payment_method: true,
      receipt: {
        customer: { email: params.email },
        items: [
          {
            description: `Подписка ${PLANS[params.planId]?.name ?? params.planId}, ${params.period}`,
            quantity: "1",
            amount: { value: params.amountRub.toFixed(2), currency: "RUB" },
            vat_code: VAT_CODE,
            payment_mode: "full_payment",
            payment_subject: "service",
          },
        ],
      },
      metadata: {
        userId: params.userId,
        planId: params.planId,
        period: params.period,
      },
    };

    logger.info(
      { idempotencyKey, amount: params.amountRub, planId: params.planId },
      "yookassa createSubscriptionPayment",
    );

    const payment = await this.client.createPayment(body, idempotencyKey);

    return {
      paymentId: payment.id,
      confirmationUrl: payment.confirmation?.confirmation_url ?? "",
      status: payment.status === "succeeded" ? "succeeded" : "pending",
    };
  }

  async chargeSavedPayment(params: {
    paymentMethodId: string;
    amountRub: number;
    description: string;
    email: string;
    userId: string;
    planId: string;
    period: string;
  }): Promise<{ paymentId: string; status: "pending" | "succeeded" }> {
    const idempotencyKey = generateIdempotencyKey();
    const body = {
      amount: { value: params.amountRub.toFixed(2), currency: "RUB" },
      capture: true,
      payment_method_id: params.paymentMethodId,
      description: params.description,
      save_payment_method: true,
      receipt: {
        customer: { email: params.email },
        items: [
          {
            description: `Подписка, продление`,
            quantity: "1",
            amount: { value: params.amountRub.toFixed(2), currency: "RUB" },
            vat_code: VAT_CODE,
            payment_mode: "full_payment",
            payment_subject: "service",
          },
        ],
      },
      metadata: {
        userId: params.userId,
        planId: params.planId,
        period: params.period,
        type: "recurring",
      },
    };

    logger.info(
      {
        idempotencyKey,
        amount: params.amountRub,
        paymentMethodId: params.paymentMethodId,
      },
      "yookassa chargeSavedPayment",
    );

    const payment = await this.client.createPayment(body, idempotencyKey);
    return {
      paymentId: payment.id,
      status: payment.status === "succeeded" ? "succeeded" : "pending",
    };
  }

  async cancelSubscription(
    providerSubId: string,
  ): Promise<{ canceledAt: Date }> {
    logger.info({ providerSubId }, "yookassa cancelSubscription");

    try {
      await this.client.deactivateSavedPayment(
        providerSubId,
        generateIdempotencyKey(),
      );
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      logger.warn(
        { providerSubId, error: msg },
        "saved payment method deactivate failed, continuing",
      );
    }

    return { canceledAt: new Date() };
  }

  async getSubscriptionStatus(
    providerSubId: string,
  ): Promise<SubscriptionStatus> {
    try {
      const payment = await this.client.getPayment(providerSubId);
      return {
        status: payment.status === "active" ? "active" : "expired",
      };
    } catch {
      return { status: "expired" };
    }
  }

  async parseWebhook(req: unknown): Promise<WebhookEvent> {
    const r = req as { ip?: string; body?: Record<string, unknown> };
    const ip = r?.ip ?? "unknown";
    assertYookassaIp(ip);

    const body = (r?.body ?? {}) as Record<string, unknown>;
    const event = String(body.event ?? "");
    const object = (body.object ?? {}) as Record<string, unknown>;
    const paymentMethod = (object.payment_method ?? {}) as Record<
      string,
      unknown
    >;

    const id = `${event}:${String(object.id ?? "")}`;
    const type = mapYooKassaEvent(event);
    const providerSubId = paymentMethod?.id
      ? String(paymentMethod.id)
      : undefined;
    const amountVal = (object.amount as Record<string, unknown>)?.value;
    const amountRub = amountVal ? parseFloat(String(amountVal)) : undefined;
    const createdAt = object.created_at
      ? new Date(String(object.created_at))
      : new Date();

    return {
      id,
      type,
      providerSubId,
      amountRub,
      occurredAt: createdAt,
      raw: body,
    };
  }
}
