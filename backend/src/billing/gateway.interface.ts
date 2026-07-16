import { z } from "zod";

export type SubscriptionPlan = "free" | "pro" | "team";

export const CreatePaymentSchema = z.object({
  userId: z.string(),
  planId: z.enum(["pro", "team"]),
  period: z.enum(["month", "year"]),
  amountRub: z.number().positive(),
  description: z.string(),
  returnUrl: z.string().url(),
  email: z.string().email(),
});

export interface PaymentGateway {
  readonly name: "yookassa" | "stripe" | "cloudpayments" | "tinkoff";

  createSubscriptionPayment(
    params: z.infer<typeof CreatePaymentSchema>,
  ): Promise<{
    paymentId: string;
    confirmationUrl: string;
    status: "pending" | "succeeded";
  }>;

  cancelSubscription(providerSubId: string): Promise<{ canceledAt: Date }>;
  getSubscriptionStatus(providerSubId: string): Promise<SubscriptionStatus>;
  parseWebhook(req: unknown): Promise<WebhookEvent>;
  chargeSavedPayment(params: {
    paymentMethodId: string;
    amountRub: number;
    description: string;
    email: string;
    userId: string;
    planId: string;
    period: string;
  }): Promise<{
    paymentId: string;
    status: "pending" | "succeeded";
  }>;
}

export interface SubscriptionStatus {
  status: "active" | "canceled" | "past_due" | "expired";
  currentPeriodEnd?: Date;
  nextPaymentDate?: Date;
}

export interface WebhookEvent {
  id: string;
  type:
    | "payment.succeeded"
    | "payment.failed"
    | "subscription.renewed"
    | "subscription.canceled"
    | "refund.succeeded";
  providerSubId?: string;
  amountRub?: number;
  occurredAt: Date;
  raw: unknown;
}
