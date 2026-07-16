import { z } from "zod";
import { v4 as uuid } from "uuid";
import { Kysely } from "kysely";
import type { DB } from "../db/kysely";
import {
  PaymentGateway,
  WebhookEvent,
  SubscriptionPlan,
} from "./gateway.interface";
import * as userRepo from "../db/repositories/user-repo";
import pino from "pino";

const logger = pino({ level: process.env.LOG_LEVEL || "info" });

const PLANS: Record<
  string,
  { name: string; price: { month: number; year: number }; dailyLimit: number }
> = {
  free: { name: "Free", price: { month: 0, year: 0 }, dailyLimit: 10 },
  pro: {
    name: "Pro",
    price: { month: 29900, year: 299000 },
    dailyLimit: 500,
  },
  team: {
    name: "Team",
    price: { month: 79900, year: 799000 },
    dailyLimit: 2000,
  },
};

const APP_URL = process.env.APP_URL || "http://localhost:3000";

export class SubscriptionService {
  constructor(
    private gateway: PaymentGateway,
    private db: Kysely<DB>,
  ) {}

  async subscribe(
    userId: string,
    planId: "pro" | "team",
    period: "month" | "year",
  ): Promise<{ confirmationUrl: string }> {
    const user = await userRepo.findById(userId);
    if (!user) {
      throw Object.assign(new Error("User not found"), { statusCode: 404 });
    }

    const plan = PLANS[planId];
    const priceKopecks = plan.price[period];
    const priceRub = priceKopecks / 100;

    const payment = await this.gateway.createSubscriptionPayment({
      userId,
      planId,
      period,
      amountRub: priceRub,
      description: `Подписка ${plan.name} (${period})`,
      returnUrl: `${APP_URL}/settings?from=payment`,
      email: user.email,
    });

    await this.db
      .insertInto("subscriptions")
      .values({
        id: uuid(),
        user_id: userId,
        plan: planId,
        status: "pending",
        provider: this.gateway.name,
        provider_payment_id: payment.paymentId,
        period,
        price_kopecks: priceKopecks,
        created_at: new Date(),
        updated_at: new Date(),
      })
      .execute();

    logger.info(
      { userId, planId, period, paymentId: payment.paymentId },
      "subscription created",
    );

    return { confirmationUrl: payment.confirmationUrl };
  }

  async cancelSubscription(userId: string): Promise<{ canceledAt: Date }> {
    const sub = await this.db
      .selectFrom("subscriptions")
      .selectAll()
      .where("user_id", "=", userId)
      .where("status", "in", ["active", "pending"])
      .executeTakeFirst();

    if (!sub) {
      throw Object.assign(new Error("No active subscription found"), {
        statusCode: 404,
      });
    }

    const result = await this.gateway.cancelSubscription(
      sub.provider_sub_id ?? sub.provider_payment_id,
    );

    await this.db
      .updateTable("subscriptions")
      .set({ status: "canceled", updated_at: new Date() })
      .where("id", "=", sub.id)
      .execute();

    logger.info({ userId, subId: sub.id }, "subscription canceled");

    return result;
  }

  async handleWebhook(event: WebhookEvent): Promise<void> {
    const existing = await this.db
      .selectFrom("webhook_events")
      .selectAll()
      .where("id", "=", event.id)
      .executeTakeFirst();

    if (existing) {
      logger.info({ eventId: event.id }, "duplicate webhook, skipped");
      return;
    }

    try {
      switch (event.type) {
        case "payment.succeeded":
          await this.activateSubscription(event);
          break;
        case "subscription.renewed":
          await this.renewSubscription(event);
          break;
        case "subscription.canceled":
        case "payment.failed":
          await this.markFailed(event);
          break;
      }
    } catch (err) {
      logger.error({ eventId: event.id, error: err }, "webhook handler error");
      throw err;
    }

    await this.db
      .insertInto("webhook_events")
      .values({
        id: event.id,
        type: event.type,
        occurred_at: event.occurredAt,
        raw: JSON.stringify(event.raw),
        created_at: new Date(),
      })
      .execute();
  }

  getPlanList(): Array<{
    id: string;
    name: string;
    priceMonth: number;
    priceYear: number;
    dailyLimit: number;
  }> {
    return Object.entries(PLANS).map(([id, plan]) => ({
      id,
      name: plan.name,
      priceMonth: plan.price.month,
      priceYear: plan.price.year,
      dailyLimit: plan.dailyLimit,
    }));
  }

  private async activateSubscription(event: WebhookEvent): Promise<void> {
    const sub = await this.db
      .selectFrom("subscriptions")
      .selectAll()
      .where((eb) =>
        eb.or([
          eb("provider_payment_id", "=", event.providerSubId ?? ""),
          eb("id", "=", event.id),
        ]),
      )
      .executeTakeFirst();

    if (!sub) {
      logger.warn(
        { eventId: event.id },
        "no subscription found for payment webhook",
      );
      return;
    }

    await this.db
      .updateTable("subscriptions")
      .set({
        status: "active",
        provider_sub_id: event.providerSubId,
        updated_at: new Date(),
      })
      .where("id", "=", sub.id)
      .execute();

    await this.db
      .updateTable("users")
      .set({ tier: sub.plan })
      .where("id", "=", sub.user_id)
      .execute();

    logger.info(
      { userId: sub.user_id, plan: sub.plan },
      "subscription activated",
    );
  }

  private async renewSubscription(event: WebhookEvent): Promise<void> {
    await this.activateSubscription(event);
  }

  async renewDueSubscriptions(): Promise<number> {
    const now = new Date();
    const deadline = new Date(now.getTime() + 86400000);
    const due = await this.db
      .selectFrom("subscriptions")
      .selectAll()
      .where("status", "=", "active")
      .where("current_period_end", "<", deadline)
      .where("current_period_end", ">", now)
      .where("provider_sub_id", "is not", null)
      .execute();

    for (const sub of due) {
      try {
        const plan = PLANS[sub.plan];
        const user = await userRepo.findById(sub.user_id);
        if (!user || !sub.provider_sub_id) continue;

        const priceKopecks =
          plan.price[sub.period as "month" | "year"] ?? plan.price.month;
        const priceRub = priceKopecks / 100;

        await this.gateway.chargeSavedPayment({
          paymentMethodId: sub.provider_sub_id,
          amountRub: priceRub,
          description: `Подписка ${plan.name}, продление`,
          email: user.email,
          userId: sub.user_id,
          planId: sub.plan,
          period: sub.period ?? "month",
        });

        await this.db
          .updateTable("subscriptions")
          .set({
            current_period_end: new Date(
              now.getTime() + (sub.period === "year" ? 365 : 30) * 86400000,
            ),
            updated_at: now,
          })
          .where("id", "=", sub.id)
          .execute();

        logger.info(
          { subId: sub.id, userId: sub.user_id },
          "recurring charge initiated",
        );
      } catch (err) {
        logger.error({ subId: sub.id, error: err }, "recurring charge failed");
      }
    }

    return due.length;
  }

  private async markFailed(event: WebhookEvent): Promise<void> {
    const sub = await this.db
      .selectFrom("subscriptions")
      .selectAll()
      .where((eb) =>
        eb.or([
          eb("provider_payment_id", "=", event.providerSubId ?? ""),
          eb("id", "=", event.id),
        ]),
      )
      .executeTakeFirst();

    if (!sub) return;

    await this.db
      .updateTable("subscriptions")
      .set({ status: "past_due", updated_at: new Date() })
      .where("id", "=", sub.id)
      .execute();

    logger.warn(
      { userId: sub.user_id, eventType: event.type },
      "subscription payment failed",
    );
  }
}
