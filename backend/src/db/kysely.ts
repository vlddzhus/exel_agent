import { Kysely, PostgresDialect, Generated } from "kysely";
import { Pool } from "pg";
import pino from "pino";
import { env } from "../config/env";

const logger = pino({ level: env.LOG_LEVEL });

export interface UsersTable {
  id: string;
  email: string;
  password: string;
  tier: Generated<string>;
  created_at: Generated<Date>;
  updated_at: Generated<Date>;
}

export interface RefreshTokensTable {
  id: string;
  user_id: string;
  token: string;
  expires_at: Date;
  created_at: Generated<Date>;
}

export interface ChatSessionsTable {
  id: string;
  user_id: string;
  title: string;
  date: string;
  preview: string;
  step_count: number;
  token_count: number;
  messages: unknown;
  created_at: Generated<Date>;
  updated_at: Generated<Date>;
}

export interface UsageEventsTable {
  id: string;
  user_id: string;
  provider: string;
  model: string;
  input_tokens: number;
  output_tokens: number;
  cost: string;
  created_at: Generated<Date>;
}

export interface SubscriptionsTable {
  id: string;
  user_id: string;
  plan: string;
  status: string;
  provider: string;
  provider_payment_id: string;
  provider_sub_id?: string;
  period: string;
  price_kopecks: number;
  current_period_start?: Date;
  current_period_end?: Date;
  created_at: Generated<Date>;
  updated_at: Generated<Date>;
}

export interface WebhookEventsTable {
  id: string;
  type: string;
  occurred_at: Date;
  raw: string;
  created_at: Generated<Date>;
}

export interface DB {
  users: UsersTable;
  refresh_tokens: RefreshTokensTable;
  chat_sessions: ChatSessionsTable;
  usage_events: UsageEventsTable;
  subscriptions: SubscriptionsTable;
  webhook_events: WebhookEventsTable;
}

let db: Kysely<DB> | null = null;

export function getDb(): Kysely<DB> {
  if (!db) {
    const pool = new Pool({
      connectionString: env.DATABASE_URL,
      max: 10,
    });

    db = new Kysely<DB>({
      dialect: new PostgresDialect({ pool }),
    });

    pool.on("error", (err) => {
      logger.error(err, "postgres pool error");
    });
  }
  return db;
}

export async function closeDb(): Promise<void> {
  if (db) {
    await db.destroy();
    db = null;
  }
}
