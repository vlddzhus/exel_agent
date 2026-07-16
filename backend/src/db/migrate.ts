import { Kysely, sql } from "kysely";
import { DB } from "./kysely";

async function createTableIfNotExists(
  db: Kysely<DB>,
  tableName: string,
  builder: (db: Kysely<DB>) => Promise<void>,
): Promise<void> {
  const result = await sql<{ exists: boolean }>`
    SELECT EXISTS (
      SELECT FROM information_schema.tables WHERE table_name = ${tableName}
    ) AS exists
  `.execute(db);
  if (result.rows[0]?.exists) return;
  await builder(db);
}

export async function runMigrations(db: Kysely<DB>): Promise<void> {
  await createTableIfNotExists(db, "users", async (db) => {
    await db.schema
      .createTable("users")
      .addColumn("id", "text", (c) => c.primaryKey())
      .addColumn("email", "text", (c) => c.notNull().unique())
      .addColumn("password", "text", (c) => c.notNull())
      .addColumn("tier", "text", (c) => c.notNull().defaultTo("free"))
      .addColumn("created_at", "timestamptz", (c) =>
        c.notNull().defaultTo("now()"),
      )
      .addColumn("updated_at", "timestamptz", (c) =>
        c.notNull().defaultTo("now()"),
      )
      .execute();
  });

  await createTableIfNotExists(db, "refresh_tokens", async (db) => {
    await db.schema
      .createTable("refresh_tokens")
      .addColumn("id", "text", (c) => c.primaryKey())
      .addColumn("user_id", "text", (c) =>
        c.notNull().references("users.id").onDelete("cascade"),
      )
      .addColumn("token", "text", (c) => c.notNull().unique())
      .addColumn("expires_at", "timestamptz", (c) => c.notNull())
      .addColumn("created_at", "timestamptz", (c) =>
        c.notNull().defaultTo("now()"),
      )
      .execute();
  });

  await createTableIfNotExists(db, "chat_sessions", async (db) => {
    await db.schema
      .createTable("chat_sessions")
      .addColumn("id", "text", (c) => c.primaryKey())
      .addColumn("user_id", "text", (c) =>
        c.notNull().references("users.id").onDelete("cascade"),
      )
      .addColumn("title", "text", (c) => c.notNull().defaultTo(""))
      .addColumn("date", "text", (c) => c.notNull())
      .addColumn("preview", "text", (c) => c.notNull().defaultTo(""))
      .addColumn("step_count", "integer", (c) => c.notNull().defaultTo(0))
      .addColumn("token_count", "integer", (c) => c.notNull().defaultTo(0))
      .addColumn("messages", "jsonb", (c) => c.notNull().defaultTo("[]"))
      .addColumn("created_at", "timestamptz", (c) =>
        c.notNull().defaultTo("now()"),
      )
      .addColumn("updated_at", "timestamptz", (c) =>
        c.notNull().defaultTo("now()"),
      )
      .execute();
  });

  await createTableIfNotExists(db, "usage_events", async (db) => {
    await db.schema
      .createTable("usage_events")
      .addColumn("id", "text", (c) => c.primaryKey())
      .addColumn("user_id", "text", (c) =>
        c.notNull().references("users.id").onDelete("cascade"),
      )
      .addColumn("provider", "text", (c) => c.notNull())
      .addColumn("model", "text", (c) => c.notNull())
      .addColumn("input_tokens", "integer", (c) => c.notNull().defaultTo(0))
      .addColumn("output_tokens", "integer", (c) => c.notNull().defaultTo(0))
      .addColumn("cost", "text", (c) => c.notNull().defaultTo("0"))
      .addColumn("created_at", "timestamptz", (c) =>
        c.notNull().defaultTo("now()"),
      )
      .execute();
  });

  await createTableIfNotExists(db, "subscriptions", async (db) => {
    await db.schema
      .createTable("subscriptions")
      .addColumn("id", "text", (c) => c.primaryKey())
      .addColumn("user_id", "text", (c) =>
        c.notNull().references("users.id").onDelete("cascade"),
      )
      .addColumn("plan", "text", (c) => c.notNull())
      .addColumn("status", "text", (c) => c.notNull().defaultTo("pending"))
      .addColumn("provider", "text", (c) => c.notNull())
      .addColumn("provider_payment_id", "text", (c) => c.notNull())
      .addColumn("provider_sub_id", "text")
      .addColumn("period", "text", (c) => c.notNull())
      .addColumn("price_kopecks", "integer", (c) => c.notNull())
      .addColumn("current_period_start", "timestamptz")
      .addColumn("current_period_end", "timestamptz")
      .addColumn("created_at", "timestamptz", (c) =>
        c.notNull().defaultTo("now()"),
      )
      .addColumn("updated_at", "timestamptz", (c) =>
        c.notNull().defaultTo("now()"),
      )
      .execute();
  });

  await createTableIfNotExists(db, "webhook_events", async (db) => {
    await db.schema
      .createTable("webhook_events")
      .addColumn("id", "text", (c) => c.primaryKey())
      .addColumn("type", "text", (c) => c.notNull())
      .addColumn("occurred_at", "timestamptz", (c) => c.notNull())
      .addColumn("raw", "text", (c) => c.notNull())
      .addColumn("created_at", "timestamptz", (c) =>
        c.notNull().defaultTo("now()"),
      )
      .execute();
  });
}
