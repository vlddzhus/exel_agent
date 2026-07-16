import { getDb } from "../kysely";
import { v4 as uuid } from "uuid";

export interface SessionRow {
  id: string;
  user_id: string;
  title: string;
  date: string;
  preview: string;
  step_count: number;
  token_count: number;
  messages: unknown;
  created_at: Date;
  updated_at: Date;
}

export type SessionMeta = Pick<
  SessionRow,
  "id" | "title" | "date" | "preview" | "step_count" | "token_count"
>;

export async function listSessions(
  userId: string,
  limit = 50,
): Promise<SessionMeta[]> {
  const db = getDb();
  return db
    .selectFrom("chat_sessions")
    .select(["id", "title", "date", "preview", "step_count", "token_count"])
    .where("user_id", "=", userId)
    .orderBy("date", "desc")
    .limit(limit)
    .execute();
}

export async function getSession(
  sessionId: string,
): Promise<SessionRow | undefined> {
  const db = getDb();
  return db
    .selectFrom("chat_sessions")
    .selectAll()
    .where("id", "=", sessionId)
    .executeTakeFirst();
}

export async function upsertSession(
  userId: string,
  data: {
    id: string;
    title?: string;
    date?: string;
    preview?: string;
    step_count?: number;
    token_count?: number;
    messages?: unknown;
  },
): Promise<void> {
  const db = getDb();
  const existing = await db
    .selectFrom("chat_sessions")
    .select("id")
    .where("id", "=", data.id)
    .executeTakeFirst();

  if (existing) {
    await db
      .updateTable("chat_sessions")
      .set({ ...data, updated_at: new Date() })
      .where("id", "=", data.id)
      .execute();
  } else {
    await db
      .insertInto("chat_sessions")
      .values({
        id: data.id,
        user_id: userId,
        title: data.title || "",
        date: data.date || new Date().toISOString(),
        preview: data.preview || "",
        step_count: data.step_count || 0,
        token_count: data.token_count || 0,
        messages: data.messages || [],
      })
      .execute();
  }
}

export async function deleteSession(sessionId: string): Promise<boolean> {
  const db = getDb();
  const result = await db
    .deleteFrom("chat_sessions")
    .where("id", "=", sessionId)
    .executeTakeFirst();
  return result.numDeletedRows > 0n;
}

export async function renameSession(
  sessionId: string,
  title: string,
): Promise<boolean> {
  const db = getDb();
  const result = await db
    .updateTable("chat_sessions")
    .set({ title, updated_at: new Date() })
    .where("id", "=", sessionId)
    .executeTakeFirst();
  return result.numUpdatedRows > 0n;
}
