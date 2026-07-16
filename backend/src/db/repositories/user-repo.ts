import { getDb } from "../kysely";
import { v4 as uuid } from "uuid";
import { hashPassword } from "../../utils/auth";

export interface User {
  id: string;
  email: string;
  tier: string;
  created_at: Date;
  updated_at: Date;
}

export interface UserWithPassword extends User {
  password: string;
}

export async function createUser(
  email: string,
  password: string,
): Promise<User> {
  const id = uuid();
  const hashed = await hashPassword(password);
  const db = getDb();
  const row = await db
    .insertInto("users")
    .values({ id, email, password: hashed, tier: "free" })
    .returningAll()
    .executeTakeFirstOrThrow();
  return row;
}

export async function findByEmail(
  email: string,
): Promise<UserWithPassword | undefined> {
  const db = getDb();
  return db
    .selectFrom("users")
    .selectAll()
    .where("email", "=", email)
    .executeTakeFirst();
}

export async function findById(id: string): Promise<User | undefined> {
  const db = getDb();
  return db
    .selectFrom("users")
    .select(["id", "email", "tier", "created_at", "updated_at"])
    .where("id", "=", id)
    .executeTakeFirst();
}

export async function updateTier(id: string, tier: string): Promise<void> {
  const db = getDb();
  await db.updateTable("users").set({ tier }).where("id", "=", id).execute();
}
