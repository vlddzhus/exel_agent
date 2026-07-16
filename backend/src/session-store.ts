import fs from "fs";
import path from "path";

export interface ChatMessageData {
  role: string;
  content: string;
}

export interface ChatSession {
  id: string;
  title: string;
  date: string;
  preview: string;
  stepCount: number;
  tokenCount: number;
  messages: ChatMessageData[];
}

const DATA_DIR = path.join(__dirname, "..", "data", "sessions");
const MAX_SESSIONS = 50;

function ensureDir(): void {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
  }
}

const SESSION_ID_RE = /^[a-zA-Z0-9_-]+$/;

function filePath(id: string): string {
  if (!SESSION_ID_RE.test(id)) {
    throw new Error(`Invalid session ID: ${id}`);
  }
  return path.join(DATA_DIR, `${id}.json`);
}

export class SessionStore {
  private cache: Map<string, ChatSession> = new Map();
  private listCache: ChatSession[] = [];

  constructor() {
    this.loadAll();
  }

  private loadAll(): void {
    ensureDir();
    this.cache.clear();
    this.listCache = [];
    try {
      const files = fs
        .readdirSync(DATA_DIR)
        .filter((f) => f.endsWith(".json"))
        .sort()
        .reverse()
        .slice(0, MAX_SESSIONS);
      for (const file of files) {
        try {
          const session = JSON.parse(
            fs.readFileSync(path.join(DATA_DIR, file), "utf-8"),
          ) as ChatSession;
          if (session && session.id) {
            this.cache.set(session.id, session);
            this.listCache.push(session);
          }
        } catch {}
      }
      this.listCache.sort(
        (a, b) => new Date(b.date).getTime() - new Date(a.date).getTime(),
      );
    } catch {}
  }

  getAll(): ChatSession[] {
    return this.listCache;
  }

  getAllMeta(): Array<Omit<ChatSession, "messages">> {
    return this.listCache.map((s) => ({
      id: s.id,
      title: s.title,
      date: s.date,
      preview: s.preview,
      stepCount: s.stepCount,
      tokenCount: s.tokenCount,
    }));
  }

  getById(id: string): ChatSession | undefined {
    if (this.cache.has(id)) return this.cache.get(id);
    try {
      if (fs.existsSync(filePath(id))) {
        const session = JSON.parse(
          fs.readFileSync(filePath(id), "utf-8"),
        ) as ChatSession;
        this.cache.set(id, session);
        return session;
      }
    } catch {}
    return undefined;
  }

  save(session: ChatSession): void {
    ensureDir();
    this.cache.set(session.id, session);

    const idx = this.listCache.findIndex((s) => s.id === session.id);
    if (idx >= 0) {
      this.listCache[idx] = session;
    } else {
      this.listCache.unshift(session);
    }
    if (this.listCache.length > MAX_SESSIONS) {
      const removed = this.listCache.splice(MAX_SESSIONS);
      for (const r of removed) {
        this.cache.delete(r.id);
        try {
          fs.unlinkSync(filePath(r.id));
        } catch {}
      }
    }

    try {
      fs.writeFileSync(filePath(session.id), JSON.stringify(session, null, 2));
    } catch {}
  }

  delete(id: string): boolean {
    this.cache.delete(id);
    this.listCache = this.listCache.filter((s) => s.id !== id);
    try {
      if (fs.existsSync(filePath(id))) {
        fs.unlinkSync(filePath(id));
        return true;
      }
    } catch {}
    return false;
  }

  rename(id: string, title: string): boolean {
    const session = this.cache.get(id) || this.getById(id);
    if (!session) return false;
    session.title = title;
    this.save(session);
    return true;
  }
}

export const sessionStore = new SessionStore();
