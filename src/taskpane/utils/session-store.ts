// HTTPS обязательно: taskpane грузится с https://localhost:3000, HTTP-fetch
// заблокируется как mixed content. См. backend/.env: USE_HTTPS=true.
const DEFAULT_BACKEND_URL = 'https://localhost:4000';
const STORAGE_KEY = 'excel_ai_sessions';
const MAX_SESSIONS = 20;

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

function getBackendUrl(): string {
  const stored = localStorage.getItem('backend_url');
  // Нормализуем http→https для localhost, иначе будет NETWORK_ERROR (mixed content).
  // См. подробный комментарий в useAgent.ts.
  if (stored && /^http:\/\/localhost/i.test(stored)) {
    const fixed = stored.replace(/^http:/i, 'https:');
    localStorage.setItem('backend_url', fixed);
    return fixed;
  }
  return stored || DEFAULT_BACKEND_URL;
}

async function apiGet<T>(path: string): Promise<T | null> {
  try {
    const res = await fetch(`${getBackendUrl()}${path}`);
    if (!res.ok) return null;
    return (await res.json()) as T;
  } catch {
    return null;
  }
}

async function apiPost(path: string, body: unknown): Promise<boolean> {
  try {
    const res = await fetch(`${getBackendUrl()}${path}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    return res.ok;
  } catch {
    return false;
  }
}

async function apiDelete(path: string): Promise<boolean> {
  try {
    const res = await fetch(`${getBackendUrl()}${path}`, {
      method: 'DELETE',
    });
    return res.ok;
  } catch {
    return false;
  }
}

// ── LocalStorage helpers ──

function lsGetSessions(): ChatSession[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return [];
    return JSON.parse(raw) as ChatSession[];
  } catch {
    return [];
  }
}

function lsSaveSession(session: ChatSession): void {
  const sessions = lsGetSessions();
  const idx = sessions.findIndex(s => s.id === session.id);
  if (idx >= 0) {
    sessions[idx] = session;
  } else {
    sessions.unshift(session);
  }
  if (sessions.length > MAX_SESSIONS) {
    sessions.splice(MAX_SESSIONS);
  }
  localStorage.setItem(STORAGE_KEY, JSON.stringify(sessions));
}

function lsDeleteSession(id: string): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(lsGetSessions().filter(s => s.id !== id)));
}

// ── Public API (API-first, local-fallback) ──

export async function getSessionsAsync(): Promise<ChatSession[]> {
  const data = await apiGet<{ sessions: ChatSession[] }>('/api/sessions');
  if (data && data.sessions) return data.sessions;
  return lsGetSessions();
}

export function getSessions(): ChatSession[] {
  return lsGetSessions();
}

export async function saveSessionAsync(session: ChatSession): Promise<void> {
  const ok = await apiPost('/api/sessions', { session });
  if (!ok) {
    lsSaveSession(session);
  }
}

export function saveSession(session: ChatSession): void {
  lsSaveSession(session);
  saveSessionAsync(session).catch(() => {});
}

export async function deleteSessionAsync(id: string): Promise<void> {
  const ok = await apiDelete(`/api/sessions/${encodeURIComponent(id)}`);
  if (!ok) {
    lsDeleteSession(id);
  }
}

export function deleteSession(id: string): void {
  lsDeleteSession(id);
  deleteSessionAsync(id).catch(() => {});
}

export async function loadSessionAsync(id: string): Promise<ChatSession | undefined> {
  const data = await apiGet<{ session: ChatSession }>(`/api/sessions/${encodeURIComponent(id)}`);
  if (data && data.session) return data.session;
  return lsGetSessions().find(s => s.id === id);
}

export function renameSession(id: string, title: string): void {
  const sessions = lsGetSessions();
  const session = sessions.find(s => s.id === id);
  if (session) {
    session.title = title;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(sessions));
  }
  fetch(`${getBackendUrl()}/api/sessions/${encodeURIComponent(id)}/rename`, {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ title }),
  }).catch(() => {});
}

export function exportSessionText(id: string): string {
  const sessions = lsGetSessions();
  const session = sessions.find(s => s.id === id);
  if (!session) return '';
  return session.messages.map(m => `[${m.role.toUpperCase()}]\n${m.content}`).join('\n\n---\n\n');
}

export function generateSessionId(): string {
  return Date.now().toString(36) + Math.random().toString(36).substring(2, 8);
}

export function countTokens(text: string): number {
  return Math.ceil(text.length / 4);
}
