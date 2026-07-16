import request from "supertest";
import jwt from "jsonwebtoken";

jest.mock("ai", () => ({
  streamText: jest.fn(),
  generateText: jest.fn(),
  tool: jest.fn((def: unknown) => def),
  jsonSchema: jest.fn((s: unknown) => s),
}));

jest.mock("../src/utils/auth", () => {
  const actual = jest.requireActual("../src/utils/auth") as Record<
    string,
    unknown
  >;
  return {
    signAccessToken: actual.signAccessToken as typeof actual.signAccessToken,
    signRefreshToken: actual.signRefreshToken as typeof actual.signRefreshToken,
    verifyToken: actual.verifyToken as typeof actual.verifyToken,
    hashPassword: jest.fn((pw: string) => Promise.resolve(pw)),
    comparePassword: jest.fn((pw: string, hash: string) =>
      Promise.resolve(pw === hash),
    ),
  };
});

jest.mock("../src/utils/provider-factory", () => ({
  getProviderEntries: jest.fn(() => []),
  getModelChain: jest.fn(() => []),
  convertTools: jest.fn(() => ({})),
}));

jest.mock("../src/db/kysely", () => ({
  getDb: jest.fn(() => ({})),
  closeDb: jest.fn(),
}));

jest.mock("../src/db/migrate", () => ({
  runMigrations: jest.fn(() => Promise.resolve()),
}));

const mockUsers: Array<{
  id: string;
  email: string;
  password: string;
  tier: string;
}> = [];

jest.mock("../src/db/repositories/user-repo", () => ({
  findByEmail: jest.fn((email: string) =>
    Promise.resolve(mockUsers.find((u) => u.email === email) || null),
  ),
  findById: jest.fn((id: string) =>
    Promise.resolve(mockUsers.find((u) => u.id === id) || null),
  ),
  createUser: jest.fn((email: string, password: string) => {
    const user = { id: `user_${Date.now()}`, email, password, tier: "free" };
    mockUsers.push(user);
    return Promise.resolve(user);
  }),
}));

jest.mock("../src/db/repositories/session-repo", () => ({
  findByUserId: jest.fn(() => Promise.resolve([])),
  findById: jest.fn(() => Promise.resolve(null)),
  create: jest.fn(() => Promise.resolve()),
  updateTitle: jest.fn(() => Promise.resolve()),
  remove: jest.fn(() => Promise.resolve()),
}));

process.env.NODE_ENV = "test";
process.env.JWT_SECRET =
  "test-secret-64-chars-long-at-least-for-hs256-validation";

const { app } = require("../src/server");

const testToken = jwt.sign(
  { userId: "test-user", email: "test@example.com", tier: "free" },
  process.env.JWT_SECRET,
  { expiresIn: "15m" },
);

const authHeader = { Authorization: `Bearer ${testToken}` };

const mockSession = {
  id: "test-001",
  title: "Test Session",
  date: new Date().toISOString(),
  preview: "Test data analysis",
  stepCount: 5,
  tokenCount: 100,
  messages: [
    { role: "user", content: "analyze data" },
    { role: "assistant", content: "done" },
  ],
};

describe("Backend API", () => {
  beforeEach(() => {
    mockUsers.length = 0;
  });

  afterAll(async () => {
    try {
      await request(app)
        .delete(`/api/sessions/${mockSession.id}`)
        .set(authHeader);
    } catch {}
  });

  describe("GET /api/health", () => {
    it("returns ok status", async () => {
      const res = await request(app).get("/api/health");
      expect(res.status).toBe(200);
      expect(res.body).toHaveProperty("status", "ok");
    });
  });

  describe("POST /api/auth/register", () => {
    it("registers a new user", async () => {
      const res = await request(app)
        .post("/api/auth/register")
        .send({ email: "new@test.com", password: "password123" });
      expect(res.status).toBe(201);
      expect(res.body).toHaveProperty("accessToken");
      expect(res.body).toHaveProperty("refreshToken");
      expect(res.body.user).toHaveProperty("tier", "free");
    });

    it("returns 400 for invalid email", async () => {
      const res = await request(app)
        .post("/api/auth/register")
        .send({ email: "not-an-email", password: "password123" });
      expect(res.status).toBe(400);
      expect(res.body).toHaveProperty("code", "VALIDATION_ERROR");
    });

    it("returns 409 for duplicate email", async () => {
      await request(app)
        .post("/api/auth/register")
        .send({ email: "dup@test.com", password: "password123" });
      const res = await request(app)
        .post("/api/auth/register")
        .send({ email: "dup@test.com", password: "password123" });
      expect(res.status).toBe(409);
      expect(res.body).toHaveProperty("code", "EMAIL_EXISTS");
    });
  });

  describe("POST /api/auth/login", () => {
    it("logs in with valid credentials", async () => {
      const email = `login_${Date.now()}@test.com`;
      await request(app)
        .post("/api/auth/register")
        .send({ email, password: "password123" });
      const res = await request(app)
        .post("/api/auth/login")
        .send({ email, password: "password123" });
      expect(res.status).toBe(200);
      expect(res.body).toHaveProperty("accessToken");
    });

    it("returns 401 for wrong password", async () => {
      const res = await request(app)
        .post("/api/auth/login")
        .send({ email: "noone@test.com", password: "wrong" });
      expect(res.status).toBe(401);
      expect(res.body).toHaveProperty("code", "INVALID_CREDENTIALS");
    });
  });

  describe("POST /api/auth/refresh", () => {
    it("returns a new access token", async () => {
      const email = `refresh_${Date.now()}@test.com`;
      const reg = await request(app)
        .post("/api/auth/register")
        .send({ email, password: "password123" });
      const res = await request(app)
        .post("/api/auth/refresh")
        .send({ refreshToken: reg.body.refreshToken });
      expect(res.status).toBe(200);
      expect(res.body).toHaveProperty("accessToken");
    });

    it("returns 401 for invalid refresh token", async () => {
      const res = await request(app)
        .post("/api/auth/refresh")
        .send({ refreshToken: "invalid" });
      expect(res.status).toBe(401);
    });
  });

  describe("GET /api/sessions", () => {
    it("returns 401 without auth", async () => {
      const res = await request(app).get("/api/sessions");
      expect(res.status).toBe(401);
    });

    it("returns sessions list", async () => {
      const res = await request(app).get("/api/sessions").set(authHeader);
      expect(res.status).toBe(200);
      expect(res.body).toHaveProperty("sessions");
      expect(Array.isArray(res.body.sessions)).toBe(true);
    });
  });

  describe("POST /api/sessions", () => {
    it("creates a session", async () => {
      const res = await request(app)
        .post("/api/sessions")
        .set(authHeader)
        .send({ session: mockSession });
      expect(res.status).toBe(200);
      expect(res.body).toHaveProperty("status", "ok");
    });

    it("returns 400 without session id", async () => {
      const res = await request(app)
        .post("/api/sessions")
        .set(authHeader)
        .send({ session: { title: "no-id" } });
      expect(res.status).toBe(400);
      expect(res.body).toHaveProperty("error");
    });
  });

  describe("GET /api/sessions/:id", () => {
    it("returns a session by id", async () => {
      const res = await request(app)
        .get(`/api/sessions/${mockSession.id}`)
        .set(authHeader);
      expect(res.status).toBe(200);
      expect(res.body.session).toBeDefined();
      expect(res.body.session.id).toBe(mockSession.id);
    });

    it("returns 404 for unknown session", async () => {
      const res = await request(app)
        .get("/api/sessions/nonexistent")
        .set(authHeader);
      expect(res.status).toBe(404);
      expect(res.body).toHaveProperty("error");
    });
  });

  describe("PATCH /api/sessions/:id/rename", () => {
    it("renames a session", async () => {
      const res = await request(app)
        .patch(`/api/sessions/${mockSession.id}/rename`)
        .set(authHeader)
        .send({ title: "Renamed Session" });
      expect(res.status).toBe(200);
      expect(res.body).toHaveProperty("status", "renamed");
    });

    it("returns 400 without title", async () => {
      const res = await request(app)
        .patch(`/api/sessions/${mockSession.id}/rename`)
        .set(authHeader)
        .send({});
      expect(res.status).toBe(400);
      expect(res.body).toHaveProperty("error");
    });

    it("returns not_found for unknown session", async () => {
      const res = await request(app)
        .patch("/api/sessions/nonexistent/rename")
        .set(authHeader)
        .send({ title: "New" });
      expect(res.status).toBe(200);
      expect(res.body).toHaveProperty("status", "not_found");
    });
  });

  describe("DELETE /api/sessions/:id", () => {
    it("deletes a session", async () => {
      const res = await request(app)
        .delete(`/api/sessions/${mockSession.id}`)
        .set(authHeader);
      expect(res.status).toBe(200);
      expect(res.body).toHaveProperty("status", "deleted");
    });

    it("returns not_found for already deleted", async () => {
      const res = await request(app)
        .delete("/api/sessions/nonexistent")
        .set(authHeader);
      expect(res.status).toBe(200);
      expect(res.body).toHaveProperty("status", "not_found");
    });
  });

  describe("POST /api/feedback validation", () => {
    it("returns 400 for invalid rating", async () => {
      const res = await request(app)
        .post("/api/feedback")
        .set(authHeader)
        .send({ rating: "invalid" });
      expect(res.status).toBe(400);
      expect(res.body).toHaveProperty("code", "INVALID_RATING");
    });

    it("accepts valid feedback", async () => {
      const res = await request(app)
        .post("/api/feedback")
        .set(authHeader)
        .send({
          rating: "helpful",
          messageId: "msg-1",
          comment: "works great",
        });
      expect(res.status).toBe(200);
      expect(res.body).toHaveProperty("status", "ok");
    });
  });

  describe("POST /api/agent/complete validation", () => {
    it("returns 400 for empty messages", async () => {
      const res = await request(app)
        .post("/api/agent/complete")
        .set(authHeader)
        .send({ messages: [] });
      expect(res.status).toBe(400);
      expect(res.body).toHaveProperty("code", "INVALID_MESSAGES");
    });
  });

  describe("POST /api/agent/stream validation", () => {
    it("returns 400 for empty messages", async () => {
      const res = await request(app)
        .post("/api/agent/stream")
        .set(authHeader)
        .send({ messages: [] });
      expect(res.status).toBe(400);
      expect(res.body).toHaveProperty("code", "INVALID_MESSAGES");
    });
  });

  describe("POST /api/agent/tool-result", () => {
    it("returns 400 without requestId", async () => {
      const res = await request(app)
        .post("/api/agent/tool-result")
        .set(authHeader)
        .send({ toolCallId: "tc-1", result: {} });
      expect(res.status).toBe(400);
      expect(res.body).toHaveProperty("code", "INVALID_TOOL_RESULT");
    });

    it("returns 400 without toolCallId", async () => {
      const res = await request(app)
        .post("/api/agent/tool-result")
        .set(authHeader)
        .send({ requestId: "req-1", result: {} });
      expect(res.status).toBe(400);
      expect(res.body).toHaveProperty("code", "INVALID_TOOL_RESULT");
    });

    it("returns 404 for unknown request", async () => {
      const res = await request(app)
        .post("/api/agent/tool-result")
        .set(authHeader)
        .send({ requestId: "nonexistent", toolCallId: "tc-1", result: {} });
      expect(res.status).toBe(404);
      expect(res.body).toHaveProperty("code", "TOOL_NOT_FOUND");
    });

    it("returns 404 for unknown toolCallId", async () => {
      const res = await request(app)
        .post("/api/agent/tool-result")
        .set(authHeader)
        .send({ requestId: "req-1", toolCallId: "nonexistent", result: {} });
      expect(res.status).toBe(404);
      expect(res.body).toHaveProperty("code", "TOOL_NOT_FOUND");
    });
  });
});
