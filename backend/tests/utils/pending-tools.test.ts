import {
  registerRequest,
  registerToolCall,
  resolveToolCall,
  cleanupRequest,
  getStepCount,
} from "../../src/utils/pending-tools";

const REQ_ID = "test-req-001";

function ignoreRejection(promise: Promise<unknown>): void {
  promise.catch(() => {});
}

describe("pending-tools", () => {
  afterEach(() => {
    cleanupRequest(REQ_ID);
  });

  describe("registerRequest", () => {
    it("registers a new request", () => {
      registerRequest(REQ_ID);
      expect(getStepCount(REQ_ID)).toBe(0);
    });
  });

  describe("registerToolCall", () => {
    it("registers and resolves a tool call", async () => {
      registerRequest(REQ_ID);
      const promise = registerToolCall(REQ_ID, "tc-1", "testTool", "{}", 5000);

      const ok = resolveToolCall(REQ_ID, "tc-1", { result: "ok" });
      expect(ok).toBe(true);

      const result = await promise;
      expect(result).toEqual({ result: "ok" });
    });

    it("rejects on timeout", async () => {
      registerRequest(REQ_ID);
      const promise = registerToolCall(REQ_ID, "tc-2", "testTool", "{}", 50);

      await expect(promise).rejects.toThrow("TOOL_TIMEOUT");
    }, 200);

    it("detects loop: same tool+args 3 times", async () => {
      registerRequest(REQ_ID);

      const p1 = ignoreRejection(
        registerToolCall(REQ_ID, "a", "sameTool", '"sameArgs"', 500),
      );
      resolveToolCall(REQ_ID, "a", "ok");

      const p2 = ignoreRejection(
        registerToolCall(REQ_ID, "b", "sameTool", '"sameArgs"', 500),
      );
      resolveToolCall(REQ_ID, "b", "ok");

      const p3 = registerToolCall(REQ_ID, "c", "sameTool", '"sameArgs"', 500);

      await expect(p3).rejects.toThrow("LOOP_DETECTED");
    });

    it("does not detect loop when args differ", async () => {
      registerRequest(REQ_ID);

      const p1 = registerToolCall(REQ_ID, "a", "sameTool", '"args1"', 500);
      resolveToolCall(REQ_ID, "a", "ok");

      const p2 = registerToolCall(REQ_ID, "b", "sameTool", '"args2"', 500);
      resolveToolCall(REQ_ID, "b", "ok");

      const p3 = registerToolCall(REQ_ID, "c", "sameTool", '"args3"', 500);
      resolveToolCall(REQ_ID, "c", "ok");

      await expect(p1).resolves.toBe("ok");
      await expect(p2).resolves.toBe("ok");
      await expect(p3).resolves.toBe("ok");
    });

    it("does not detect loop when tool name differs", async () => {
      registerRequest(REQ_ID);

      const p1 = registerToolCall(REQ_ID, "a", "toolA", '"args"', 500);
      resolveToolCall(REQ_ID, "a", "ok");

      const p2 = registerToolCall(REQ_ID, "b", "toolB", '"args"', 500);
      resolveToolCall(REQ_ID, "b", "ok");

      const p3 = registerToolCall(REQ_ID, "c", "toolC", '"args"', 500);
      resolveToolCall(REQ_ID, "c", "ok");

      await expect(p1).resolves.toBe("ok");
      await expect(p2).resolves.toBe("ok");
      await expect(p3).resolves.toBe("ok");
    });

    it("throws for unknown request", () => {
      cleanupRequest(REQ_ID);
      expect(() => registerToolCall(REQ_ID, "x", "tool", "{}", 500)).toThrow(
        "Unknown request",
      );
    });
  });

  describe("resolveToolCall", () => {
    it("returns false for unknown request", () => {
      cleanupRequest(REQ_ID);
      const ok = resolveToolCall(REQ_ID, "x", {});
      expect(ok).toBe(false);
    });

    it("returns false for unknown toolCallId", () => {
      registerRequest(REQ_ID);
      const ok = resolveToolCall(REQ_ID, "nonexistent", {});
      expect(ok).toBe(false);
    });

    it("returns true once, false the second time", async () => {
      registerRequest(REQ_ID);
      const promise = registerToolCall(REQ_ID, "tc-id", "tool", "{}", 500);

      expect(resolveToolCall(REQ_ID, "tc-id", "ok")).toBe(true);
      expect(resolveToolCall(REQ_ID, "tc-id", "again")).toBe(false);

      await expect(promise).resolves.toBe("ok");
    });
  });

  describe("cleanupRequest", () => {
    it("cleans up and rejects pending tools", async () => {
      registerRequest(REQ_ID);
      const promise = registerToolCall(REQ_ID, "tc-x", "tool", "{}", 5000);
      ignoreRejection(promise);

      cleanupRequest(REQ_ID);

      await expect(promise).rejects.toThrow("REQUEST_CANCELLED");
      expect(getStepCount(REQ_ID)).toBe(0);
    });

    it("is idempotent", () => {
      registerRequest(REQ_ID);
      cleanupRequest(REQ_ID);
      cleanupRequest(REQ_ID);
    });
  });

  describe("getStepCount", () => {
    it("increments on each registerToolCall", () => {
      registerRequest(REQ_ID);
      expect(getStepCount(REQ_ID)).toBe(0);

      ignoreRejection(registerToolCall(REQ_ID, "a", "tool", "{}", 500));
      expect(getStepCount(REQ_ID)).toBe(1);

      ignoreRejection(registerToolCall(REQ_ID, "b", "tool", "{}", 500));
      expect(getStepCount(REQ_ID)).toBe(2);
    });

    it("returns 0 after cleanup", () => {
      registerRequest(REQ_ID);
      ignoreRejection(registerToolCall(REQ_ID, "a", "tool", "{}", 500));
      cleanupRequest(REQ_ID);
      expect(getStepCount(REQ_ID)).toBe(0);
    });
  });
});
