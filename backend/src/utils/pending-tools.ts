interface PendingToolCall {
  id: string;
  resolve: (result: unknown) => void;
  reject: (error: Error) => void;
  timeout: NodeJS.Timeout;
}

interface PendingRequest {
  tools: Map<string, PendingToolCall>;
  createdAt: number;
  stepCount: number;
  lastToolSig: string | null;
  sameSigCount: number;
}

const pending = new Map<string, PendingRequest>();

export function registerRequest(requestId: string): void {
  pending.set(requestId, {
    tools: new Map(),
    createdAt: Date.now(),
    stepCount: 0,
    lastToolSig: null,
    sameSigCount: 0,
  });
}

export function registerToolCall(
  requestId: string,
  toolCallId: string,
  name: string,
  args: string,
  timeoutMs: number,
): Promise<unknown> {
  const req = pending.get(requestId);
  if (!req) throw new Error(`Unknown request: ${requestId}`);

  req.stepCount++;

  const sig = `${name}:${args}`;
  if (req.lastToolSig === sig) {
    req.sameSigCount++;
    if (req.sameSigCount >= 3) {
      return Promise.reject(new Error("LOOP_DETECTED"));
    }
  } else {
    req.lastToolSig = sig;
    req.sameSigCount = 1;
  }

  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      req.tools.delete(toolCallId);
      reject(new Error("TOOL_TIMEOUT"));
    }, timeoutMs);

    req.tools.set(toolCallId, { id: toolCallId, resolve, reject, timeout });
  });
}

export function resolveToolCall(
  requestId: string,
  toolCallId: string,
  result: unknown,
): boolean {
  const req = pending.get(requestId);
  if (!req) return false;
  const tool = req.tools.get(toolCallId);
  if (!tool) return false;

  clearTimeout(tool.timeout);
  req.tools.delete(toolCallId);
  tool.resolve(result);
  return true;
}

export function cleanupRequest(requestId: string): void {
  const req = pending.get(requestId);
  if (!req) return;
  for (const tool of req.tools.values()) {
    clearTimeout(tool.timeout);
    tool.reject(new Error("REQUEST_CANCELLED"));
  }
  pending.delete(requestId);
}

export function getStepCount(requestId: string): number {
  return pending.get(requestId)?.stepCount ?? 0;
}
