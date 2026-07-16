// ponytail: rough character-based token estimation, ~3 chars/token typical
function estimateTokens(text: string): number {
  return Math.ceil(text.length / 3);
}

export function pruneMessages<T extends { role?: string; content?: string }>(
  messages: T[],
  maxTokens: number = 8000,
): T[] {
  if (messages.length === 0) return messages;

  const tokenCount = (m: T) => estimateTokens(m.content || "") + 4;

  let total = 0;
  const systemMessages = messages.filter((m) => m.role === "system");
  const nonSystem = messages.filter((m) => m.role !== "system");

  for (const m of systemMessages) total += tokenCount(m);
  for (const m of nonSystem) total += tokenCount(m);

  if (total <= maxTokens) return messages;

  const keepLast = 6;
  const keepLastMessages = nonSystem.slice(-keepLast);
  let prunedTotal = 0;
  for (const m of systemMessages) prunedTotal += tokenCount(m);
  for (const m of keepLastMessages) prunedTotal += tokenCount(m);

  if (prunedTotal <= maxTokens) return [...systemMessages, ...keepLastMessages];

  for (let i = keepLastMessages.length - 1; i >= 0; i--) {
    const msg = keepLastMessages[i];
    const t = tokenCount(msg);
    if (prunedTotal - t <= maxTokens || i === 0) {
      keepLastMessages.splice(0, i);
      break;
    }
  }

  return [...systemMessages, ...keepLastMessages];
}
