/**
 * _shared/performance.ts — производительность и лимиты инструментов.
 *
 * См. docs/03-TOOLS-SPEC.md §3.3, docs/02-ARCHITECTURE.md §8.
 */

import { parseRangeAddress, splitFullAddress } from "./address";

// Максимальное количество ячеек для одной операции записи.
export const SAFE_CELL_LIMIT = 50_000;

// Размер чанка для batched-чтения/записи больших диапазонов.
export const CHUNK_CELL_SIZE = 15_000;

/**
 * Выполняет callback в Excel.run с отключёнными calc + events,
 * гарантированно восстанавливая их в finally.
 */
export async function withPerformanceGuard<T>(
  callback: (context: Excel.RequestContext) => Promise<T>,
): Promise<T> {
  let result: T;
  await Excel.run(async (context) => {
    context.application.calculationMode = Excel.CalculationMode.manual;
    try {
      context.runtime.enableEvents = false;
    } catch {
      // older Excel versions: ignore
    }
    await context.sync();

    try {
      result = await callback(context);
    } finally {
      context.application.calculationMode = Excel.CalculationMode.automatic;
      try {
        context.runtime.enableEvents = true;
      } catch {
        // older Excel versions: ignore
      }
      await context.sync();
    }
  });
  return result!;
}

/**
 * Оценивает количество ячеек в адресе диапазона.
 * "A1:B10" → 20, "A1" → 1, "Лист!A1:Z100" → 2600.
 */
export function estimateRangeSize(address: string): number {
  try {
    const { rangeAddress } = splitFullAddress(address);
    if (!rangeAddress.includes(":")) return 1;
    const { startCol, startRow, endCol, endRow } =
      parseRangeAddress(rangeAddress);
    const colStart = colLetterToNum(startCol);
    const colEnd = colLetterToNum(endCol);
    const cols = Math.abs(colEnd - colStart) + 1;
    const rows = Math.abs(endRow - startRow) + 1;
    return cols * rows;
  } catch {
    return 1;
  }
}

function colLetterToNum(letter: string): number {
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.toUpperCase().charCodeAt(i) - 64);
  }
  return result;
}

/**
 * Разбивает массив на чанки заданного размера и вызывает fn для каждого чанка.
 * Возвращает aggregated-результат.
 */
export async function chunkLarge<T, R>(
  items: T[],
  chunkSize: number,
  fn: (chunk: T[], chunkIndex: number, totalChunks: number) => Promise<R>,
): Promise<R[]> {
  if (!Number.isInteger(chunkSize) || chunkSize <= 0) {
    throw new Error(`chunkSize must be positive integer, got: ${chunkSize}`);
  }
  if (items.length === 0) return [];

  const results: R[] = [];
  const totalChunks = Math.ceil(items.length / chunkSize);

  for (let i = 0, idx = 0; i < items.length; i += chunkSize, idx++) {
    const chunk = items.slice(i, i + chunkSize);
    const result = await fn(chunk, idx, totalChunks);
    results.push(result);
  }
  return results;
}

/**
 * Проверяет, что количество ячеек не превышает SAFE_CELL_LIMIT.
 * Бросает Error с понятным сообщением при превышении.
 */
export function assertCellLimit(cellCount: number, operation = "operation"): void {
  if (cellCount > SAFE_CELL_LIMIT) {
    throw new Error(
      `Превышен лимит ячеек для операции "${operation}": ${cellCount} > ${SAFE_CELL_LIMIT}. ` +
        `Разбейте задачу на части или уменьшите диапазон.`,
    );
  }
}
