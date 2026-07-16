/**
 * Тесты для _shared/performance.ts — лимиты, chunking, оценка размера диапазона.
 *
 * withPerformanceGuard тестируется отдельно в integration-тестах (требует mock
 * Excel.run с context.application.calculationMode и context.runtime.enableEvents).
 * Здесь покрываем чистые функции: estimateRangeSize, chunkLarge, assertCellLimit.
 */
import {
  SAFE_CELL_LIMIT,
  CHUNK_CELL_SIZE,
  estimateRangeSize,
  chunkLarge,
  assertCellLimit,
} from "../../src/taskpane/tools/_shared/performance";

describe("константы", () => {
  test("SAFE_CELL_LIMIT = 50 000", () => {
    expect(SAFE_CELL_LIMIT).toBe(50_000);
  });

  test("CHUNK_CELL_SIZE > 0", () => {
    expect(CHUNK_CELL_SIZE).toBeGreaterThan(0);
  });
});

// ---------------------------------------------------------------------------
// estimateRangeSize
// ---------------------------------------------------------------------------

describe("estimateRangeSize", () => {
  test("одиночная ячейка", () => {
    expect(estimateRangeSize("A1")).toBe(1);
    expect(estimateRangeSize("B5")).toBe(1);
    expect(estimateRangeSize("Лист!C10")).toBe(1);
  });

  test("простой диапазон", () => {
    expect(estimateRangeSize("A1:B10")).toBe(20); // 2 cols × 10 rows
    expect(estimateRangeSize("A1:Z100")).toBe(2600); // 26 × 100
  });

  test("диапазон с именем листа", () => {
    expect(estimateRangeSize("Лист1!A1:C3")).toBe(9);
    expect(estimateRangeSize("'Мой лист'!A1:E10")).toBe(50);
  });

  test("multi-letter колонки", () => {
    expect(estimateRangeSize("AA1:AB10")).toBe(20); // 2 cols × 10 rows
    expect(estimateRangeSize("A1:AA100")).toBe(2700); // 27 × 100
  });

  test("обратный порядок (start > end) — берём abs", () => {
    expect(estimateRangeSize("B10:A1")).toBe(20); // 2 × 10
  });

  test("невалидный адрес → 1 (не бросает)", () => {
    expect(estimateRangeSize("")).toBe(1);
    expect(estimateRangeSize("garbage")).toBe(1);
  });
});

// ---------------------------------------------------------------------------
// chunkLarge
// ---------------------------------------------------------------------------

describe("chunkLarge", () => {
  test("разбивает массив на чанки заданного размера", async () => {
    const items = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
    const sizes: number[] = [];
    await chunkLarge(items, 3, async (chunk: number[]) => {
      sizes.push(chunk.length);
      return chunk.length;
    });
    expect(sizes).toEqual([3, 3, 3, 1]);
  });

  test("вызывает fn с правильными индексами", async () => {
    const indices: number[] = [];
    await chunkLarge([1, 2, 3, 4, 5], 2, async (_: number[], idx: number) => {
      indices.push(idx);
      return idx;
    });
    expect(indices).toEqual([0, 1, 2]);
  });

  test("передаёт totalChunks", async () => {
    const totals: number[] = [];
    await chunkLarge([1, 2, 3, 4, 5, 6, 7], 3, async (_: number[], __: number, total: number) => {
      totals.push(total);
      return null;
    });
    expect(totals).toEqual([3, 3, 3]); // ceil(7/3) = 3 chunks
  });

  test("возвращает aggregated результаты в порядке", async () => {
    const items = [1, 2, 3, 4, 5];
    const results = await chunkLarge(items, 2, async (chunk: number[]) =>
      chunk.reduce((a, b) => a + b, 0),
    );
    expect(results).toEqual([3, 7, 5]); // [1+2, 3+4, 5]
  });

  test("пустой массив → пустой результат, fn не вызывается", async () => {
    const fn = jest.fn().mockResolvedValue("x");
    const results = await chunkLarge([], 10, fn);
    expect(results).toEqual([]);
    expect(fn).not.toHaveBeenCalled();
  });

  test("размер чанка больше массива → один чанк", async () => {
    const calls: number[] = [];
    await chunkLarge([1, 2, 3], 100, async (chunk: number[]) => {
      calls.push(chunk.length);
      return null;
    });
    expect(calls).toEqual([3]);
  });

  test("бросает Error для невалидного chunkSize", async () => {
    await expect(chunkLarge([1, 2], 0, async () => null)).rejects.toThrow();
    await expect(chunkLarge([1, 2], -1, async () => null)).rejects.toThrow();
    await expect(chunkLarge([1, 2], 1.5, async () => null)).rejects.toThrow();
  });

  test("обрабатывает чанки последовательно (не параллельно)", async () => {
    const order: string[] = [];
    await chunkLarge([1, 2, 3], 1, async (item: number[]) => {
      order.push(`start-${item[0]}`);
      await new Promise((r) => setTimeout(r, 10));
      order.push(`end-${item[0]}`);
      return null;
    });
    // Последовательность: каждый чанк полностью завершается до следующего
    expect(order).toEqual([
      "start-1",
      "end-1",
      "start-2",
      "end-2",
      "start-3",
      "end-3",
    ]);
  });
});

// ---------------------------------------------------------------------------
// assertCellLimit
// ---------------------------------------------------------------------------

describe("assertCellLimit", () => {
  test("не бросает для значений в пределах лимита", () => {
    expect(() => assertCellLimit(1)).not.toThrow();
    expect(() => assertCellLimit(50_000)).not.toThrow();
  });

  test("бросает Error при превышении лимита", () => {
    expect(() => assertCellLimit(50_001)).toThrow();
    expect(() => assertCellLimit(100_000)).toThrow();
  });

  test("сообщение об ошибке содержит операцию и числа", () => {
    try {
      assertCellLimit(100_000, "setValues");
      fail("Должно бросить");
    } catch (e) {
      const msg = (e as Error).message;
      expect(msg).toContain("setValues");
      expect(msg).toContain("100000");
      expect(msg).toContain("50000");
    }
  });
});
