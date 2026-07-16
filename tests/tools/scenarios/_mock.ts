/**
 * _mock.ts — shared mock helpers for scenario acceptance tests.
 *
 * Provides a configurable Excel mock with address-aware ranges,
 * undoManager mock, and withPerformanceGuard mock.
 */

import {
  parseCellAddress,
  parseRangeAddress,
} from "../../../src/taskpane/tools/_shared/address";

// ===========================================================================
// Address helpers for mocks
// ===========================================================================

/** Parse "A1:C10" → { rows, cols }, or "A1" → { rows: 1, cols: 1 } */
export function parseAddressDimensions(address: string): {
  rows: number;
  cols: number;
} {
  if (!address) return { rows: 1, cols: 1 };

  // Strip sheet name if present
  const bangIdx = address.indexOf("!");
  const rangePart = bangIdx >= 0 ? address.slice(bangIdx + 1) : address;

  if (!rangePart.includes(":")) {
    // Single cell or column (e.g. "A1" or "A:A")
    if (/^[A-Za-z]+$/.test(rangePart)) {
      return { rows: 100, cols: 1 }; // whole column
    }
    return { rows: 1, cols: 1 };
  }

  try {
    const parsed = parseRangeAddress(rangePart);
    const colStart = parsed.startCol.toUpperCase().charCodeAt(0) - 65; // A=0
    const colEnd = parsed.endCol.toUpperCase().charCodeAt(0) - 65;
    const cols = Math.abs(colEnd - colStart) + 1;
    const rows = Math.abs(parsed.endRow - parsed.startRow) + 1;
    return { rows, cols };
  } catch {
    return { rows: 10, cols: 3 }; // fallback
  }
}

// ===========================================================================
// Tool state interface
// ===========================================================================

export interface ScenarioState {
  writtenValues: unknown[][];
  writtenFormulas: string[][];
  fillValues: unknown[][];
  tableCreated: boolean;
  syncCalls: number;
}

export function createScenarioState(): ScenarioState {
  return {
    writtenValues: [],
    writtenFormulas: [],
    fillValues: [],
    tableCreated: false,
    syncCalls: 0,
  };
}

// ===========================================================================
// Range factory
// ===========================================================================

export function createMockRange(
  address: string,
  state: ScenarioState,
  overrides?: { rows?: number; cols?: number },
) {
  const { rows, cols } = overrides?.rows
    ? { rows: overrides.rows, cols: overrides.cols ?? 1 }
    : parseAddressDimensions(address);

  const range: Record<string, unknown> = {
    rowCount: rows,
    columnCount: cols,
    address,
    load: jest.fn(),
    values: [] as unknown[][],
    formulas: [] as string[][],
    numberFormat: [["General"]],
    style: "Normal",
    removeDuplicates: jest
      .fn()
      .mockReturnValue({ removed: 2, load: jest.fn() }),
    textToColumns: jest.fn(),
    clear: jest.fn(),
    conditionalFormats: {
      add: jest.fn().mockImplementation(() => ({
        load: jest.fn(),
        getName: jest.fn(),
        colorScale: { criteria: {} },
        bar: { fill: { color: "#000" }, showBarOnly: false },
        top10: { rank: 10, bottom: false, percent: false },
        cellValue: {
          rule: { operator: "", formula1: "", formula2: "" },
          format: { fill: { color: "#000" }, font: { color: "#000" } },
        },
        iconSet: {
          iconSet: "ThreeTrafficLights1",
          reverseIconOrder: false,
          showIconOnly: false,
          format: { fill: { color: "#000" }, font: { color: "#000" } },
        },
        customRule: {
          formula: "",
          format: { fill: { color: "#000" }, font: { color: "#000" } },
        },
        presetCriteria: {
          rule: { type: "DuplicateValues" },
          format: { fill: { color: "#000" }, font: { color: "#000" } },
        },
      })),
    },
    format: {
      font: {
        bold: false,
        italic: false,
        size: 11,
        color: "#000000",
        name: "Calibri",
        underline: "None",
        strikethrough: false,
      },
      fill: { color: "#FFFFFF" },
      horizontalAlignment: "General",
      verticalAlignment: "Bottom",
      wrapText: false,
      indentLevel: 0,
      rowHeight: 15,
      columnWidth: 8.43,
      protection: { locked: true },
      autofitColumns: jest.fn(),
      autofitRows: jest.fn(),
      borders: {
        getItem: jest.fn().mockReturnValue({ style: "None", color: "#000000" }),
      },
    },
    sort: { apply: jest.fn() },
    getCell: (_r: number, _c: number) =>
      createMockRange(`${address}!${_r}x${_c}`, state, { rows: 1, cols: 1 }),
    getResizedRange: (_dr: number, _dc: number) =>
      createMockRange(`${address}~${_dr}x${_dc}`, state, {
        rows: _dr + 1,
        cols: _dc + 1,
      }),
  };

  // Intercept values setter
  Object.defineProperty(range, "values", {
    get: () => state.writtenValues,
    set: (v: unknown[][]) => {
      state.writtenValues = v;
    },
    configurable: true,
  });

  // Intercept formulas setter
  Object.defineProperty(range, "formulas", {
    get: () => state.writtenFormulas,
    set: (v: string[][]) => {
      state.writtenFormulas = v;
    },
    configurable: true,
  });

  return range;
}

// ===========================================================================
// Full Excel mock setup
// ===========================================================================

export interface ExcelMockConfig {
  /** Pre-set values for getRangeStats etc. */
  rangeValues?: unknown[][];
  /** Force rowCount for the first range returned (for testing). */
  forceRows?: number;
  /** Force colCount. */
  forceCols?: number;
}

export function setupExcelMock(state: ScenarioState, config?: ExcelMockConfig) {
  const syncMock = jest.fn().mockImplementation(async () => {
    state.syncCalls++;
  });

  const defaultRange = createMockRange("TestSheet!A1:Z100", state, {
    rows: config?.forceRows ?? 100,
    cols: config?.forceCols ?? 26,
  });
  if (config?.rangeValues) {
    defaultRange.values = config.rangeValues;
  }

  const usedRange = createMockRange("TestSheet!A1:E10", state, {
    rows: 10,
    cols: 5,
  });
  Object.assign(usedRange, { isNullObject: false });

  const tableMock = {
    name: "Table1",
    style: "TableStyleLight1",
    load: jest.fn(),
    delete: jest.fn(),
    getDataBodyRange: jest.fn().mockReturnValue({
      load: jest.fn(),
      clear: jest.fn(),
    }),
  };

  const sheet = {
    name: "TestSheet",
    load: jest.fn(),
    tabColor: "",
    getRange: jest.fn().mockImplementation((addr: string) => {
      return createMockRange(addr, state);
    }),
    getUsedRangeOrNullObject: jest
      .fn()
      .mockReturnValue(Object.assign({}, usedRange, { isNullObject: false })),
    tables: {
      add: jest.fn().mockImplementation(() => {
        state.tableCreated = true;
        return tableMock;
      }),
      getItem: jest.fn().mockReturnValue(tableMock),
      load: jest.fn(),
    },
    autoFilter: {
      apply: jest.fn(),
      clear: jest.fn(),
    },
    delete: jest.fn(),
  };

  (globalThis as { Excel?: unknown }).Excel = {
    ClearApplyTo: { contents: "Contents", formats: "Formats" },
    CalculationMode: { manual: "manual", automatic: "automatic" },
    run: jest
      .fn()
      .mockImplementation(async (fn: (ctx: unknown) => Promise<unknown>) => {
        const ctx = {
          workbook: {
            worksheets: {
              getActiveWorksheet: jest.fn().mockReturnValue(sheet),
              getItem: jest.fn().mockReturnValue(sheet),
            },
          },
          sync: syncMock,
          runtime: { enableEvents: false },
          application: { calculationMode: "" },
        };
        return fn(ctx);
      }),
  };

  return { syncCount: syncMock, sheet };
}

// ===========================================================================
// Cleanup
// ===========================================================================

export function cleanupExcelMock() {
  delete (globalThis as { Excel?: unknown }).Excel;
}
