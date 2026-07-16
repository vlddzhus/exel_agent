/**
 * UndoManager — data-safety core for Excel AI Agent.
 *
 * Features (Batch 3):
 *  - Chunked backups (no 50k hard limit; reads in ~15k-cell chunks with progress)
 *  - Transactional batches (group backups; rollback whole group on error)
 *  - Human-readable action descriptions for undo preview
 *  - IndexedDB persistence with localStorage fallback (survives Task Pane / Excel restart)
 *  - Auto-cleanup of backups older than 24 hours
 */
import {
  columnToLetter,
  letterToColumn,
  parseRangeAddress,
  parseCellAddress,
} from "./_shared/address";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface BackupChunk {
  startRow: number; // 1-based absolute row in sheet
  startCol: number; // 1-based absolute col index (0-based letterToColumn + 1)
  rowCount: number;
  colCount: number;
  values: unknown[][];
  formulas: unknown[][];
}

export interface BackupEntry {
  id: string;
  timestamp: number;
  toolName: string;
  address: string;
  sheetName: string;
  description: string;
  cellCount: number;
  chunks: BackupChunk[];
  transactionId?: string;
  // Фаза 1: метаданные для восстановления удалённых листов (фикс бага
  // восстановления в активный лист вместо пересоздания).
  sheetDeleted?: boolean;
  sheetPosition?: number;
}

export interface Transaction {
  id: string;
  startedAt: number;
  description: string;
  backupIds: string[];
}

export interface BackupResult {
  backupId: string | null;
  backupSkipped: boolean;
  reason?: string;
  cellCount: number;
}

export type ProgressCallback = (
  done: number,
  total: number,
  label: string,
) => void;

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const CHUNK_CELL_TARGET = 15000; // ~15k cells per chunk
const MAX_STACK_SIZE = 20; // keep last 20 undoable actions
const MAX_AGE_MS = 24 * 60 * 60 * 1000; // 24 hours
const IDB_NAME = "excel_ai_agent";
const IDB_STORE = "backups";
const LS_KEY = "excel_undo_stack";
const LS_FALLBACK_MAX = 3; // localStorage fallback: keep last 3 (size-limited)

// ---------------------------------------------------------------------------
// IndexedDB wrapper (minimal, promise-based)
// ---------------------------------------------------------------------------

function idbAvailable(): boolean {
  return typeof indexedDB !== "undefined" && indexedDB !== null;
}

function idbOpen(): Promise<IDBDatabase | null> {
  return new Promise((resolve) => {
    if (!idbAvailable()) return resolve(null);
    try {
      const req = indexedDB.open(IDB_NAME, 1);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(IDB_STORE)) {
          db.createObjectStore(IDB_STORE, { keyPath: "id" });
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
}

async function idbPutAll(entries: BackupEntry[]): Promise<void> {
  const db = await idbOpen();
  if (!db) return;
  await new Promise<void>((resolve) => {
    try {
      const tx = db.transaction(IDB_STORE, "readwrite");
      const store = tx.objectStore(IDB_STORE);
      store.clear();
      for (const e of entries) store.put(e);
      tx.oncomplete = () => resolve();
      tx.onerror = () => resolve();
    } catch {
      resolve();
    }
  });
  db.close();
}

async function idbGetAll(): Promise<BackupEntry[]> {
  const db = await idbOpen();
  if (!db) return [];
  const result = await new Promise<BackupEntry[]>((resolve) => {
    try {
      const tx = db.transaction(IDB_STORE, "readonly");
      const store = tx.objectStore(IDB_STORE);
      const req = store.getAll();
      req.onsuccess = () => resolve((req.result as BackupEntry[]) || []);
      req.onerror = () => resolve([]);
    } catch {
      resolve([]);
    }
  });
  db.close();
  return result;
}

async function idbClear(): Promise<void> {
  const db = await idbOpen();
  if (!db) return;
  await new Promise<void>((resolve) => {
    try {
      const tx = db.transaction(IDB_STORE, "readwrite");
      tx.objectStore(IDB_STORE).clear();
      tx.oncomplete = () => resolve();
      tx.onerror = () => resolve();
    } catch {
      resolve();
    }
  });
  db.close();
}

// ---------------------------------------------------------------------------
// Address helpers
// ---------------------------------------------------------------------------

/** Split "Sheet1!A1:B10" → { sheetName: "Sheet1", range: "A1:B10" }. Sheet optional. */
export function splitSheetFromAddress(address: string): {
  sheetName: string | null;
  range: string;
} {
  const m = address.match(/^(?:'[^']*'|[^!']+)!(.+)$/);
  if (m)
    return {
      sheetName: address.slice(0, address.indexOf("!")).replace(/^'|'$/g, ""),
      range: m[1],
    };
  return { sheetName: null, range: address };
}

/** Build chunk sub-range address from absolute start row/col and dimensions. */
export function chunkRangeAddress(
  startRow: number,
  startCol0: number,
  rowCount: number,
  colCount: number,
): string {
  const startCol = columnToLetter(startCol0);
  const endCol = columnToLetter(startCol0 + colCount - 1);
  return `${startCol}${startRow}:${endCol}${startRow + rowCount - 1}`;
}

/** Compute how many rows each chunk should cover for a given column count. */
export function computeRowsPerChunk(
  totalCols: number,
  target = CHUNK_CELL_TARGET,
): number {
  return Math.max(1, Math.floor(target / totalCols));
}

// ---------------------------------------------------------------------------
// UndoManager
// ---------------------------------------------------------------------------

class UndoManager {
  private stack: BackupEntry[] = [];
  private transactions = new Map<string, Transaction>();
  private counter = 0;
  private restored = false;

  constructor() {
    // Restore async; fire-and-forget — first call will await if not ready.
    this.restoreFromStorage();
  }

  private async restoreFromStorage(): Promise<void> {
    if (this.restored) return;
    this.restored = true;
    try {
      let entries = await idbGetAll();
      if (entries.length === 0) {
        // try localStorage fallback
        const raw = localStorage.getItem(LS_KEY);
        if (raw) {
          const parsed = JSON.parse(raw);
          if (Array.isArray(parsed)) entries = parsed as BackupEntry[];
        }
      }
      const now = Date.now();
      const valid = entries
        .filter(
          (e) =>
            e &&
            typeof e.timestamp === "number" &&
            now - e.timestamp < MAX_AGE_MS,
        )
        .sort((a, b) => a.timestamp - b.timestamp);
      this.stack = valid.slice(-MAX_STACK_SIZE);
      this.counter = this.stack.reduce((max, e) => {
        const num = parseInt(
          e.id.replace(/^backup_/, "").replace(/_\d+$/, ""),
          10,
        );
        return isNaN(num) ? max : Math.max(max, num);
      }, 0);
      // Rebuild transactions index
      for (const e of this.stack) {
        if (e.transactionId) {
          const tx = this.transactions.get(e.transactionId);
          if (tx) tx.backupIds.push(e.id);
        }
      }
      // If we filtered out expired entries, persist cleaned state.
      if (valid.length !== entries.length) await this.persist();
    } catch {
      // ignore — start with empty stack
    }
  }

  private async persist(): Promise<void> {
    try {
      await idbPutAll(this.stack);
      // localStorage fallback: keep last 3 (small entries only)
      const fallback = this.stack.slice(-LS_FALLBACK_MAX);
      try {
        localStorage.setItem(LS_KEY, JSON.stringify(fallback));
      } catch {
        // localStorage quota — drop oldest fallback entry and retry once
        try {
          localStorage.setItem(LS_KEY, JSON.stringify(fallback.slice(-1)));
        } catch {}
      }
    } catch {}
  }

  /**
   * Create a backup of the target range before a destructive operation.
   * Reads in chunks of ~15k cells. No hard upper limit.
   */
  async createBackup(
    address: string,
    toolName: string,
    options?: {
      sheetName?: string;
      description?: string;
      transactionId?: string;
      wholeSheet?: boolean; // for clearWorksheet / deleteWorksheet
      onProgress?: ProgressCallback;
    },
  ): Promise<BackupResult> {
    await this.restoreFromStorage();

    const sheetNameOpt = options?.sheetName;
    const wholeSheet =
      options?.wholeSheet ??
      (toolName === "clearWorksheet" || toolName === "deleteWorksheet");
    const transactionId = options?.transactionId;

    return Excel.run(async (context) => {
      const sheet = sheetNameOpt
        ? context.workbook.worksheets.getItem(sheetNameOpt)
        : context.workbook.worksheets.getActiveWorksheet();

      let baseRange: any;
      let resolvedSheetName: string;
      let rangeAddress: string;

      if (wholeSheet) {
        baseRange = sheet.getUsedRange();
        sheet.load("name, position");
        baseRange.load("address, rowCount, columnCount");
        await context.sync();
        resolvedSheetName = sheet.name;
        rangeAddress = baseRange.address;
      } else {
        const { sheetName: addrSheet, range } = splitSheetFromAddress(address);
        baseRange = sheet.getRange(range);
        sheet.load("name, position");
        baseRange.load("address, rowCount, columnCount");
        await context.sync();
        resolvedSheetName = sheet.name;
        rangeAddress = `${resolvedSheetName}!${range}`;
      }

      const totalRows = baseRange.rowCount;
      const totalCols = baseRange.columnCount;
      const cellCount = totalRows * totalCols;

      if (cellCount === 0 || totalRows < 1 || totalCols < 1) {
        return {
          backupId: null,
          backupSkipped: true,
          reason: "Empty range",
          cellCount: 0,
        };
      }

      // Determine chunk row size so each chunk ≈ CHUNK_CELL_TARGET cells.
      const rowsPerChunk = computeRowsPerChunk(totalCols);
      const totalChunks = Math.ceil(totalRows / rowsPerChunk);
      const chunks: BackupChunk[] = [];

      // Absolute origin of the range in the sheet.
      // baseRange.address like "Sheet1!A1:F50" — parse start cell.
      const { range: addrRange } = splitSheetFromAddress(baseRange.address);
      let originRow = 1;
      let originCol0 = 0;
      try {
        if (addrRange.includes(":")) {
          const pr = parseRangeAddress(addrRange);
          originRow = pr.startRow;
          originCol0 = letterToColumn(pr.startCol);
        } else {
          const pc = parseCellAddress(addrRange);
          originRow = pc.row;
          originCol0 = letterToColumn(pc.col);
        }
      } catch {}

      for (let i = 0; i < totalChunks; i++) {
        const startRowIdx = i * rowsPerChunk;
        const chunkRowCount = Math.min(rowsPerChunk, totalRows - startRowIdx);
        const subRange = baseRange
          .getCell(startRowIdx, 0)
          .getResizedRange(chunkRowCount - 1, totalCols - 1);
        subRange.load("values, formulas");
        await context.sync();

        chunks.push({
          startRow: originRow + startRowIdx,
          startCol: originCol0,
          rowCount: chunkRowCount,
          colCount: totalCols,
          values: subRange.values as unknown[][],
          formulas: subRange.formulas as unknown[][],
        });

        options?.onProgress?.(
          i + 1,
          totalChunks,
          `Backup chunk ${i + 1}/${totalChunks}`,
        );
      }

      const id = `backup_${++this.counter}_${Date.now()}`;
      const description =
        options?.description ||
        this.defaultDescription(
          toolName,
          resolvedSheetName,
          cellCount,
          address,
        );

      // Фаза 1: метаданные для корректного восстановления удалённых листов.
      // sheetDeleted=true => restore создаст новый лист через worksheets.add
      // вместо записи в активный (фикс бага потери данных).
      const sheetDeleted = toolName === "deleteWorksheet";
      const sheetPosition = (sheet as { position?: number }).position;

      const entry: BackupEntry = {
        id,
        timestamp: Date.now(),
        toolName,
        address: rangeAddress,
        sheetName: resolvedSheetName,
        description,
        cellCount,
        chunks,
        transactionId,
        sheetDeleted,
        sheetPosition,
      };

      this.stack.push(entry);
      if (this.stack.length > MAX_STACK_SIZE) this.stack.shift();

      if (transactionId) {
        let tx = this.transactions.get(transactionId);
        if (!tx) {
          tx = {
            id: transactionId,
            startedAt: Date.now(),
            description: "",
            backupIds: [],
          };
          this.transactions.set(transactionId, tx);
        }
        tx.backupIds.push(id);
      }

      await this.persist();
      return { backupId: id, backupSkipped: false, cellCount };
    });
  }

  private defaultDescription(
    toolName: string,
    sheetName: string,
    cellCount: number,
    address: string,
  ): string {
    const { range } = splitSheetFromAddress(address);
    switch (toolName) {
      case "clearWorksheet":
        return `Очищен лист «${sheetName}» (${cellCount} ячеек)`;
      case "deleteWorksheet":
        return `Удалён лист «${sheetName}» (${cellCount} ячеек)`;
      case "clearRange":
        return `Очищен диапазон ${range} (${cellCount} ячеек)`;
      case "setValues":
        return `Записаны значения в ${range} (${cellCount} ячеек)`;
      case "setFormula":
        return `Записана формула в ${range}`;
      case "fillFormula":
        return `Заполнены формулы в ${range} (${cellCount} ячеек)`;
      case "applyFormat":
        return `Применён числовой формат к ${range} (${cellCount} ячеек)`;
      case "setCellFormat":
        return `Применено форматирование к ${range} (${cellCount} ячеек)`;
      case "mergeCells":
        return `Объединены ячейки ${range}`;
      case "sortTable":
        return `Отсортирована таблица (${cellCount} ячеек)`;
      case "filterTable":
        return `Применён фильтр к таблице (${cellCount} ячеек)`;
      case "createPivotTable":
        return `Создана сводная таблица в ${range}`;
      case "createChart":
        return `Создан график по данным ${range}`;
      case "createTable":
        return `Создана таблица в ${range} (${cellCount} ячеек)`;
      case "addTableRow":
        return `Добавлена строка в таблицу`;
      case "deleteChart":
        return `Удалён график`;
      default:
        return `${toolName} → ${range} (${cellCount} ячеек)`;
    }
  }

  /**
   * Restore a single backup by id. Removes it from the stack.
   */
  async restoreBackup(
    backupId: string,
  ): Promise<{ success: boolean; error?: string; description?: string }> {
    await this.restoreFromStorage();
    const index = this.stack.findIndex((e) => e.id === backupId);
    if (index === -1)
      return { success: false, error: `Backup ${backupId} not found` };

    const entry = this.stack[index];
    this.stack.splice(index, 1);
    await this.persist();

    try {
      await this.writeBackEntry(entry);
      return { success: true, description: entry.description };
    } catch (err) {
      // Re-push on failure so user can retry
      this.stack.splice(index, 0, entry);
      await this.persist();
      return {
        success: false,
        error: err instanceof Error ? err.message : String(err),
      };
    }
  }

  /**
   * Roll back a whole transaction: restore all its backups in reverse order.
   */
  async rollbackTransaction(
    transactionId: string,
  ): Promise<{ success: boolean; restored: number; errors: string[] }> {
    await this.restoreFromStorage();
    const tx = this.transactions.get(transactionId);
    if (!tx)
      return { success: false, restored: 0, errors: ["Transaction not found"] };

    const errors: string[] = [];
    let restored = 0;
    // Reverse order — LIFO
    const ids = [...tx.backupIds].reverse();

    for (const id of ids) {
      const idx = this.stack.findIndex((e) => e.id === id);
      if (idx === -1) continue;
      const entry = this.stack[idx];
      try {
        await this.writeBackEntry(entry);
        this.stack.splice(idx, 1);
        restored++;
      } catch (err) {
        errors.push(
          `${entry.id}: ${err instanceof Error ? err.message : String(err)}`,
        );
      }
    }

    this.transactions.delete(transactionId);
    await this.persist();
    return { success: errors.length === 0, restored, errors };
  }

  private async writeBackEntry(entry: BackupEntry): Promise<void> {
    await Excel.run(async (context) => {
      let sheet: Excel.Worksheet;

      if (entry.sheetDeleted) {
        // Фаза 1 фикс: лист был удалён — пересоздаём через worksheets.add
        // с тем же именем и позицией. До этого использовался fallback на
        // активный лист, что приводило к тихой потере данных.
        sheet = context.workbook.worksheets.add(entry.sheetName);
        if (typeof entry.sheetPosition === "number") {
          sheet.position = entry.sheetPosition;
        }
      } else {
        // Лист не был удалён — ищем. Используем getItemOrNullObject
        // (getItem не бросает до sync, поэтому try/catch бесполезен).
        const maybe = context.workbook.worksheets.getItemOrNullObject(
          entry.sheetName,
        );
        await context.sync();
        if (maybe.isNullObject) {
          // Лист переименован/удалён в процессе — создаём с тем же именем,
          // чтобы данные не потерялись (но позиция уже неизвестна).
          sheet = context.workbook.worksheets.add(entry.sheetName);
        } else {
          sheet = maybe;
        }
      }

      for (const chunk of entry.chunks) {
        const addr = chunkRangeAddress(
          chunk.startRow,
          chunk.startCol,
          chunk.rowCount,
          chunk.colCount,
        );
        const range = sheet.getRange(addr);
        range.values = chunk.values as unknown[][];
        range.formulas = chunk.formulas as unknown[][];
      }
      await context.sync();
    });
  }

  // ── Transaction lifecycle ──

  beginTransaction(description = ""): string {
    const id = `tx_${++this.counter}_${Date.now()}`;
    this.transactions.set(id, {
      id,
      startedAt: Date.now(),
      description,
      backupIds: [],
    });
    return id;
  }

  async commitTransaction(transactionId: string): Promise<void> {
    this.transactions.delete(transactionId);
  }

  // ── Accessors ──

  getLatestBackup(): BackupEntry | undefined {
    return this.stack[this.stack.length - 1];
  }

  getStack(): BackupEntry[] {
    return [...this.stack];
  }

  getTransaction(id: string): Transaction | undefined {
    return this.transactions.get(id);
  }

  async clear(): Promise<void> {
    this.stack = [];
    this.transactions.clear();
    try {
      localStorage.removeItem(LS_KEY);
    } catch {}
    await idbClear();
  }
}

export const undoManager = new UndoManager();
