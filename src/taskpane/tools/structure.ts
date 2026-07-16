/**
 * structure.ts — Structure-инструменты агента (категория S в docs/03-TOOLS-SPEC.md §1).
 *
 * Неделя 3 Фазы 1.
 *
 * Инструменты:
 *   S1 manageSheets       — add/rename/delete/copy листов.
 *   S2 manageTable        — list/create/delete именованных таблиц.
 *   S3 createPivotTable   — сводная: rows/columns/values/filters.
 *   S4 createChart        — график: type/source/title/position.
 *   S5 freezePanes        — закрепить заголовки/первую колонку.
 */
import { defineTool, toolRegistry, type ToolResult } from "./registry";
import { getRangeSafe } from "./address-helper";

// ============================================================================
// S1. manageSheets — add/rename/delete/copy
// (docs/03-TOOLS-SPEC.md §1 S1)
// ============================================================================

export const manageSheetsTool = defineTool({
  name: "manageSheets",
  description: `Управление листами книги: add (создать), rename (переименовать), delete (удалить), copy (копировать).
  - add: name (имя нового листа), position (опционально, 0-based)
  - rename: name (текущее имя), newName (новое имя)
  - delete: name (имя листа для удаления) — удаление необратимо, требует подтверждения
  - copy: name (исходный лист), newName (имя копии)
Используй для "создай лист", "переименуй лист", "удали лист", "скопируй лист".`,
  parameters: {
    type: "object",
    properties: {
      action: {
        type: "string",
        enum: ["add", "rename", "delete", "copy"],
        description: "Действие с листом",
      },
      name: {
        type: "string",
        description: "Имя листа (для add/rename/delete/copy)",
      },
      newName: {
        type: "string",
        description: "Новое имя (для rename/copy)",
      },
      position: {
        type: "number",
        description: "Позиция листа 0-based (для add, опционально)",
      },
    },
    required: ["action", "name"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: () => 0,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const action = String(args.action ?? "");
    const name = String(args.name ?? "");
    if (!action || !name) {
      return {
        ok: false,
        summary: "action и name обязательны",
        error: {
          code: "MISSING_ARGS",
          message: "action и name обязательны",
          retryable: false,
        },
      };
    }

    return Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;

      switch (action) {
        case "add": {
          const sheet = sheets.add(name);
          const position =
            typeof args.position === "number" ? args.position : undefined;
          if (position !== undefined) sheet.position = position;
          sheet.load("name, position");
          await context.sync();
          return {
            ok: true,
            summary: `Создан лист "${sheet.name}"`,
            data: { name: sheet.name, position: sheet.position },
          };
        }

        case "rename": {
          const newName = String(args.newName ?? "");
          if (!newName) {
            return {
              ok: false,
              summary: "newName обязателен для rename",
              error: {
                code: "MISSING_NEW_NAME",
                message: "newName обязателен",
                retryable: false,
              },
            };
          }
          const sheet = sheets.getItem(name);
          sheet.name = newName;
          await context.sync();
          return {
            ok: true,
            summary: `Лист "${name}" переименован в "${newName}"`,
            data: { oldName: name, newName },
          };
        }

        case "delete": {
          const sheet = sheets.getItem(name);
          sheet.delete();
          await context.sync();
          return {
            ok: true,
            summary: `Лист "${name}" удалён`,
            data: { deleted: name },
          };
        }

        case "copy": {
          const newName = String(args.newName ?? "");
          if (!newName) {
            return {
              ok: false,
              summary: "newName обязателен для copy",
              error: {
                code: "MISSING_NEW_NAME",
                message: "newName обязателен",
                retryable: false,
              },
            };
          }
          const source = sheets.getItem(name);
          const copy = source.copy("After", source);
          copy.name = newName;
          await context.sync();
          return {
            ok: true,
            summary: `Лист "${name}" скопирован в "${newName}"`,
            data: { source: name, newName },
          };
        }

        default:
          return {
            ok: false,
            summary: `Неизвестное действие: ${action}`,
            error: {
              code: "INVALID_ACTION",
              message: "action должен быть add/rename/delete/copy",
              retryable: false,
            },
          };
      }
    });
  },
});

toolRegistry.registerDefinition(manageSheetsTool);

// ============================================================================
// S2. manageTable — list/create/delete именованных таблиц
// (docs/03-TOOLS-SPEC.md §1 S2)
// ============================================================================

export const manageTableTool = defineTool({
  name: "manageTable",
  description: `Управление именованными таблицами: list (список), create (создать), delete (удалить).
  - list: без параметров — возвращает список всех таблиц активного листа
  - create: address (диапазон), hasHeaders (по умолчанию true), tableName (опционально), style (опционально)
  - delete: tableName (имя таблицы для удаления)
Используй для "создай таблицу", "удали таблицу", "покажи таблицы".`,
  parameters: {
    type: "object",
    properties: {
      action: {
        type: "string",
        enum: ["list", "create", "delete"],
        description: "Действие с таблицей",
      },
      address: {
        type: "string",
        description: 'Адрес диапазона для create: "A1:D100"',
      },
      hasHeaders: {
        type: "boolean",
        description: "Первая строка — заголовки (для create)",
      },
      tableName: {
        type: "string",
        description: "Имя таблицы (для create/delete)",
      },
      style: {
        type: "string",
        description: "Стиль таблицы (для create, опционально)",
      },
    },
    required: ["action"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: () => 0,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const action = String(args.action ?? "");
    if (!action) {
      return {
        ok: false,
        summary: "action обязателен",
        error: {
          code: "MISSING_ACTION",
          message: "action обязателен",
          retryable: false,
        },
      };
    }

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      if (action === "list") {
        const tables = sheet.tables;
        tables.load("items/name, items/range/address");
        await context.sync();
        const list = tables.items.map((t) => ({
          name: t.name,
          address: (t as unknown as { range: { address: string } }).range
            .address,
        }));
        return {
          ok: true,
          summary:
            list.length > 0
              ? `Найдено таблиц: ${list.length}`
              : "Нет таблиц на активном листе",
          data: { tables: list, count: list.length },
        };
      }

      if (action === "create") {
        const address = String(args.address ?? "");
        if (!address) {
          return {
            ok: false,
            summary: "address обязателен для create",
            error: {
              code: "MISSING_ADDRESS",
              message: "address обязателен",
              retryable: false,
            },
          };
        }
        const hasHeaders = args.hasHeaders !== false;
        const range = sheet.getRange(address);
        const table = sheet.tables.add(range, hasHeaders);
        const tableName =
          typeof args.tableName === "string" ? args.tableName : undefined;
        if (tableName) table.name = tableName;
        const style = typeof args.style === "string" ? args.style : undefined;
        if (style) table.style = style;
        table.load("name, style");
        await context.sync();
        const tableAddr =
          (table as unknown as { range: { address: string } }).range?.address ??
          address;
        return {
          ok: true,
          summary: `Создана таблица "${table.name}" (${tableAddr})`,
          data: {
            name: table.name,
            address: tableAddr,
            style: table.style,
          },
        };
      }

      if (action === "delete") {
        const tableName = String(args.tableName ?? "");
        if (!tableName) {
          return {
            ok: false,
            summary: "tableName обязателен для delete",
            error: {
              code: "MISSING_TABLE_NAME",
              message: "tableName обязателен",
              retryable: false,
            },
          };
        }
        sheet.tables.getItem(tableName).delete();
        await context.sync();
        return {
          ok: true,
          summary: `Таблица "${tableName}" удалена`,
          data: { deleted: tableName },
        };
      }

      return {
        ok: false,
        summary: `Неизвестное действие: ${action}`,
        error: {
          code: "INVALID_ACTION",
          message: "action должен быть list/create/delete",
          retryable: false,
        },
      };
    });
  },
});

toolRegistry.registerDefinition(manageTableTool);

// ============================================================================
// S3. createPivotTable — сводная таблица (без as any)
// (docs/03-TOOLS-SPEC.md §1 S3, §2 S3)
// ============================================================================

const AGG_MAP: Record<string, string> = {
  sum: "Sum",
  count: "Count",
  average: "Average",
  max: "Max",
  min: "Min",
  product: "Product",
  countNums: "CountNumbers",
  stdDev: "StandardDeviation",
  var: "Variance",
};

export const createPivotTableTool = defineTool({
  name: "createPivotTable",
  description: `Создаёт сводную таблицу из исходного диапазона.
Параметры:
  - sourceAddress: исходный диапазон данных
  - destinationAddress: левая верхняя ячейка для сводной
  - rows: массив имён колонок для строк (row labels)
  - columns: массив имён колонок для колонок (column labels, опционально)
  - values: массив полей [{column: "Имя", agg: "sum"|"count"|"average"|"max"|"min"}]
  - filters: массив имён колонок для фильтров (опционально)
  - name: имя сводной таблицы
Используй для "сделай сводную", "сумма продаж по месяцам", "средний чек по менеджерам".`,
  parameters: {
    type: "object",
    properties: {
      sourceAddress: {
        type: "string",
        description: 'Исходный диапазон: "A1:D100"',
      },
      destinationAddress: {
        type: "string",
        description: 'Левая верхняя ячейка: "F1"',
      },
      name: { type: "string", description: "Имя сводной таблицы" },
      rows: {
        type: "array",
        items: { type: "string" },
        description: "Колонки для строк",
      },
      columns: {
        type: "array",
        items: { type: "string" },
        description: "Колонки для колонок (опционально)",
      },
      values: {
        type: "array",
        items: {
          type: "object",
          properties: {
            column: { type: "string", description: "Имя колонки" },
            agg: {
              type: "string",
              enum: Object.keys(AGG_MAP),
              description: "Агрегация: sum/count/average/max/min",
            },
          },
          required: ["column", "agg"],
        },
        description: "Поля значений",
      },
      filters: {
        type: "array",
        items: { type: "string" },
        description: "Колонки для фильтров (опционально)",
      },
    },
    required: ["sourceAddress", "destinationAddress", "name", "rows", "values"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: () => 0,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const sourceAddress = String(args.sourceAddress ?? "");
    const destinationAddress = String(args.destinationAddress ?? "");
    const name = String(args.name ?? "");
    if (!sourceAddress || !destinationAddress || !name) {
      return {
        ok: false,
        summary: "sourceAddress, destinationAddress и name обязательны",
        error: {
          code: "MISSING_ARGS",
          message: "Обязательные параметры отсутствуют",
          retryable: false,
        },
      };
    }
    const rows = (args.rows ?? []) as string[];
    const valuesList = (args.values ?? []) as { column: string; agg: string }[];
    const columns = (args.columns ?? []) as string[];
    const filters = (args.filters ?? []) as string[];

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const sourceRange = sheet.getRange(sourceAddress);

      let destinationRange: Excel.Range;
      try {
        destinationRange = sheet.getRange(destinationAddress);
      } catch {
        return {
          ok: false,
          summary: "Неверный destinationAddress",
          error: {
            code: "INVALID_DESTINATION",
            message: "Проверьте destinationAddress",
            retryable: false,
          },
        };
      }

      const pivotTable = sheet.pivotTables.add(
        name,
        sourceRange,
        destinationRange,
      );

      for (const row of rows) {
        try {
          const hierarchy = pivotTable.hierarchies.getItem(row);
          pivotTable.rowHierarchies.add(hierarchy);
        } catch {
          return {
            ok: false,
            summary: `Колонка "${row}" не найдена в источнике`,
            error: {
              code: "COLUMN_NOT_FOUND",
              message: "Проверьте имена колонок",
              retryable: false,
            },
          };
        }
      }

      for (const col of columns) {
        try {
          const hierarchy = pivotTable.hierarchies.getItem(col);
          pivotTable.columnHierarchies.add(hierarchy);
        } catch {
          return {
            ok: false,
            summary: `Колонка "${col}" не найдена в источнике`,
            error: {
              code: "COLUMN_NOT_FOUND",
              message: `Колонка "${col}" не найдена`,
              retryable: false,
            },
          };
        }
      }

      for (const v of valuesList) {
        try {
          const hierarchy = pivotTable.hierarchies.getItem(v.column);
          const dataField = pivotTable.dataFields.add(hierarchy);
          const agg = AGG_MAP[v.agg];
          if (agg) {
            (dataField as unknown as { summarizeBy: string }).summarizeBy = agg;
          }
        } catch {
          return {
            ok: false,
            summary: `Ошибка добавления значения "${v.column}"`,
            error: {
              code: "VALUE_ERROR",
              message: `Колонка "${v.column}" не найдена`,
              retryable: false,
            },
          };
        }
      }

      for (const f of filters) {
        try {
          const hierarchy = pivotTable.hierarchies.getItem(f);
          pivotTable.filterHierarchies.add(hierarchy);
        } catch {
          return {
            ok: false,
            summary: `Колонка фильтра "${f}" не найдена`,
            error: {
              code: "COLUMN_NOT_FOUND",
              message: `Колонка "${f}" не найдена`,
              retryable: false,
            },
          };
        }
      }

      pivotTable.load("name");
      await context.sync();

      return {
        ok: true,
        summary: `Создана сводная таблица "${pivotTable.name}"`,
        data: {
          name: pivotTable.name,
          rows,
          columns,
          values: valuesList,
          filters,
        },
      };
    });
  },
});

toolRegistry.registerDefinition(createPivotTableTool);

// ============================================================================
// S4. createChart — график
// (docs/03-TOOLS-SPEC.md §1 S4, §2 S4)
// ============================================================================

const CHART_TYPES = [
  "ColumnClustered",
  "ColumnStacked",
  "ColumnStacked100",
  "BarClustered",
  "BarStacked",
  "BarStacked100",
  "Line",
  "LineMarkers",
  "LineStacked",
  "LineStackedMarkers",
  "Pie",
  "Doughnut",
  "PieOfPie",
  "BarOfPie",
  "Area",
  "AreaStacked",
  "AreaStacked100",
  "XYScatter",
  "ScatterSmoothLine",
  "ScatterSmoothLineMarkers",
  "Bubble",
];

export const createChartTool = defineTool({
  name: "createChart",
  description: `Создаёт график/диаграмму из диапазона данных.
Типы: ColumnClustered, ColumnStacked, BarClustered, BarStacked, Line, Pie, Doughnut, Area, XYScatter, Bubble и другие.
Позиция: newSheet (на отдельном листе "График_...") или Object (на текущем листе).
Используй для "построй график", "диаграмма по месяцам", "круговая диаграмма".`,
  parameters: {
    type: "object",
    properties: {
      address: { type: "string", description: 'Диапазон данных: "A1:D10"' },
      chartType: {
        type: "string",
        enum: CHART_TYPES,
        description: "Тип диаграммы",
      },
      title: {
        type: "string",
        description: "Заголовок диаграммы (опционально)",
      },
      position: {
        type: "string",
        enum: ["Object", "newSheet"],
        description: "Object — на листе, newSheet — на отдельном листе",
      },
    },
    required: ["address", "chartType"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: () => 0,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const address = String(args.address ?? "");
    const chartType = String(args.chartType ?? "");
    if (!address || !chartType) {
      return {
        ok: false,
        summary: "address и chartType обязательны",
        error: {
          code: "MISSING_ARGS",
          message: "address и chartType обязательны",
          retryable: false,
        },
      };
    }
    const title = typeof args.title === "string" ? args.title : undefined;
    const position = String(args.position ?? "Object");

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = getRangeSafe(context, address);
      const chart = sheet.charts.add(
        chartType as Excel.ChartType,
        range,
        "Auto",
      );

      if (title) {
        chart.title.text = title;
        chart.title.visible = true;
      }

      if (position === "newSheet") {
        const newSheet = context.workbook.worksheets.add(
          `График_${Date.now()}`,
        );
        const copiedChart = newSheet.charts.add(
          chartType as Excel.ChartType,
          range,
          "Auto",
        );
        if (title) {
          copiedChart.title.text = title;
          copiedChart.title.visible = true;
        }
        sheet.charts.getItem(chart.name).delete();
        copiedChart.load("id, name");
        await context.sync();
        return {
          ok: true,
          summary: `Создан график "${copiedChart.name}" на новом листе`,
          data: {
            chartId: copiedChart.id,
            name: copiedChart.name,
            position: "newSheet",
          },
        };
      }

      chart.load("id, name");
      await context.sync();
      return {
        ok: true,
        summary: `Создан график "${chart.name}" на листе "${sheet.name}"`,
        data: {
          chartId: chart.id,
          name: chart.name,
          position: "Object",
          sheetName: sheet.name,
        },
      };
    });
  },
});

toolRegistry.registerDefinition(createChartTool);

// ============================================================================
// S5. freezePanes — закрепить заголовки
// (docs/03-TOOLS-SPEC.md §1 S5)
// ============================================================================

export const freezePanesTool = defineTool({
  name: "freezePanes",
  description: `Закрепляет области листа: первую строку (freezeRow), первую колонку (freezeColumn), или по конкретной ячейке (at).
  - "firstRow" — закрепляет первую строку
  - "firstColumn" — закрепляет первую колонку
  - "A2" / "B2" — закрепляет всё выше и левее указанной ячейки
  - "none" — отменяет закрепление
Используй для "закрепи шапку", "закрепи первую строку", "закрепи колонку A".`,
  parameters: {
    type: "object",
    properties: {
      target: {
        type: "string",
        description:
          '"firstRow", "firstColumn", адрес ячейки ("A2"), или "none" для отмены',
      },
      sheetName: {
        type: "string",
        description: "Имя листа (опционально, по умолчанию активный)",
      },
    },
    required: ["target"],
  },
  riskLevel: "safe",
  requiresUndo: false,
  estimateCells: () => 0,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const target = String(args.target ?? "");
    if (!target) {
      return {
        ok: false,
        summary: "target обязателен",
        error: {
          code: "MISSING_TARGET",
          message: "target обязателен",
          retryable: false,
        },
      };
    }

    return Excel.run(async (context) => {
      const sheet =
        typeof args.sheetName === "string"
          ? context.workbook.worksheets.getItem(args.sheetName)
          : context.workbook.worksheets.getActiveWorksheet();
      const fp = sheet.freezePanes;

      if (target === "none") {
        fp.unfreeze();
        await context.sync();
        return {
          ok: true,
          summary: "Закрепление областей отменено",
          data: { target: "none" },
        };
      }

      if (target === "firstRow") {
        fp.freezeRows(1);
        await context.sync();
        return {
          ok: true,
          summary: `Закреплена первая строка на листе "${sheet.name}"`,
          data: { target: "firstRow" },
        };
      }

      if (target === "firstColumn") {
        fp.freezeColumns(1);
        await context.sync();
        return {
          ok: true,
          summary: `Закреплена первая колонка на листе "${sheet.name}"`,
          data: { target: "firstColumn" },
        };
      }

      const cell = sheet.getRange(target);
      cell.load("address");
      await context.sync();
      fp.freezeAt(cell);
      await context.sync();
      return {
        ok: true,
        summary: `Закрепление по ячейке ${cell.address} на листе "${sheet.name}"`,
        data: { target: cell.address },
      };
    });
  },
});

toolRegistry.registerDefinition(freezePanesTool);
