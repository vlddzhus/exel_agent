/**
 * format-extended.ts — Расширенные инструменты форматирования (категория F+).
 *
 * Итерация «Профессиональное форматирование». Новые инструменты:
 *   F7  setRowHeights   — явная установка высоты строк (зеркало setColumnWidths).
 *   F8  autoFitRows     — авто-высота строк по содержимому (зеркало autoFitColumns).
 *   F9  copyFormat      — копирование стиля (font/fill/border/numberFormat) между
 *                         диапазонами («Design by Example»).
 *   F10 applyNamedStyle — применение встроенного именованного стиля Excel
 *                         ("Good"/"Bad"/"Neutral"/"Input"/"Output"/...).
 *   F11 setSheetTabColor — цвет ярлычка листа для визуальной навигации.
 *
 * Обратная совместимость: все инструменты новые, существующие не трогаются.
 * Регистрация через defineTool + toolRegistry.registerDefinition (единый API).
 */
import { defineTool, toolRegistry, type ToolResult } from "./registry";
import { getRangeSafe } from "./address-helper";
import { undoManager } from "./backup";

// ============================================================================
// F7. setRowHeights — явная установка высоты строк
// (зеркало setColumnWidths из format.ts, для строк)
// ============================================================================

export const setRowHeightsTool = defineTool({
  name: "setRowHeights",
  description: `Устанавливает высоту одной или нескольких строк вручную.
Укажите address как диапазон: "A1:A1" для одной строки, "A1:A5" для строк 1-5 (все получат одинаковую высоту). Колонка в адресе игнорируется — высота применяется к строкам.
Используйте когда autoFitRows даёт слишком высокие/низкие строки, или для заголовков (обычно 25-30 для жирного заголовка).
Высота указывается в пунктах (стандартная ~15, для заголовков 25-30).`,
  parameters: {
    type: "object",
    properties: {
      address: {
        type: "string",
        description:
          'Диапазон строк: "A1:A1" (одна строка) или "A1:A5" (строки 1-5)',
      },
      height: {
        type: "number",
        description:
          "Высота в пунктах (стандартная ~15, для заголовков 25-30, для многострочного текста 30-45)",
      },
    },
    required: ["address", "height"],
  },
  riskLevel: "safe",
  requiresUndo: false,
  estimateCells: () => 0,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const address = String(args.address ?? "");
    if (!address) {
      return {
        ok: false,
        summary: "Не указан адрес",
        error: {
          code: "MISSING_ADDRESS",
          message: "address обязателен",
          retryable: false,
        },
      };
    }
    const height = Number(args.height);
    if (!Number.isFinite(height) || height <= 0) {
      return {
        ok: false,
        summary: "Некорректная высота",
        error: {
          code: "INVALID_HEIGHT",
          message: "height должен быть положительным числом",
          retryable: false,
        },
      };
    }

    return Excel.run(async (context) => {
      const range = getRangeSafe(context, address);
      range.format.rowHeight = height;
      await context.sync();

      return {
        ok: true,
        summary: `Установлена высота строк ${address} = ${height}`,
        data: { address, height },
      };
    });
  },
});

toolRegistry.registerDefinition(setRowHeightsTool);

// ============================================================================
// F8. autoFitRows — авто-высота строк по содержимому
// (зеркало autoFitColumns из format.ts)
// ============================================================================

export const autoFitRowsTool = defineTool({
  name: "autoFitRows",
  description: `Авто-высота строк по содержимому.
Можно указать диапазон или весь лист (без address).
Используй для "подгони высоту строк", "авто-высота".
Часто нужно после включения wrapText, чтобы multiline-текст отображался полностью.`,
  parameters: {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: 'Диапазон: "A1:D10" (опционально — без адреса весь лист)',
      },
    },
    required: [],
  },
  riskLevel: "safe",
  requiresUndo: false,
  estimateCells: () => 0,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const rawAddress =
        typeof args.address === "string" ? args.address : undefined;

      if (rawAddress) {
        const range = getRangeSafe(context, rawAddress);
        range.format.autofitRows();
        await context.sync();
        return {
          ok: true,
          summary: `Авто-высота строк в ${rawAddress}`,
          data: { address: rawAddress },
        };
      }

      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("isNullObject, address");
      await context.sync();
      if (!usedRange.isNullObject) {
        usedRange.format.autofitRows();
        await context.sync();
        return {
          ok: true,
          summary: `Авто-высота по всему листу "${sheet.name}"`,
          data: { sheetName: sheet.name },
        };
      }
      return {
        ok: true,
        summary: `Лист "${sheet.name}" пуст`,
        data: { sheetName: sheet.name },
      };
    });
  },
});

toolRegistry.registerDefinition(autoFitRowsTool);

// ============================================================================
// F9. copyFormat — копирование стиля между диапазонами («Design by Example»)
// ============================================================================

/** Что копировать. */
type CopyFormatAspect =
  | "font"
  | "fill"
  | "border"
  | "numberFormat"
  | "alignment";

/** Поддержанные аспекты стиля. */
const ALL_ASPECTS: CopyFormatAspect[] = [
  "font",
  "fill",
  "border",
  "numberFormat",
  "alignment",
];

export const copyFormatTool = defineTool({
  name: "copyFormat",
  description: `Копирует формат (стиль) с одного диапазона на другой — "Design by Example".
Что копировать: font (шрифт), fill (заливка), border (границы), numberFormat (числовой формат), alignment (выравнивание).
По умолчанию копирует всё. Укажите what=["font","fill"] для выборочного копирования.
Размеры диапазонов могут различаться — формат берётся с top-left ячейки source и применяется ко всему target.
Используй для "сделай как в этом диапазоне", "скопируй стиль с заголовка".`,
  parameters: {
    type: "object",
    properties: {
      sourceAddress: {
        type: "string",
        description: 'Откуда копировать стиль: "Sheet1!A1:F1" (шаблон)',
      },
      targetAddress: {
        type: "string",
        description: 'Куда применять стиль: "Sheet2!A1:F1"',
      },
      what: {
        type: "array",
        items: {
          type: "string",
          enum: ["font", "fill", "border", "numberFormat", "alignment"],
        },
        description:
          'Что копировать (по умолчанию всё). Пример: ["font","fill"]',
      },
    },
    required: ["sourceAddress", "targetAddress"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: (args: Record<string, unknown>) => {
    const t = String(args.targetAddress ?? "");
    return t.includes(":") ? 100 : 1;
  },

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const sourceAddress = String(args.sourceAddress ?? "");
    const targetAddress = String(args.targetAddress ?? "");
    if (!sourceAddress || !targetAddress) {
      return {
        ok: false,
        summary: "Не указан source или target адрес",
        error: {
          code: "MISSING_ADDRESS",
          message: "sourceAddress и targetAddress обязательны",
          retryable: false,
        },
      };
    }
    const whatRaw = Array.isArray(args.what) ? args.what : ALL_ASPECTS;
    const aspects = (whatRaw.filter((w) =>
      ALL_ASPECTS.includes(w as CopyFormatAspect),
    ) as CopyFormatAspect[]) ?? ALL_ASPECTS;
    if (aspects.length === 0) {
      return {
        ok: false,
        summary: "Не указано что копировать",
        error: {
          code: "MISSING_WHAT",
          message: 'what должен содержать хотя бы один из: font/fill/border/numberFormat/alignment',
          retryable: false,
        },
      };
    }

    return Excel.run(async (context) => {
      const target = getRangeSafe(context, targetAddress);
      target.load("rowCount, columnCount, address");
      await undoManager.createBackup(targetAddress, "copyFormat", {
        description: `Копирование стиля ${sourceAddress} → ${targetAddress}`,
      });

      // Читаем стиль с source (одним батчем load).
      const source = getRangeSafe(context, sourceAddress);
      const loadList: string[] = [];
      if (aspects.includes("font"))
        loadList.push(
          "format/font/name, format/font/size, format/font/color, format/font/bold, format/font/italic, format/font/underline, format/font/strikethrough",
        );
      if (aspects.includes("fill")) loadList.push("format/fill/color");
      if (aspects.includes("border"))
        loadList.push(
          "format/borders/EdgeTop, format/borders/EdgeBottom, format/borders/EdgeLeft, format/borders/EdgeRight",
        );
      if (aspects.includes("numberFormat"))
        loadList.push("numberFormat");
      if (aspects.includes("alignment"))
        loadList.push(
          "format/horizontalAlignment, format/verticalAlignment, format/wrapText, format/indentLevel",
        );

      if (loadList.length > 0) source.load(loadList.join(", "));
      await context.sync();

      // Применяем на target.
      const tf = target.format;
      if (aspects.includes("font")) {
        const sf = source.format.font as any;
        tf.font.name = sf.name;
        tf.font.size = sf.size;
        tf.font.color = sf.color;
        tf.font.bold = sf.bold;
        tf.font.italic = sf.italic;
        if (sf.underline !== undefined) tf.font.underline = sf.underline;
        if (sf.strikethrough !== undefined) {
          (tf.font as any).strikethrough = sf.strikethrough;
        }
      }
      if (aspects.includes("fill")) {
        (tf.fill as any).color = (source.format.fill as any).color;
      }
      if (aspects.includes("border")) {
        const edges = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];
        for (const edge of edges) {
          const src = (source.format.borders.getItem(edge as any) as any);
          const dst = tf.borders.getItem(edge as any);
          dst.style = src.style;
          dst.color = src.color;
        }
      }
      if (aspects.includes("numberFormat")) {
        target.numberFormat = (source as any).numberFormat;
      }
      if (aspects.includes("alignment")) {
        const sf = source.format as any;
        tf.horizontalAlignment = sf.horizontalAlignment;
        tf.verticalAlignment = sf.verticalAlignment;
        tf.wrapText = sf.wrapText;
        if (sf.indentLevel !== undefined) {
          (tf as any).indentLevel = sf.indentLevel;
        }
      }

      await context.sync();

      const cellCount = target.rowCount * target.columnCount;
      return {
        ok: true,
        summary: `Скопирован стиль (${aspects.join(", ")}) с ${sourceAddress} на ${target.address} (${cellCount} ячеек)`,
        data: {
          sourceAddress,
          targetAddress: target.address,
          aspects,
          cellCount,
        },
        cellsAffected: cellCount,
      };
    });
  },
});

toolRegistry.registerDefinition(copyFormatTool);

// ============================================================================
// F10. applyNamedStyle — встроенные именованные стили Excel
// ============================================================================

/**
 * Встроенные стили Excel. Имена остаются латиницей в Office.js API
 * (локализация имени работает только в UI Excel, программно — английские имена).
 */
const NAMED_STYLES = [
  "Good",
  "Bad",
  "Neutral",
  "Input",
  "Output",
  "Calculation",
  "Check Cell",
  "Explanatory Text",
  "Warning Text",
  "Total",
  "Note",
  "Hyperlink",
] as const;

export const applyNamedStyleTool = defineTool({
  name: "applyNamedStyle",
  description: `Применяет встроенный именованный стиля Excel к диапазону.
Стили: "Good" (зелёный), "Bad" (красный), "Neutral" (жёлтый), "Input" (голубой — для ввода), "Output" (оранжевый — результат), "Calculation", "Total", "Note", "Warning Text".
Используй для семантической подсветки: "отметь как Good/Bad", "выдели ячейки ввода".
Преимущество: пользователь может потом изменить стиль через UI Excel — все ячейки обновятся.`,
  parameters: {
    type: "object",
    properties: {
      address: { type: "string", description: 'Адрес: "A1:D10"' },
      style: {
        type: "string",
        enum: [...NAMED_STYLES],
        description: 'Имя стиля: "Good", "Bad", "Neutral", "Input", "Output"',
      },
    },
    required: ["address", "style"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: (args: Record<string, unknown>) => {
    const a = String(args.address ?? "");
    return a.includes(":") ? 100 : 1;
  },

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const address = String(args.address ?? "");
    if (!address) {
      return {
        ok: false,
        summary: "Не указан адрес",
        error: {
          code: "MISSING_ADDRESS",
          message: "address обязателен",
          retryable: false,
        },
      };
    }
    const style = String(args.style ?? "");
    if (!style) {
      return {
        ok: false,
        summary: "Не указан стиль",
        error: {
          code: "MISSING_STYLE",
          message: "style обязателен",
          retryable: false,
        },
      };
    }

    return Excel.run(async (context) => {
      const range = getRangeSafe(context, address);
      range.load("rowCount, columnCount, address");
      await undoManager.createBackup(address, "applyNamedStyle", {
        description: `Стиль ${style} → ${address}`,
      });

      try {
        range.style = style;
        await context.sync();
      } catch {
        // Стиль может отсутствовать в локализованной версии — fallback.
        return await fallbackStyle(context, address, style);
      }

      const cellCount = range.rowCount * range.columnCount;
      return {
        ok: true,
        summary: `Применён стиль "${style}" к ${range.address} (${cellCount} ячеек)`,
        data: { address: range.address, style, cellCount },
        cellsAffected: cellCount,
      };
    });
  },
});

toolRegistry.registerDefinition(applyNamedStyleTool);

/**
 * Fallback для applyNamedStyle: если встроенный стиль недоступен,
 * применяем эквивалентное ручное форматирование по семантике стиля.
 * Обеспечивает кросс-локальную устойчивость.
 */
async function fallbackStyle(
  context: Excel.RequestContext,
  address: string,
  style: string,
): Promise<ToolResult> {
  const range = getRangeSafe(context, address);
  range.load("rowCount, columnCount, address");

  // Семантический map: стиль → заливка + цвет шрифта + bold
  const STYLE_EQUIVALENTS: Record<
    string,
    { fill: string; fontColor: string; bold?: boolean }
  > = {
    Good: { fill: "#C6EFCE", fontColor: "#006100" },
    Bad: { fill: "#FFC7CE", fontColor: "#9C0006" },
    Neutral: { fill: "#FFEB9C", fontColor: "#9C6500" },
    Input: { fill: "#BDD7EE", fontColor: "#1F4E79" },
    Output: { fill: "#FCE4D6", fontColor: "#843C0C" },
    Calculation: { fill: "#D9E1F2", fontColor: "#1F3864", bold: true },
    "Check Cell": { fill: "#A5A5A5", fontColor: "#FFFFFF", bold: true },
    "Explanatory Text": { fill: "#FFFFFF", fontColor: "#7F7F7F", bold: false },
    "Warning Text": { fill: "#FFC7CE", fontColor: "#9C0006", bold: true },
    Total: { fill: "#A5A5A5", fontColor: "#FFFFFF", bold: true },
    Note: { fill: "#FFFFCC", fontColor: "#000000", bold: false },
    Hyperlink: { fill: "#FFFFFF", fontColor: "#0563C1" },
  };

  const equiv = STYLE_EQUIVALENTS[style];
  await context.sync();

  if (equiv) {
    const rf = range.format;
    rf.fill.color = equiv.fill;
    rf.font.color = equiv.fontColor;
    if (equiv.bold !== undefined) rf.font.bold = equiv.bold;
    await context.sync();
  }

  const cellCount = range.rowCount * range.columnCount;
  return {
    ok: true,
    summary: `Применён стиль "${style}" (fallback) к ${range.address} (${cellCount} ячеек)`,
    data: {
      address: range.address,
      style,
      cellCount,
      fallback: true,
    },
    cellsAffected: cellCount,
  };
}

// ============================================================================
// F11. setSheetTabColor — цвет ярлычка листа
// ============================================================================

export const setSheetTabColorTool = defineTool({
  name: "setSheetTabColor",
  description: `Устанавливает цвет ярлычка листа для визуальной навигации.
Используй для организации: красный — важное, синий — справочники, зелёный — данные, серый — черновики.
Сброс цвета: color="none" или "auto".`,
  parameters: {
    type: "object",
    properties: {
      sheetName: {
        type: "string",
        description: "Имя листа (опционально — по умолчанию активный лист)",
      },
      color: {
        type: "string",
        description:
          'CSS hex цвет "#FF0000" или "none"/"auto" для сброса',
      },
    },
    required: ["color"],
  },
  riskLevel: "safe",
  requiresUndo: false,
  estimateCells: () => 0,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const colorRaw = String(args.color ?? "");
    if (!colorRaw) {
      return {
        ok: false,
        summary: "Не указан цвет",
        error: {
          code: "MISSING_COLOR",
          message: "color обязателен",
          retryable: false,
        },
      };
    }
    const sheetName =
      typeof args.sheetName === "string" ? args.sheetName : undefined;

    return Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");

      // "none"/"auto" → сброс цвета (передаём undefined)
      const reset = /^(none|auto)$/i.test(colorRaw);
      if (reset) {
        (sheet as any).tabColor = "";
      } else {
        (sheet as any).tabColor = colorRaw;
      }

      await context.sync();
      return {
        ok: true,
        summary: reset
          ? `Сброшен цвет ярлычка листа "${sheet.name}"`
          : `Установлен цвет ярлычка "${sheet.name}": ${colorRaw}`,
        data: { sheetName: sheet.name, color: reset ? "none" : colorRaw },
      };
    });
  },
});

toolRegistry.registerDefinition(setSheetTabColorTool);
