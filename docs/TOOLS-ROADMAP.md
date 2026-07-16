# Tools Roadmap — все инструменты Excel AI Agent

**Легенда:** ✅ реализовано | 🔨 в работе | 📅 запланировано

---

## Tier 0 — базовые (уже работают)

| Инструмент | Файл | Статус |
|---|---|---|
| `getWorkbookOverview` | read.ts | ✅ |
| `getRange` | read.ts | ✅ |
| `getFormula` | read.ts | ✅ |
| `getRangeStats` | read.ts | ✅ |
| `detectDataTypes` | read.ts | ✅ |
| `findAnomalies` | read.ts | ✅ |
| `setValues` | write.ts | ✅ |
| `setFormula` | write.ts | ✅ |
| `fillRange` | write.ts | ✅ |
| `appendRows` | write.ts | ✅ |
| `clearRange` | write.ts | ✅ |
| `applyCellFormat` | format.ts | ✅ |
| `applyNumberFormat` | format.ts | ✅ |
| `applyConditionalFormat` | format.ts | ✅ |
| `formatAsTable` | format.ts | ✅ |
| `autoFitColumns` | format.ts | ✅ |
| `setColumnWidths` | format.ts | ✅ |
| `manageSheets` | structure.ts | ✅ |
| `manageTable` | structure.ts | ✅ |
| `createPivotTable` | structure.ts | ✅ |
| `createChart` | structure.ts | ✅ |
| `freezePanes` | structure.ts | ✅ |
| `addDataValidation` | data-validation.ts | ✅ |
| `listTables`, `createTable`, `addTableRow`, `sortTable`, `filterTable` | table-tools.ts | ✅ |
| `sortData`, `filterData`, `removeDuplicates`, `splitTextToColumns`, `normalizeText`, `lookup` | transform.ts | ✅ |

## Tier 1 — сейчас (сделаны в этой сессии)

| Инструмент | Файл | Статус |
|---|---|---|
| `manageRowsColumns` (insertRows/deleteRows/insertColumns/deleteColumns) | advanced-structure.ts | ✅ |
| `mergeCells` | advanced-structure.ts | ✅ |
| `unmergeCells` | advanced-structure.ts | ✅ |
| `manageSheetProtection` (protect/unprotect) | advanced-structure.ts | ✅ |
| `setRowHeights` (высота строк) | format-extended.ts | ✅ |
| `autoFitRows` (авто-высота строк) | format-extended.ts | ✅ |
| `copyFormat` (копирование стиля — Design by Example) | format-extended.ts | ✅ |
| `applyNamedStyle` (Excel named styles: Good/Bad/Neutral/Input/Output) | format-extended.ts | ✅ |
| `setSheetTabColor` (цвет ярлычка листа) | format-extended.ts | ✅ |
| `applyAutoDesign` (авто-дизайн таблиц по типу данных) | auto-designer.ts | ✅ |
| `applyCellFormat` расширение (fontName/underline/strikethrough/indentLevel/locked) | format.ts | ✅ |
| `applyConditionalFormat` расширение (iconSet/customFormula/duplicates) | format.ts | ✅ |
| Knowledge `excel-templates.md` (Design System с палитрами) | knowledge/sections/ | ✅ |

## Tier 2 — следующие (сделать в ближайшее время)

| Инструмент | Файл | Статус |
|---|---|---|
| `findAndReplace` | search-tools.ts | ✅ |
| `manageComments` (add/get/delete/clear) | comments-charts.ts | ✅ |
| `manageNamedRanges` (add/list/delete/get) | named-ranges.ts | ✅ |
| `manageHyperlinks` (add/get/remove) | hyperlinks.ts | ✅ |

## Tier 3 — скоро

| Инструмент | Файл | Статус |
|---|---|---|
| `groupRows` / `groupColumns` / `ungroup` / `clearOutline` | grouping.ts (manageGrouping) | ✅ |
| `addSubtotal` | (через manageGrouping + SUBTOTAL формулы) | ✅ (эмулируется) |
| `addSparklines` | — | 🚫 Office.js не предоставляет прямой API |
| `formatChart` (пост-форматирование: оси/легенда/подписи) | comments-charts.ts | ✅ |
| `manageSheetView` (gridlines/headings/zoom) | sheet-view.ts | ✅ |
| `managePageSetup` (поля/ориентация/область печати) | sheet-view.ts | ✅ |

## Tier 4 — отложено (отдельный спринт)

| Инструмент | Описание |
|---|---|
| Events (onChanged/onSelectionChanged/onSingleClicked) | Реактивный режим — отдельная архитектура (event lifecycle), не tool-pattern. Отдельный спринт. |
| Shapes (textbox/arrow) / Images | Низкий приоритет. |
| Import CSV/JSON / Export PDF | Низкий приоритет (через backend/Office.js getFileAsync). |
| Theme/palette управление | 🚫 Office.js не предоставляет. |
| Formula auditing (precedents/dependents) | Низкая ROI, ограниченный API. |

## Как добавить новый инструмент

1. Создать файл `src/taskpane/tools/<name>.ts`
2. Использовать `defineTool({...})` с riskLevel/requiresUndo/estimateCells/execute
3. Вызвать `toolRegistry.registerDefinition(tool)`
4. Импортировать в `src/taskpane/tools/index.ts`
5. Обновить `agent-prompt.ts` — добавить в список инструментов и алгоритм
6. Обновить этот файл
