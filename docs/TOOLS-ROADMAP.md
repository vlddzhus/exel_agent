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

## Tier 2 — следующие (сделать в ближайшее время)

| Инструмент | Файл | Описание |
|---|---|---|
| `findAndReplace` | (новый) | Поиск и замена по тексту/значениям |
| `addComment` / `getComments` / `deleteComment` | (новый) | Комментарии к ячейкам |
| `createNamedRange` / `listNamedRanges` / `deleteNamedRange` | (новый) | Именованные диапазоны |

## Tier 3 — скоро

| Инструмент | Файл | Описание |
|---|---|---|
| `transposeRange` | (новый) | Транспонирование (строки↔колонки) |
| `groupRows` / `groupColumns` / `ungroup` | (новый) | Группировка/структура |
| `addSubtotal` | (новый) | Промежуточные итоги |
| `addSparklines` | (новый) | Спарклайны (мини-графики) |

## Tier 4 — доделать

| Инструмент | Файл | Описание |
|---|---|---|
| `addHyperlink` | (новый) | Гиперссылки |
| `setRowHeight` | (новый) | Высота строк |

## Как добавить новый инструмент

1. Создать файл `src/taskpane/tools/<name>.ts`
2. Использовать `defineTool({...})` с riskLevel/requiresUndo/estimateCells/execute
3. Вызвать `toolRegistry.registerDefinition(tool)`
4. Импортировать в `src/taskpane/tools/index.ts`
5. Обновить `agent-prompt.ts` — добавить в список инструментов и алгоритм
6. Обновить этот файл
