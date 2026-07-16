import { toolRegistry } from './registry';

// ──────────────────────────────────────────────
// EMBEDDED KNOWLEDGE BASE (always available)
// Hybrid approach: embedded constants + fetch from dist/knowledge/
// ──────────────────────────────────────────────

interface KnowledgeEntry {
  keywords: string[];
  content: string;
}

const EMBEDDED_KNOWLEDGE: Record<string, KnowledgeEntry[]> = {
  'excel-formulas': [
    {
      keywords: ['multiplication', 'operator', '*', 'cell reference', 'умножение', 'b8b6', 'name error'],
      content: `## CRITICAL: Cell References MUST have operators between them
✅ =B8*B6          → B8 multiplied by B6
❌ =B8B6           → #NAME! error (missing *)
✅ =B8*B6+B9*B7    → Correct compound formula
❌ =B8B6+B9B7      → #NAME! error (both pairs missing *)
✅ =(B8-B6)*B9/B7  → Correct with parentheses`,
    },
    {
      keywords: ['sum', 'сумм', 'total', 'итог'],
      content: `## SUM / СУММ
English: =SUM(range)        e.g., =SUM(D3:D10)
Russian: =СУММ(диапазон)    e.g., =СУММ(D3:D10)`,
    },
    {
      keywords: ['average', 'срзнач', 'mean', 'среднее'],
      content: `## AVERAGE / СРЗНАЧ
English: =AVERAGE(range)        e.g., =AVERAGE(D3:D10)
Russian: =СРЗНАЧ(диапазон)      e.g., =СРЗНАЧ(D3:D10)`,
    },
    {
      keywords: ['if', 'если', 'conditional', 'условие'],
      content: `## IF / ЕСЛИ
English: =IF(test, trueVal, falseVal)      e.g., =IF(B3>100,"High","Low")
Russian: =ЕСЛИ(тест; да; нет)              e.g., =ЕСЛИ(B3>100;"High";"Low")`,
    },
    {
      keywords: ['vlookup', 'впр', 'lookup', 'поиск'],
      content: `## VLOOKUP / ВПР
English: =VLOOKUP(val, table, col, exact)  e.g., =VLOOKUP(A2,E:F,2,FALSE)
Russian: =ВПР(знач; табл; номер; ложь)     e.g., =ВПР(A2;E:F;2;ЛОЖЬ)
Note: FALSE/ЛОЖЬ for exact match, TRUE/ИСТИНА for approximate`,
    },
    {
      keywords: ['xlookup', 'просмотрх', 'modern lookup'],
      content: `## XLOOKUP / ПРОСМОТРХ
English: =XLOOKUP(val, arr, ret)           e.g., =XLOOKUP(A2,E:E,F:F)
Russian: =ПРОСМОТРХ(знач; массив; возврат) e.g., =ПРОСМОТРХ(A2;E:E;F:F)
More powerful than VLOOKUP — no column index needed.`,
    },
    {
      keywords: ['countif', 'счётесли', 'count', 'счёт'],
      content: `## COUNTIF / СЧЁТЕСЛИ
English: =COUNTIF(range, criteria)         e.g., =COUNTIF(A:A,"Yes")
Russian: =СЧЁТЕСЛИ(диапазон; критерий)    e.g., =СЧЁТЕСЛИ(A:A;"Да")`,
    },
    {
      keywords: ['sumif', 'суммесли', 'conditional sum'],
      content: `## SUMIF / СУММЕСЛИ
English: =SUMIF(range, criteria, sumRange)     e.g., =SUMIF(A:A,"Yes",B:B)
Russian: =СУММЕСЛИ(диапазон; критерий; сумма)  e.g., =СУММЕСЛИ(A:A;"Да";B:B)`,
    },
    {
      keywords: ['iferror', 'еслиошбика', 'error handling', 'ошибка'],
      content: `## IFERROR / ЕСЛИОШИБКА
English: =IFERROR(val, default)              e.g., =IFERROR(A1/B1,"")
Russian: =ЕСЛИОШИБКА(знач; по_умолч)        e.g., =ЕСЛИОШИБКА(A1/B1;"")
Use to handle DIV/0, N/A, VALUE errors gracefully.`,
    },
    {
      keywords: ['date', 'дата', 'today', 'сегодня', 'datedif', 'разндат'],
      content: `## Date Functions
TODAY: =TODAY() / =СЕГОДНЯ()
NOW:   =NOW() / =ТДАТА()
DATE:  =DATE(2026,12,31) / =ДАТА(2026;12;31)
DATEDIF: =DATEDIF(A1,B1,"d") / =РАЗНДАТ(A1;B1;"d")  — "d", "m", "y"
YEAR:  =YEAR(A1) / =ГОД(A1)
MONTH: =MONTH(A1) / =МЕСЯЦ(A1)
DAY:   =DAY(A1) / =ДЕНЬ(A1)
WEEKDAY: =WEEKDAY(A1,2) / =ДЕНЬНЕД(A1;2)  — 2=Mon=1..Sun=7`,
    },
    {
      keywords: ['text', 'текст', 'concatenate', 'сцепить', 'left', 'левсимв', 'right', 'правсимв'],
      content: `## Text Functions
CONCATENATE: =A1&" "&B1 or =CONCATENATE(A1," ",B1) / =СЦЕПИТЬ(A1;" ";B1)
LEFT:  =LEFT(A1,3) / =ЛЕВСИМВ(A1;3)
RIGHT: =RIGHT(A1,3) / =ПРАВСИМВ(A1;3)
MID:   =MID(A1,2,3) / =ПСТР(A1;2;3)
LEN:   =LEN(A1) / =ДЛСТР(A1)
TRIM:  =TRIM(A1) / =СЖПРОБЕЛЫ(A1)
UPPER: =UPPER(A1) / =ПРОПИСН(A1)
LOWER: =LOWER(A1) / =СТРОЧН(A1)`,
    },
    {
      keywords: ['rank', 'ранг', 'large', 'наибольший', 'small', 'наименьший', 'statistical'],
      content: `## Statistical / Ranking
RANK:   =RANK(A1,A$1:A$100) / =РАНГ(A1;A$1:A$100)
RANK desc: =RANK(A1,A$1:A$100,0) / =РАНГ(A1;A$1:A$100;0)
LARGE:  =LARGE(A1:A100,3) / =НАИБОЛЬШИЙ(A1:A100;3)
SMALL:  =SMALL(A1:A100,3) / =НАИМЕНЬШИЙ(A1:A100;3)
MEDIAN: =MEDIAN(A1:A100) / =МЕДИАНА(A1:A100)
STDEV:  =STDEV(A1:A100) / =СТАНДОТКЛОН(A1:A100)`,
    },
    {
      keywords: ['running total', 'cumulative', 'накопление', 'итог нарастающий'],
      content: `## Running Total
Formula: =SUM($A$1:A1)
Copy this down the column.
$A$1 is absolute (anchor stays fixed), A1 is relative (adjusts as you copy down).
Example: =SUM($B$3:B3) copied from B3 down to B100.`,
    },
    {
      keywords: ['percentage', 'процент', 'percent'],
      content: `## Percentages
Part/Total: =D3/SUM(D$3:D$100)  — with $ on total range
Percentage display: =D3/C3*100  — multiply by 100
Format as %: use applyFormat with "0%" or "0.00%"`,
    },
  ],

  'office-js-patterns': [
    {
      keywords: ['range', 'getrange', 'setvalues', 'write', 'read'],
      content: `## Range Operations
Read:    getRange(address) → returns values, formulas, rowCount, columnCount
Write:   setValues(address, values) → 2D array, address = top-left cell
Clear:   clearRange(address) → REQUIRES CONFIRMATION
Format:  applyFormat(address, format) → e.g., "#,##0.00"`,
    },
    {
      keywords: ['table', 'createtable', 'addtablerow', 'sort', 'filter'],
      content: `## Table Operations
Create:  createTable(address, hasHeaders, tableName) → creates from range
Add row: addTableRow(tableName, values) → appends row
Sort:    sortTable(tableName, columnIndex, ascending) → 0-based column
Filter:  filterTable(tableName, columnIndex, filterType, values)`,
    },
    {
      keywords: ['formula', 'setformula', 'fillformula', 'autofill'],
      content: `## Formula Operations
Set:       setFormula(cellAddress, formula) → no leading "="
Fill down: fillFormula(sourceCell, targetRange, formula)
  → Writes formula in sourceCell, auto-fills to targetRange
  → Relative references adjust automatically
  → Example: fillFormula("B2", "B3:B100", "A2*C2")`,
    },
    {
      keywords: ['undo', 'backup', 'confirmation', 'dangerous'],
      content: `## Undo & Confirmation
- Destructive actions (clearRange, clearWorksheet, setValues on existing data) are backed up automatically
- User sees "↩ Undo" button after destructive actions
- Some operations require user confirmation before executing
- Always ask before deleting or overwriting data`,
    },
  ],

  'agent-workflows': [
    {
      keywords: ['workflow', 'template', 'report', 'create', 'создание', 'отчет'],
      content: `## Workflow: Creating a Report
1. getWorkbookOverview() → check current state
2. Ask user confirmation if clearing data
3. setValues for headers (row 1)
4. setValues for data (starting row 2)
5. createTable with hasHeaders: true
6. applyFormat for numbers
7. Use fillFormula for calculated columns`,
    },
    {
      keywords: ['analyze', 'analysis', 'анализ', 'анализировать'],
      content: `## Workflow: Data Analysis
1. getWorkbookOverview() → find data
2. getRange() → read specific data range
3. Identify columns, data types, row count
4. Execute transformations (sort, filter, formula)
5. Report findings to user`,
    },
    {
      keywords: ['anti-pattern', 'never', 'dont', 'clear', 'delete', 'нельзя'],
      content: `## CRITICAL ANTI-PATTERNS
❌ NEVER clearWorksheet before reading data
  → Call getWorkbookOverview + getRange FIRST
❌ NEVER write formulas without * between cell refs
  → =B8*B6 not =B8B6
❌ NEVER guess range sizes — calculate them
  → N rows starting at row R = range R:(R+N-1)`,
    },
  ],
};

// ──────────────────────────────────────────────
// FETCH-BASED KNOWLEDGE RETRIEVAL
// Falls back to embedded knowledge if fetch fails
// ──────────────────────────────────────────────

const MANIFEST_PATH = '/knowledge/knowledge-manifest.json';

interface ManifestCategory {
  id: string;
  title: string;
  description: string;
  files: string[];
}

interface KnowledgeManifest {
  version: string;
  lastUpdated: string;
  categories: ManifestCategory[];
}

/** In-memory cache: categoryId -> fileName -> content */
const knowledgeCache = new Map<string, string>();

let manifestCache: KnowledgeManifest | null = null;

/**
 * Load the knowledge manifest from the server.
 * Returns null if fetch fails (browser offline, dev server not serving /knowledge/).
 */
async function loadManifest(): Promise<KnowledgeManifest | null> {
  if (manifestCache) return manifestCache;

  try {
    const response = await fetch(MANIFEST_PATH);
    if (!response.ok) return null;
    manifestCache = (await response.json()) as KnowledgeManifest;
    return manifestCache;
  } catch {
    // Fetch failed — knowledge directory may not be deployed yet
    return null;
  }
}

/**
 * Load a specific knowledge section file.
 * Tries fetch() first, falls back gracefully.
 */
async function loadSectionFile(categoryId: string, fileName: string): Promise<string | null> {
  const cacheKey = `${categoryId}/${fileName}`;
  if (knowledgeCache.has(cacheKey)) {
    return knowledgeCache.get(cacheKey)!;
  }

  try {
    const response = await fetch(`/knowledge/sections/${fileName}`);
    if (!response.ok) return null;
    const content = await response.text();
    knowledgeCache.set(cacheKey, content);
    return content;
  } catch {
    return null;
  }
}

// ──────────────────────────────────────────────
// EXPOSED HELPERS for testing / direct embedded access
// ──────────────────────────────────────────────

/**
 * Get embedded knowledge content for a specific category (no fetch required).
 * Uses only the embedded constant data — useful for testing and offline use.
 *
 * @param category - Category name: "excel-formulas", "office-js-patterns", "agent-workflows", or "all"
 * @returns Formatted markdown string of all matching entries
 */
export async function getEmbeddedKnowledge(category: string): Promise<string> {
  if (category === 'all') {
    const parts: string[] = [];
    for (const cat of Object.keys(EMBEDDED_KNOWLEDGE)) {
      for (const entry of EMBEDDED_KNOWLEDGE[cat]) {
        parts.push(entry.content);
      }
    }
    return parts.join('\n');
  }

  const entries = EMBEDDED_KNOWLEDGE[category];
  if (!entries) return '';

  return entries.map((e) => e.content).join('\n');
}

/**
 * Search embedded knowledge for entries matching a query string.
 *
 * @param query - Search term (case-insensitive, matched against keywords)
 * @returns Array of content strings that match the query
 */
export async function searchKnowledge(query: string): Promise<string[]> {
  if (!query || query.trim().length === 0) return [];

  const lowerQuery = query.toLowerCase();
  const results: string[] = [];

  for (const cat of Object.keys(EMBEDDED_KNOWLEDGE)) {
    for (const entry of EMBEDDED_KNOWLEDGE[cat]) {
      const matchesKeyword = entry.keywords.some((kw) => kw.toLowerCase().includes(lowerQuery));
      if (matchesKeyword) {
        results.push(entry.content);
      }
    }
  }

  return results;
}

// ──────────────────────────────────────────────
// PUBLIC API
// ──────────────────────────────────────────────

/**
 * Get knowledge about Excel formulas, Office.js patterns, or agent workflows.
 *
 * @param category - Category to search: "excel-formulas", "office-js-patterns", "agent-workflows", or "all"
 * @param query   - Optional keyword to filter results (e.g., "SUM", "VLOOKUP", "СУММ")
 * @returns Formatted markdown string (max 3000 chars).
 */
export async function getKnowledge(category?: string, query?: string): Promise<string> {
  const parts: string[] = [];
  let remainingBudget = 3000; // character limit

  function appendToResult(text: string): boolean {
    if (text.length > remainingBudget) {
      parts.push(text.substring(0, remainingBudget) + '\n\n_(results truncated — use a more specific query)_');
      return false;
    }
    parts.push(text);
    remainingBudget -= text.length;
    return true;
  }

  // ── Step 1: Try loading from fetched knowledge ──
  let fetchedContent = '';
  const manifest = await loadManifest();

  if (manifest) {
    const categoriesToLoad = manifest.categories.filter(
      (c) => !category || category === 'all' || c.id === category
    );

    for (const cat of categoriesToLoad) {
      for (const fileName of cat.files) {
        const content = await loadSectionFile(cat.id, fileName);
        if (content) {
          fetchedContent += content + '\n\n';
        }
      }
    }
  }

  // ── Step 2: Search fetched content by query ──
  if (fetchedContent && query) {
    const lowerQuery = query.toLowerCase();
    const lines = fetchedContent.split('\n');
    const matchingLines: string[] = [];
    let inMatch = false;

    for (const line of lines) {
      if (line.startsWith('#') || line.startsWith('##')) {
        inMatch = line.toLowerCase().includes(lowerQuery);
      }
      if (inMatch) {
        matchingLines.push(line);
      }
    }

    if (matchingLines.length > 0) {
      appendToResult(matchingLines.join('\n'));
    }
  } else if (fetchedContent) {
    // No query — return first 3000 chars of fetched content
    const truncated = fetchedContent.length > 3000
      ? fetchedContent.substring(0, 3000) + '\n\n... (use a specific query for more detail)'
      : fetchedContent;
    appendToResult(truncated);
  }

  // ── Step 3: Fallback to embedded knowledge ──
  const categoriesToSearch = category && category !== 'all'
    ? [category]
    : Object.keys(EMBEDDED_KNOWLEDGE);

  for (const cat of categoriesToSearch) {
    const entries = EMBEDDED_KNOWLEDGE[cat];
    if (!entries) continue;

    for (const entry of entries) {
      if (query) {
        const lowerQuery = query.toLowerCase();
        const matchesKeyword = entry.keywords.some((kw) => kw.toLowerCase().includes(lowerQuery));
        if (!matchesKeyword) continue;
      }
      if (!appendToResult(entry.content + '\n')) {
        break;
      }
    }
    if (remainingBudget <= 0) break;
  }

  // ── Step 4: Handle empty results ──
  if (parts.length === 0) {
    if (query) {
      return `No knowledge found for "${query}" in category "${category || 'all'}". Try a different search term or use getKnowledge("all") to see available topics.`;
    }
    return `No knowledge available for category "${category || 'all'}". Available categories: excel-formulas, office-js-patterns, agent-workflows.`;
  }

  return parts.join('\n');
}

// ──────────────────────────────────────────────
// REGISTER TOOL
// ──────────────────────────────────────────────

toolRegistry.register(
  'getKnowledge',
  'Get documentation about Excel formulas (EN/RU function names, syntax), Office.js patterns (range operations, table ops, autofill), or agent workflow templates. Call this when you need formula syntax, function reference with Russian names, or step-by-step workflow guidance.',
  {
    type: 'object',
    properties: {
      category: {
        type: 'string',
        enum: ['excel-formulas', 'office-js-patterns', 'agent-workflows', 'all'],
        description: 'Documentation category. "excel-formulas" for formula syntax and function reference. "office-js-patterns" for Office.js API patterns. "agent-workflows" for multi-step workflow templates. "all" to search everything.',
      },
      query: {
        type: 'string',
        description: 'Search keyword or topic. Examples: "SUM", "VLOOKUP", "СУММ", "multiplication", "create table", "auto fill", "воркфлоу"',
      },
    },
    required: [],
  },
  async (args) => {
    const category = args.category as string | undefined;
    const query = args.query as string | undefined;
    const result = await getKnowledge(category, query);
    return JSON.stringify({ success: true, content: result });
  }
);
