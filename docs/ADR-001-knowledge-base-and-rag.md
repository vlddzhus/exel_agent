# ADR-001: Knowledge Base and RAG Integration for Excel AI Agent

**Status:** Accepted  
**Date:** 2026-06-17  
**Author:** Architect Agent  
**Deciders:** Architect, Backend, Frontend  
**Discussion:** 2026-06-16  

---

## Context

The Excel AI Agent uses Groq API (specifically `meta-llama/llama-4-scout-17b-16e-instruct`) to translate natural language instructions into Office.js tool calls. In production usage, the LLM exhibits four categories of failure:

### Failure Category 1: Invalid Excel Formula Syntax
The LLM generates formulas missing required operators between cell references:
- **Incorrect:** `=B8B6+B9B7` (missing `*`)
- **Incorrect:** `=A1A2` instead of `=A1*A2`
- **Root cause:** LLM treats cell references like `B8B6` as a single token rather than two cell references requiring multiplication.

### Failure Category 2: Wrong Range Dimensions
When writing data, the LLM miscalculates range sizes:
- If `values` array has 9 rows starting at A3 → correct range is `A3:E11` (rows 3..11 inclusive = 9 rows)
- But it frequently calculates `A3:E10` (off by one) or `A3:E12` (off by one)
- When creating tables on ranges that include merged title cells, the operation fails

### Failure Category 3: Destructive Actions Without Reading
The most severe category — seen in `error.txt`:
1. User requests improvements to an existing table
2. Agent calls `getWorkbookOverview()` — sees the table exists
3. Agent calls `clearWorksheet()` **without reading the actual data first**
4. All original user data is destroyed
5. Agent creates a simplistic replacement with only 2 data rows
6. User becomes frustrated: "вы всё удалили и ничего не улучшили"

### Failure Category 4: Context Window Limits on Large Workbooks
When workbooks have many sheets, the LLM context fills with data, leaving insufficient room for reasoning about formulas and operations.

### Current Mitigation
The existing `SYSTEM_PROMPT` is 30 lines of flat text with:
- No formula syntax rules
- No range calculation rules
- No anti-pattern guidance
- No structured workflows
- No Russian/English function name mappings

---

## Decision

Implement a **three-tier knowledge system**:

### Layer 1: Static Prompt Injection (`knowledge-injector.ts`)
**Compile-time injection** of high-signal knowledge into the system prompt:
- Anti-pattern rules (from real production failures)
- Formula operator rules (cell reference multiplication requires `*`)
- Range calculation formulas
- Mandatory first-step enforcement
- Russian/English function name mappings

**Selection criteria:** Only rules that are:
1. Small in size (< 300 tokens total)
2. High frequency (needed on every interaction)
3. Critical for correctness (preventing data loss)

### Layer 2: Tool-Based Retrieval (`getKnowledge` tool)
**On-demand retrieval** via a new tool registered in the tool registry:
- Category-based access: `excel-formulas`, `office-js-patterns`, `agent-workflows`
- Keyword filtering within categories
- Returns formatted markdown snippets (max 3000 chars)
- Cached in memory after first load
- Falls back to embedded knowledge if fetch fails

**Selection criteria:** Material that is:
1. Too large for static injection (> 300 tokens per section)
2. Needed only for specific tasks (formula construction, pivot tables, chart types)
3. Reference material (function signatures, parameter details)

### Layer 3: Knowledge Directory (`knowledge/`)
**Source of truth** as markdown files organized by category:
- `knowledge/sections/excel-formulas.md` — 50+ formula patterns, Russian/English names
- `knowledge/sections/office-js-patterns.md` — Office.js best practices, range operations
- `knowledge/sections/agent-workflows.md` — Multi-step workflow templates
- `knowledge/knowledge-manifest.json` — Metadata and file listing

These files are:
- Copied to `dist/knowledge/` by webpack's CopyWebpackPlugin
- Fetched at runtime from the taskpane via `fetch('/knowledge/...')`
- Also have their core content embedded in TypeScript as fallback

---

## Options Considered

### Option 1: LangChain.js RAG Pipeline
- **Approach:** Vector store with Chroma, LangChain's document loaders, and Groq embeddings
- **Pros:** Full semantic search, automatic chunking, reranking
- **Cons:** Requires server-side processing (browser-only taskpane), 500KB+ bundle overhead, over-engineered for 3 knowledge categories
- **Verdict:** Rejected

### Option 2: Vercel AI SDK with RAG
- **Approach:** `@ai-sdk/core` with tool calls and context retrieval
- **Pros:** Streaming support, tool integration
- **Cons:** Not Office.js specific, adds dependency, same browser-environment limitations
- **Verdict:** Rejected

### Option 3: Static Only (Expand System Prompt)
- **Approach:** Put all knowledge into a 200-line system prompt
- **Pros:** Simplest to implement, no new tools needed
- **Cons:** Context window pressure, especially on large workbooks; cannot dynamically select relevant sections; LLM attention degrades with long prompts
- **Verdict:** Works as first layer but insufficient alone

### Option 4: Prompt Caching (Groq Feature)
- **Approach:** Use Groq's prompt caching to reuse system prompt prefix
- **Pros:** Reduces cost, faster response on cached prompts
- **Cons:** Still consumes context, doesn't solve context overflow, Groq caching has constraints
- **Verdict:** Worth using if available but orthogonal to knowledge quality

### Chosen: Hybrid Static + Tool Retrieval (Option 3 + Layer 2)
- **Approach:** Static injection for critical rules + tool-based retrieval for reference docs
- **Pros:** Balances simplicity and depth; keeps critical rules always visible; defers large docs to on-demand retrieval; no server dependency
- **Cons:** Tool-based retrieval requires the LLM to know when to call `getKnowledge`; keyword matching is less flexible than vector search
- **Mitigation:** The system prompt explicitly instructs the LLM to call `getKnowledge` when unsure about formula syntax or Office.js patterns

---

## Consequences

### Positive
1. **Structured knowledge** — model has categorized, human-readable reference material
2. **Context efficiency** — only critical rules are always injected; full docs retrieved on demand
3. **Prevents data loss** — anti-pattern rules explicitly prohibit `clearWorksheet` before reading
4. **Formula correctness** — operator rules embedded at the top of every interaction
5. **Russian/English support** — bilingual formula function name mapping
6. **Extensible** — new knowledge sections can be added without changing system prompt
7. **No server dependency** — all knowledge is compiled into the add-in
8. **Dual-path retrieval** — fetch from `knowledge/` directory with Typescript embedded fallback

### Negative
1. **Static prompt growth** — system prompt increases from 30 to ~200 lines
2. **Retrieval latency** — `getKnowledge` tool adds one round-trip when model needs docs
3. **Maintenance burden** — keeping knowledge markdown in sync with TypeScript embeddings
4. **No semantic search** — keyword matching may miss relevant results with different phrasing

### Mitigation Strategy
| Risk | Mitigation |
|---|---|
| Prompt too large | Only inject rules under 300 tokens; defer bulk to tools |
| Model doesn't call getKnowledge | System prompt explicitly lists categories and when to use each |
| Knowledge drifts from code | Single source of truth in TypeScript with markdown as secondary |
| Keyword matching misses | Include synonyms and multiple search terms; embedded fallback |
| Fetch fails in taskpane | Embedded TypeScript fallback ensures core knowledge always available |

---

## Implementation Plan

### Phase 1: Foundation (This ADR — Delivered)
1. ✅ **ADR-001** — This document
2. ✅ **Knowledge Directory** — markdown files in `knowledge/sections/`
3. ✅ **`knowledge/knowledge-manifest.json`** — category metadata
4. ✅ **`src/taskpane/tools/knowledge-tools.ts`** — retrieval tool with embedded + fetch
5. ✅ **`src/taskpane/agent/knowledge-injector.ts`** — static injection module
6. ✅ **`src/taskpane/agent/system-prompt.ts`** — v2 structured prompt
7. ✅ **`webpack.config.js`** — CopyWebpackPlugin for knowledge directory
8. ✅ **`docs/system-prompt-v2.md`** — designed system prompt reference

### Phase 2: Refinement (Future)
- [ ] Add usage analytics (which categories are queried most)
- [ ] Add feedback loop — if `getKnowledge` returns empty, suggest missing docs
- [ ] Evaluate vector search if knowledge base grows beyond 50 sections

### Phase 3: Advanced (Optional)
- [ ] Auto-generate knowledge from TypeScript tool definitions
- [ ] Knowledge diffing — detect when tool parameters change and update docs
- [ ] Community-contributed knowledge snippets

---

## Technical Architecture

```
excel-ai-agent/
├── knowledge/                                    ← Source of truth (markdown)
│   ├── sections/
│   │   ├── excel-formulas.md                    ← 50+ formula patterns (EN/RU)
│   │   ├── office-js-patterns.md                ← Office.js best practices
│   │   └── agent-workflows.md                   ← Multi-step workflow templates
│   └── knowledge-manifest.json                  ← Metadata for categories
│
├── src/taskpane/
│   ├── agent/
│   │   ├── system-prompt.ts                     ← v2 structured prompt
│   │   ├── knowledge-injector.ts                ← Static injection logic
│   │   └── react-loop.ts                        ← Unchanged (imports SYSTEM_PROMPT)
│   │
│   ├── tools/
│   │   ├── registry.ts                          ← Tool registry (unchanged)
│   │   ├── knowledge-tools.ts                   ← getKnowledge() tool (embedded + fetch)
│   │   └── [...existing tools...]
│   │
│   └── chat/
│       └── groq-client.ts                       ← Unchanged
│
├── docs/
│   ├── ADR-001-knowledge-base-and-rag.md        ← This document
│   └── system-prompt-v2.md                      ← Designed prompt reference
│
└── webpack.config.js                            ← Updated: copies knowledge/ to dist/
```

### Data Flow

```
User Request
    │
    ▼
┌──────────────────────────────────────────────────┐
│  ReAct Loop receives SYSTEM_PROMPT               │
│  (base prompt + static knowledge injection)       │
│  ┌─────────────────────┐  ┌───────────────────┐  │
│  │ BASE_PROMPT         │  │ KNOWLEDGE_INJECT  │  │
│  │ • Role              │  │ • Formula rules   │  │
│  │ • Workflows         │  │ • Anti-patterns   │  │
│  │ • Tool definitions  │  │ • EN/RU mappings  │  │
│  │ • Error handling    │  │ • Quick reference │  │
│  └─────────────────────┘  └───────────────────┘  │
└──────────────────────────────────────────────────┘
    │
    ▼
LLM decides if it needs more knowledge
    │
    ├── No: Proceeds with tools
    │
    └── Yes: Calls getKnowledge(category, query)
                  │
                  ▼
         ┌────────────────────┐
         │  knowledge-tools   │
         │  • Check cache     │
         │  • Try fetch()     │
         │    from dist/      │
         │  • Fallback to     │
         │    embedded        │
         │  • Return markdown │
         └────────────────────┘
```

---

## References

- [SPEC.md](../SPEC.md) — Project specification
- [system-prompt.ts](../src/taskpane/agent/system-prompt.ts) — Current system prompt
- [react-loop.ts](../src/taskpane/agent/react-loop.ts) — ReAct loop implementation
- [error.txt](../error.txt) — Recorded failure case
- [system-prompt-v2.md](system-prompt-v2.md) — New prompt design reference
- [knowledge-tools.ts](../src/taskpane/tools/knowledge-tools.ts) — Implementation
- [knowledge-injector.ts](../src/taskpane/agent/knowledge-injector.ts) — Implementation
