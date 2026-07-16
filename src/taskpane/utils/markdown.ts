// ── Markdown Renderer with Syntax Highlighting ──

export interface MarkdownOptions {
  highlightCode?: boolean;
  renderTables?: boolean;
  renderLatex?: boolean;
  renderCollapsible?: boolean;
}

const ESCAPE_MAP: Record<string, string> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&#39;',
};

function escapeHtml(text: string): string {
  return text.replace(/[&<>"']/g, (ch) => ESCAPE_MAP[ch] || ch);
}

function escapeAttr(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ── Inline renderers ──

type InlineToken = { type: string; content: string; lang?: string };

function tokenizeInline(text: string): InlineToken[] {
  const tokens: InlineToken[] = [];
  let remaining = text;

  const patterns: { regex: RegExp; type: string }[] = [
    { regex: /^!\[([^\]]*)\]\(([^)]+)\)/, type: 'image' },
    { regex: /^\[([^\]]+)\]\(([^)]+)\)/, type: 'link' },
    { regex: /^`([^`]+)`/, type: 'code' },
    { regex: /^\$\$(.+?)\$\$/, type: 'latex-block' },
    { regex: /^\$(.+?)\$/, type: 'latex' },
    { regex: /^\*\*(.+?)\*\*/, type: 'bold' },
    { regex: /^__(.+?)__/, type: 'bold' },
    { regex: /^\*(.+?)\*/, type: 'italic' },
    { regex: /^_(.+?)_/, type: 'italic' },
    { regex: /^~~(.+?)~~/, type: 'strikethrough' },
    { regex: /^`(.+?)`/, type: 'code' },
  ];

  while (remaining.length > 0) {
    let matched = false;
    for (const { regex, type } of patterns) {
      const m = regex.exec(remaining);
      if (m) {
        tokens.push({ type, content: m[1], lang: m[2] });
        remaining = remaining.substring(m[0].length);
        matched = true;
        break;
      }
    }
    if (!matched) {
      // Take next character
      const next = remaining[0];
      tokens.push({ type: 'text', content: next });
      remaining = remaining.substring(1);
    }
  }
  return tokens;
}

function renderInlineTokens(tokens: InlineToken[]): string {
  return tokens.map(t => {
    switch (t.type) {
      case 'text':
        return escapeHtml(t.content);
      case 'bold':
        return `<strong>${renderInline(t.content)}</strong>`;
      case 'italic':
        return `<em>${renderInline(t.content)}</em>`;
      case 'strikethrough':
        return `<del>${renderInline(t.content)}</del>`;
      case 'code':
        return `<code>${escapeHtml(t.content)}</code>`;
      case 'link':
        return `<a href="${escapeAttr(t.lang || '#')}" target="_blank" rel="noopener">${escapeHtml(t.content)}</a>`;
      case 'image':
        return `<img src="${escapeAttr(t.lang || '')}" alt="${escapeAttr(t.content)}" loading="lazy" />`;
      case 'latex':
        return `<span class="latex-inline">\(${escapeHtml(t.content)}\)</span>`;
      case 'latex-block':
        return `<div class="latex-block">\[${escapeHtml(t.content)}\]</div>`;
      default:
        return escapeHtml(t.content);
    }
  }).join('');
}

function renderInline(text: string): string {
  return renderInlineTokens(tokenizeInline(text));
}

// ── Syntax Highlighting ──

const LANG_HANDLERS: Record<string, (code: string) => string> = {
  json: highlightJson,
  javascript: highlightJs,
  js: highlightJs,
  typescript: highlightJs,
  ts: highlightJs,
  python: highlightPython,
  py: highlightPython,
  html: highlightHtml,
  xml: highlightHtml,
  css: highlightCss,
  sql: highlightSql,
  bash: highlightBash,
  shell: highlightBash,
  sh: highlightBash,
  plain: (code: string) => escapeHtml(code),
  text: (code: string) => escapeHtml(code),
};

function highlightJson(code: string): string {
  return escapeHtml(code)
    .replace(/(&quot;[^&]+&quot;)(\s*:)/g, '<span class="hl-key">$1</span>$2')
    .replace(/("(?:[^"\\]|\\.)*")/g, (m) => {
      if (m.includes('hl-key')) return m;
      return `<span class="hl-str">${m}</span>`;
    })
    .replace(/\b(true|false|null)\b/g, '<span class="hl-bool">$1</span>')
    .replace(/\b(\d+\.?\d*)\b/g, '<span class="hl-num">$1</span>');
}

function highlightJs(code: string): string {
  const keywords = /\b(async|await|break|case|catch|class|const|continue|debugger|default|delete|do|else|export|extends|finally|for|function|if|import|in|instanceof|let|new|of|return|static|super|switch|this|throw|try|typeof|var|void|while|with|yield)\b/g;
  const builtins = /\b(console|Math|JSON|Array|Object|String|Number|Boolean|Map|Set|Promise|Symbol|RegExp|Date|Error)\b/g;
  return escapeHtml(code)
    .replace(/(\/\/.*)/g, '<span class="hl-comment">$1</span>')
    .replace(/(\/\*[\s\S]*?\*\/)/g, '<span class="hl-comment">$1</span>')
    .replace(keywords, '<span class="hl-keyword">$1</span>')
    .replace(builtins, '<span class="hl-builtin">$1</span>')
    .replace(/("(?:[^"\\]|\\.)*")/g, '<span class="hl-str">$1</span>')
    .replace(/('(?:[^'\\]|\\.)*')/g, '<span class="hl-str">$1</span>')
    .replace(/`(?:[^`\\]|\\.)*`/g, '<span class="hl-str">$1</span>')
    .replace(/\b(\d+\.?\d*)\b/g, '<span class="hl-num">$1</span>');
}

function highlightPython(code: string): string {
  const keywords = /\b(False|None|True|and|as|assert|async|await|break|class|continue|def|del|elif|else|except|finally|for|from|global|if|import|in|is|lambda|nonlocal|not|or|pass|raise|return|try|while|with|yield)\b/g;
  const builtins = /\b(print|len|range|int|str|float|list|dict|set|tuple|type|open|input|map|filter|zip|enumerate|sorted|reversed|abs|max|min|sum|any|all)\b/g;
  return escapeHtml(code)
    .replace(/(#.*)/g, '<span class="hl-comment">$1</span>')
    .replace(/(""".*?""")/gs, '<span class="hl-str">$1</span>')
    .replace(/('''.*?''')/gs, '<span class="hl-str">$1</span>')
    .replace(keywords, '<span class="hl-keyword">$1</span>')
    .replace(builtins, '<span class="hl-builtin">$1</span>')
    .replace(/("(?:[^"\\]|\\.)*")/g, '<span class="hl-str">$1</span>')
    .replace(/('(?:[^'\\]|\\.)*')/g, '<span class="hl-str">$1</span>')
    .replace(/f"(?:[^"\\]|\\.)*"/g, '<span class="hl-str">$1</span>')
    .replace(/\b(\d+\.?\d*)\b/g, '<span class="hl-num">$1</span>');
}

function highlightHtml(code: string): string {
  return escapeHtml(code)
    .replace(/(&lt;\/?[a-zA-Z][^&]*&gt;)/g, '<span class="hl-tag">$1</span>')
    .replace(/\b([a-zA-Z-]+)(=)(&quot;[^&]*&quot;)/g, '<span class="hl-attr">$1</span>$2<span class="hl-str">$3</span>');
}

function highlightCss(code: string): string {
  return escapeHtml(code)
    .replace(/(\/\*[\s\S]*?\*\/)/g, '<span class="hl-comment">$1</span>')
    .replace(/(@{1,2}[a-zA-Z-]+)/g, '<span class="hl-keyword">$1</span>')
    .replace(/([a-zA-Z-]+)(\s*:)/g, '<span class="hl-prop">$1</span>$2')
    .replace(/!(important)/g, '<span class="hl-keyword">!$1</span>')
    .replace(/(#[\da-fA-F]{3,8})\b/g, '<span class="hl-num">$1</span>')
    .replace(/(\d+\.?\d*(px|em|rem|vh|vw|%|s|ms)?)/g, '<span class="hl-num">$1</span>');
}

function highlightSql(code: string): string {
  const keywords = /\b(SELECT|FROM|WHERE|INSERT|INTO|VALUES|UPDATE|SET|DELETE|CREATE|TABLE|ALTER|DROP|INDEX|JOIN|LEFT|RIGHT|INNER|OUTER|ON|AND|OR|NOT|IN|LIKE|BETWEEN|IS|NULL|AS|ORDER|BY|GROUP|HAVING|LIMIT|OFFSET|UNION|ALL|DISTINCT|COUNT|SUM|AVG|MIN|MAX|CASE|WHEN|THEN|ELSE|END|EXISTS|WITH|RECURSIVE|PRIMARY|KEY|FOREIGN|REFERENCES|CASCADE|BEGIN|COMMIT|ROLLBACK)\b/g;
  return escapeHtml(code)
    .replace(/(--.*)/g, '<span class="hl-comment">$1</span>')
    .replace(keywords, '<span class="hl-keyword">$1</span>')
    .replace(/('(?:[^'\\]|\\.)*')/g, '<span class="hl-str">$1</span>')
    .replace(/\b(\d+\.?\d*)\b/g, '<span class="hl-num">$1</span>');
}

function highlightBash(code: string): string {
  return escapeHtml(code)
    .replace(/(#.*)/g, '<span class="hl-comment">$1</span>')
    .replace(/("(?:[^"\\]|\\.)*")/g, '<span class="hl-str">$1</span>')
    .replace(/('(?:[^'\\]|\\.)*')/g, '<span class="hl-str">$1</span>')
    .replace(/\b(\d+)\b/g, '<span class="hl-num">$1</span>');
}

function highlightDefault(code: string): string {
  // Generic highlighting for unknown languages
  return escapeHtml(code)
    .replace(/(\/\/.*)/g, '<span class="hl-comment">$1</span>')
    .replace(/(#.*)/g, '<span class="hl-comment">$1</span>')
    .replace(/("(?:[^"\\]|\\.)*")/g, '<span class="hl-str">$1</span>')
    .replace(/('(?:[^'\\]|\\.)*')/g, '<span class="hl-str">$1</span>')
    .replace(/\b(\d+\.?\d*)\b/g, '<span class="hl-num">$1</span>');
}

function highlightCode(code: string, lang: string): string {
  const handler = LANG_HANDLERS[lang.toLowerCase()];
  if (handler) return handler(code);
  return highlightDefault(code);
}

// ── Block-level parsing ──

export interface BlockToken {
  type: string;
  content: string;
  lang?: string;
  items?: string[];
  cells?: string[][];
  isOrdered?: boolean;
}

function tokenizeBlocks(text: string): BlockToken[] {
  const lines = text.split('\n');
  const blocks: BlockToken[] = [];
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];

    // Collapsible section
    if (/^<details>/.test(line.trim())) {
      let summary = '';
      let bodyLines: string[] = [];
      let j = i + 1;
      let inSummary = false;
      while (j < lines.length && !/^<\/details>/.test(lines[j].trim())) {
        if (/^<summary>/.test(lines[j].trim())) {
          inSummary = true;
          summary = lines[j].replace(/^<summary>\s*/, '').replace(/\s*<\/summary>\s*$/, '');
          j++;
          continue;
        }
        if (inSummary) {
          bodyLines.push(lines[j]);
        }
        j++;
      }
      blocks.push({ type: 'collapsible', content: bodyLines.join('\n'), lang: summary });
      i = j + 1;
      continue;
    }

    // Code block
    const codeMatch = line.match(/^```(\w*)/);
    if (codeMatch) {
      const lang = codeMatch[1];
      const codeLines: string[] = [];
      i++;
      while (i < lines.length && !/^```/.test(lines[i].trim())) {
        codeLines.push(lines[i]);
        i++;
      }
      blocks.push({ type: 'code', content: codeLines.join('\n'), lang });
      i++;
      continue;
    }

    // Horizontal rule
    if (/^---\s*$/.test(line) || /^\*\*\*\s*$/.test(line) || /^___\s*$/.test(line)) {
      blocks.push({ type: 'hr', content: '' });
      i++;
      continue;
    }

    // Heading
    const headingMatch = line.match(/^(#{1,6})\s+(.+)/);
    if (headingMatch) {
      blocks.push({ type: `h${headingMatch[1].length}`, content: headingMatch[2] });
      i++;
      continue;
    }

    // Blockquote
    if (/^>\s/.test(line)) {
      const quoteLines: string[] = [];
      while (i < lines.length && /^>\s/.test(lines[i])) {
        quoteLines.push(lines[i].replace(/^>\s?/, ''));
        i++;
      }
      blocks.push({ type: 'blockquote', content: quoteLines.join('\n') });
      continue;
    }

    // Table
    if (/^\|.+\|/.test(line) && i + 2 < lines.length && /^\|[-| :]+\|/.test(lines[i + 1])) {
      const headerCells = line.split('|').map(c => c.trim()).filter(c => c);
      const rows: string[][] = [headerCells];
      let j = i + 2;
      while (j < lines.length && /^\|.+\|/.test(lines[j])) {
        const cells = lines[j].split('|').map(c => c.trim()).filter(c => c);
        if (cells.length > 0) rows.push(cells);
        j++;
      }
      blocks.push({ type: 'table', content: '', cells: rows });
      i = j;
      continue;
    }

    // Unordered list
    if (/^[-*+]\s/.test(line)) {
      const items: string[] = [];
      while (i < lines.length && /^[-*+]\s/.test(lines[i])) {
        items.push(lines[i].replace(/^[-*+]\s+/, ''));
        i++;
      }
      blocks.push({ type: 'list', content: '', items, isOrdered: false });
      continue;
    }

    // Ordered list
    if (/^\d+[.)]\s/.test(line)) {
      const items: string[] = [];
      while (i < lines.length && /^\d+[.)]\s/.test(lines[i])) {
        items.push(lines[i].replace(/^\d+[.)]\s+/, ''));
        i++;
      }
      blocks.push({ type: 'list', content: '', items, isOrdered: true });
      continue;
    }

    // Empty line
    if (/^\s*$/.test(line)) {
      i++;
      continue;
    }

    // Paragraph
    const paraLines: string[] = [];
    while (i < lines.length && !/^\s*$/.test(lines[i]) &&
           !/^#{1,6}\s/.test(lines[i]) && !/^```/.test(lines[i]) &&
           !/^\d+[.)]\s/.test(lines[i]) && !/^[-*+]\s/.test(lines[i]) &&
           !/^\|/.test(lines[i]) && !/^>/.test(lines[i]) &&
           !/^---\s*$/.test(lines[i]) && !/^\*\*\*\s*$/.test(lines[i]) &&
           !/^___{1,}\s*$/.test(lines[i])) {
      paraLines.push(lines[i]);
      i++;
    }
    if (paraLines.length > 0) {
      blocks.push({ type: 'paragraph', content: paraLines.join('\n') });
    } else {
      i++;
    }
  }

  return blocks;
}

function renderBlock(block: BlockToken): string {
  switch (block.type) {
    case 'h1': return `<h1>${renderInline(block.content)}</h1>`;
    case 'h2': return `<h2>${renderInline(block.content)}</h2>`;
    case 'h3': return `<h3>${renderInline(block.content)}</h3>`;
    case 'h4': return `<h4>${renderInline(block.content)}</h4>`;
    case 'h5': return `<h5>${renderInline(block.content)}</h5>`;
    case 'h6': return `<h6>${renderInline(block.content)}</h6>`;
    case 'paragraph':
      return `<p>${renderInline(block.content.replace(/\n/g, '<br>'))}</p>`;
    case 'blockquote':
      return `<blockquote>${renderInline(block.content)}</blockquote>`;
    case 'hr':
      return '<hr>';
    case 'code': {
      const highlighted = highlightCode(block.content, block.lang || '');
      const langLabel = block.lang ? `<div class="code-lang">${escapeHtml(block.lang)}</div>` : '';
      return `<div class="code-block">
        ${langLabel}
        <pre><code class="lang-${escapeHtml(block.lang || 'none')}">${highlighted}</code></pre>
        <button class="code-copy" data-code="${escapeAttr(block.content)}">Copy</button>
        <button class="code-insert" data-code="${escapeAttr(block.content)}" title="Insert into sheet">📋 Insert</button>
      </div>`;
    }
    case 'table': {
      if (!block.cells || block.cells.length < 2) return '';
      const thead = block.cells[0];
      const tbody = block.cells.slice(1);
      let html = '<div class="table-wrapper"><table><thead><tr>';
      for (const th of thead) {
        html += `<th>${renderInline(th)}</th>`;
      }
      html += '</tr></thead><tbody>';
      for (const row of tbody) {
        html += '<tr>';
        for (const cell of row) {
          html += `<td>${renderInline(cell)}</td>`;
        }
        html += '</tr>';
      }
      html += '</tbody></table></div>';
      return html;
    }
    case 'list': {
      const tag = block.isOrdered ? 'ol' : 'ul';
      let html = `<${tag}>`;
      if (block.items) {
        for (const item of block.items) {
          html += `<li>${renderInline(item)}</li>`;
        }
      }
      html += `</${tag}>`;
      return html;
    }
    case 'collapsible': {
      const summaryText = block.lang || 'Details';
      return `<details class="collapsible-section">
        <summary>${renderInline(summaryText)}</summary>
        <div class="collapsible-body">${renderMarkdown(block.content)}</div>
      </details>`;
    }
    default:
      return escapeHtml(block.content);
  }
}

// ── Main renderer ──

export function renderMarkdown(md: string, options?: MarkdownOptions): string {
  const opts: MarkdownOptions = {
    highlightCode: true,
    renderTables: true,
    renderLatex: true,
    renderCollapsible: true,
    ...options,
  };

  const blocks = tokenizeBlocks(md);
  const parts = blocks.map(b => renderBlock(b));

  // Render inline LaTeX in text nodes (post-processing)
  let html = parts.join('\n');

  // Wrap standalone tables for Excel-style export
  html = html.replace(/<div class="table-wrapper">/g, '<div class="table-wrapper" data-table="true">');

  return html;
}

// ── Helper: extract table as TSV ──

export function extractTableAsTsv(html: string): string {
  const tableMatch = html.match(/<table>[\s\S]*?<\/table>/);
  if (!tableMatch) return '';
  const tableHtml = tableMatch[0];

  const rows: string[][] = [];
  const rowRegex = /<tr>([\s\S]*?)<\/tr>/g;
  let rowMatch: RegExpExecArray | null;

  while ((rowMatch = rowRegex.exec(tableHtml)) !== null) {
    const cells: string[] = [];
    const cellRegex = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/g;
    let cellMatch: RegExpExecArray | null;
    while ((cellMatch = cellRegex.exec(rowMatch[1])) !== null) {
      cells.push(cellMatch[1].replace(/<[^>]+>/g, '').trim());
    }
    rows.push(cells);
  }

  return rows.map(r => r.join('\t')).join('\n');
}
