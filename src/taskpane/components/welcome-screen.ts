export interface SuggestionGroup {
  label: string;
  items: { icon: string; text: string }[];
}

export interface WelcomeScreenOptions {
  onSuggestionClick: (text: string) => void;
  suggestionGroups?: SuggestionGroup[];
}

const DEFAULT_GROUPS: SuggestionGroup[] = [
  {
    label: 'Анализ данных',
    items: [
      { icon: '📊', text: 'Проанализировать выделенный диапазон' },
      { icon: '🔍', text: 'Найти дубликаты в данных' },
      { icon: '📈', text: 'Показать сводную статистику' },
    ],
  },
  {
    label: 'Диаграммы и графики',
    items: [
      { icon: '📊', text: 'Создать столбчатую диаграмму' },
      { icon: '🧹', text: 'Форматировать как таблицу' },
    ],
  },
  {
    label: 'Форматирование',
    items: [
      { icon: '🎨', text: 'Применить условное форматирование' },
      { icon: '📋', text: 'Сортировать данные по столбцу' },
    ],
  },
];

export interface WelcomeScreenAPI {
  element: HTMLElement;
  destroy(): void;
}

export function createWelcomeScreen(options: WelcomeScreenOptions): WelcomeScreenAPI {
  const groups = options.suggestionGroups || DEFAULT_GROUPS;

  const container = document.createElement('div');
  container.className = 'welcome-screen';
  container.id = 'welcome-screen';

  let suggestionsHtml = '';
  for (const group of groups) {
    suggestionsHtml += `<div class="welcome-suggestion-group">`;
    suggestionsHtml += `<div class="welcome-group-label">${escapeHtml(group.label)}</div>`;
    for (const item of group.items) {
      suggestionsHtml += `<button class="welcome-chip" data-prompt="${escapeHtml(item.text)}">
        <span class="chip-icon">${escapeHtml(item.icon)}</span>
        <span>${escapeHtml(item.text)}</span>
      </button>`;
    }
    suggestionsHtml += `</div>`;
  }

  container.innerHTML = `
    <div class="welcome-avatar">🤖</div>
    <div class="welcome-title">AI Агент</div>
    <div class="welcome-subtitle">Что вы хотите сделать в таблице?</div>
    <div class="welcome-suggestions">${suggestionsHtml}</div>
  `;

  container.addEventListener('click', (e) => {
    const chip = (e.target as HTMLElement).closest('.welcome-chip') as HTMLElement;
    if (chip) {
      const prompt = chip.dataset.prompt;
      if (prompt) {
        options.onSuggestionClick(prompt);
      }
    }
  });

  return {
    element: container,
    destroy() {
      container.remove();
    },
  };
}

function escapeHtml(s: string): string {
  const div = document.createElement('div');
  div.textContent = s;
  return div.innerHTML;
}
