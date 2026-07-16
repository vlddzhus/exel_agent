export interface PromptCategory {
  id: string;
  label: string;
  icon: string;
  prompts: PromptItem[];
}

export interface PromptItem {
  text: string;
  description?: string;
}

export interface PromptLibraryOptions {
  onSelectPrompt: (text: string, autoSend?: boolean) => void;
}

export interface PromptLibraryAPI {
  element: HTMLElement;
  refresh: () => void;
}

const CUSTOM_PROMPTS_KEY = 'saved_prompts';

const DEFAULT_CATEGORIES: PromptCategory[] = [
  {
    id: 'analysis',
    label: 'Анализ данных',
    icon: '📊',
    prompts: [
      { text: 'Проанализировать выделенный диапазон и показать статистику', description: 'Среднее, медиана, количество, выбросы' },
      { text: 'Найти и выделить дубликаты в выделении', description: 'Обнаружить дублирующиеся строки' },
      { text: 'Показать корреляцию между столбцом A и B', description: 'Корреляционный анализ' },
      { text: 'Определить топ-10 значений по сумме продаж', description: 'Сортировка и ранжирование' },
    ],
  },
  {
    id: 'charts',
    label: 'Диаграммы и графики',
    icon: '📈',
    prompts: [
      { text: 'Создать столбчатую диаграмму продаж по регионам', description: 'Визуализация столбчатой диаграммы' },
      { text: 'Создать круговую диаграмму распределения категорий', description: 'Круговая диаграмма' },
      { text: 'Создать линейную диаграмму тренда по времени', description: 'Линейная диаграмма для временных рядов' },
    ],
  },
  {
    id: 'cleaning',
    label: 'Очистка данных',
    icon: '🧹',
    prompts: [
      { text: 'Удалить все пустые строки из выделения', description: 'Очистка пустых строк' },
      { text: 'Стандартизировать формат даты в ДД/ММ/ГГГГ в столбце A', description: 'Нормализация формата дат' },
      { text: 'Обрезать пробелы во всех ячейках выделения', description: 'Удаление лишних пробелов' },
      { text: 'Найти и заменить "старое" на "новое" в выделении', description: 'Поиск и замена' },
    ],
  },
  {
    id: 'formatting',
    label: 'Форматирование',
    icon: '🎨',
    prompts: [
      { text: 'Отформатировать выделение как таблицу с заголовками', description: 'Преобразовать в таблицу' },
      { text: 'Применить чередование цветов строк в выделении', description: 'Чередование строк' },
      { text: 'Сделать строку заголовка жирной и добавить границы', description: 'Форматирование заголовка' },
      { text: 'Применить условное форматирование: подсветить ячейки > 100 зелёным', description: 'Условное форматирование' },
    ],
  },
  {
    id: 'reports',
    label: 'Отчёты и экспорт',
    icon: '📝',
    prompts: [
      { text: 'Сгенерировать сводный отчёт по этому листу', description: 'Автоотчёт' },
      { text: 'Создать сводную таблицу с группировкой по региону', description: 'Создание сводной таблицы' },
      { text: 'Экспортировать выделение как форматированный текст', description: 'Текстовый экспорт' },
    ],
  },
];

function loadCustomPrompts(): PromptItem[] {
  try {
    const data = JSON.parse(localStorage.getItem(CUSTOM_PROMPTS_KEY) || '[]');
    return data.map((p: { text: string; date: string }) => ({
      text: p.text,
      description: `Сохранено ${new Date(p.date).toLocaleDateString()}`,
    }));
  } catch {
    return [];
  }
}

export function createPromptLibrary(options: PromptLibraryOptions): PromptLibraryAPI {
  const container = document.createElement('div');
  container.id = 'panel-prompts';
  container.className = 'tab-panel active';

  function render() {
    let html = '<div class="prompts-content">';

    // Custom (saved) prompts
    const custom = loadCustomPrompts();
    if (custom.length > 0) {
      html += `<div class="prompt-category">`;
      html += `<div class="prompt-category-header">
        <span class="prompt-category-icon">💾</span>
        <span class="prompt-category-label">Сохранённые запросы (${custom.length})</span>
      </div>`;
      html += `<div class="prompt-category-items">`;
      for (const prompt of custom) {
        html += `<button class="prompt-item" data-prompt="${escapeAttr(prompt.text)}" title="${escapeAttr(prompt.description || '')}">
          <span class="prompt-item-text">${escapeHtml(prompt.text)}</span>
          <span class="prompt-item-desc">${escapeHtml(prompt.description || '')}</span>
        </button>`;
      }
      html += `</div></div>`;
    }

    for (const category of DEFAULT_CATEGORIES) {
      html += `<div class="prompt-category">`;
      html += `<div class="prompt-category-header">
        <span class="prompt-category-icon">${escapeHtml(category.icon)}</span>
        <span class="prompt-category-label">${escapeHtml(category.label)}</span>
      </div>`;
      html += `<div class="prompt-category-items">`;

      for (const prompt of category.prompts) {
        html += `<button class="prompt-item" data-prompt="${escapeAttr(prompt.text)}" title="${escapeAttr(prompt.description || '')}">
          <span class="prompt-item-text">${escapeHtml(prompt.text)}</span>
          ${prompt.description ? `<span class="prompt-item-desc">${escapeHtml(prompt.description)}</span>` : ''}
        </button>`;
      }

      html += `</div></div>`;
    }

    html += '</div>';
    container.innerHTML = html;

    container.querySelectorAll('.prompt-item').forEach((item) => {
      item.addEventListener('click', () => {
        const text = (item as HTMLElement).dataset.prompt;
        if (text) {
          options.onSelectPrompt(text, true);
        }
      });
    });
  }

  render();

  return {
    element: container,
    refresh: render,
  };
}

function escapeHtml(s: string): string {
  const div = document.createElement('div');
  div.textContent = s;
  return div.innerHTML;
}

function escapeAttr(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
