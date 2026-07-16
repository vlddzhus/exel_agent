import { ChatSession, getSessions, getSessionsAsync, deleteSession, deleteSessionAsync, renameSession, exportSessionText } from '../utils/session-store';
import { showToast } from './toast';

export interface HistoryOptions {
  onSelectSession: (session: ChatSession) => void;
  onNewChat: () => void;
}

export interface HistoryAPI {
  element: HTMLElement;
  refresh(): void;
}

export function createHistory(options: HistoryOptions): HistoryAPI {
  const container = document.createElement('div');
  container.id = 'panel-history';
  container.className = 'tab-panel active';

  let currentSessions: ChatSession[] = [];

  async function loadSessions(): Promise<void> {
    currentSessions = await getSessionsAsync();
  }

  function render() {
    if (currentSessions.length === 0) {
      container.innerHTML = `
        <div class="tab-panel-placeholder">
          <div class="placeholder-icon">📋</div>
          <div class="placeholder-title">История</div>
          <div class="placeholder-desc">Ваши прошлые диалоги появятся здесь. Начните чат, чтобы сохранить первую сессию.</div>
        </div>
      `;
      return;
    }

    let html = '<div class="history-header">';
    html += `<span class="history-count">${currentSessions.length} сесси${currentSessions.length !== 1 ? 'й' : 'я'}</span>`;
    html += '</div>';
    html += '<div class="history-list">';

    for (const session of currentSessions) {
      const date = new Date(session.date);
      const dateStr = date.toLocaleDateString(undefined, {
        month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit',
      });
      const preview = session.preview.length > 80
        ? session.preview.substring(0, 80) + '...'
        : session.preview;

      html += `
        <div class="history-item" data-session-id="${escapeHtml(session.id)}">
          <div class="history-item-main">
            <div class="history-item-title">${escapeHtml(session.title || 'Без названия')}</div>
            <div class="history-item-preview">${escapeHtml(preview)}</div>
            <div class="history-item-meta">
              <span>${dateStr}</span>
              <span>·</span>
              <span>${session.stepCount} шаг${session.stepCount !== 1 ? 'а' : ''}</span>
              <span>·</span>
              <span>~${session.tokenCount} токенов</span>
            </div>
          </div>
          <div class="history-item-actions">
            <button class="history-action-btn" data-action="rename" title="Переименовать" aria-label="Переименовать сессию">✏️</button>
            <button class="history-action-btn" data-action="export" title="Экспорт" aria-label="Экспорт сессии">📥</button>
            <button class="history-action-btn" data-action="delete" title="Удалить" aria-label="Удалить сессию">🗑</button>
          </div>
        </div>
      `;
    }

    html += '</div>';
    container.innerHTML = html;

    container.querySelectorAll('.history-item').forEach((item) => {
      const sessionId = (item as HTMLElement).dataset.sessionId;
      if (!sessionId) return;

      item.addEventListener('click', (e) => {
        const target = e.target as HTMLElement;
        const actionBtn = target.closest('.history-action-btn') as HTMLElement;
        if (actionBtn) return;

        const session = currentSessions.find(s => s.id === sessionId);
        if (session) {
          options.onSelectSession(session);
        }
      });

      item.querySelectorAll('.history-action-btn').forEach((btn) => {
        btn.addEventListener('click', (e) => {
          e.stopPropagation();
          const action = (btn as HTMLElement).dataset.action;
          const session = currentSessions.find(s => s.id === sessionId);
          if (!session) return;

          switch (action) {
            case 'rename': {
              const newTitle = prompt('Переименовать сессию:', session.title);
              if (newTitle && newTitle.trim()) {
                renameSession(sessionId, newTitle.trim());
                session.title = newTitle.trim();
                render();
                showToast({ message: 'Сессия переименована', type: 'success' });
              }
              break;
            }
            case 'export': {
              const text = exportSessionText(sessionId);
              navigator.clipboard.writeText(text).then(() => {
                showToast({ message: 'Сессия скопирована в буфер обмена', type: 'success' });
              });
              break;
            }
            case 'delete': {
              if (confirm(`Удалить "${session.title || 'Без названия'}"?`)) {
                currentSessions = currentSessions.filter(s => s.id !== sessionId);
                deleteSession(sessionId);
                deleteSessionAsync(sessionId).catch(() => {});
                render();
                showToast({ message: 'Сессия удалена', type: 'info' });
              }
              break;
            }
          }
        });
      });
    });
  }

  loadSessions().then(render);

  return {
    element: container,
    async refresh() {
      await loadSessions();
      render();
    },
  };
}

function escapeHtml(s: string): string {
  const div = document.createElement('div');
  div.textContent = s;
  return div.innerHTML;
}
