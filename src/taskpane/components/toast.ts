export type ToastType = 'success' | 'error' | 'warning' | 'info';

export interface ToastOptions {
  message: string;
  type?: ToastType;
  duration?: number;
}

const DEFAULT_DURATION = 7000;

export function showToast(options: ToastOptions) {
  const { message, type = 'info', duration = DEFAULT_DURATION } = options;

  let container = document.getElementById('toast-container');
  if (!container) {
    container = document.createElement('div');
    container.id = 'toast-container';
    document.body.appendChild(container);
  }

  const icons: Record<ToastType, string> = {
    success: '✓',
    error: '✕',
    warning: '⚠',
    info: 'ℹ',
  };

  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.innerHTML = `
    <span class="toast-icon">${icons[type]}</span>
    <span class="toast-message">${escapeHtml(message)}</span>
    <button class="toast-dismiss" aria-label="Закрыть">✕</button>
  `;

  container.appendChild(toast);

  let dismissTimer: ReturnType<typeof setTimeout> | null = null;
  let removalTimer: ReturnType<typeof setTimeout> | null = null;

  function scheduleDismiss() {
    dismissTimer = setTimeout(() => {
      toast.classList.add('removing');
      removalTimer = setTimeout(() => {
        toast.remove();
        if (container && container.children.length === 0) {
          container.remove();
        }
      }, 200);
    }, duration);
  }

  function cancelDismiss() {
    if (dismissTimer) { clearTimeout(dismissTimer); dismissTimer = null; }
    if (removalTimer) { clearTimeout(removalTimer); removalTimer = null; }
  }

  scheduleDismiss();

  toast.querySelector('.toast-dismiss')?.addEventListener('click', () => {
    cancelDismiss();
    toast.classList.add('removing');
    setTimeout(() => {
      toast.remove();
      if (container && container.children.length === 0) {
        container.remove();
      }
    }, 200);
  });

  toast.addEventListener('mouseenter', cancelDismiss);
  toast.addEventListener('mouseleave', scheduleDismiss);
}

function escapeHtml(s: string): string {
  const div = document.createElement('div');
  div.textContent = s;
  return div.innerHTML;
}
