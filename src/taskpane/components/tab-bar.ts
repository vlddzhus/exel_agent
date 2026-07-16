export interface TabDefinition {
  id: string;
  label: string;
  icon: string;
  badge?: number;
}

export interface TabBarOptions {
  tabs: TabDefinition[];
  activeTab: string;
  onTabChange: (tabId: string) => void;
}

export interface TabBarAPI {
  element: HTMLElement;
  setActiveTab(tabId: string): void;
  setBadge(tabId: string, count: number): void;
}

export function createTabBar(options: TabBarOptions): TabBarAPI {
  const nav = document.createElement("nav");
  nav.id = "tab-bar";
  nav.setAttribute("role", "tablist");
  nav.setAttribute("aria-label", "Main navigation");

  const buttons: Map<string, HTMLButtonElement> = new Map();

  for (const tab of options.tabs) {
    const btn = document.createElement("button");
    btn.className = `tab-btn ${tab.id === options.activeTab ? "active" : ""}`;
    btn.setAttribute("role", "tab");
    btn.setAttribute("aria-selected", String(tab.id === options.activeTab));
    btn.setAttribute("aria-controls", `panel-${tab.id}`);
    btn.dataset.tab = tab.id;

    let labelHtml = `<span class="tab-icon">${escapeHtml(tab.icon)}</span><span class="tab-label">${escapeHtml(tab.label)}</span>`;
    if (tab.badge !== undefined && tab.badge > 0) {
      labelHtml += `<span class="tab-badge">${tab.badge}</span>`;
    }
    btn.innerHTML = labelHtml;

    btn.addEventListener("click", () => {
      if (tab.id !== options.activeTab) {
        options.onTabChange(tab.id);
      }
    });

    nav.appendChild(btn);
    buttons.set(tab.id, btn);
  }

  return {
    element: nav,
    setActiveTab(tabId: string) {
      options.activeTab = tabId;
      nav.querySelectorAll(".tab-btn").forEach((btn) => {
        const isActive = (btn as HTMLElement).dataset.tab === tabId;
        btn.classList.toggle("active", isActive);
        btn.setAttribute("aria-selected", String(isActive));
      });
    },
    setBadge(tabId: string, count: number) {
      const btn = buttons.get(tabId);
      if (!btn) return;
      const existingBadge = btn.querySelector(".tab-badge");
      if (existingBadge) {
        if (count > 0) {
          existingBadge.textContent = String(count);
        } else {
          existingBadge.remove();
        }
      } else if (count > 0) {
        const badge = document.createElement("span");
        badge.className = "tab-badge";
        badge.textContent = String(count);
        btn.appendChild(badge);
      }
    },
  };
}

function escapeHtml(s: string): string {
  const div = document.createElement("div");
  div.textContent = s;
  return div.innerHTML;
}
