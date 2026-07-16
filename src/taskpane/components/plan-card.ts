export interface PlanStep {
  number: number;
  summary: string;
  details?: string;
  icon: string;
  status: "pending" | "approved" | "executing" | "completed" | "failed";
  riskLevel?: "safe" | "moderate" | "dangerous";
  canRunInParallel?: boolean;
  toolName?: string;
}

export type PlanCardState =
  | "pending"
  | "approved"
  | "executing"
  | "completed"
  | "failed";

export interface PlanCardResult {
  approved: boolean;
  editedPlan?: string;
}

export interface PlanCardOptions {
  planText: string;
  steps: PlanStep[];
  onResult: (result: PlanCardResult) => void;
}

export interface PlanCardAPI {
  element: HTMLElement;
  setState(state: PlanCardState): void;
  setStepStatus(stepNumber: number, status: PlanStep["status"]): void;
  destroy(): void;
}

const MAX_VISIBLE_STEPS = 3;
const STEP_ICONS: Record<string, string> = {
  pending: "○",
  approved: "○",
  executing: "⏳",
  completed: "✓",
  failed: "✕",
};

export function parsePlanSteps(planText: string): PlanStep[] {
  const lines = planText.split("\n");
  const steps: PlanStep[] = [];
  let currentStep: Partial<PlanStep> | null = null;

  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line) {
      if (currentStep?.details !== undefined) {
        currentStep.details += "\n";
      }
      continue;
    }

    const stepMatch = line.match(/^(\d+)[.)]\s*(.+)/);
    const checkboxMatch = line.match(/^[-*]\s*\[.?\]\s*(.+)/);

    if (stepMatch || checkboxMatch) {
      if (currentStep && currentStep.summary) {
        steps.push({
          number: currentStep.number || steps.length + 1,
          summary: currentStep.summary,
          details: currentStep.details?.trim() || undefined,
          icon: currentStep.icon || assignIcon(currentStep.summary),
          status: "pending",
        });
      }

      const summary = (stepMatch ? stepMatch[2] : checkboxMatch![1]).trim();
      currentStep = {
        number: stepMatch ? parseInt(stepMatch[1], 10) : steps.length + 1,
        summary: summary.replace(/^[📊📈🔧🧹🎨📋📝]\s*/, ""),
        icon: assignIcon(summary),
        details: "",
      };
    } else if (currentStep) {
      currentStep.details = (currentStep.details || "") + line + "\n";
    }
  }

  if (currentStep && currentStep.summary) {
    steps.push({
      number: currentStep.number || steps.length + 1,
      summary: currentStep.summary,
      details: currentStep.details?.trim() || undefined,
      icon: currentStep.icon || assignIcon(currentStep.summary),
      status: "pending",
    });
  }

  if (steps.length === 0) {
    steps.push({
      number: 1,
      summary: planText.trim().substring(0, 120),
      icon: "📋",
      status: "pending",
    });
  }

  return steps;
}

function assignIcon(text: string): string {
  const lower = text.toLowerCase();
  if (/\b(chart|graph|plot|visualize|dashboard)\b/.test(lower)) return "📈";
  if (/\b(analy|summar|statistic|mean|avg|total|count)\b/.test(lower))
    return "📊";
  if (/\b(clean|duplicate|remove|delete|empty|null|missing)\b/.test(lower))
    return "🧹";
  if (/\b(format|style|color|bold|header|border|align)\b/.test(lower))
    return "🎨";
  if (/\b(sort|filter|order|group)\b/.test(lower)) return "📋";
  if (/\b(report|export|save|output|print)\b/.test(lower)) return "📝";
  if (/\b(formula|calculate|compute|sum|if|vlookup)\b/.test(lower)) return "🔢";
  if (/\b(pivot|table|crosstab|aggregate)\b/.test(lower)) return "📊";
  if (/\b(read|load|import|fetch|get|extract)\b/.test(lower)) return "📂";
  if (/\b(update|set|write|put|change|modify|edit)\b/.test(lower)) return "✏️";
  if (/\b(check|verify|validate|test|review|confirm)\b/.test(lower))
    return "🔍";
  return "🔧";
}

export function createPlanCard(options: PlanCardOptions): PlanCardAPI {
  const state: { current: PlanCardState } = { current: "pending" };
  const hasManySteps = options.steps.length > MAX_VISIBLE_STEPS;
  let allVisible = false;

  const card = document.createElement("div");
  card.className = "plan-card";
  card.id = "plan-card";
  card.setAttribute("role", "region");
  card.setAttribute(
    "aria-label",
    `Action plan (${options.steps.length} step${options.steps.length !== 1 ? "s" : ""})`,
  );

  function renderSteps(visible: boolean): string {
    const showAll = !hasManySteps || visible;
    return options.steps
      .map((step, idx) => {
        const isHidden = hasManySteps && !showAll && idx >= MAX_VISIBLE_STEPS;
        const riskBadge =
          step.riskLevel === "dangerous"
            ? '<span class="plan-risk-badge danger" title="Dangerous">⚠️</span>'
            : step.riskLevel === "moderate"
              ? '<span class="plan-risk-badge moderate" title="Modifies data">✏️</span>'
              : "";
        const parallelBadge = step.canRunInParallel
          ? '<span class="plan-parallel-badge" title="Can run in parallel">⚡</span>'
          : "";
        const toolBadge = step.toolName
          ? `<span class="plan-tool-badge">${escapeHtml(step.toolName)}</span>`
          : "";
        return `
    <div class="plan-step${isHidden ? " plan-step-hidden" : ""}" data-step="${step.number}">
      <div class="plan-step-header">
        <span class="plan-step-icon pending">${escapeHtml(step.icon)}</span>
        <span class="plan-step-summary">${escapeHtml(step.summary)}</span>
        ${riskBadge}${parallelBadge}
        ${step.details ? '<span class="plan-step-chevron">▶</span>' : ""}
      </div>
      ${step.details ? `<div class="plan-step-details">${escapeHtml(step.details)}${toolBadge ? `<br>${toolBadge}` : ""}</div>` : ""}
    </div>`;
      })
      .join("");
  }

  card.innerHTML = `
    <div class="plan-card-header">
      <div class="plan-card-title">
        <span>📋 ${options.steps.length} step${options.steps.length !== 1 ? "s" : ""}</span>
        <button class="plan-card-edit-btn" id="plan-edit-btn" title="Edit">✏️</button>
      </div>
    </div>
    <div class="plan-steps">${renderSteps(false)}</div>
    ${hasManySteps ? `<button class="plan-expand-btn" id="plan-expand-btn">Show all ${options.steps.length} steps ↓</button>` : ""}
    <div class="plan-actions">
      <button class="plan-btn plan-btn-approve" id="plan-approve-btn">✓ Execute</button>
      <button class="plan-btn plan-btn-cancel" id="plan-cancel-btn" title="Cancel">✕</button>
    </div>
  `;

  // ── Step expand/collapse ──

  const stepsElements = card.querySelectorAll(".plan-step-header");
  stepsElements.forEach((header) => {
    header.setAttribute("role", "button");
    header.setAttribute("tabindex", "0");
    header.setAttribute("aria-expanded", "false");
    const stepEl = header.closest(".plan-step") as HTMLElement;

    function toggleStep() {
      if (!stepEl) return;
      const details = stepEl.querySelector(".plan-step-details") as HTMLElement;
      const chevron = header.querySelector(".plan-step-chevron") as HTMLElement;
      if (details && chevron) {
        const isOpen = details.classList.contains("open");
        details.classList.toggle("open");
        chevron.classList.toggle("open");
        header.setAttribute("aria-expanded", String(!isOpen));
      }
    }

    header.addEventListener("click", toggleStep);
    header.addEventListener("keydown", (evt: Event) => {
      const ke = evt as KeyboardEvent;
      if (ke.key === "Enter" || ke.key === " ") {
        ke.preventDefault();
        toggleStep();
      }
    });
  });

  // ── Show all steps toggle ──

  const expandBtn = card.querySelector("#plan-expand-btn") as HTMLButtonElement;
  if (expandBtn) {
    expandBtn.addEventListener("click", () => {
      allVisible = true;
      const hidden = card.querySelectorAll(".plan-step-hidden");
      hidden.forEach((el) => {
        el.classList.remove("plan-step-hidden");
        el.classList.add("plan-step-revealed");
      });
      expandBtn.textContent = "Collapse ↑";
      expandBtn.addEventListener(
        "click",
        () => {
          allVisible = false;
          const revealed = card.querySelectorAll(".plan-step-revealed");
          revealed.forEach((el) => {
            el.classList.add("plan-step-hidden");
            el.classList.remove("plan-step-revealed");
          });
          expandBtn.textContent = `Show all ${options.steps.length} steps ↓`;
        },
        { once: true },
      );
    });
  }

  // ── Buttons ──

  const approveBtn = card.querySelector(
    "#plan-approve-btn",
  ) as HTMLButtonElement;
  const cancelBtn = card.querySelector("#plan-cancel-btn") as HTMLButtonElement;
  const editBtn = card.querySelector("#plan-edit-btn") as HTMLButtonElement;

  approveBtn.addEventListener("click", () => {
    if (state.current === "executing" || state.current === "completed") return;
    state.current = "executing";
    approveBtn.disabled = true;
    cancelBtn.disabled = true;
    editBtn.disabled = true;
    options.onResult({ approved: true });
  });

  cancelBtn.addEventListener("click", () => {
    options.onResult({ approved: false });
  });

  editBtn.addEventListener("click", () => {
    showEditOverlay(options.planText, (newText) => {
      options.onResult({ approved: true, editedPlan: newText });
    });
  });

  return {
    element: card,
    setState(newState: PlanCardState) {
      state.current = newState;
    },
    setStepStatus(stepNumber: number, status: PlanStep["status"]) {
      const stepEl = card.querySelector(
        `.plan-step[data-step="${stepNumber}"]`,
      ) as HTMLElement;
      if (!stepEl) return;
      const icon = stepEl.querySelector(".plan-step-icon") as HTMLElement;
      if (icon) {
        icon.className = `plan-step-icon ${status}`;
      }
    },
    destroy() {
      card.remove();
    },
  };
}

function showEditOverlay(planText: string, onSave: (newText: string) => void) {
  const overlay = document.createElement("div");
  overlay.className = "plan-edit-overlay";

  overlay.innerHTML = `
    <div class="plan-edit-box">
      <h3>✏️ Edit plan</h3>
      <p>Modify the steps below. Changes will be sent to the agent.</p>
      <textarea id="plan-edit-textarea">${escapeHtml(planText)}</textarea>
      <div class="edit-actions">
        <button class="btn-discard" id="plan-edit-discard">Cancel</button>
        <button class="btn-save" id="plan-edit-save">Save</button>
      </div>
    </div>
  `;

  document.body.appendChild(overlay);

  const textarea = overlay.querySelector(
    "#plan-edit-textarea",
  ) as HTMLTextAreaElement;
  textarea.focus();
  textarea.setSelectionRange(textarea.value.length, textarea.value.length);

  overlay.querySelector("#plan-edit-save")?.addEventListener("click", () => {
    overlay.remove();
    onSave(textarea.value);
  });

  overlay.querySelector("#plan-edit-discard")?.addEventListener("click", () => {
    overlay.remove();
  });

  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) overlay.remove();
  });
}

function escapeHtml(s: string): string {
  const div = document.createElement("div");
  div.textContent = s;
  return div.innerHTML;
}
