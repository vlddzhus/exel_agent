/**
 * performance-guard.ts — legacy re-export.
 *
 * Реальная реализация: ./_shared/performance.ts (см. docs/03-TOOLS-SPEC.md §3.3).
 * Этот файл сохранён для обратной совместимости с кодом, импортирующим
 * withPerformanceGuard из performance-guard. Новый код должен импортировать
 * напрямую из ./_shared/performance.
 */
export { withPerformanceGuard } from "./_shared/performance";
