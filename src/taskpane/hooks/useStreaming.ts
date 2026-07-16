import { useCallback, useRef } from "react";
import { useLiveActivityStore } from "../stores/liveActivityStore";
import { useChatStore } from "../stores/chatStore";

// Человекочитаемые названия инструментов для Live-ленты (для незнакомых — как есть).
const TOOL_LABELS: Record<string, string> = {
  setValues: "Запись значений",
  setFormula: "Запись формул",
  appendRows: "Добавление строк",
  fillRange: "Заполнение диапазона",
  clearRange: "Очистка диапазона",
  getWorkbookOverview: "Обзор книги",
  getRange: "Чтение диапазона",
  getFormula: "Чтение формул",
  getRangeStats: "Статистика по диапазону",
  detectDataTypes: "Определение типов данных",
  findAnomalies: "Поиск аномалий",
  applyCellFormat: "Форматирование ячеек",
  applyNumberFormat: "Числовой формат",
  applyConditionalFormat: "Условное форматирование",
  formatAsTable: "Оформление в таблицу",
  autoFitColumns: "Авторазмер колонок",
  manageSheets: "Операции с листами",
  manageTable: "Операции с таблицами",
  createPivotTable: "Сводная таблица",
  createChart: "Создание диаграммы",
  freezePanes: "Закрепление областей",
  sortData: "Сортировка",
  filterData: "Фильтрация",
  removeDuplicates: "Удаление дубликатов",
  splitTextToColumns: "Разбивка текста по колонкам",
  normalizeText: "Нормализация текста",
  lookup: "Поиск (lookup)",
  listTables: "Список таблиц",
  createTable: "Создание таблицы",
  addTableRow: "Добавление строки в таблицу",
  sortTable: "Сортировка таблицы",
  filterTable: "Фильтр таблицы",
};

export function useStreaming() {
  const addThought = useLiveActivityStore((s) => s.addThought);
  const updateProgress = useLiveActivityStore((s) => s.updateProgress);
  const setStats = useLiveActivityStore((s) => s.setStats);
  const finish = useLiveActivityStore((s) => s.finish);
  const fail = useLiveActivityStore((s) => s.fail);
  const setStatus = useLiveActivityStore.setState;
  const addAssistantMessage = useChatStore((s) => s.addAssistantMessage);
  const setProcessing = useChatStore((s) => s.setProcessing);

  const lastThoughtTime = useRef(0);

  const handleEvent = useCallback(
    (event: string, data: any) => {
      switch (event) {
        case "thinking":
          const now = Date.now();
          if (now - lastThoughtTime.current > 300) {
            addThought(data.text);
            lastThoughtTime.current = now;
          }
          break;
        case "status":
          if (data.provider) {
            setStats({ provider: data.provider });
          }
          break;
        case "tool_call": {
          // Показываем в Live-ленте, что агент выполняет инструмент.
          // Раньше тут был мёртвый код: считали неиспользуемое имя и звали
          // markStepDone для несуществующего шага плана. Теперь — реальная
          // обратная связь: статус "executing" + человекочитаемая метка.
          const toolName = data.function?.name || data.name || "инструмент";
          const label = TOOL_LABELS[toolName] || toolName;
          setStatus({ status: "executing" });
          updateProgress(0, 0, label);
          addThought(`▶ Выполняю: ${label}`);
          break;
        }
        case "done":
          if (data.usage) {
            setStats({
              tokensIn: data.usage.input_tokens,
              tokensOut: data.usage.output_tokens,
            });
          }
          if (data.choices?.[0]?.message?.content) {
            addAssistantMessage(data.choices[0].message.content);
          }
          finish();
          setProcessing(false);
          break;
        case "error":
          fail(data.code, data.message);
          setProcessing(false);
          break;
        case "model_fallback":
          setStats({ provider: data.to });
          break;
      }
    },
    [
      addThought,
      updateProgress,
      setStats,
      finish,
      fail,
      addAssistantMessage,
      setProcessing,
    ],
  );

  return { handleEvent };
}
