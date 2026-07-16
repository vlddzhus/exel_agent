import { useCallback } from "react";
import { undoManager } from "../tools/backup";

export function useUndo() {
  const snapshotIfRequired = useCallback(
    async (toolName: string, address: string) => {
      const writeTools = new Set([
        "setValues",
        "setFormula",
        "fillRange",
        "clearRange",
        "applyCellFormat",
        "applyConditionalFormat",
        "formatAsTable",
        "manageSheets",
        "removeDuplicates",
        "sortData",
      ]);
      if (writeTools.has(toolName) && address) {
        await undoManager.createBackup(address, toolName);
      }
    },
    [],
  );

  const undoAll = useCallback(async () => {
    const stack = undoManager.getStack();
    for (const entry of [...stack].reverse()) {
      await undoManager.restoreBackup(entry.id);
    }
    undoManager.clear();
  }, []);

  const commit = useCallback(() => {}, []);

  return { snapshotIfRequired, undoAll, commit };
}
