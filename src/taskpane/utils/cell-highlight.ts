const HIGHLIGHT_COLOR = "#E8F5E9";
const RESET_DELAY = 2000;

let activeHighlights: string[] = [];

export async function highlightRange(address: string): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(address);
      range.load(["address", "format/fill/color"]);
      await context.sync();

      const originalColor = range.format.fill.color || "None";
      range.format.fill.color = HIGHLIGHT_COLOR;
      activeHighlights.push(range.address);

      setTimeout(async () => {
        try {
          await Excel.run(async (ctx) => {
            const reset = sheet.getRange(range.address);
            reset.format.fill.color = originalColor;
            await ctx.sync();
          });
        } catch {}
        activeHighlights = activeHighlights.filter((a) => a !== range.address);
      }, RESET_DELAY);

      await context.sync();
    });
  } catch {}
}

export function clearHighlights(): void {
  if (activeHighlights.length === 0) return;
  activeHighlights.forEach((addr) => {
    Excel.run(async (context) => {
      try {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(addr);
        range.format.fill.color = "None";
        await context.sync();
      } catch {}
    });
  });
  activeHighlights = [];
}
