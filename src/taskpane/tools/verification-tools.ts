import { toolRegistry } from "./registry";
import { getRangeSafe } from "./address-helper";
import { columnToLetter, letterToColumn } from "./_shared/address";

const EXCEL_ERRORS = [
  "#REF!",
  "#VALUE!",
  "#DIV/0!",
  "#NAME?",
  "#N/A",
  "#NULL!",
  "#NUM!",
  "#SPILL!",
  "#CALC!",
];

export interface VerificationResult {
  hasErrors: boolean;
  errors: { address: string; value: string }[];
  values: unknown[][];
}

export async function verifyRange(
  address: string,
): Promise<VerificationResult> {
  return Excel.run(async (context) => {
    const range = getRangeSafe(context, address);
    range.load("values, address");
    await context.sync();

    const values = range.values as unknown[][];
    const errors: { address: string; value: string }[] = [];

    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < (values[r]?.length ?? 0); c++) {
        const cellValue = String(values[r][c] ?? "");
        if (EXCEL_ERRORS.some((err) => cellValue.startsWith(err))) {
          const cellAddress = getCellAddress(range.address, r, c);
          errors.push({ address: cellAddress, value: cellValue });
        }
      }
    }

    return { hasErrors: errors.length > 0, errors, values };
  });
}

function getCellAddress(
  rangeAddress: string,
  rowOffset: number,
  colOffset: number,
): string {
  const match = rangeAddress.match(/([A-Z]+)(\d+):?/);
  if (!match) return `${rangeAddress}!R${rowOffset + 1}C${colOffset + 1}`;
  const startCol = match[1];
  const startRow = parseInt(match[2], 10);
  const col = columnToLetter(letterToColumn(startCol) + colOffset);
  const row = startRow + rowOffset;
  return `${col}${row}`;
}

export async function verifyToolResult(
  toolName: string,
  args: Record<string, unknown>,
  _resultJson: string,
): Promise<{ verified: boolean; message?: string }> {
  const mutationTools = [
    "setValues",
    "setFormula",
    "fillFormula",
    "clearRange",
    "clearWorksheet",
  ];

  if (!mutationTools.includes(toolName)) {
    return { verified: true };
  }

  let address: string | undefined;

  if (toolName === "clearWorksheet") {
    const sheetName = args.name as string | undefined;
    if (sheetName) {
      address = `'${sheetName}'!A1:Z10000`;
    } else {
      address = "A1:Z10000";
    }
  } else {
    address = (args.address ?? args.cellAddress ?? args.targetRange) as
      string | undefined;
  }

  if (!address) {
    return { verified: true };
  }

  const result = await verifyRange(address);

  if (result.hasErrors) {
    const errorList = result.errors
      .map((e) => `${e.address}: ${e.value}`)
      .join("; ");
    return {
      verified: false,
      message: `Tool ${toolName} completed but Excel errors found in range ${address}: ${errorList}. Please analyze the data and fix these errors.`,
    };
  }

  return { verified: true };
}
