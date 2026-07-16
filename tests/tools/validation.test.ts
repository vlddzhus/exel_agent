/**
 * Тесты для _shared/validation.ts — валидация аргументов инструментов.
 *
 * Покрытие: address validation, sanitizeTableName, sensitive data detection.
 * Критично для безопасности (path traversal) и data protection (ПДн).
 */
import {
  assertFullAddress,
  assertSheetName,
  sanitizeTableName,
  detectSensitiveData,
  hasSensitiveData,
  validateAndSplitAddress,
} from "../../src/taskpane/tools/_shared/validation";

// ---------------------------------------------------------------------------
// assertFullAddress — path traversal защита
// ---------------------------------------------------------------------------

describe("assertFullAddress", () => {
  test.each([
    ["A1"],
    ["A1:B2"],
    ["Лист1!A1"],
    ["'Мой лист'!A1:B2"],
    ["XFD1048576"],
  ])("не бросает для валидного %j", (addr) => {
    expect(() => assertFullAddress(addr)).not.toThrow();
  });

  test.each([
    ["../../etc/passwd"],
    ["Лист!../../secret"],
    ["A1/../../etc"],
    [""],
    ["!A1"],
    ["Лист!"],
    ["A1:B2:C3"],
  ])("бросает Error для невалидного %j", (addr) => {
    expect(() => assertFullAddress(addr)).toThrow();
  });

  test("сообщение об ошибке содержит подсказку", () => {
    try {
      assertFullAddress("../etc/passwd");
      fail("Должно бросить");
    } catch (e) {
      expect((e as Error).message).toContain("Невалидный");
      expect((e as Error).message).toContain("A1");
    }
  });
});

// ---------------------------------------------------------------------------
// assertSheetName
// ---------------------------------------------------------------------------

describe("assertSheetName", () => {
  test.each([["Лист1"], ["МойЛист"], ["'Мой лист'"], ["Sheet1"], ["Данные"]])(
    "не бросает для валидного %j",
    (name) => {
      expect(() => assertSheetName(name)).not.toThrow();
    },
  );

  test.each([
    [""],
    ["Лист:1"],
    ["Лист/1"],
    ["Лист\\1"],
    ["Лист?1"],
    ["Лист*1"],
    ["[Лист]"],
    ["a".repeat(32)],
  ])("бросает для невалидного %j", (name) => {
    expect(() => assertSheetName(name)).toThrow();
  });
});

// ---------------------------------------------------------------------------
// sanitizeTableName
// ---------------------------------------------------------------------------

describe("sanitizeTableName", () => {
  test("обрезает пробелы", () => {
    expect(sanitizeTableName("  Продажи  ")).toBe("Продажи");
  });

  test("удаляет запрещённые символы", () => {
    expect(sanitizeTableName("Прода/жи:1")).toBe("Продажи1");
    expect(sanitizeTableName("Лист[1]")).toBe("Лист1");
  });

  test("ограничивает длину 31 символом", () => {
    const long = "a".repeat(50);
    const result = sanitizeTableName(long);
    expect(result.length).toBe(31);
  });

  test("бросает для пустого", () => {
    expect(() => sanitizeTableName("")).toThrow();
    expect(() => sanitizeTableName("   ")).toThrow();
    expect(() => sanitizeTableName("///")).toThrow(); // после очистки пусто
  });
});

// ---------------------------------------------------------------------------
// detectSensitiveData / hasSensitiveData — ПДн
// ---------------------------------------------------------------------------

describe("detectSensitiveData", () => {
  test("распознаёт паспорт РФ", () => {
    const text = "Мой паспорт 4510 123456 выдан...";
    const findings = detectSensitiveData(text);
    expect(findings.length).toBeGreaterThan(0);
    expect(findings.some((f) => f.type === "passport")).toBe(true);
  });

  test("распознаёт номер банковской карты", () => {
    const text = "Карта 4111 1111 1111 1111";
    const findings = detectSensitiveData(text);
    expect(findings.some((f) => f.type === "card")).toBe(true);
  });

  test("распознаёт ИНН", () => {
    const text = "ИНН организации 7707083893";
    const findings = detectSensitiveData(text);
    expect(findings.some((f) => f.type === "inn")).toBe(true);
  });

  test("распознаёт СНИЛС", () => {
    const text = "СНИЛС: 112-233-445 95";
    const findings = detectSensitiveData(text);
    expect(findings.some((f) => f.type === "snils")).toBe(true);
  });

  test("распознаёт российский телефон", () => {
    const text = "Звоните +7 (495) 123-45-67";
    const findings = detectSensitiveData(text);
    expect(findings.some((f) => f.type === "phone_ru")).toBe(true);
  });

  test("распознаёт email", () => {
    const text = "Пишите на ivan@example.com";
    const findings = detectSensitiveData(text);
    expect(findings.some((f) => f.type === "email")).toBe(true);
  });

  test("возвращает пустой массив для обычного текста", () => {
    expect(detectSensitiveData("Обычный текст про продажи за месяц")).toEqual([]);
  });

  test("возвращает пустой массив для пустого/невалидного ввода", () => {
    expect(detectSensitiveData("")).toEqual([]);
    expect(detectSensitiveData(null as unknown as string)).toEqual([]);
  });

  test("находит несколько типов ПДн в одном тексте", () => {
    const text = "Иван, ivan@example.com, +7 999 123 45 67";
    const findings = detectSensitiveData(text);
    const types = findings.map((f) => f.type);
    expect(types).toContain("email");
    expect(types).toContain("phone_ru");
  });
});

describe("hasSensitiveData", () => {
  test("true для текста с ПДн", () => {
    expect(hasSensitiveData("Карта 4111 1111 1111 1111")).toBe(true);
  });

  test("false для обычного текста", () => {
    expect(hasSensitiveData("Обычный текст")).toBe(false);
  });

  test("false для пустого", () => {
    expect(hasSensitiveData("")).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// validateAndSplitAddress
// ---------------------------------------------------------------------------

describe("validateAndSplitAddress", () => {
  test("возвращает распарсенный адрес для валидного ввода", () => {
    const result = validateAndSplitAddress("Лист1!A1:B2");
    expect(result).toEqual({ sheetName: "Лист1", rangeAddress: "A1:B2" });
  });

  test("работает без имени листа", () => {
    const result = validateAndSplitAddress("A1");
    expect(result).toEqual({ rangeAddress: "A1" });
    expect(result.sheetName).toBeUndefined();
  });

  test("бросает для path traversal", () => {
    expect(() => validateAndSplitAddress("../etc/passwd")).toThrow();
  });
});
