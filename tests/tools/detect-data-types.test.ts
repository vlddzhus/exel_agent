/**
 * Тесты для detectValueType и детекторов RU-форматов из read.ts.
 *
 * Критично: правильное распознавание ИНН/СНИЛС/телефона/дат определяет качество
 * всей последующей обработки данных (docs/03-TOOLS-SPEC.md §2 R4).
 */
import {
  isInn,
  isKpp,
  isOgrn,
  isSnils,
  isPhoneRu,
  isEmail,
  isUrl,
  isDateString,
  detectValueType,
  detectColumnType,
  type ColumnType,
} from "../../src/taskpane/tools/read";

// ---------------------------------------------------------------------------
// isInn — ИНН 10 или 12 цифр
// ---------------------------------------------------------------------------

describe("isInn", () => {
  test.each([
    ["7707083893", true], // ИНН юрлица 10 цифр (Сбербанк)
    ["500100732259", true], // ИНН физлица 12 цифр
    [7707083893, true], // number
    ["123", false], // мало цифр
    ["1234567890123", false], // 13 цифр
    ["", false],
    ["hello", false],
    [null, false],
  ])("isInn(%j) === %s", (input, expected) => {
    expect(isInn(input)).toBe(expected);
  });
});

// ---------------------------------------------------------------------------
// isKpp — 9 цифр
// ---------------------------------------------------------------------------

describe("isKpp", () => {
  test.each([
    ["123456789", true],
    ["773301001", true],
    [123456789, true],
    ["12345678", false], // 8 цифр
    ["1234567890", false], // 10 цифр
    ["", false],
  ])("isKpp(%j) === %s", (input, expected) => {
    expect(isKpp(input)).toBe(expected);
  });
});

// ---------------------------------------------------------------------------
// isOgrn — 13 (ОГРН) или 15 (ОГРНИП) цифр
// ---------------------------------------------------------------------------

describe("isOgrn", () => {
  test.each([
    ["1027700132195", true], // ОГРН 13 цифр
    ["315774600200100", true], // ОГРНИП 15 цифр
    ["123456789012", false], // 12 цифр
    ["12345678901234", false], // 14 цифр
    ["", false],
  ])("isOgrn(%j) === %s", (input, expected) => {
    expect(isOgrn(input)).toBe(expected);
  });
});

// ---------------------------------------------------------------------------
// isSnils — 11 цифр
// ---------------------------------------------------------------------------

describe("isSnils", () => {
  test.each([
    ["112-233-445 95", true],
    ["11223344595", true],
    [11223344595, true],
    ["123-456-789", false], // 9 цифр
    ["1234567890", false], // 10 цифр
    ["", false],
  ])("isSnils(%j) === %s", (input, expected) => {
    expect(isSnils(input)).toBe(expected);
  });
});

// ---------------------------------------------------------------------------
// isPhoneRu — +7/8 + 10 цифр
// ---------------------------------------------------------------------------

describe("isPhoneRu", () => {
  test.each([
    ["+7 (495) 123-45-67", true],
    ["+74951234567", true],
    ["89991234567", true],
    ["8 999 123 45 67", true],
    [84951234567, true],
    ["+7-999-123-45-67", true],
    ["12345", false], // слишком короткий
    ["hello", false],
    ["", false],
  ])("isPhoneRu(%j) === %s", (input, expected) => {
    expect(isPhoneRu(input)).toBe(expected);
  });
});

// ---------------------------------------------------------------------------
// isEmail
// ---------------------------------------------------------------------------

describe("isEmail", () => {
  test.each([
    ["ivan@example.com", true],
    ["test.mail+tag@sub.example.co.uk", true],
    ["ivan@", false],
    ["@example.com", false],
    ["ivanexample.com", false],
    ["", false],
    [42, false],
  ])("isEmail(%j) === %s", (input, expected) => {
    expect(isEmail(input)).toBe(expected);
  });
});

// ---------------------------------------------------------------------------
// isUrl
// ---------------------------------------------------------------------------

describe("isUrl", () => {
  test.each([
    ["https://example.com", true],
    ["http://example.com/path?q=1", true],
    ["example.com", false], // без схемы
    ["ftp://example.com", false], // не http(s)
    ["", false],
  ])("isUrl(%j) === %s", (input, expected) => {
    expect(isUrl(input)).toBe(expected);
  });
});

// ---------------------------------------------------------------------------
// isDateString — поддержка RU и ISO
// ---------------------------------------------------------------------------

describe("isDateString", () => {
  test.each([
    ["01.01.2024", true],
    ["1.1.2024", true],
    ["31.12.24", true],
    ["2024-01-15", true],
    ["01/01/2024", true],
    ["01-01-2024", true],
    ["15.06.2024 14:30", true],
    ["15.06.2024 14:30:00", true],
    [44197, true], // Excel serial date
    [1, true],
    ["hello", false],
    ["123", false], // не похож на дату
    ["", false],
    [null, false],
  ])("isDateString(%j) === %s", (input, expected) => {
    expect(isDateString(input)).toBe(expected);
  });
});

// ---------------------------------------------------------------------------
// detectValueType — приоритеты типов
// ---------------------------------------------------------------------------

describe("detectValueType — приоритеты", () => {
  test("empty для null/undefined/''", () => {
    expect(detectValueType(null)).toBe("empty");
    expect(detectValueType(undefined)).toBe("empty");
    expect(detectValueType("")).toBe("empty");
  });

  test("inn приоритетнее general number (10 цифр)", () => {
    expect(detectValueType("7707083893")).toBe("inn");
  });

  test("phone приоритетнее text для +7...", () => {
    expect(detectValueType("+7 999 123 45 67")).toBe("phone");
  });

  test("email", () => {
    expect(detectValueType("ivan@example.com")).toBe("email");
  });

  test("date для строки DD.MM.YYYY", () => {
    expect(detectValueType("01.01.2024")).toBe("date");
  });

  test("integer для числа без дробной части", () => {
    expect(detectValueType(42)).toBe("integer");
    expect(detectValueType("42")).toBe("integer");
  });

  test("float для числа с дробью", () => {
    expect(detectValueType(3.14)).toBe("float");
    expect(detectValueType("3,14")).toBe("float");
  });

  test("percent для '50%'", () => {
    expect(detectValueType("50%")).toBe("percent");
  });

  test("text для обычной строки", () => {
    expect(detectValueType("Иван Иванов")).toBe("text");
  });
});

// ---------------------------------------------------------------------------
// detectColumnType — доминирующий тип колонки
// ---------------------------------------------------------------------------

describe("detectColumnType — доминирующий тип", () => {
  test("пустая колонка → empty", () => {
    expect(detectColumnType([null, undefined, "", null])).toEqual({
      type: "empty",
      confidence: 1,
    });
  });

  test("все числа → integer", () => {
    const result = detectColumnType([1, 2, 3, 4, 5]);
    expect(result.type).toBe("integer");
    expect(result.confidence).toBe(1);
  });

  test("все даты → date", () => {
    const result = detectColumnType([
      "01.01.2024",
      "15.02.2024",
      "30.03.2024",
      "01.04.2024",
    ]);
    expect(result.type).toBe("date");
    expect(result.confidence).toBe(1);
  });

  test("все email → email", () => {
    const result = detectColumnType([
      "ivan@example.com",
      "maria@example.com",
      "petr@example.com",
    ]);
    expect(result.type).toBe("email");
    expect(result.confidence).toBe(1);
  });

  test("mixed для разнородных значений", () => {
    const result = detectColumnType([
      "01.01.2024",
      "Иван",
      "ivan@example.com",
      "+7 999 123 45 67",
      "123",
    ]);
    expect(result.type).toBe("mixed");
  });

  test("confidence игнорирует empty значения", () => {
    const result = detectColumnType([1, 2, null, 3, undefined, ""]);
    // 3 из 3 не-empty — integer, confidence = 1
    expect(result.type).toBe("integer");
    expect(result.confidence).toBe(1);
  });

  test("доминирующий phone при большинстве телефонов", () => {
    const result = detectColumnType([
      "+7 999 123 45 67",
      "+7 495 111 22 33",
      "+7 911 555 66 77",
      "не телефон",
      "+7 921 999 88 77",
    ]);
    expect(result.type).toBe("phone");
  });
});
