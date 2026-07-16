jest.mock("../../src/utils/provider-factory", () => ({
  getProviderEntries: jest.fn(),
}));

import { getProviderEntries } from "../../src/utils/provider-factory";
import {
  pickModelName,
  getFilteredProviderChain,
  TaskType,
} from "../../src/utils/model-router";

const mockGetProviderEntries = getProviderEntries as jest.Mock;

describe("pickModelName", () => {
  it("returns claude-haiku-4-5 for free/simple", () => {
    expect(pickModelName("free", TaskType.SIMPLE)).toBe("claude-haiku-4-5");
  });

  it("returns claude-haiku-4-5 for free/complex", () => {
    expect(pickModelName("free", TaskType.COMPLEX)).toBe("claude-haiku-4-5");
  });

  it("returns claude-haiku-4-5 for pro/simple", () => {
    expect(pickModelName("pro", TaskType.SIMPLE)).toBe("claude-haiku-4-5");
  });

  it("returns gpt-5.2 for pro/complex", () => {
    expect(pickModelName("pro", TaskType.COMPLEX)).toBe("gpt-5.2");
  });

  it("returns claude-haiku-4-5 for team/simple", () => {
    expect(pickModelName("team", TaskType.SIMPLE)).toBe("claude-haiku-4-5");
  });

  it("returns gpt-5.2 for team/complex", () => {
    expect(pickModelName("team", TaskType.COMPLEX)).toBe("gpt-5.2");
  });

  it("falls back to free.simple for unknown tier", () => {
    expect(pickModelName("bogus" as any, TaskType.SIMPLE)).toBe(
      "claude-haiku-4-5",
    );
  });
});

describe("getFilteredProviderChain", () => {
  beforeEach(() => {
    mockGetProviderEntries.mockReset();
  });

  it("returns matching providers when model matches", () => {
    const entries = [
      {
        name: "openmodel",
        priority: 0,
        defaultModel: "claude-haiku-4-5",
        model: {} as any,
        available: true,
      },
      {
        name: "openai",
        priority: 5,
        defaultModel: "gpt-5.2",
        model: {} as any,
        available: true,
      },
    ];
    mockGetProviderEntries.mockReturnValue(entries);

    const result = getFilteredProviderChain("pro", TaskType.COMPLEX);
    expect(result).toHaveLength(1);
    expect(result[0].name).toBe("openai");
  });

  it("returns empty when no providers match", () => {
    const entries = [
      {
        name: "openai",
        priority: 5,
        defaultModel: "gpt-5.2",
        model: {} as any,
        available: true,
      },
    ];
    mockGetProviderEntries.mockReturnValue(entries);

    const result = getFilteredProviderChain("free", TaskType.SIMPLE);
    expect(result).toHaveLength(1);
    expect(result[0].name).toBe("openai");
  });

  it("sorts by priority when no match", () => {
    const entries = [
      {
        name: "anthropic",
        priority: 20,
        defaultModel: "claude-sonnet-4-20250514",
        model: {} as any,
        available: true,
      },
      {
        name: "groq",
        priority: 10,
        defaultModel: "llama-3.3-70b-versatile",
        model: {} as any,
        available: true,
      },
    ];
    mockGetProviderEntries.mockReturnValue(entries);

    const result = getFilteredProviderChain("free", TaskType.SIMPLE);
    expect(result).toHaveLength(2);
    expect(result[0].name).toBe("groq");
    expect(result[1].name).toBe("anthropic");
  });

  it("returns empty array when no providers exist", () => {
    mockGetProviderEntries.mockReturnValue([]);
    const result = getFilteredProviderChain("free", TaskType.SIMPLE);
    expect(result).toEqual([]);
  });

  it("ignores unknown priority names", () => {
    const entries = [
      {
        name: "custom-vendor",
        priority: 99,
        defaultModel: "custom-model",
        model: {} as any,
        available: true,
      },
    ];
    mockGetProviderEntries.mockReturnValue(entries);

    const result = getFilteredProviderChain("free", TaskType.SIMPLE);
    expect(result).toHaveLength(1);
    expect(result[0].name).toBe("custom-vendor");
  });
});
