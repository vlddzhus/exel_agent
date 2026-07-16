/**
 * Тесты для registry.ts — единый реестр инструментов.
 */
import {
  toolRegistry,
  defineTool,
  type ToolDefinition,
  type ToolResult,
} from '../../src/taskpane/tools/registry';

describe('defineTool helper', () => {
  test('возвращает определение', () => {
    const def = {
      name: 'testTool',
      description: 'Test',
      parameters: { type: 'object', properties: {} },
      riskLevel: 'safe' as const,
      requiresUndo: false,
      estimateCells: () => 0,
      execute: async () => ({ ok: true, summary: 'ok' }),
    };
    expect(defineTool(def)).toBe(def);
  });
});

describe('ToolRegistry registerDefinition', () => {
  test('register + getDefinition', () => {
    const def = {
      name: 'uniqueTestTool1',
      description: 'test',
      parameters: { type: 'object' },
      riskLevel: 'moderate' as const,
      requiresUndo: true,
      estimateCells: (args: Record<string, unknown>) => ((args.count as number) || 0),
      execute: async () => ({ ok: true, summary: 'ok' }),
    };
    toolRegistry.registerDefinition(def);
    expect(toolRegistry.getDefinition('uniqueTestTool1')).toBe(def);
  });

  test('бросает при повторной регистрации', () => {
    const name = 'duplicateName1';
    const make = () => ({
      name, description: 'test', parameters: { type: 'object' },
      riskLevel: 'safe' as const, requiresUndo: false, estimateCells: () => 0,
      execute: async () => ({ ok: true, summary: 'ok' }),
    });
    toolRegistry.registerDefinition(make());
    expect(() => toolRegistry.registerDefinition(make())).toThrow(/already registered/);
  });

  test('riskLevel доступен', () => {
    toolRegistry.registerDefinition({
      name: 'riskTestTool1', description: 'test', parameters: { type: 'object' },
      riskLevel: 'dangerous' as const, requiresUndo: true, estimateCells: () => 0,
      execute: async () => ({ ok: true, summary: 'ok' }),
    });
    expect(toolRegistry.riskLevel('riskTestTool1')).toBe('dangerous');
    expect(toolRegistry.requiresUndo('riskTestTool1')).toBe(true);
    expect(toolRegistry.requiresConfirmation('riskTestTool1')).toBe(true);
  });

  test('estimateCells', () => {
    toolRegistry.registerDefinition({
      name: 'estimateTestTool1', description: 'test', parameters: { type: 'object' },
      riskLevel: 'safe' as const, requiresUndo: false,
      estimateCells: (args: Record<string, unknown>) => ((args.rows as number) || 0) * ((args.cols as number) || 0),
      execute: async () => ({ ok: true, summary: 'ok' }),
    });
    expect(toolRegistry.estimateCells('estimateTestTool1', { rows: 10, cols: 5 })).toBe(50);
  });

  test('execute возвращает JSON ToolResult', async () => {
    toolRegistry.registerDefinition({
      name: 'execTestTool1', description: 'test', parameters: { type: 'object' },
      riskLevel: 'safe' as const, requiresUndo: false, estimateCells: () => 0,
      execute: async () => ({ ok: true, summary: 'Done', data: { x: 1 } }),
    });
    const result = await toolRegistry.execute('execTestTool1', {});
    const parsed = JSON.parse(result);
    expect(parsed.ok).toBe(true);
    expect(parsed.data).toEqual({ x: 1 });
  });

  test('execute бросает для незарегистрированного', async () => {
    await expect(toolRegistry.execute('doesNotExist999', {})).rejects.toThrow(/Unknown tool/);
  });
});

describe('getSchemas', () => {
  test('формат function-schema', () => {
    toolRegistry.registerDefinition({
      name: 'schemaTestTool1', description: 'Schema test',
      parameters: { type: 'object', properties: { x: { type: 'number' } } },
      riskLevel: 'safe' as const, requiresUndo: false, estimateCells: () => 0,
      execute: async () => ({ ok: true, summary: 'ok' }),
    });
    const schemas = toolRegistry.getSchemas();
    const our = schemas.find((s) => s.function.name === 'schemaTestTool1');
    if (!our) throw new Error('schema not found');
    expect(our.type).toBe('function');
    expect(our.function.description).toBe('Schema test');
  });
});

describe('Legacy API register', () => {
  test('register + getTool', () => {
    toolRegistry.register('legacyTestTool1', 'legacy', { type: 'object' }, async () => 'legacy-result', true);
    expect(toolRegistry.getTool('legacyTestTool1')).toBeDefined();
    expect(toolRegistry.requiresConfirmation('legacyTestTool1')).toBe(true);
  });

  test('legacy moderate', () => {
    toolRegistry.register('legacyModerate1', 'test', { type: 'object' }, async () => 'x');
    expect(toolRegistry.riskLevel('legacyModerate1')).toBe('moderate');
  });

  test('legacy execute возвращает строку', async () => {
    toolRegistry.register('legacyRawResult1', 'test', { type: 'object' }, async () => 'raw-string-result');
    expect(await toolRegistry.execute('legacyRawResult1', {})).toBe('raw-string-result');
  });
});
