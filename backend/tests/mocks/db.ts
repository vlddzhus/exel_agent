export function createMockDb() {
  const thenable = { execute: jest.fn(() => Promise.resolve()) };

  const whereBuilder = {
    or: jest.fn(() => ({
      executeTakeFirst: jest.fn(() => Promise.resolve(null)),
    })),
  };

  const selectAllResult = {
    executeTakeFirst: jest.fn(() => Promise.resolve(null)),
    execute: jest.fn(() => Promise.resolve()),
    where: jest.fn((...args: unknown[]) => {
      if (typeof args[0] === "function") {
        (args[0] as (eb: unknown) => unknown)(whereBuilder);
        return {
          executeTakeFirst: jest.fn(() => Promise.resolve(null)),
        };
      }
      return {
        executeTakeFirst: jest.fn(() => Promise.resolve(null)),
        or: jest.fn(() => ({
          executeTakeFirst: jest.fn(() => Promise.resolve(null)),
        })),
      };
    }),
  };

  return {
    insertInto: jest.fn(() => ({
      values: jest.fn(() => thenable),
    })),
    selectFrom: jest.fn(() => ({
      selectAll: jest.fn(() => selectAllResult),
    })),
    updateTable: jest.fn(() => ({
      set: jest.fn(() => ({
        where: jest.fn(() => thenable),
      })),
    })),
    destroy: jest.fn(),
  };
}
