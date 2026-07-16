module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'node',
  roots: ['<rootDir>/tests'],
  transform: {
    '^.+\\.tsx?$': ['ts-jest', {
      tsconfig: 'tsconfig.test.json',
    }],
  },
  testMatch: ['**/*.test.ts'],
  moduleFileExtensions: ['ts', 'tsx', 'js', 'jsx', 'json'],
  // NOTE: office-addin-mock тянется к office-addin-manifest (@microsoft/app-manifest),
  // который не компилируется в CJS без transformIgnorePatterns. Для unit-тестов
  // tools мы используем тонкий ручной mock Office.js (см. tests/smoke-office-mock.test.ts),
  // поэтому transformIgnorePatterns не нужен. Когда в Фазе 1+ понадобится
  // office-addin-mock для integration-тестов, расскомментировать:
  // transformIgnorePatterns: [
  //   '/node_modules/(?!office-addin-mock|@microsoft|strip-bom|jszip)',
  // ],
};
