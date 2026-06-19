import type { Config } from 'jest';

const config: Config = {
  testEnvironment: 'node',
  preset: 'ts-jest',
  moduleNameMapper: {
    '^@/(.*)$': '<rootDir>/$1',
  },
  testMatch: ['**/__tests__/**/*.test.ts', '**/__tests__/**/*.test.tsx'],
  transform: {
    '^.+\\.tsx?$': ['ts-jest', {
      tsconfig: { module: 'commonjs' },
    }],
    // Transform ESM packages (otplib uses @scure/base which is pure ESM)
    '^.+\\.js$': ['ts-jest', {
      tsconfig: { module: 'commonjs' },
      diagnostics: false,
    }],
  },
  // Allow jest to transform these ESM-only node_modules
  transformIgnorePatterns: [
    'node_modules/(?!(@scure|@noble|otplib)/)',
  ],
  collectCoverageFrom: ['lib/**/*.ts', 'app/api/**/*.ts'],
};

export default config;
