import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    projects: ["packages/*", "apps/*"],
    coverage: {
      provider: "v8",
      include: ["packages/core/src/**/*.ts"],
      exclude: ["packages/core/src/**/*.test.ts", "packages/core/src/__tests__/**"],
      thresholds: {
        lines: 83,
        functions: 78,
        branches: 75,
        statements: 82,
      },
    },
  },
});
