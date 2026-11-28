import shared from "@pptxdk/tooling-config/vitest";
import { defineConfig } from "vitest/config";

import { unitTestInclude } from "./vitest.unit.config.ts";

export default defineConfig({
  ...shared,
  test: {
    ...shared.test,
    include: ["src/**/test/**/*.test.{ts,tsx}", "test/**/*.test.{ts,tsx}"],
    exclude: [
      "**/node_modules/**",
      ...unitTestInclude.filter((pattern) => !pattern.startsWith("!")),
    ],
    // Limit worker fan-out and allow cold WASM validator startup under parallel workspace gates.
    maxWorkers: 4,
    testTimeout: 15_000,
  },
});
