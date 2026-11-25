import shared from "@pptxdk/tooling-config/vitest";
import { defineConfig } from "vitest/config";

export default defineConfig({
  ...shared,
  test: {
    ...shared.test,
    include: ["src/**/test/**/*.test.{ts,tsx}", "test/**/*.test.{ts,tsx}"],
    // Limit worker fan-out and allow cold WASM validator startup under parallel workspace gates.
    maxWorkers: 4,
    testTimeout: 15_000,
  },
});
