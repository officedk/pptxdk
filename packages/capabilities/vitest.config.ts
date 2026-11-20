import shared from "@react-pptx/tooling-config/vitest";
import { defineConfig } from "vitest/config";

export default defineConfig({
  ...shared,
  test: {
    ...shared.test,
    // Limit worker fan-out and allow cold WASM validator startup under parallel workspace gates.
    maxWorkers: 8,
    testTimeout: 15_000,
  },
});
