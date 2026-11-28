import shared from "@pptxdk/tooling-config/vitest";
import { defineConfig } from "vitest/config";

export const unitTestInclude = [
  "src/**/test/*-codegen.test.ts",
  "!src/masters/test/brand-master-codegen.test.ts",
  "src/line/test/line-schemas.test.ts",
  "src/table/test/table-schemas.test.ts",
  "src/text/test/text-schemas.test.ts",
  "test/schema-closure.test.tsx",
] as const;

export default defineConfig({
  ...shared,
  test: {
    ...shared.test,
    include: [...unitTestInclude],
    maxWorkers: 4,
    testTimeout: 15_000,
  },
});
