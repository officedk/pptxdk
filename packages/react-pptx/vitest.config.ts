import shared from "@pptxdk/tooling-config/vitest";
import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    projects: [
      { ...shared, test: { ...shared.test, name: "development" } },
      // React adaptation reads private reconciler state; ADR-0018 requires real integration acceptance against the production build users run.
      {
        ...shared,
        test: {
          ...shared.test,
          name: "production",
          include: ["test/async-build.test.tsx"],
          env: { NODE_ENV: "production" },
        },
      },
    ],
  },
});
