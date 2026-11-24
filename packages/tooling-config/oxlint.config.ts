import { defineConfig } from "oxlint";

// Oxc's erasing-op rule does not reject enums, runtime namespaces, or parameter properties yet;
// erasable syntax relies on TypeScript 7's erasableSyntaxOnly check.
export default defineConfig({
  categories: { correctness: "error", suspicious: "warn" },
  // Package graph allows validation tools as devDependencies; production source must not depend on file inspection implementations.
  overrides: [
    // Package entrypoints may be pre-created; business files still undergo empty file checks.
    { files: ["**/src/index.ts"], rules: { "unicorn/no-empty-file": "off" } },
    {
      files: ["**/src/**"],
      rules: {
        "no-restricted-imports": [
          "error",
          {
            patterns: [
              {
                group: ["@pptxdk/validation", "@pptxdk/validation/*"],
                message:
                  "File validation is for testing only; consume actual build artifacts in tests.",
              },
            ],
          },
        ],
      },
    },
    {
      files: ["packages/codegen/src/**", "**/packages/codegen/src/**"],
      rules: {
        "no-restricted-imports": [
          "error",
          {
            patterns: [
              {
                group: [
                  "@pptxdk/capabilities",
                  "@pptxdk/capabilities/*",
                  "@pptxdk/emit",
                  "@pptxdk/emit/*",
                  "@pptxdk/import",
                  "@pptxdk/import/*",
                  "@pptxdk/preview",
                  "@pptxdk/preview/*",
                  "@pptxdk/react",
                  "@pptxdk/react/*",
                  "@pptxdk/validation",
                  "@pptxdk/validation/*",
                  "pptxdk",
                  "pptxdk/*",
                  "react",
                  "react/*",
                ],
                message:
                  "Codegen stage source must remain a pure stage package depending only on @pptxdk/core and zod.",
              },
            ],
          },
        ],
      },
    },
    {
      files: [
        "packages/capabilities/src/**",
        "**/packages/capabilities/src/**",
      ],
      rules: {
        "no-restricted-imports": [
          "error",
          {
            patterns: [
              {
                group: ["@pptxdk/validation", "@pptxdk/validation/*"],
                message:
                  "File validation is for testing only; consume actual build artifacts in tests.",
              },
              {
                group: [
                  "@pptxdk/preview",
                  "@pptxdk/preview/*",
                  "@pptxdk/react",
                  "@pptxdk/react/*",
                  "pptxdk",
                  "pptxdk/*",
                ],
                message:
                  "Capabilities production source must only depend on core/emit/import/codegen stage contracts; preview, react, and sdk are excluded from capabilities runtime.",
              },
            ],
          },
        ],
      },
    },
    {
      files: [
        "packages/capabilities/src/**/test/**",
        "**/packages/capabilities/src/**/test/**",
      ],
      rules: {
        "no-restricted-imports": "off",
      },
    },
  ],
});
