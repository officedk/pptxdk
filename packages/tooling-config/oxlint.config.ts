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
                group: ["@react-pptx/validation", "@react-pptx/validation/*"],
                message:
                  "File validation is for testing only; consume actual build artifacts in tests.",
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
                group: ["@react-pptx/validation", "@react-pptx/validation/*"],
                message:
                  "File validation is for testing only; consume actual build artifacts in tests.",
              },
              {
                group: [
                  "@react-pptx/preview",
                  "@react-pptx/preview/*",
                  "@react-pptx/react",
                  "@react-pptx/react/*",
                ],
                message:
                  "Capabilities production source must only depend on core/emit stage contracts; preview and react are test-only devDependencies.",
              },
            ],
          },
        ],
      },
    },
  ],
});
