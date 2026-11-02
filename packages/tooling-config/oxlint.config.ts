import { defineConfig } from "oxlint";

// Oxc 的 erasing-op 规则尚不拒绝枚举、运行时命名空间和参数属性；
// 可擦除语法以 TypeScript 7 的 erasableSyntaxOnly 检查为准。
export default defineConfig({
  categories: { correctness: "error", suspicious: "warn" },
  // 包图允许校验工具作为开发依赖；生产源码不能反向依赖文件检查实现。
  overrides: [
    // 包入口允许预先建立；业务文件仍接受空文件检查。
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
                message: "文件校验只供测试使用，请在 test 中消费实际构建产物。",
              },
            ],
          },
        ],
      },
    },
  ],
});
