import { spawnSync } from "node:child_process";
import { mkdtemp, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

const packageNames = [
  "@react-pptx/core",
  "@react-pptx/capabilities",
  "@react-pptx/react-pptx",
  "@react-pptx/validation",
] as const;

interface PnpmPackage {
  readonly packageManager?: string;
}

interface DependencyEntry {
  readonly version: string;
}

interface PnpmListEntry {
  readonly dependencies?: Readonly<Record<string, DependencyEntry>>;
  readonly devDependencies?: Readonly<Record<string, DependencyEntry>>;
}

function run(
  command: string,
  args: readonly string[],
  cwd = process.cwd(),
): string {
  const result = spawnSync(command, [...args], {
    cwd,
    encoding: "utf8",
    maxBuffer: 1024 * 1024,
  });
  if (result.error !== undefined || result.status !== 0) {
    const reason =
      result.error?.message ?? result.stderr.trim() ?? "子进程失败";
    throw new Error(`${command} ${args.join(" ")} 失败：${reason}`);
  }
  return result.stdout.trim();
}

function installedVersion(entry: PnpmListEntry, packageName: string): string {
  const version =
    entry.dependencies?.[packageName]?.version ??
    entry.devDependencies?.[packageName]?.version;
  if (version === undefined) {
    throw new Error(`工作区没有为制品消费配置 ${packageName}`);
  }
  return version;
}

const workspacePackage = JSON.parse(
  await readFile("package.json", "utf8"),
) as PnpmPackage;
if (workspacePackage.packageManager === undefined) {
  throw new Error(
    "根 package.json 缺少 packageManager，无法固定临时消费者的 pnpm 版本",
  );
}
const coreDependencies = JSON.parse(
  run("pnpm", [
    "--filter",
    "@react-pptx/core",
    "list",
    "react",
    "@types/react",
    "typescript-7",
    "--depth",
    "0",
    "--json",
  ]),
) as readonly PnpmListEntry[];
const core = coreDependencies[0];
if (core === undefined) {
  throw new Error("pnpm 未返回 @react-pptx/core 的安装依赖");
}

const temporaryDirectory = await mkdtemp(
  join(tmpdir(), "react-pptx-consumer-"),
);
try {
  const dependencies: Record<string, string> = {};
  const overrides: Record<string, string> = {};
  for (const name of packageNames) {
    const filesBeforePack = new Set(await readdir(temporaryDirectory));
    run("pnpm", [
      "--filter",
      name,
      "pack",
      "--pack-destination",
      temporaryDirectory,
    ]);
    const packedFiles = (await readdir(temporaryDirectory)).filter(
      (file) => file.endsWith(".tgz") && !filesBeforePack.has(file),
    );
    if (packedFiles.length !== 1) {
      throw new Error(
        `${name} 应生成一个 tarball，实际生成 ${packedFiles.length} 个`,
      );
    }
    const localPackage = `file:./${packedFiles[0]}`;
    dependencies[name] = localPackage;
    overrides[name] = localPackage;
  }

  await writeFile(
    join(temporaryDirectory, "package.json"),
    JSON.stringify(
      {
        name: "react-pptx-artifact-consumer",
        private: true,
        type: "module",
        packageManager: workspacePackage.packageManager,
        dependencies: {
          ...dependencies,
          react: installedVersion(core, "react"),
        },
        devDependencies: {
          "@types/react": installedVersion(core, "@types/react"),
          typescript: installedVersion(core, "typescript-7"),
        },
      },
      null,
      2,
    ),
  );

  await writeFile(
    join(temporaryDirectory, "tsconfig.json"),
    JSON.stringify(
      {
        compilerOptions: {
          target: "ES2022",
          module: "NodeNext",
          moduleResolution: "NodeNext",
          strict: true,
          skipLibCheck: true,
          noEmit: true,
        },
        include: ["consumer-types.ts"],
      },
      null,
      2,
    ),
  );
  await writeFile(
    join(temporaryDirectory, "pnpm-workspace.yaml"),
    [
      "packages:",
      '  - "."',
      "overrides:",
      ...Object.entries(overrides).map(
        ([name, localPackage]) =>
          `  ${JSON.stringify(name)}: ${JSON.stringify(localPackage)}`,
      ),
      "",
    ].join("\n"),
  );
  await writeFile(
    join(temporaryDirectory, "consumer-types.ts"),
    [
      'import { createElement } from "react";',
      'import { Deck, font, P, Shape, Slide, Span, Text, build, type BuildResult, type ShapeProps, type TextProps } from "@react-pptx/react-pptx";',
      'const shape: ShapeProps = { testId: "typed", x: 0, y: 0, width: 40, height: 30 };',
      'const text: TextProps = { testId: "typed-text", x: 0, y: 36, width: 100, height: 36, style: { fontFamily: font("Consumer Sans"), lineHeight: "22pt" } };',
      'const result: Promise<BuildResult> = build(createElement(Deck, {}, createElement(Slide, { slideId: "typed" }, createElement(Shape, shape), createElement(Text, text, createElement(P, { style: { list: { type: "bullet" } } }, createElement(Span, { style: { fontWeight: "bold" } }, "Installed 文本"))))));',
      "void result;",
    ].join("\n"),
  );
  await writeFile(
    join(temporaryDirectory, "consumer-runtime.ts"),
    [
      'import { createElement } from "react";',
      'import { Deck, font, P, Shape, Slide, Span, Text, build } from "@react-pptx/react-pptx";',
      'import { validatePptx } from "@react-pptx/validation";',
      'import { writeFile } from "node:fs/promises";',
      'const result = await build(createElement(Deck, { width: 120, height: 96 }, createElement(Slide, { slideId: "installed" }, createElement(Shape, { testId: "packed-shape", geometry: "hexagon", x: 10, y: 12, width: 45, height: 30, fill: "#123456" }), createElement(Text, { testId: "packed-text", x: 10, y: 48, width: 90, height: 36, style: { fontFamily: font("Consumer Sans"), lineHeight: "22pt" } }, createElement(P, { style: { list: { type: "bullet" } } }, createElement(Span, { style: { fontWeight: "bold" } }, "Installed 文本"))))));',
      "const file = await validatePptx(result.pptx);",
      'if (file.errors.length !== 0 || file.slides[0]?.objects[0]?.name !== "packed-shape" || file.slides[0]?.objects[1]?.text?.paragraphs[0]?.text !== "Installed 文本" || file.slides[0]?.objects[1]?.text?.paragraphs[0]?.list?.type !== "bullet") throw new Error(JSON.stringify(file.errors));',
      'await writeFile("artifact.pptx", result.pptx);',
      "process.stdout.write(JSON.stringify({ standard: file.profile.standard, slideId: file.slides[0]?.id, object: file.slides[0]?.objects[0]?.name, text: file.slides[0]?.objects[1]?.text?.text, bytes: result.pptx.byteLength }));",
    ].join("\n"),
  );

  run("pnpm", ["install", "--offline", "--ignore-scripts"], temporaryDirectory);
  run(
    "pnpm",
    ["exec", "tsc", "--project", "tsconfig.json"],
    temporaryDirectory,
  );
  const summaryText = run(
    process.execPath,
    ["consumer-runtime.ts"],
    temporaryDirectory,
  );
  const summary = JSON.parse(summaryText) as {
    readonly standard?: string;
    readonly slideId?: string;
    readonly object?: string;
    readonly text?: string;
    readonly bytes?: number;
  };
  if (
    summary.standard !== "ECMA-376 5th edition Transitional" ||
    summary.slideId !== "installed" ||
    summary.object !== "packed-shape" ||
    summary.text !== "Installed 文本" ||
    summary.bytes === undefined ||
    summary.bytes <= 0
  ) {
    throw new Error(`安装制品返回了意外结果：${summaryText}`);
  }
  process.stdout.write(`正式制品消费通过：${summaryText}\n`);
} finally {
  await rm(temporaryDirectory, { recursive: true, force: true });
}
