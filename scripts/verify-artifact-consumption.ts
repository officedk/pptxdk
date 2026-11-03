import { spawnSync } from "node:child_process";
import {
  mkdir,
  mkdtemp,
  readdir,
  readFile,
  rm,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

const packageNames = [
  "@react-pptx/core",
  "@react-pptx/capabilities",
  "@react-pptx/react-pptx",
  "@react-pptx/validation",
  "@react-pptx/cli",
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
      result.error?.message ??
      [result.stderr.trim(), result.stdout.trim()].find(
        (output) => output !== "",
      ) ??
      "子进程失败";
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
    "react-reconciler",
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
if (
  installedVersion(core, "react") !== "19.3.1" ||
  installedVersion(core, "react-reconciler") !== "0.34.0"
) {
  throw new Error("制品消费验收要求 React 19.3.1 与 react-reconciler 0.34.0");
}
const reactPptxDependencies = JSON.parse(
  run("pnpm", [
    "--filter",
    "@react-pptx/react-pptx",
    "list",
    "fflate",
    "esbuild",
    "--depth",
    "0",
    "--json",
  ]),
) as readonly PnpmListEntry[];
const reactPptx = reactPptxDependencies[0];
if (reactPptx === undefined) {
  throw new Error("pnpm 未返回 @react-pptx/react-pptx 的安装依赖");
}

const temporaryDirectory = await mkdtemp(
  join(tmpdir(), "react-pptx-consumer-"),
);
try {
  await writeFile(
    join(temporaryDirectory, "jost-latin-400-normal.woff2"),
    await readFile(
      new URL(
        import.meta
          .resolve("@fontsource/jost/files/jost-latin-400-normal.woff2"),
      ),
    ),
  );
  await writeFile(
    join(temporaryDirectory, "Jost-400-Book.ttf"),
    await readFile("scripts/assets/fonts/Jost-400-Book.ttf"),
  );
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
          fflate: installedVersion(reactPptx, "fflate"),
          react: installedVersion(core, "react"),
          esbuild: installedVersion(reactPptx, "esbuild"),
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
      'import { createElement, Suspense, use, useEffect, useState } from "react";',
      'import { after, Cell, Chart, Deck, fade, font, Group, Image, MotionScope, P, Row, Shape, Slide, Span, spin, Step, Table, Text, useMotionScope, build, transitions, type BuildOptions, type BuildResult, type CellProps, type ChartAxis, type ChartData, type ChartDataLabels, type ChartKind, type ChartLine, type ChartPoint, type ChartProps as DefaultChartProps, type ChartSeries, type ColumnOrBarChartProps, type GroupProps, type ImageProps, type LineOrAreaChartProps, type MotionCompletionState, type MotionGranularity, type MotionMoment, type MotionObservation, type MotionProps, type MotionScopeHandle, type MotionScopeMode, type MotionScopeProps, type NativeTableObservation, type RowProps, type ShapeProps, type StepProps, type TableColumn, type TableProps, type TableRowStyle, type TableCellStyle, type TableSpacing, type TableStyle, type TableTextStyle, type TableVerticalAlign, type TextProps, type Transition } from "@react-pptx/react-pptx";',
      'import { build as buildChartOnly, Chart as ChartOnly, Deck as ChartDeck, Slide as ChartSlide, type ChartProps } from "@react-pptx/react-pptx/chart";',
      'import { type ValidationIssue, type ValidationResult } from "@react-pptx/validation";',
      'import { buildFromSource, type BuildFromSourceOptions } from "@react-pptx/react-pptx/node";',
      'const shape: ShapeProps = { testId: "typed", x: 0, y: 0, width: 40, height: 30 };',
      "const image: ImageProps = { src: new Uint8Array([0x89, 0x50, 0x4e, 0x47]), width: 40, height: 30 };",
      'const group: GroupProps = { name: "typed-group" };',
      'const text: TextProps = { testId: "typed-text", x: 0, y: 36, width: 100, height: 36, style: { fontFamily: font("Consumer Sans"), lineHeight: "22pt" } };',
      "const motion: MotionProps = { enter: fade({ duration: 800 }) };",
      'const textGranularity: MotionGranularity = "letter";',
      'const granularText: TextProps = { testId: "typed-granular-text", x: 0, y: 0, width: 100, height: 32, enter: fade({ by: textGranularity, intervalMs: 80 }) };',
      "declare const scopeHandle: MotionScopeHandle;",
      'const scopeMode: MotionScopeMode = "sequence";',
      "const scopeMoment: MotionMoment = scopeHandle.end;",
      'const scopeCompletion: MotionCompletionState = "no-natural-completion";',
      'const scopeProps: MotionScopeProps = { scope: scopeHandle, mode: scopeMode, publishEnd: scopeHandle.cue("done") };',
      "const step: StepProps = { n: 1 };",
      'const contentSizedText: TextProps = { testId: "typed-content-sized-text", x: 0, y: 72, width: 100, height: "auto", style: { fontFamily: font("Consumer Sans") } };',
      'const transition: Transition = transitions.morph({ by: "word", duration: 1200, fallback: transitions.fade() });',
      'const defaultChart: DefaultChartProps = { kind: "line", x: 0, y: 0, width: 100, height: 60, data: { categories: ["Q1"], series: [{ name: "Revenue", values: [1] }] } };',
      'const chartOnly: ChartProps = { kind: "column", x: 0, y: 0, width: 100, height: 60, data: { categories: ["Q1"], series: [{ name: "Revenue", values: [1] }] } };',
      'const granularChart: ChartProps = { ...chartOnly, enter: fade({ by: "category" }) };',
      'const tableColumns: readonly TableColumn[] = [28, "2fr", "auto"];',
      'const table: TableProps = { testId: "typed-table", x: 8, y: 8, width: 104, columns: tableColumns };',
      'const row: RowProps = { testId: "typed-row", style: { height: 24 } };',
      'const cell: CellProps = { testId: "typed-cell", rowSpan: 2, style: { padding: [2, 4], verticalAlign: "middle" } };',
      'const tableObservation: NativeTableObservation = { columnWidths: [28, 52, 24], rowHeights: [24, 24], cells: [["North", "Q1", "100"], ["", "Q2", "120"]], merges: [{ row: 0, column: 0, rowSpan: 2, colSpan: 1 }] };',
      'const validationIssue: ValidationIssue = { part: "/ppt/slides/slide1.xml", code: "schema-error", message: "invalid", xpath: "/p:sld[1]" };',
      'const validatorProfile: ValidationResult["profile"]["schemaValidator"] = { package: "@ooxml-tools/validate", packageVersion: "0.4.6", sdkVersion: "3.0.2", officeVersion: "Microsoft365" };',
      "function AsyncText({ promise }: { promise: Promise<string> }) { return createElement(Text, text, use(promise)); }",
      "function AsyncSalesTable({ source }: { source: string }) { const [promise, setPromise] = useState<Promise<string> | null>(null); useEffect(() => { setPromise(fetch(source).then(response => response.text())); }, [source]); return promise === null ? null : createElement(Suspense, { fallback: null }, createElement(AsyncText, { promise })); }",
      'const asyncTable = createElement(AsyncSalesTable, { source: "https://example.invalid/data" });',
      "const signalOptions: BuildOptions = { signal: new AbortController().signal };",
      'const fontEmbeddingOptions: BuildOptions = { embedFonts: "full" };',
      "const sourceOptions: BuildFromSourceOptions = { signal: new AbortController().signal };",
      'const sourceBuild = buildFromSource("./consumer-deck.tsx", sourceOptions);',
      'const result: Promise<BuildResult> = build(createElement(Deck, {}, createElement(Slide, { slideId: "typed", transition }, createElement(Group, group, createElement(Shape, { ...shape, ...motion }), createElement(Text, text, createElement(P, { style: { list: { type: "bullet" } } }, createElement(Span, { style: { fontWeight: "bold" } }, "Installed 文本"))), createElement(Image, image), createElement(Step, { ...step }, createElement(Shape, { ...shape, testId: "step-shape", enter: fade(), emphasis: spin() }))))));',
      'const defaultChartResult: Promise<BuildResult> = build(createElement(Deck, {}, createElement(Slide, { slideId: "typed-chart" }, createElement(Chart, defaultChart))));',
      'const chartOnlyResult: Promise<BuildResult> = buildChartOnly(createElement(ChartDeck, {}, createElement(ChartSlide, { slideId: "typed-chart-only" }, createElement(ChartOnly, chartOnly))));',
      'const tableElement = createElement(Table, table, createElement(Row, row, createElement(Cell, cell, "North")));',
      "void [result, defaultChartResult, chartOnlyResult, contentSizedText, granularText, granularChart, validationIssue, validatorProfile, asyncTable, signalOptions, fontEmbeddingOptions, sourceBuild, tableElement, tableObservation, scopeMoment, scopeCompletion, scopeProps];",
    ].join("\n"),
  );
  await writeFile(
    join(temporaryDirectory, "consumer-deck-helper.ts"),
    'export const sourceTitle = "安装 CLI 源码标题";\n',
  );
  await writeFile(
    join(temporaryDirectory, "consumer-deck.tsx"),
    [
      'import { Deck, Shape, Slide } from "@react-pptx/react-pptx";',
      'import { sourceTitle } from "./consumer-deck-helper.ts";',
      "export default function ConsumerDeck() {",
      "  return (",
      '    <Deck width={120} height={96}><Slide slideId="source-cli">',
      '      <Shape testId="source-cli" name={sourceTitle} x={8} y={8} width={80} height={32} text={sourceTitle} />',
      "    </Slide></Deck>",
      "  );",
      "}",
    ].join("\n"),
  );
  await writeFile(
    join(temporaryDirectory, "consumer-runtime.ts"),
    [
      'import { createElement, Suspense, use, useEffect, useState } from "react";',
      'import { after, BuildError, Cell, Chart, Deck, fade, font, Group, Image, MotionScope, P, Row, Shape, Slide, Span, spin, Step, Table, Text, useMotionScope, build, targets, transitions, wipe } from "@react-pptx/react-pptx";',
      'import { build as buildChartOnly, Chart as ChartOnly, Deck as ChartDeck, Slide as ChartSlide } from "@react-pptx/react-pptx/chart";',
      'import { validatePptx } from "@react-pptx/validation";',
      'import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";',
      'import { getEventListeners } from "node:events";',
      'import { readFile, writeFile } from "node:fs/promises";',
      'import { buildFromSource } from "@react-pptx/react-pptx/node";',
      'const jostMetrics = new Uint8Array(await readFile(new URL("./jost-latin-400-normal.woff2", import.meta.url)));',
      'const embeddedFontBytes = new Uint8Array(await readFile(new URL("./Jost-400-Book.ttf", import.meta.url)));',
      'const installedCliBytes = await readFile("cli-output/installed.pptx");',
      "const installedCliFile = await validatePptx(new Uint8Array(installedCliBytes));",
      'const sourceBuildResult = await buildFromSource("./consumer-deck.tsx");',
      "const installedCliEquivalent = Buffer.from(installedCliBytes).equals(Buffer.from(sourceBuildResult.pptx));",
      'if (installedCliFile.errors.length !== 0 || installedCliFile.slides[0]?.id !== "source-cli" || installedCliFile.slides[0]?.objects[0]?.text?.text !== "安装 CLI 源码标题" || !installedCliEquivalent || sourceBuildResult.report.slides[0]?.id !== "source-cli") throw new Error(JSON.stringify({ errors: installedCliFile.errors, slide: installedCliFile.slides[0], installedCliEquivalent }));',
      'const transition = transitions.morph({ by: "word", duration: 1200, fallback: transitions.fade() });',
      'function InstalledScopeDeck() { const source = useMotionScope(); const follower = useMotionScope(); return createElement(Deck, { width: 120, height: 96 }, createElement(Slide, { slideId: "installed-motion-scope" }, createElement(MotionScope, { scope: source, mode: "sequence" }, createElement(Shape, { testId: "scope-first", x: 4, y: 4, width: 20, height: 20, enter: fade({ duration: 200 }) }), createElement(Shape, { testId: "scope-second", x: 28, y: 4, width: 20, height: 20, enter: fade({ duration: 100 }) })), createElement(MotionScope, { scope: follower, startAt: after(source.end, 25) }, createElement(Shape, { testId: "scope-follower", x: 52, y: 4, width: 20, height: 20, enter: fade({ duration: 50 }) })))); }',
      "const scopeResult = await build(createElement(InstalledScopeDeck));",
      "const scopeFile = await validatePptx(scopeResult.pptx);",
      "const scopeSlide = scopeResult.report.slides[0];",
      "const scopeMotion = scopeSlide?.capabilities?.motion;",
      "const motionScopeStarts = scopeMotion?.steps[0]?.effects.map(({ target, startMs, endMs }) => [target.testId, startMs, endMs]);",
      "const scopeFileStarts = scopeFile.slides[0]?.animations?.[0]?.effects.map(({ targetId, delayMs }) => [scopeSlide?.objects.find(({ id }) => id === targetId)?.testId, delayMs]);",
      'if (scopeFile.errors.length !== 0 || JSON.stringify(motionScopeStarts) !== JSON.stringify([["scope-first", 0, 200], ["scope-second", 200, 300], ["scope-follower", 325, 375]]) || JSON.stringify(scopeFileStarts) !== JSON.stringify([["scope-first", 0], ["scope-second", 200], ["scope-follower", 325]])) throw new Error(JSON.stringify({ errors: scopeFile.errors, motionScopeStarts, scopeFileStarts }));',
      'function InstalledGranularMotionDeck() { const source = useMotionScope(); const follower = useMotionScope(); const data = { categories: ["Q1", "Q2"], series: [{ name: "North", values: [10, 20] }, { name: "South", values: [12, 18] }] }; return createElement(Deck, { width: 240, height: 160 }, createElement(Slide, { slideId: "installed-granular-motion" }, createElement(MotionScope, { scope: source }, createElement(Text, { testId: "installed-letters", x: 4, y: 4, width: 120, height: 32, enter: fade({ by: "letter", duration: 400, intervalMs: 200 }) }, "LATE EARLY")), createElement(MotionScope, { scope: follower, startAt: source.end }, createElement(Shape, { testId: "installed-after-letters", x: 4, y: 48, width: 40, height: 24, enter: fade({ duration: 100 }) })), createElement(Chart, { testId: "installed-series-chart", kind: "column", x: 130, y: 4, width: 100, height: 100, data, enter: fade({ by: "series", duration: 150 }) }))); }',
      "const granularResult = await build(createElement(InstalledGranularMotionDeck));",
      "const granularFile = await validatePptx(granularResult.pptx);",
      "const granularMotion = granularResult.report.slides[0]?.capabilities?.motion as MotionObservation;",
      "const granularEffects = granularMotion.steps[0]?.effects ?? [];",
      'const letterEffect = granularEffects.find(({ target }) => target.testId === "installed-letters");',
      'const afterLettersEffect = granularEffects.find(({ target }) => target.testId === "installed-after-letters");',
      'const seriesEffects = granularEffects.filter(({ target }) => target.testId === "installed-series-chart");',
      "const granularFileEffects = granularFile.slides[0]?.animations?.[0]?.effects ?? [];",
      "const fileLetterEffect = granularFileEffects.find(({ targetId }) => targetId === letterEffect?.target.id);",
      "const fileAfterLettersEffect = granularFileEffects.find(({ targetId }) => targetId === afterLettersEffect?.target.id);",
      'if (granularFile.errors.length !== 0 || letterEffect?.durationMs !== 400 || letterEffect.intervalMs !== 200 || letterEffect.completion !== "unknown" || letterEffect.endMs !== undefined || afterLettersEffect?.startMs !== undefined || afterLettersEffect?.startAfter?.targetId !== letterEffect.target.id || fileLetterEffect?.by !== "letter" || fileLetterEffect.intervalMs !== 200 || fileLetterEffect.completion !== "unknown" || fileAfterLettersEffect?.startAfter?.targetId !== letterEffect.target.id || fileAfterLettersEffect.startAfter.point !== "end" || seriesEffects.length !== 3 || seriesEffects.some(({ target }) => target.range?.type !== "chart") || granularFileEffects.filter(({ targetId, targetRange }) => targetId === seriesEffects[0]?.target.id && targetRange?.type === "chart").length !== 3) throw new Error(JSON.stringify({ errors: granularFile.errors, letterEffect, afterLettersEffect, seriesEffects, fileLetterEffect, fileAfterLettersEffect, fileEffects: granularFileEffects }));',
      'const chartData = { categories: ["Installed Q1", "Installed Q2"], series: [{ name: "North", values: [10, 20], color: "#336699" }, { name: "South", values: [12, 18], color: "#993366" }] };',
      'const defaultChartResult = await build(createElement(Deck, { width: 120, height: 96 }, createElement(Slide, { slideId: "installed-default-chart" }, createElement(Chart, { testId: "default-chart", kind: "line", x: 4, y: 4, width: 100, height: 80, data: chartData }))));',
      "const defaultChartFile = await validatePptx(defaultChartResult.pptx);",
      'if (defaultChartFile.errors.length !== 0 || defaultChartFile.slides[0]?.objects[0]?.kind !== "chart" || defaultChartFile.slides[0]?.objects[0]?.chart?.types[0] !== "lineChart") throw new Error(JSON.stringify({ errors: defaultChartFile.errors, chart: defaultChartFile.slides[0]?.objects[0] }));',
      'const chartResult = await buildChartOnly(createElement(ChartDeck, { width: 640, height: 480 }, createElement(ChartSlide, { slideId: "installed-charts" }, createElement(ChartOnly, { testId: "installed-column", kind: "column", x: 0, y: 0, width: 300, height: 210, data: chartData, dataLabels: { value: true, category: true } }), createElement(ChartOnly, { testId: "installed-bar", kind: "bar", x: 320, y: 0, width: 300, height: 210, data: chartData, stacking: "stacked" }), createElement(ChartOnly, { testId: "installed-line", kind: "line", x: 0, y: 230, width: 300, height: 210, data: chartData }), createElement(ChartOnly, { testId: "installed-area", kind: "area", x: 320, y: 230, width: 300, height: 210, data: chartData })), createElement(ChartSlide, { slideId: "installed-charts-second" }, createElement(ChartOnly, { testId: "installed-second-page", kind: "column", x: 0, y: 0, width: 200, height: 120, data: chartData }))));',
      "const chartFile = await validatePptx(chartResult.pptx);",
      'const chartObjects = chartFile.slides.flatMap(({ objects }) => objects).filter(({ kind }) => kind === "chart");',
      "const embeddedChartSeries = chartObjects.map(({ chart }) => chart?.embeddedWorkbook?.series);",
      'if (chartFile.errors.length !== 0 || chartObjects.length !== 5 || chartObjects.map(({ chart }) => chart?.types[0]).join(",") !== "barChart,barChart,lineChart,areaChart,barChart" || chartObjects[1]?.chart?.direction !== "bar" || chartObjects[0]?.chart?.direction !== "col" || chartObjects[4]?.chart?.embeddedWorkbook?.part !== "ppt/embeddings/Microsoft_Excel_Worksheet5.xlsx" || embeddedChartSeries.some((series) => JSON.stringify(series) !== JSON.stringify(chartObjects[0]?.chart?.series))) throw new Error(JSON.stringify({ errors: chartFile.errors, chartObjects }));',
      'const imageBytes = new Uint8Array(Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/n1sAAAAASUVORK5CYII=", "base64"));',
      'const result = await build(createElement(Deck, { width: 120, height: 96 }, createElement(Slide, { slideId: "installed", transition }, createElement(Group, { testId: "packed-group", enter: fade({ duration: 800, delay: 100, easing: { decelerate: 0.5 } }) }, createElement(Step, { n: 1 }, createElement(Shape, { testId: "packed-shape", geometry: "hexagon", x: 10, y: 12, width: 45, height: 30, fill: "#123456", text: "Installed shape text", enter: fade({ duration: 250 }), emphasis: spin({ from: 180 }) }), createElement(Image, { testId: "packed-image", src: imageBytes, x: 60, y: 12, width: 40, height: 30, fit: "cover", enter: wipe({ from: "right" }) })), createElement(Text, { testId: "packed-text", x: 10, y: 48, width: 90, height: 36, style: { fontFamily: font("Consumer Sans"), lineHeight: "22pt" }, emphasis: spin({ from: 90 }) }, createElement(P, { style: { list: { type: "bullet" } } }, createElement(Span, { style: { fontWeight: "bold" } }, "Installed 文本"))), createElement(Group, { testId: "nested-group" }, createElement(Shape, { testId: "nested-shape", x: 80, y: 60, width: 12, height: 10, text: "Nested member" }))))));',
      'const file = await validatePptx(result.pptx, { allowedExtensions: ["p14", "p159"] });',
      'const tableResult = await build(createElement(Deck, { width: 120, height: 96 }, createElement(Slide, { slideId: "installed-table" }, createElement(Table, { testId: "packed-table", x: 8, y: 8, width: 104, columns: [28, "2fr", "auto"], style: { fontFamily: font("Jost", [{ weight: 400, data: jostMetrics }]), fontSize: 14 } }, createElement(Row, { style: { height: 24 } }, createElement(Cell, { colSpan: 2 }, "Period"), createElement(Cell, null, "YoY")), createElement(Row, { style: { height: 24 } }, createElement(Cell, { rowSpan: 2 }, "N"), createElement(Cell, null, "Q1"), createElement(Cell, null, "100")), createElement(Row, { style: { height: 24 } }, createElement(Cell, null, "Q2"), createElement(Cell, null, "120"))))));',
      "const tableFile = await validatePptx(tableResult.pptx);",
      "const tableObject = tableFile.slides[0]?.objects[0];",
      'if (tableFile.errors.length !== 0 || tableObject?.kind !== "table" || tableObject.name !== "packed-table" || tableObject.table?.cells[0]?.[0] !== "Period" || tableObject.table.cells[1]?.[0] !== "N" || tableObject.table.cells[2]?.[1] !== "Q2" || JSON.stringify(tableObject.table.merges) !== JSON.stringify([{ row: 0, column: 0, rowSpan: 1, colSpan: 2 }, { row: 1, column: 0, rowSpan: 2, colSpan: 1 }]) || tableObject.table.columnWidthsEmu[0] !== 266700 || JSON.stringify(tableResult.report.slides[0]?.objects[0]?.table?.merges) !== JSON.stringify([{ row: 0, column: 0, rowSpan: 1, colSpan: 2 }, { row: 1, column: 0, rowSpan: 2, colSpan: 1 }]) || tableResult.report.diagnostics.length !== 0) throw new Error(JSON.stringify({ errors: tableFile.errors, tableObject, report: tableResult.report }));',
      'await writeFile("table-artifact.pptx", tableResult.pptx);',
      "const group = file.slides[0]?.objects[0];",
      'const shape = group?.members?.find(({ name }) => name === "packed-shape");',
      'const text = group?.members?.find(({ name }) => name === "packed-text");',
      'const image = group?.members?.find(({ kind }) => kind === "image");',
      'const nestedGroup = group?.members?.find(({ kind }) => kind === "group");',
      'const nestedShape = nestedGroup?.members?.find(({ name }) => name === "nested-shape");',
      "const animations = file.slides[0]?.animations?.map(({ trigger, effects }) => [trigger, effects.map(({ targetId, preset, durationMs, delayMs, from, easing }) => [targetId, preset, durationMs, delayMs, from ?? null, easing ?? null])]);",
      'const expectedAnimations = [["auto", [[text?.id, "spin", 1000, 0, 90, null], [group?.id, "fade", 800, 100, null, { decelerate: 0.5 }]]], ["click", [[shape?.id, "fade", 250, 0, null, null], [shape?.id, "spin", 1000, 0, 180, null], [image?.id, "wipe", 500, 0, "right", null]]]];',
      'if (file.errors.length !== 0 || file.extensions.join(",") !== "p14,p159" || file.slides[0]?.transition?.kind !== "morph" || file.slides[0]?.transition?.by !== "word" || file.slides[0]?.transition?.durationMs !== 1200 || group?.kind !== "group" || group.name !== "packed-group" || !group.boxEmu || group.boxEmu.width <= 0 || shape?.text?.text !== "Installed shape text" || !shape.boxEmu || shape.boxEmu.x <= 0 || text?.text?.paragraphs[0]?.text !== "Installed 文本" || text.text.paragraphs[0]?.list?.type !== "bullet" || image?.name !== "packed-image" || image.image?.contentType !== "image/png" || !image.boxEmu || nestedGroup?.name !== "nested-group" || nestedShape?.text?.text !== "Nested member" || !nestedShape.boxEmu || JSON.stringify(animations) !== JSON.stringify(expectedAnimations)) throw new Error(JSON.stringify({ errors: file.errors, group, shape, text, image, nestedGroup, animations }));',
      "const asyncResource = Promise.withResolvers<string>();",
      "const asyncStarted = Promise.withResolvers<void>();",
      'function AsyncText({ promise }) { const content = use(promise); return createElement(Text, { testId: "installed-async", x: 0, y: 0, width: 90, height: 36 }, content); }',
      "function AsyncSalesTable() { const [promise, setPromise] = useState(null); useEffect(() => { asyncStarted.resolve(); setPromise(asyncResource.promise); }, []); return promise ? createElement(Suspense, { fallback: null }, createElement(AsyncText, { promise })) : null; }",
      'const asyncBuildPromise = build(createElement(Deck, { width: 120, height: 96 }, createElement(Slide, { slideId: "installed-async" }, createElement(AsyncSalesTable))));',
      "let asyncSettled = false; void asyncBuildPromise.then(() => { asyncSettled = true; }, () => { asyncSettled = true; });",
      "await asyncStarted.promise;",
      'if (asyncSettled) throw new Error("installed build returned before its React dependency completed");',
      'asyncResource.resolve("安装制品的最终异步内容");',
      "const asyncResult = await asyncBuildPromise;",
      "const asyncFile = await validatePptx(asyncResult.pptx);",
      'if (asyncFile.errors.length !== 0 || asyncResult.report.slides[0]?.objects[0]?.text !== "安装制品的最终异步内容" || asyncFile.slides[0]?.objects[0]?.text?.text !== "安装制品的最终异步内容") throw new Error(JSON.stringify({ errors: asyncFile.errors, report: asyncResult.report.slides }));',
      "const cancellationController = new AbortController();",
      "const cancellationStarted = Promise.withResolvers<void>();",
      "const neverResource = new Promise<string>(() => {});",
      "let cancellationCleanups = 0;",
      'function CancelledContent() { useEffect(() => { cancellationStarted.resolve(); return () => { cancellationCleanups += 1; }; }, []); return createElement(Suspense, { fallback: createElement(Text, { x: 0, y: 0, width: 90, height: 36 }, "安装制品加载占位") }, createElement(AsyncText, { promise: neverResource })); }',
      'const cancellationReason = new Error("安装制品调用方取消");',
      'const cancellationBuild = build(createElement(Deck, { width: 120, height: 96 }, createElement(Slide, { slideId: "installed-cancelled" }, createElement(CancelledContent))), { signal: cancellationController.signal });',
      "let cancellationObserved = false;",
      'const cancellationRejected = cancellationBuild.then(() => { throw new Error("安装制品取消后仍返回了产物"); }, error => { if (error !== cancellationReason) throw error; cancellationObserved = true; });',
      "await cancellationStarted.promise;",
      'if (getEventListeners(cancellationController.signal, "abort").length !== 1) throw new Error("安装制品没有监听调用方取消");',
      "cancellationController.abort(cancellationReason);",
      "await cancellationRejected;",
      'if (!cancellationObserved || cancellationCleanups !== 1 || getEventListeners(cancellationController.signal, "abort").length !== 0) throw new Error(JSON.stringify({ cancellationObserved, cancellationCleanups, abortListeners: getEventListeners(cancellationController.signal, "abort").length }));',
      "const mutatedParts = unzipSync(result.pptx);",
      'const masterPart = "ppt/slideMasters/slideMaster1.xml";',
      "const master = strFromU8(mutatedParts[masterPart]);",
      'const layoutIdMatch = master.match(/(<p:sldLayoutId\\b[^>]*\\bid=")(\\d+)(")/u);',
      'if (layoutIdMatch === null) throw new Error("generated master layout id was not found");',
      'if (Number(layoutIdMatch[2]) < 2147483648 || Number(layoutIdMatch[2]) > 4294967295) throw new Error("generated master layout id is outside the legal range");',
      "mutatedParts[masterPart] = strToU8(master.replace(layoutIdMatch[0], `${layoutIdMatch[1]}1${layoutIdMatch[3]}`));",
      'const invalidFile = await validatePptx(zipSync(mutatedParts), { allowedExtensions: ["p14", "p159"] });',
      'const schemaIssue = invalidFile.errors.find(({ code, part, xpath }) => code === "Sch_AttributeValueDataTypeDetailed" && part === "/ppt/slideMasters/slideMaster1.xml" && xpath !== undefined);',
      "if (schemaIssue === undefined) throw new Error(JSON.stringify({ errors: invalidFile.errors }));",
      "let forbidden;",
      'try { await build(createElement(Deck, {}, createElement(Slide, { slideId: "forbidden", transition })), { target: targets.baselineOnly }); } catch (error) { forbidden = error; }',
      'if (!(forbidden instanceof BuildError) || forbidden.diagnostics[0]?.code !== "forbidden-mechanism") throw new Error("installed baseline target did not reject the transition mechanism");',
      'const contentSizedResult = await build(createElement(Deck, { width: 120, height: 96 }, createElement(Slide, { slideId: "installed-content-sized" }, createElement(Text, { testId: "packed-content-sized", x: 10, y: 10, width: 70, height: "auto", style: { fontFamily: font("Jost", [{ weight: 400, data: jostMetrics }]), fontSize: 20 } }, "Installed font"))));',
      "const contentSizedFile = await validatePptx(contentSizedResult.pptx);",
      "const contentSizedText = contentSizedFile.slides[0]?.objects[0];",
      'if (contentSizedFile.errors.length !== 0 || contentSizedText?.text?.text !== "Installed font" || contentSizedText.boxEmu?.height !== 550672 || contentSizedText.text.paragraphs[0]?.lineSpacingPoints !== 21.68 || contentSizedText.text.paragraphs[0]?.runs[0]?.fontFamily !== "Jost" || contentSizedResult.report.slides[0]?.objects[0]?.box?.height !== 57.81333333333333) throw new Error(JSON.stringify({ errors: contentSizedFile.errors, text: contentSizedText, report: contentSizedResult.report.slides[0]?.objects[0] }));',
      'const fontEmbeddingResult = await build(createElement(Deck, { width: 120, height: 96 }, createElement(Slide, { slideId: "installed-font-embedding" }, createElement(Text, { testId: "installed-embedded-font", x: 8, y: 8, width: 96, height: 36, style: { fontFamily: font("Jost*", [{ weight: 400, data: embeddedFontBytes }]) } }, "Installed embedded font"))), { embedFonts: "full" });',
      "const fontEmbeddingFile = await validatePptx(fontEmbeddingResult.pptx);",
      "const fontEmbeddingParts = unzipSync(fontEmbeddingResult.pptx);",
      'const fontPresentation = strFromU8(fontEmbeddingParts["ppt/presentation.xml"]);',
      'const fontRelationships = strFromU8(fontEmbeddingParts["ppt/_rels/presentation.xml.rels"]);',
      'const fontRelation = fontRelationships.split("/>").find((relation) => relation.includes("/relationships/font"));',
      'const fontRelationId = fontRelation?.match(/Id="([^"]+)"/u)?.[1];',
      'const fontRelationTarget = fontRelation?.match(/Target="([^"]+)"/u)?.[1];',
      "const embeddedFontPart = fontRelationTarget === undefined ? undefined : fontEmbeddingParts[`ppt/${fontRelationTarget}`];",
      "const embeddedFontPayload = embeddedFontPart?.subarray(embeddedFontPart.byteLength - embeddedFontBytes.byteLength);",
      "const fontEmbeddingPayloadMatches = embeddedFontPayload !== undefined && Buffer.from(embeddedFontPayload).equals(Buffer.from(embeddedFontBytes));",
      'const fontEmbeddingStyleReferenceMatches = fontRelationId !== undefined && fontPresentation.includes(`<p:regular r:id="${fontRelationId}"/>`);',
      'const fontEmbeddingXmlMatches = fontPresentation.includes("embedTrueTypeFonts=\\"1\\"") && fontPresentation.includes("<p:font typeface=\\"Jost*\\"") && fontEmbeddingStyleReferenceMatches;',
      'if (fontEmbeddingFile.errors.length !== 0 || fontEmbeddingFile.slides[0]?.objects[0]?.text?.text !== "Installed embedded font" || !fontEmbeddingXmlMatches || fontRelationId === undefined || embeddedFontPayload === undefined || !fontEmbeddingPayloadMatches) throw new Error(JSON.stringify({ errors: fontEmbeddingFile.errors, presentation: fontPresentation, relationships: fontRelationships, fontRelationId, fontRelationTarget, payloadLength: embeddedFontPayload?.byteLength, sourceFontLength: embeddedFontBytes.byteLength, fontEmbeddingXmlMatches, fontEmbeddingStyleReferenceMatches, fontEmbeddingPayloadMatches }));',
      'await writeFile("font-embedding-installed.pptx", fontEmbeddingResult.pptx);',
      "let missingMetrics;",
      'try { await build(createElement(Deck, {}, createElement(Slide, { slideId: "missing-metrics" }, createElement(Text, { testId: "content-sized-text", x: 0, y: 0, width: 100, height: "auto", style: { fontFamily: font("Consumer Sans") } }, "Content sized text")))); } catch (error) { missingMetrics = error; }',
      'if (!(missingMetrics instanceof BuildError) || missingMetrics.diagnostics[0]?.code !== "font-measurement-unavailable") throw new Error("installed content-sized text did not require explicit font metrics");',
      'await writeFile("artifact.pptx", result.pptx);',
      'await writeFile("artifact-chart.pptx", chartResult.pptx);',
      'process.stdout.write(JSON.stringify({ installedCliSlideId: installedCliFile.slides[0]?.id, installedCliText: installedCliFile.slides[0]?.objects[0]?.text?.text, installedCliEquivalent, standard: file.profile.standard, schemaPackage: file.profile.schemaValidator.package, schemaPackageVersion: file.profile.schemaValidator.packageVersion, sdkVersion: file.profile.schemaValidator.sdkVersion, officeVersion: file.profile.schemaValidator.officeVersion, applicationEvidence: file.profile.applicationEvidence, slideId: file.slides[0]?.id, group: group.name, groupBox: group.boxEmu, shape: shape.name, shapeText: shape.text.text, shapeBox: shape.boxEmu, text: text.name, textContent: text.text.text, bullet: text.text.paragraphs[0]?.list?.type, image: image.name, mediaType: image.image.contentType, imageBox: image.boxEmu, table: tableObject.name, tableCells: tableObject.table.cells, tableMerges: tableObject.table.merges, tableColumnWidthEmu: tableObject.table.columnWidthsEmu[0], tableBytes: tableResult.pptx.byteLength, contentSizedText: contentSizedText.text.text, contentSizedTextHeightEmu: contentSizedText.boxEmu.height, contentSizedTextLineSpacingPoints: contentSizedText.text.paragraphs[0].lineSpacingPoints, nestedGroup: nestedGroup.name, nestedShape: nestedShape.name, nestedShapeText: nestedShape.text.text, nestedShapeBox: nestedShape.boxEmu, missingMetrics: missingMetrics.diagnostics[0].code, objectIds: [group.id, shape.id, text.id, image.id, nestedGroup.id, nestedShape.id], animations, transition: file.slides[0]?.transition?.kind, durationMs: file.slides[0]?.transition?.durationMs, extensions: file.extensions, forbiddenCode: forbidden.diagnostics[0]?.code, schemaGate: schemaIssue.code, schemaPart: schemaIssue.part, schemaXPath: schemaIssue.xpath, bytes: result.pptx.byteLength, contentSizedBytes: contentSizedResult.pptx.byteLength, embeddedFontPartBytes: embeddedFontPart?.byteLength, embeddedFontPayloadBytes: embeddedFontPayload?.byteLength, embeddedFontFamily: fontEmbeddingFile.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0]?.fontFamily, chartKinds: chartObjects.map(({ chart }) => chart?.types[0]), chartEmbeddedWorkbookParts: chartObjects.map(({ chart }) => chart?.embeddedWorkbook?.part), chartBytes: chartResult.pptx.byteLength, cancellationObserved, cancellationCleanups, cancellationListeners: getEventListeners(cancellationController.signal, "abort").length }));',
    ].join("\n"),
  );
  const cropLoader = join(temporaryDirectory, "reject-text-box.mjs");
  await writeFile(
    cropLoader,
    [
      "export async function resolve(specifier, context, nextResolve) {",
      "  const resolved = await nextResolve(specifier, context);",
      '  if (resolved.url.includes("/capabilities/dist/text.js")) throw new Error("installed text-box implementation was loaded");',
      "  return resolved;",
      "}",
    ].join("\n"),
  );
  await writeFile(
    join(temporaryDirectory, "consumer-cropped.mjs"),
    [
      'import { createElement } from "react";',
      'import { BuildError, createBuilder } from "@react-pptx/core";',
      'import { Text as TextDeclaration } from "@react-pptx/capabilities/text-elements";',
      'import { Cell, Row, Table, tableCapability } from "@react-pptx/capabilities/table";',
      'import { Deck, Shape, Slide, shapeCapability } from "@react-pptx/react-pptx/shape";',
      'import { validatePptx } from "@react-pptx/validation";',
      "const build = createBuilder([shapeCapability, tableCapability]);",
      'const result = await build(createElement(Deck, { width: 120, height: 96 }, createElement(Slide, { slideId: "installed-cropped" }, createElement(Shape, { testId: "installed-shape-text", x: 0, y: 0, width: 90, height: 30, text: "Installed shape text" }), createElement(Table, { testId: "installed-table", x: 0, y: 32, columns: [120], style: { fontSize: 14 } }, createElement(Row, { style: { height: 24 } }, createElement(Cell, null, "Installed table"))))));',
      "const file = await validatePptx(result.pptx);",
      'if (file.errors.length !== 0 || file.slides[0]?.objects[0]?.text?.text !== "Installed shape text" || file.slides[0]?.objects[1]?.kind !== "table" || file.slides[0]?.objects[1]?.table?.cells[0]?.[0] !== "Installed table") throw new Error(JSON.stringify(file.errors));',
      'function NestedText() { return createElement(TextDeclaration, { testId: "missing-installed-text", x: 0, y: 32, width: 90, height: 30 }, "not registered"); }',
      "let missingText;",
      'try { await build(createElement(Deck, {}, createElement(Slide, { slideId: "nested" }, createElement(NestedText)))); } catch (error) { missingText = error; }',
      'if (!(missingText instanceof BuildError) || missingText.diagnostics[0]?.code !== "missing-capability" || missingText.diagnostics[0]?.path !== "deck/slide[nested]/text[missing-installed-text]") throw new Error("installed nested Text did not report the cropped capability");',
      "process.stdout.write(JSON.stringify({ shapeText: file.slides[0]?.objects[0]?.text?.text, tableKind: file.slides[0]?.objects[1]?.kind, tableText: file.slides[0]?.objects[1]?.table?.cells[0]?.[0], missingCode: missingText.diagnostics[0]?.code, missingPath: missingText.diagnostics[0]?.path, textBoxImplementationLoaded: false }));",
    ].join("\n"),
  );
  const devtoolsSentinel = join(temporaryDirectory, "devtools-sentinel.mjs");
  await writeFile(
    devtoolsSentinel,
    [
      "let injections = 0;",
      "let notifications = 0;",
      "const hook = Object.freeze({ supportsFiber: true, inject() { injections += 1; return 73; }, onCommitFiberRoot() { notifications += 1; } });",
      "Object.defineProperty(globalThis, '__REACT_DEVTOOLS_GLOBAL_HOOK__', { value: hook, enumerable: false, configurable: false, writable: false });",
      "process.on('exit', () => { if (globalThis.__REACT_DEVTOOLS_GLOBAL_HOOK__ !== hook || injections !== 0 || notifications !== 0) throw new Error('React build changed or used the existing DevTools hook'); });",
    ].join("\n"),
  );

  run("pnpm", ["install", "--offline", "--ignore-scripts"], temporaryDirectory);
  run(
    "pnpm",
    ["exec", "tsc", "--project", "tsconfig.json"],
    temporaryDirectory,
  );
  const cliOutput = run(
    "pnpm",
    [
      "exec",
      "react-pptx",
      "consumer-deck.tsx",
      "--out",
      "cli-output/installed.pptx",
    ],
    temporaryDirectory,
  );
  if (!cliOutput.includes("已导出文稿：")) {
    throw new Error(`安装后的 CLI 没有报告成功导出：${cliOutput}`);
  }
  const summaryText = run(
    process.execPath,
    ["--import", devtoolsSentinel, "consumer-runtime.ts"],
    temporaryDirectory,
  );
  const summary = JSON.parse(summaryText) as {
    readonly standard?: string;
    readonly schemaPackage?: string;
    readonly schemaPackageVersion?: string;
    readonly sdkVersion?: string;
    readonly officeVersion?: string;
    readonly applicationEvidence?: string;
    readonly slideId?: string;
    readonly group?: string;
    readonly shape?: string;
    readonly shapeText?: string;
    readonly text?: string;
    readonly textContent?: string;
    readonly bullet?: string;
    readonly image?: string;
    readonly mediaType?: string;
    readonly nestedGroup?: string;
    readonly nestedShape?: string;
    readonly nestedShapeText?: string;
    readonly objectIds?: readonly number[];
    readonly animations?: readonly unknown[];
    readonly transition?: string;
    readonly durationMs?: number;
    readonly extensions?: readonly string[];
    readonly forbiddenCode?: string;
    readonly schemaGate?: string;
    readonly schemaPart?: string;
    readonly schemaXPath?: string;
    readonly bytes?: number;
    readonly contentSizedText?: string;
    readonly contentSizedTextHeightEmu?: number;
    readonly contentSizedTextLineSpacingPoints?: number;
    readonly table?: string;
    readonly tableCells?: readonly (readonly string[])[];
    readonly tableMerges?: readonly {
      readonly row: number;
      readonly column: number;
      readonly rowSpan: number;
      readonly colSpan: number;
    }[];
    readonly tableColumnWidthEmu?: number;
    readonly tableBytes?: number;
    readonly missingMetrics?: string;
    readonly contentSizedBytes?: number;
    readonly cancellationObserved?: boolean;
    readonly cancellationCleanups?: number;
    readonly cancellationListeners?: number;
    readonly installedCliSlideId?: string;
    readonly installedCliText?: string;
    readonly installedCliEquivalent?: boolean;
    readonly embeddedFontPartBytes?: number;
    readonly embeddedFontPayloadBytes?: number;
    readonly embeddedFontFamily?: string;
  };
  if (
    summary.installedCliSlideId !== "source-cli" ||
    summary.installedCliText !== "安装 CLI 源码标题" ||
    summary.installedCliEquivalent !== true ||
    summary.standard !== "ECMA-376 5th edition Transitional" ||
    summary.schemaPackage !== "@ooxml-tools/validate" ||
    summary.schemaPackageVersion !== "0.4.6" ||
    summary.sdkVersion !== "3.0.2" ||
    summary.officeVersion !== "Microsoft365" ||
    summary.applicationEvidence !== "not-validated" ||
    summary.slideId !== "installed" ||
    summary.group !== "packed-group" ||
    summary.shape !== "packed-shape" ||
    summary.shapeText !== "Installed shape text" ||
    summary.text !== "packed-text" ||
    summary.textContent !== "Installed 文本" ||
    summary.bullet !== "bullet" ||
    summary.image !== "packed-image" ||
    summary.mediaType !== "image/png" ||
    summary.nestedGroup !== "nested-group" ||
    summary.nestedShape !== "nested-shape" ||
    summary.nestedShapeText !== "Nested member" ||
    summary.objectIds?.length !== 6 ||
    new Set(summary.objectIds).size !== 6 ||
    JSON.stringify(summary.animations) !==
      JSON.stringify([
        [
          "auto",
          [
            [summary.objectIds?.[2], "spin", 1000, 0, 90, null],
            [
              summary.objectIds?.[0],
              "fade",
              800,
              100,
              null,
              { decelerate: 0.5 },
            ],
          ],
        ],
        [
          "click",
          [
            [summary.objectIds?.[1], "fade", 250, 0, null, null],
            [summary.objectIds?.[1], "spin", 1000, 0, 180, null],
            [summary.objectIds?.[3], "wipe", 500, 0, "right", null],
          ],
        ],
      ]) ||
    summary.transition !== "morph" ||
    summary.durationMs !== 1200 ||
    summary.extensions?.join(",") !== "p14,p159" ||
    summary.forbiddenCode !== "forbidden-mechanism" ||
    summary.schemaGate !== "Sch_AttributeValueDataTypeDetailed" ||
    summary.schemaPart !== "/ppt/slideMasters/slideMaster1.xml" ||
    summary.schemaXPath === undefined ||
    summary.contentSizedText !== "Installed font" ||
    summary.contentSizedTextHeightEmu !== 550672 ||
    summary.contentSizedTextLineSpacingPoints !== 21.68 ||
    summary.table !== "packed-table" ||
    JSON.stringify(summary.tableCells) !==
      JSON.stringify([
        ["Period", "", "YoY"],
        ["N", "Q1", "100"],
        ["", "Q2", "120"],
      ]) ||
    JSON.stringify(summary.tableMerges) !==
      JSON.stringify([
        { row: 0, column: 0, rowSpan: 1, colSpan: 2 },
        { row: 1, column: 0, rowSpan: 2, colSpan: 1 },
      ]) ||
    summary.tableColumnWidthEmu !== 266700 ||
    summary.tableBytes === undefined ||
    summary.tableBytes <= 0 ||
    summary.missingMetrics !== "font-measurement-unavailable" ||
    summary.bytes === undefined ||
    summary.bytes <= 0 ||
    summary.contentSizedBytes === undefined ||
    summary.contentSizedBytes <= 0 ||
    summary.cancellationObserved !== true ||
    summary.cancellationCleanups !== 1 ||
    summary.cancellationListeners !== 0 ||
    summary.embeddedFontPayloadBytes !== 90780 ||
    summary.embeddedFontPartBytes === undefined ||
    summary.embeddedFontPartBytes <= summary.embeddedFontPayloadBytes ||
    summary.embeddedFontFamily !== "Jost*"
  ) {
    throw new Error(`安装制品返回了意外结果：${summaryText}`);
  }
  const artifactDirectory = join(
    process.cwd(),
    ".scratch/react-pptx-formal/12-artifacts",
  );
  await mkdir(artifactDirectory, { recursive: true });
  await writeFile(
    join(artifactDirectory, "installed-font-embedding.pptx"),
    await readFile(join(temporaryDirectory, "font-embedding-installed.pptx")),
  );
  const croppedText = run(
    process.execPath,
    ["--loader", cropLoader, "consumer-cropped.mjs"],
    temporaryDirectory,
  );
  const cropped = JSON.parse(croppedText) as {
    readonly shapeText?: string;
    readonly tableKind?: string;
    readonly tableText?: string;
    readonly missingCode?: string;
    readonly missingPath?: string;
    readonly textBoxImplementationLoaded?: boolean;
  };
  if (
    cropped.shapeText !== "Installed shape text" ||
    cropped.tableKind !== "table" ||
    cropped.tableText !== "Installed table" ||
    cropped.missingCode !== "missing-capability" ||
    cropped.missingPath !== "deck/slide[nested]/text[missing-installed-text]" ||
    cropped.textBoxImplementationLoaded !== false
  ) {
    throw new Error(`安装制品裁剪入口返回了意外结果：${croppedText}`);
  }
  process.stdout.write(
    `正式制品消费通过：${summaryText}\n安装制品裁剪通过：${croppedText}\n`,
  );
} finally {
  await rm(temporaryDirectory, { recursive: true, force: true });
}
