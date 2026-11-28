import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import {
  mkdir,
  mkdtemp,
  rename,
  rm,
  symlink,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

import type { CodegenResult } from "@pptxdk/codegen";

const packageRoot = fileURLToPath(new URL("../", import.meta.url));
const tscBin = join(packageRoot, "node_modules/.bin/tsc");

export type BatchCodegenCase = {
  readonly scenarioId: string;
  readonly originLabel: "direct" | "roundtrip-imported";
  readonly result: CodegenResult;
};

const vanillaStubJs = `export {
  AuthorQuery,
  AuthorQueryError,
  buildDiagnosticSchema,
  CommonContainer,
  CommonNode,
  DeckContainer,
  RichTextContainer,
  SlideContainer,
  SlideLayoutContainer,
  SlideMasterContainer,
  darker,
  defaultColorMap,
  defaultTheme,
  defineTheme,
  isNativeColorSlot,
  isSchemeColorName,
  isSchemeColorSlot,
  isSlideLayoutDefinition,
  isTheme,
  isThemeFontSlot,
  lighter,
  nativeColorSlots,
  schemeColorNames,
  schemeColorSlots,
  themeFontSlots,
} from "@pptxdk/core";
export { createMorphScope, isMorphKey } from "@pptxdk/core/capability";
export {
  after,
  blockArc,
  box,
  breathe,
  cell,
  chart,
  createMotionScopeHandle,
  deck,
  defineSlideLayout,
  fade,
  fillColor,
  floatDown,
  floatUp,
  fly,
  font,
  fromFontCollection,
  group,
  image,
  linearGradient,
  motionScope,
  motionSchedule,
  notes,
  paragraph,
  path,
  placeholder,
  preset,
  radialGradient,
  repeat,
  row,
  section,
  shape,
  slide,
  slideLayout,
  slideMaster,
  slideNumber,
  span,
  spin,
  spinIn,
  step,
  stretch,
  table,
  text,
  transitions,
  wipe,
  zoom,
} from "@pptxdk/capabilities/document";
`;

const vanillaStubDts = `${vanillaStubJs}export type * from "@pptxdk/core";
export type {
  BoxAuthorProps,
  BoxNode,
  CellAuthorProps,
  CellNode,
  ChartAuthorProps,
  ChartNode,
  DeckAuthorProps,
  DeckNode,
  GroupAuthorProps,
  GroupNode,
  ImageAuthorProps,
  ImageNode,
  ImperativePlaceholders,
  InferPlaceholderMap,
  MotionProps,
  MotionScopeAuthorProps,
  MotionScopeNode,
  MotionStepAuthorProps,
  MotionStepNode,
  NotesAuthorProps,
  NotesNode,
  ParagraphAuthorProps,
  ParagraphContent,
  ParagraphNode,
  ParagraphStyle,
  PlaceholderAuthorProps,
  PlaceholderNode,
  RowAuthorProps,
  RowNode,
  SectionAuthorProps,
  SectionNode,
  ShapeAuthorProps,
  ShapeNode,
  ShapeTextStyle,
  ShapeVisualStyle,
  SlideAuthorProps,
  SlideLayoutAuthorProps,
  SlideLayoutNode,
  SlideLayoutSource,
  SlideMasterAuthorProps,
  SlideMasterNode,
  SlideNode,
  SlideNumberAuthorProps,
  SlideNumberNode,
  SpanAuthorProps,
  SpanContent,
  SpanNode,
  TableAuthorProps,
  TableCellStyle,
  TableColumn,
  TableNode,
  TableRowStyle,
  TableSpacing,
  TableStyle,
  TableTextStyle,
  TableVerticalAlign,
  TextAuthorProps,
  TextContent,
  TextNode,
  Transition,
} from "@pptxdk/capabilities/document";
export type {
  BreatheEffect,
  BreatheOptions,
  Easing,
  Effect,
  EffectInstance,
  EffectOptions,
  EffectPreset,
  EmphasisEffect,
  EmphasisEffectDeclaration,
  EntranceEffect,
  EntranceEffectDeclaration,
  FillColorOptions,
  MotionCompletionState,
  MotionCue,
  MotionCueRef,
  MotionDecision,
  MotionDependency,
  MotionDifference,
  MotionDirection,
  MotionEffectRef,
  MotionGranularity,
  MotionIntervalRef,
  MotionMembership,
  MotionMoment,
  MotionMomentOffset,
  MotionMomentRef,
  MotionNativeProjection,
  MotionSchedule,
  MotionScopeHandle,
  MotionScopeMode,
  MotionScopeQueryMember,
  MotionScopeRef,
  MotionScopeSchedule,
  MotionStartDependency,
  MotionStepQuery,
  MotionStepRef,
  MotionTargetRange,
  MotionTime,
  MotionTiming,
  OwnerMotion,
  OwnerMotionEffect,
  OwnerMotionStep,
  OwnerMotionTarget,
  RepeatCount,
  RepeatedEffect,
  ResolvedMotionEffect,
  ResolvedMotionMoment,
  ResolvedMotionScope,
  ResolvedMotionStep,
  SpinOptions,
} from "@pptxdk/capabilities/document";
export type {
  ChartAxis,
  ChartAxisCrossing,
  ChartAxisPosition,
  ChartAxisReferences,
  ChartCategoryAxis,
  ChartCategoryAxisReferences,
  ChartCustomErrorBars,
  ChartDataLabelOverride,
  ChartDataLabels,
  ChartDisplayBlanksAs,
  ChartErrorBarInclude,
  ChartErrorBars,
  ChartErrorBarValues,
  ChartGroup,
  ChartGroupData,
  ChartKind,
  ChartLine,
  ChartLineDash,
  ChartMarker,
  ChartMarkerSymbol,
  ChartPoint,
  ChartSeries,
  ChartSeriesDataLabelOverride,
  ChartTickLabelPosition,
  ChartTickMark,
  ChartTitle,
  ChartTrendline,
  ChartTrendlineType,
  ChartValueAxis,
  ChartXYAxisReferences,
  ColumnOrBarChartGroup,
  DoughnutChartGroup,
  IndependentCategoryData,
  IndependentCategorySeries,
  LineOrAreaChartGroup,
  PieChartGroup,
  ScatterChartGroup,
  ScatterChartPoint,
  ScatterChartSeries,
  ScatterData,
  ScatterDataLabelOverride,
  ScatterDataLabels,
  ScatterErrorBars,
  ScatterLine,
  SharedCategoryData,
} from "@pptxdk/capabilities/document";
export type {
  ArrowType,
  CustomGeometry,
  FlipTransition,
  Geometry,
  GeometryPath,
  ImageSource,
  ImageVisualStyle,
  LineEndStyle,
  MorphProps,
  MorphTransition,
  NotesDeclarationProps,
  PathCommand,
  PathOptions,
  PresetGeometry,
  PresetGeometryName,
  StandardTransition,
  StripsDirection,
  TransitionDirection,
  TransitionDuration,
  TransitionSpeed,
} from "@pptxdk/capabilities/document";
`;

const reactStubJs = `export {
  appendJSX,
  Deck,
  expandDocument,
  expandFragment,
  fromVanilla,
  Mount,
  Section,
  Slide,
  SlideLayout,
  SlideMaster,
} from "@pptxdk/react";
export { Box } from "@pptxdk/capabilities/box-elements";
export { Chart } from "@pptxdk/capabilities/chart-elements";
export { Group } from "@pptxdk/capabilities/group-elements";
export { Image } from "@pptxdk/capabilities/image-elements";
export {
  MotionScope,
  Step,
  useMotionScope,
} from "@pptxdk/capabilities/motion-elements";
export { Notes } from "@pptxdk/capabilities/notes-elements";
export { Placeholder } from "@pptxdk/capabilities/placeholder-elements";
export { Shape } from "@pptxdk/capabilities/shape-elements";
export { Cell, Row, Table } from "@pptxdk/capabilities/table-elements";
export { P, SlideNumber, Span, Text } from "@pptxdk/capabilities/text-elements";
export {
  AuthorQuery,
  AuthorQueryError,
  buildDiagnosticSchema,
  CommonContainer,
  CommonNode,
  DeckContainer,
  RichTextContainer,
  SlideContainer,
  SlideLayoutContainer,
  SlideMasterContainer,
  darker,
  defaultColorMap,
  defaultTheme,
  defineTheme,
  isNativeColorSlot,
  isSchemeColorName,
  isSchemeColorSlot,
  isSlideLayoutDefinition,
  isTheme,
  isThemeFontSlot,
  lighter,
  nativeColorSlots,
  schemeColorNames,
  schemeColorSlots,
  themeFontSlots,
} from "@pptxdk/core";
export { createMorphScope, isMorphKey } from "@pptxdk/core/capability";
export {
  after,
  blockArc,
  breathe,
  createMotionScopeHandle,
  defineSlideLayout,
  fade,
  fillColor,
  floatDown,
  floatUp,
  fly,
  font,
  fromFontCollection,
  linearGradient,
  motionSchedule,
  path,
  preset,
  radialGradient,
  repeat,
  spin,
  spinIn,
  stretch,
  transitions,
  wipe,
  zoom,
} from "@pptxdk/capabilities/document";
`;

const reactStubDts = `import type {
  InferPlaceholderMap,
  JsxPlaceholders,
  SlideLayoutAuthorProps,
  SlideLayoutSource,
  SlideMasterAuthorProps,
} from "@pptxdk/capabilities/document";
import type { Transition } from "@pptxdk/capabilities/transition";
import type {
  ElementProps,
  SlideProps as ReactSlideProps,
} from "@pptxdk/react";
import type { ReactNode } from "react";

export {
  appendJSX,
  Deck,
  expandDocument,
  expandFragment,
  fromVanilla,
  Mount,
  Section,
} from "@pptxdk/react";
export type {
  AppendJSXOptions,
  DeckProps,
  ElementProps,
  ExpandDocumentOptions,
  ExpandFragmentOptions,
  ExpandedDocument,
  MountProps,
  SectionProps,
  VanillaFactoryProps,
  VanillaSlotResult,
} from "@pptxdk/react";

type BaseSlideProps = Omit<
  ReactSlideProps,
  "transition" | "layout" | "placeholders"
> & {
  readonly transition?: Transition;
};
export type BoundSlideProps<L extends SlideLayoutSource = SlideLayoutSource> =
  BaseSlideProps & {
    readonly layout: L;
    readonly placeholders?: JsxPlaceholders<InferPlaceholderMap<L>>;
  };
export type FreeSlideProps = BaseSlideProps & {
  readonly layout?: undefined;
  readonly placeholders?: never;
};
export type SlideProps<
  L extends SlideLayoutSource | undefined = SlideLayoutSource | undefined,
> = [L] extends [undefined]
  ? FreeSlideProps
  : [L] extends [SlideLayoutSource]
    ? BoundSlideProps<L>
    : BoundSlideProps | FreeSlideProps;
export type SlideMasterProps = SlideMasterAuthorProps & ElementProps;
export type SlideLayoutProps = SlideLayoutAuthorProps & ElementProps;
export declare function Slide<
  const L extends SlideLayoutSource | undefined = undefined,
>(props: SlideProps<L>): ReactNode;
export declare const SlideMaster: (props: SlideMasterProps) => ReactNode;
export declare const SlideLayout: (props: SlideLayoutProps) => ReactNode;

export { Box, type BoxProps } from "@pptxdk/capabilities/box-elements";
export { Chart, type ChartProps } from "@pptxdk/capabilities/chart-elements";
export { Group, type GroupProps } from "@pptxdk/capabilities/group-elements";
export { Image, type ImageProps } from "@pptxdk/capabilities/image-elements";
export {
  MotionScope,
  Step,
  useMotionScope,
  type MotionProps,
  type MotionScopeProps,
  type StepProps,
} from "@pptxdk/capabilities/motion-elements";
export { Notes, type NotesProps } from "@pptxdk/capabilities/notes-elements";
export {
  Placeholder,
  type PlaceholderProps,
} from "@pptxdk/capabilities/placeholder-elements";
export { Shape, type ShapeProps } from "@pptxdk/capabilities/shape-elements";
export {
  Cell,
  Row,
  Table,
  type CellProps,
  type RowProps,
  type TableProps,
} from "@pptxdk/capabilities/table-elements";
export {
  P,
  SlideNumber,
  Span,
  Text,
  type ParagraphProps,
  type SlideNumberProps,
  type SpanProps,
  type TextProps,
} from "@pptxdk/capabilities/text-elements";
export {
  AuthorQuery,
  AuthorQueryError,
  buildDiagnosticSchema,
  CommonContainer,
  CommonNode,
  DeckContainer,
  RichTextContainer,
  SlideContainer,
  SlideLayoutContainer,
  SlideMasterContainer,
  darker,
  defaultColorMap,
  defaultTheme,
  defineTheme,
  isNativeColorSlot,
  isSchemeColorName,
  isSchemeColorSlot,
  isSlideLayoutDefinition,
  isTheme,
  isThemeFontSlot,
  lighter,
  nativeColorSlots,
  schemeColorNames,
  schemeColorSlots,
  themeFontSlots,
} from "@pptxdk/core";
export { createMorphScope, isMorphKey } from "@pptxdk/core/capability";
export type * from "@pptxdk/core";
export {
  after,
  blockArc,
  breathe,
  createMotionScopeHandle,
  defineSlideLayout,
  fade,
  fillColor,
  floatDown,
  floatUp,
  fly,
  font,
  fromFontCollection,
  linearGradient,
  motionSchedule,
  path,
  preset,
  radialGradient,
  repeat,
  spin,
  spinIn,
  stretch,
  transitions,
  wipe,
  zoom,
} from "@pptxdk/capabilities/document";
export type {
  ArrowType,
  BreatheEffect,
  BreatheOptions,
  ChartAxis,
  ChartAxisCrossing,
  ChartAxisPosition,
  ChartAxisReferences,
  ChartCategoryAxis,
  ChartCategoryAxisReferences,
  ChartCustomErrorBars,
  ChartDataLabelOverride,
  ChartDataLabels,
  ChartDisplayBlanksAs,
  ChartErrorBarInclude,
  ChartErrorBars,
  ChartErrorBarValues,
  ChartGroup,
  ChartGroupData,
  ChartKind,
  ChartLine,
  ChartLineDash,
  ChartMarker,
  ChartMarkerSymbol,
  ChartPoint,
  ChartSeries,
  ChartSeriesDataLabelOverride,
  ChartTickLabelPosition,
  ChartTickMark,
  ChartTitle,
  ChartTrendline,
  ChartTrendlineType,
  ChartValueAxis,
  ChartXYAxisReferences,
  ColumnOrBarChartGroup,
  CustomGeometry,
  DoughnutChartGroup,
  Easing,
  Effect,
  EffectInstance,
  EffectOptions,
  EffectPreset,
  EmphasisEffect,
  EmphasisEffectDeclaration,
  EntranceEffect,
  EntranceEffectDeclaration,
  FillColorOptions,
  FlipTransition,
  Geometry,
  GeometryPath,
  ImageSource,
  ImageVisualStyle,
  IndependentCategoryData,
  IndependentCategorySeries,
  InferPlaceholderMap,
  JsxPlaceholders,
  LineEndStyle,
  LineOrAreaChartGroup,
  LinkTarget,
  MorphProps,
  MorphTransition,
  MotionCompletionState,
  MotionCue,
  MotionCueRef,
  MotionDecision,
  MotionDependency,
  MotionDifference,
  MotionDirection,
  MotionEffectRef,
  MotionGranularity,
  MotionIntervalRef,
  MotionMembership,
  MotionMoment,
  MotionMomentOffset,
  MotionMomentRef,
  MotionNativeProjection,
  MotionSchedule,
  MotionScopeHandle,
  MotionScopeMode,
  MotionScopeQueryMember,
  MotionScopeRef,
  MotionScopeSchedule,
  MotionStartDependency,
  MotionStepQuery,
  MotionStepRef,
  MotionTargetRange,
  MotionTime,
  MotionTiming,
  NotesDeclarationProps,
  OwnerMotion,
  OwnerMotionEffect,
  OwnerMotionStep,
  OwnerMotionTarget,
  ParagraphStyle,
  PathCommand,
  PathOptions,
  PieChartGroup,
  PresetGeometry,
  PresetGeometryName,
  RepeatCount,
  RepeatedEffect,
  ResolvedMotionEffect,
  ResolvedMotionMoment,
  ResolvedMotionScope,
  ResolvedMotionStep,
  ScatterChartGroup,
  ScatterChartPoint,
  ScatterChartSeries,
  ScatterData,
  ScatterDataLabelOverride,
  ScatterDataLabels,
  ScatterErrorBars,
  ScatterLine,
  ShapeTextStyle,
  ShapeVisualStyle,
  SharedCategoryData,
  SlideLayoutSource,
  SpinOptions,
  StandardTransition,
  StripsDirection,
  TableCellStyle,
  TableColumn,
  TableRowStyle,
  TableSpacing,
  TableStyle,
  TableTextStyle,
  TableVerticalAlign,
  TextList,
  Transition,
  TransitionDirection,
  TransitionDuration,
  TransitionSpeed,
  TypedElement,
} from "@pptxdk/capabilities/document";
`;

async function writeCodegenFiles(
  targetDir: string,
  files: CodegenResult["files"],
): Promise<void> {
  for (const [relPath, content] of Object.entries(files)) {
    const absPath = join(targetDir, relPath);
    await mkdir(dirname(absPath), { recursive: true });
    await writeFile(absPath, content);
  }
}

async function setupConsumerNodeModules(nodeModulesDir: string): Promise<void> {
  await mkdir(join(nodeModulesDir, "@pptxdk"), { recursive: true });
  await mkdir(join(nodeModulesDir, "@officedk/pptxdk"), { recursive: true });
  await mkdir(join(nodeModulesDir, "@types"), { recursive: true });

  await symlink(
    join(packageRoot, "node_modules/@pptxdk/core"),
    join(nodeModulesDir, "@pptxdk/core"),
    "dir",
  );
  await symlink(
    packageRoot,
    join(nodeModulesDir, "@pptxdk/capabilities"),
    "dir",
  );
  await symlink(
    join(packageRoot, "node_modules/@pptxdk/emit"),
    join(nodeModulesDir, "@pptxdk/emit"),
    "dir",
  );
  await symlink(
    join(packageRoot, "node_modules/@pptxdk/react"),
    join(nodeModulesDir, "@pptxdk/react"),
    "dir",
  );
  await symlink(
    join(packageRoot, "node_modules/@pptxdk/validation"),
    join(nodeModulesDir, "@pptxdk/validation"),
    "dir",
  );
  await symlink(
    join(packageRoot, "node_modules/react"),
    join(nodeModulesDir, "react"),
    "dir",
  );
  await symlink(
    join(packageRoot, "node_modules/@types/react"),
    join(nodeModulesDir, "@types/react"),
    "dir",
  );
  await symlink(
    join(packageRoot, "node_modules/@types/node"),
    join(nodeModulesDir, "@types/node"),
    "dir",
  );

  // Local @officedk/pptxdk stub built strictly from layer:capabilities legal dependencies.
  const stubDir = join(nodeModulesDir, "@officedk/pptxdk");
  await writeFile(
    join(stubDir, "package.json"),
    JSON.stringify({
      name: "@officedk/pptxdk",
      version: "0.0.0",
      private: true,
      type: "module",
      exports: {
        "./vanilla": { types: "./vanilla.d.ts", default: "./vanilla.js" },
        "./react": { types: "./react.d.ts", default: "./react.js" },
      },
    }),
  );
  await writeFile(join(stubDir, "vanilla.js"), vanillaStubJs);
  await writeFile(join(stubDir, "vanilla.d.ts"), vanillaStubDts);
  await writeFile(join(stubDir, "react.js"), reactStubJs);
  await writeFile(join(stubDir, "react.d.ts"), reactStubDts);
}

/**
 * Batch-typechecks all generated `vanilla` and `react` sources in one temporary isolated project
 * via a single strict `tsc --noEmit` invocation, and verifies representative relocated rebuilds.
 */
export async function verifyBatchGeneratedSourcesAndRelocatedRebuilds(
  cases: readonly BatchCodegenCase[],
  representativeRelocateScenarioIds: readonly string[],
): Promise<{
  readonly typecheckedCaseCount: number;
  readonly relocatedRebuildCount: number;
}> {
  assert.ok(cases.length > 0, "Expected at least one generated case");

  const tempRoot = await mkdtemp(join(tmpdir(), "pptxdk-cap-codegen-batch-"));
  try {
    await writeFile(
      join(tempRoot, "package.json"),
      JSON.stringify({ type: "module" }),
    );
    await setupConsumerNodeModules(join(tempRoot, "node_modules"));

    const manifest: string[] = [];
    const writtenDirs: {
      readonly item: BatchCodegenCase;
      readonly caseDir: string;
    }[] = [];

    for (const [index, item] of cases.entries()) {
      const caseDir = join(tempRoot, "cases", `case-${index}`);
      await writeCodegenFiles(caseDir, item.result.files);
      manifest.push(
        `case-${index}: ${item.scenarioId} (${item.originLabel}:${item.result.target})`,
      );
      writtenDirs.push({ item, caseDir });
    }

    await writeFile(
      join(tempRoot, "tsconfig.json"),
      JSON.stringify({
        compilerOptions: {
          target: "es2024",
          module: "nodenext",
          moduleResolution: "nodenext",
          jsx: "react-jsx",
          jsxImportSource: "react",
          strict: true,
          exactOptionalPropertyTypes: true,
          noUncheckedIndexedAccess: true,
          verbatimModuleSyntax: true,
          noEmit: true,
          skipLibCheck: false,
          types: ["node"],
        },
        include: ["cases/**/*.ts", "cases/**/*.tsx"],
      }),
    );

    const tscResult = spawnSync(
      tscBin,
      ["-p", "tsconfig.json", "--pretty", "false"],
      { cwd: tempRoot, encoding: "utf8" },
    );
    assert.equal(
      tscResult.status,
      0,
      `Batch strict tsc failed across ${cases.length} generated cases:\n${manifest.join("\n")}\n\n${tscResult.stdout}\n${tscResult.stderr}`,
    );

    const representativeIds = new Set(representativeRelocateScenarioIds);
    const relocateRoot = join(tempRoot, "relocated-consumer");
    await mkdir(relocateRoot, { recursive: true });
    await writeFile(
      join(relocateRoot, "package.json"),
      JSON.stringify({ type: "module" }),
    );
    await symlink(
      join(tempRoot, "node_modules"),
      join(relocateRoot, "node_modules"),
      "dir",
    );

    const selectedRelocate = writtenDirs.filter(
      ({ item }) =>
        item.originLabel === "direct" && representativeIds.has(item.scenarioId),
    );
    for (const id of representativeIds) {
      assert.ok(
        selectedRelocate.some(({ item }) => item.scenarioId === id),
        `Representative relocate scenario ${id} was not found in batch`,
      );
    }

    const relocatedTargets: {
      readonly item: BatchCodegenCase;
      readonly dir: string;
    }[] = [];
    const reactIncludePaths: string[] = [];

    for (const [idx, entry] of selectedRelocate.entries()) {
      const relocatedDir = join(relocateRoot, "nested", `proj-${idx}`);
      await mkdir(dirname(relocatedDir), { recursive: true });
      await rename(entry.caseDir, relocatedDir);
      relocatedTargets.push({ item: entry.item, dir: relocatedDir });
      if (entry.item.result.target === "react") {
        reactIncludePaths.push(`nested/proj-${idx}/${entry.item.result.entry}`);
      }
    }

    // Emit JS modules for relocated React TSX entries so Node imports them from their relocated directories.
    if (reactIncludePaths.length > 0) {
      await writeFile(
        join(relocateRoot, "tsconfig.emit.json"),
        JSON.stringify({
          compilerOptions: {
            target: "es2024",
            module: "nodenext",
            moduleResolution: "nodenext",
            jsx: "react-jsx",
            jsxImportSource: "react",
            strict: true,
            skipLibCheck: true,
            types: ["node"],
          },
          include: reactIncludePaths,
        }),
      );
      const emitResult = spawnSync(
        tscBin,
        ["-p", "tsconfig.emit.json", "--pretty", "false"],
        { cwd: relocateRoot, encoding: "utf8" },
      );
      assert.equal(
        emitResult.status,
        0,
        `Relocated React tsc emit failed:\n${emitResult.stdout}\n${emitResult.stderr}`,
      );
    }

    const manifestEntries = relocatedTargets.map(({ item, dir }) => ({
      scenarioId: item.scenarioId,
      target: item.result.target,
      entryUrl: pathToFileURL(
        join(
          dir,
          item.result.target === "react"
            ? item.result.entry.replace(/\.tsx$/u, ".js")
            : item.result.entry,
        ),
      ).href,
    }));

    const runnerScript = `import assert from "node:assert/strict";
import { createElement } from "react";
import { documentCapabilities } from "@pptxdk/capabilities/all";
import { createResolver } from "@pptxdk/core/build";
import { CommonNode } from "@pptxdk/core/document";
import { emit } from "@pptxdk/emit";
import { expandDocument } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";

const resolve = createResolver(documentCapabilities);
const targets = ${JSON.stringify(manifestEntries)};

for (const entry of targets) {
  const mod = await import(entry.entryUrl);
  const exported = mod.default;
  const doc = exported instanceof CommonNode
    ? exported
    : await expandDocument(createElement(exported));
  const emitted = emit(resolve(doc));
  const validated = await validatePptx(emitted.pptx);
  assert.deepEqual(
    validated.errors,
    [],
    "Relocated rebuild validation errors for " + entry.scenarioId + " (" + entry.target + "): " + JSON.stringify(validated.errors),
  );
}
`;
    const runnerPath = join(relocateRoot, "run-relocated.mjs");
    await writeFile(runnerPath, runnerScript);
    const rebuildResult = spawnSync(process.execPath, [runnerPath], {
      cwd: relocateRoot,
      encoding: "utf8",
    });
    assert.equal(
      rebuildResult.status,
      0,
      `Relocated rebuild failed:\n${rebuildResult.stdout}\n${rebuildResult.stderr}`,
    );

    return {
      typecheckedCaseCount: cases.length,
      relocatedRebuildCount: relocatedTargets.length,
    };
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
}
