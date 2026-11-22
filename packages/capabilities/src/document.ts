import {
  defineSlideLayout,
  type LayoutStyle,
  type LinkTarget,
  type PaddingStyle,
  type SlideBackground,
  type SlideLayoutDataProps,
  type SlideLayoutDefinition,
  type SlideLayoutProps,
  type SlideLayoutType,
  type SlideMasterProps,
  type TextStyle,
  type ThemeInput,
  type ThemeScope,
  type VisualStyle,
} from "@pptxdk/core";
import {
  CommonContainer,
  CommonNode,
  DeckContainer,
  RichTextContainer,
  SlideLayoutContainer,
  SlideMasterContainer,
  createDeckNode,
  createSectionNode,
  createSlideLayoutNode,
  createSlideMasterNode,
  createSlideNode,
  type RichTextItem,
} from "@pptxdk/core/document";

import {
  createBoxNode,
  type BoxDeclarationProps,
} from "./box/box-declaration.ts";
import {
  createChartNode,
  type ChartNode,
  type ChartProps,
} from "./chart/chart-declaration.ts";
import { font } from "./font/font.ts";
import {
  createGroupNode,
  type GroupDeclarationProps,
} from "./group/group-declaration.ts";
import {
  createImageNode,
  type ImageDeclarationProps,
} from "./image/image-declaration.ts";
import {
  createMotionScopeNode,
  createMotionStepIdentity,
  createMotionStepNode,
  type MotionScopeAuthorProps,
  type MotionScopeHandle,
  type MotionStepAuthorProps,
} from "./motion/motion-api.ts";
import {
  createNotesNode,
  type NotesDeclarationProps,
} from "./notes/notes-declaration.ts";
import {
  createShapeNode,
  type ShapeDeclarationProps,
} from "./shape/shape-declaration.ts";
import {
  createCellNode,
  createRowNode,
  createTableNode,
  type CellDeclarationProps,
  type CellNode,
  type RowDeclarationProps,
  type RowNode,
  type TableDeclarationProps,
  type TableNode,
} from "./table/table-declaration.ts";
import {
  createParagraphNode,
  createSlideNumberNode,
  createSpanNode,
  createTextNode,
  type ParagraphDeclarationProps,
  type SlideNumberDeclarationProps,
  type SpanDeclarationProps,
  type TextDeclarationProps,
} from "./text/text-declaration.ts";
import {
  transitions,
  type FlipTransition,
  type MorphTransition,
  type StandardTransition,
  type StripsDirection,
  type Transition,
  type TransitionDirection,
  type TransitionDuration,
  type TransitionSpeed,
} from "./transition/transition-api.ts";

/**
 * External addresses such as link URLs are stored as address values at authoring time. Content
 * referenced by an address (including resource paths accepted by a capability) is read during each
 * formal build; creating an authoring node does not snapshot external file content.
 *
 * Image byte sources are isolated when accepted by `image()` or `set()`, so later changes to the
 * caller's buffer do not alter the document. Path and URL sources keep their address value instead
 * and are loaded again by every formal build.
 */

export type DeckAuthorProps = {
  readonly width?: number;
  readonly height?: number;
  readonly theme?: ThemeInput;
  readonly style?: TextStyle;
  readonly title?: string;
};

export type SlideMasterAuthorProps = SlideMasterProps;
export type SlideLayoutAuthorProps = SlideLayoutProps;
export type SlideLayoutSource = SlideLayoutDefinition | SlideLayoutContainer;

export type SlideAuthorProps = {
  readonly id: string;
  readonly layout?: SlideLayoutSource | undefined;
  readonly transition?: Transition;
  readonly background?: SlideBackground;
  readonly theme?: ThemeScope;
  readonly style?: TextStyle & Omit<LayoutStyle, "zIndex">;
  readonly hidden?: boolean;
  readonly advanceAfter?: number;
};

export type SectionAuthorProps = {
  readonly name: string;
};

export type BoxAuthorProps = BoxDeclarationProps;
export type ImageAuthorProps = ImageDeclarationProps;
export type GroupAuthorProps = GroupDeclarationProps;
export type ShapeAuthorProps = ShapeDeclarationProps;
export type TextAuthorProps = TextDeclarationProps;
export type ParagraphAuthorProps = ParagraphDeclarationProps;
export type SpanAuthorProps = SpanDeclarationProps;
export type SlideNumberAuthorProps = SlideNumberDeclarationProps;
export type ChartAuthorProps = ChartProps;
export type TableAuthorProps = TableDeclarationProps;
export type RowAuthorProps = RowDeclarationProps;
export type CellAuthorProps = CellDeclarationProps;
export type NotesAuthorProps = NotesDeclarationProps;

export type DeckNode = DeckContainer<DeckAuthorProps>;
export type SlideMasterNode = SlideMasterContainer;
export type SlideLayoutNode = SlideLayoutContainer;
export type SlideNode = CommonContainer<
  "slide",
  SlideAuthorProps,
  TextStyle & Omit<LayoutStyle, "zIndex">
>;
export type SectionNode = CommonContainer<"section", SectionAuthorProps>;
export type BoxNode = CommonContainer<
  "box",
  BoxAuthorProps,
  LayoutStyle & TextStyle & VisualStyle
>;
export type ShapeNode = RichTextContainer<
  "shape",
  ShapeDeclarationProps,
  NonNullable<ShapeDeclarationProps["style"]>
>;
export type TextNode = RichTextContainer<
  "text",
  TextAuthorProps,
  LayoutStyle & PaddingStyle & TextStyle & VisualStyle
>;
export type NotesNode = RichTextContainer<"notes", NotesAuthorProps, TextStyle>;
export type ImageNode = CommonNode<
  "image",
  ImageDeclarationProps,
  NonNullable<ImageDeclarationProps["style"]>
>;
export type GroupNode = CommonContainer<
  "group",
  GroupDeclarationProps,
  NonNullable<GroupDeclarationProps["style"]>
>;
export type ParagraphNode = RichTextContainer<
  "paragraph",
  ParagraphAuthorProps,
  NonNullable<ParagraphDeclarationProps["style"]>
>;
export type SpanNode = RichTextContainer<
  "span",
  SpanAuthorProps,
  NonNullable<SpanDeclarationProps["style"]>
>;
export type SlideNumberNode = CommonNode<
  "slide-number",
  SlideNumberAuthorProps,
  NonNullable<SlideNumberDeclarationProps["style"]>
>;
export type MotionScopeNode = CommonContainer<
  "motion-scope",
  MotionScopeAuthorProps
>;
export type MotionStepNode = CommonContainer<
  "motion-step",
  MotionStepAuthorProps
>;
export type { CellNode, ChartNode, RowNode, TableNode };

export type TextContent = RichTextItem | readonly RichTextItem[];
export type ParagraphContent = RichTextItem | readonly RichTextItem[];
export type SpanContent = RichTextItem | readonly RichTextItem[];

function isPropsRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return false;
  }
  if (value instanceof CommonNode) return false;

  // Factory props are structural records even when a caller uses a class to organize them.
  // Built-in opaque values keep their own tag and remain in the rich-text content position.
  return Object.prototype.toString.call(value) === "[object Object]";
}

function splitContent<Props extends object>(
  args: readonly (TextContent | Props | undefined)[],
): {
  readonly content: readonly RichTextItem[];
  readonly props: Props | undefined;
} {
  const last = args.at(-1);
  const hasProps = isPropsRecord(last);
  const values = hasProps ? args.slice(0, -1) : args;
  const content = values.flatMap((value) =>
    value === undefined
      ? []
      : Array.isArray(value)
        ? value
        : [value as RichTextItem],
  );
  return {
    content,
    props: hasProps ? last : undefined,
  };
}

export function deck(props: DeckAuthorProps = {}): DeckNode {
  return createDeckNode(props);
}

/** Creates an editable native SlideMaster container owning SlideLayouts. */
export function slideMaster(
  props: SlideMasterAuthorProps = {},
): SlideMasterNode {
  return createSlideMasterNode(props);
}

/** Creates an editable native SlideLayout node from inline props or a `defineSlideLayout` spec. */
export function slideLayout(
  propsOrDefinition: SlideLayoutDataProps | SlideLayoutDefinition = {},
): SlideLayoutNode {
  return createSlideLayoutNode(propsOrDefinition);
}

export function slide(props: SlideAuthorProps): SlideNode {
  return createSlideNode(props);
}

export function section(props: SectionAuthorProps): SectionNode {
  return createSectionNode(props);
}

export function box(props: BoxAuthorProps = {}): BoxNode {
  return createBoxNode(props);
}

/** Creates a native shape whose geometry, visuals, and rich text remain editable after attachment. */
export function shape(props: ShapeAuthorProps = {}): ShapeNode {
  return createShapeNode(props);
}

/** Creates a native image whose source bytes, crop, fit, and visuals remain editable after attachment. */
export function image(props: ImageAuthorProps): ImageNode {
  return createImageNode(props);
}

/** Creates a native group preserving member identity, local theme, and layout children. */
export function group(props: GroupAuthorProps = {}): GroupNode {
  return createGroupNode(props);
}

/** Creates a native table whose rows and cells remain editable after attachment. */
export function table(props: TableAuthorProps): TableNode {
  return createTableNode(props);
}

/** Creates a native table row; attach it to a table and add cells through its content operations. */
export function row(props: RowAuthorProps = {}): RowNode {
  return createRowNode(props);
}

/** Creates a rich-text table cell whose horizontal and vertical spans remain editable. */
export function cell(content?: TextContent, props?: CellAuthorProps): CellNode;
export function cell(...content: readonly TextContent[]): CellNode;
export function cell(
  ...args: readonly (TextContent | CellAuthorProps | undefined)[]
): CellNode {
  const { content, props } = splitContent<CellAuthorProps>(args);
  const node = createCellNode(props ?? {});
  node.content(...content);
  return node;
}

/** Creates a native chart whose data, options, and styles remain editable after attachment. */
export function chart(props: ChartAuthorProps): ChartNode {
  return createChartNode(props);
}

/** Creates a text box whose original reference remains editable after attachment. */
export function text(content?: TextContent, props?: TextAuthorProps): TextNode;
export function text(...content: readonly TextContent[]): TextNode;
export function text(
  ...args: readonly (TextContent | TextAuthorProps | undefined)[]
): TextNode {
  const { content, props } = splitContent(args);
  const node = createTextNode(props ?? {});
  node.content(...content);
  return node;
}

/**
 * Creates a page-affiliated notes container. The node is attached to a Slide; rich-text content
 * remains editable through the returned reference after attachment.
 */
export function notes(
  content?: TextContent,
  props?: NotesAuthorProps,
): NotesNode;
export function notes(...content: readonly TextContent[]): NotesNode;
export function notes(
  ...args: readonly (TextContent | NotesAuthorProps | undefined)[]
): NotesNode {
  const { content, props } = splitContent<NotesAuthorProps>(args);
  const node = createNotesNode(props ?? {});
  node.content(...content);
  return node;
}

export function paragraph(
  content?: ParagraphContent,
  props?: ParagraphAuthorProps,
): ParagraphNode;
export function paragraph(
  ...content: readonly ParagraphContent[]
): ParagraphNode;
export function paragraph(
  ...args: readonly (ParagraphContent | ParagraphAuthorProps | undefined)[]
): ParagraphNode {
  const { content, props } = splitContent(args);
  const node = createParagraphNode(props ?? {});
  node.content(...content);
  return node;
}

export function span(content?: SpanContent, props?: SpanAuthorProps): SpanNode;
export function span(...content: readonly SpanContent[]): SpanNode;
export function span(
  ...args: readonly (SpanContent | SpanAuthorProps | undefined)[]
): SpanNode {
  const { content, props } = splitContent(args);
  const node = createSpanNode(props ?? {});
  node.content(...content);
  return node;
}

/** Emits the native slide number field, resolved by the presentation application. */
export function slideNumber(
  props: SlideNumberAuthorProps = {},
): SlideNumberNode {
  return createSlideNumberNode(props);
}

/** Creates an orchestration scope whose parent chain and click step are derived by each build. */
export function motionScope(
  scope: MotionScopeHandle,
  props: Omit<MotionScopeAuthorProps, "scope"> = {},
): MotionScopeNode {
  return createMotionScopeNode({ ...props, scope });
}

/** Creates a click-step container; omitted n gives this Step its own independent trigger. */
export function step(props: MotionStepAuthorProps = {}): MotionStepNode {
  return createMotionStepNode({
    ...props,
    stepIdentity: createMotionStepIdentity(),
  });
}

export {
  createMotionScopeHandle,
  after,
  breathe,
  fade,
  fillColor,
  floatDown,
  floatUp,
  fly,
  repeat,
  spin,
  spinIn,
  stretch,
  wipe,
  zoom,
} from "./motion/motion-api.ts";
export { observeMotion } from "./motion/motion-query.ts";
export { blockArc, path, preset } from "./shared/geometry.ts";
export { linearGradient, radialGradient } from "./shared/visual.ts";
export { defineSlideLayout, font, transitions };
export type { SlideLayoutDataProps, SlideLayoutDefinition, SlideLayoutType };
export type {
  BreatheEffect,
  BreatheOptions,
  Easing,
  Effect,
  EffectOptions,
  EffectPreset,
  EmphasisEffect,
  EmphasisEffectDeclaration,
  EntranceEffect,
  EntranceEffectDeclaration,
  FillColorOptions,
  MotionCompletionState,
  MotionDirection,
  MotionEffectObservation,
  MotionGranularity,
  MotionMoment,
  MotionMomentOffset,
  MotionNativeProjection,
  MotionObservation,
  MotionProps,
  MotionScopeAuthorProps,
  MotionScopeHandle,
  MotionScopeMode,
  MotionScopeObservation,
  MotionStartDependency,
  MotionStepAuthorProps,
  MotionStepObservation,
  MotionTargetRange,
  MotionTiming,
  RepeatCount,
  RepeatedEffect,
  SpinOptions,
} from "./motion/motion-api.ts";
export type {
  EffectInstanceObservation,
  MotionCueObservation,
  MotionCueRef,
  MotionDecision,
  MotionDependencyObservation,
  MotionDifference,
  MotionEffectRef,
  MotionIntervalRef,
  MotionMembership,
  MotionMomentObservation,
  MotionMomentRef,
  MotionQuery,
  MotionScopeQueryMember,
  MotionScopeQueryObservation,
  MotionScopeRef,
  MotionStepQuery,
  MotionStepRef,
  MotionTime,
} from "./motion/motion-query.ts";
export type {
  CustomGeometry,
  Geometry,
  GeometryPath,
  PathCommand,
  PathOptions,
  PresetGeometry,
  PresetGeometryName,
} from "./shared/geometry.ts";
export type {
  ArrowType,
  LineEndStyle,
  ShapeVisualStyle,
} from "./shared/visual.ts";
export type {
  ImageSource,
  ImageVisualStyle,
} from "./image/image-declaration.ts";
export type {
  ParagraphStyle,
  TextAlignment,
  TextList,
  TextStyle,
} from "./shared/text-service.ts";
export type { ShapeTextStyle } from "./text/text-declaration.ts";
export type { MorphProps } from "./shared/morph-api.ts";
export type {
  FlipTransition,
  LinkTarget,
  MorphTransition,
  StandardTransition,
  StripsDirection,
  Transition,
  TransitionDirection,
  TransitionDuration,
  TransitionSpeed,
};
export type { NotesDeclarationProps } from "./notes/notes-declaration.ts";

export type {
  ChartAxis,
  ChartAxisCrossing,
  ChartAxisPosition,
  ChartAxisReferences,
  ChartCategoryAxis,
  ChartCategoryAxisReferences,
  ChartDataLabelOverride,
  ChartCustomErrorBars,
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
  ChartProps,
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
} from "./chart/chart-declaration.ts";
export type {
  TableCellStyle,
  TableColumn,
  TableRowStyle,
  TableSpacing,
  TableStyle,
  TableTextStyle,
  TableVerticalAlign,
} from "./table/table-declaration.ts";
