import type {
  LayoutStyle,
  LinkTarget,
  PaddingStyle,
  SlideBackground,
  TextStyle,
  ThemeInput,
  ThemeScope,
  VisualStyle,
} from "@react-pptx/core/capability";
import {
  CommonContainer,
  CommonNode,
  createDeckNode,
  createSectionNode,
  createSlideNode,
  RichTextContainer,
  type RichTextItem,
} from "@react-pptx/core/document";

import {
  boxCapability,
  createBoxNode,
  type BoxDeclarationProps,
} from "./box/box.ts";
import {
  chartCapability,
  createChartNode,
  type ChartNode,
  type ChartProps,
} from "./chart/chart.ts";
import { font } from "./font/font.ts";
import {
  createGroupNode,
  groupCapability,
  type GroupDeclarationProps,
} from "./group/group.ts";
import {
  createImageNode,
  imageCapability,
  type ImageDeclarationProps,
} from "./image/image.ts";
import {
  createMotionScopeHandle,
  createMotionScopeNode,
  createMotionStepIdentity,
  createMotionStepNode,
  type MotionScopeAuthorProps,
  type MotionScopeHandle,
  type MotionStepAuthorProps,
} from "./motion/motion-api.ts";
import { motionCapability } from "./motion/motion.ts";
import {
  createNotesNode,
  type NotesDeclarationProps,
} from "./notes/notes-declaration.ts";
import { notesCapability } from "./notes/notes.ts";
import {
  createShapeNode,
  shapeCapability,
  type ShapeDeclarationProps,
} from "./shape/shape.ts";
import { slideBackgroundCapability } from "./slide-background/slide-background.ts";
import {
  createCellNode,
  createRowNode,
  createTableNode,
  tableCapability,
  type CellDeclarationProps,
  type CellNode,
  type RowDeclarationProps,
  type RowNode,
  type TableDeclarationProps,
  type TableNode,
} from "./table/table.ts";
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
import { textCapability } from "./text/text.ts";
import {
  transitionCapability,
  type Transition,
} from "./transition/transition.ts";

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

export type SlideAuthorProps = {
  readonly id: string;
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

export type DeckNode = CommonContainer<"deck", DeckAuthorProps, TextStyle>;
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

/** Baseline capabilities enabled by the public document build; factories never register capabilities. */
export const documentCapabilities = Object.freeze([
  boxCapability,
  chartCapability,
  groupCapability,
  imageCapability,
  motionCapability,
  notesCapability,
  shapeCapability,
  slideBackgroundCapability,
  tableCapability,
  textCapability,
  transitionCapability,
] as const);

export { createMotionScopeHandle, font, notesCapability };
export type {
  MotionScopeAuthorProps,
  MotionScopeHandle,
  MotionStepAuthorProps,
} from "./motion/motion-api.ts";
export type { LinkTarget, Transition };
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
} from "./chart/chart.ts";
export type {
  TableCellStyle,
  TableColumn,
  TableRowStyle,
  TableSpacing,
  TableStyle,
  TableTextStyle,
  TableVerticalAlign,
} from "./table/table.ts";
