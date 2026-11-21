import type {
  ColorValue,
  LayoutStyle,
  LinkTarget,
  TextStyle,
} from "@pptxdk/core/capability";
import { CommonContainer, RichTextContainer } from "@pptxdk/core/document";

import type { MotionProps } from "../motion/motion-api.ts";
import type { MorphProps } from "../shared/morph-api.ts";
import type { FontReference, FontSize } from "../shared/text-service.ts";

/** Column width in px; Nfr allocates from remaining width, auto uses shared font measurement. */
export type TableColumn = number | `${number}fr` | "auto";

export type TableVerticalAlign = "top" | "middle" | "bottom";
export type TableSpacing =
  | number
  | readonly [vertical: number, horizontal: number]
  | readonly [top: number, right: number, bottom: number, left: number];

export type TableTextStyle = {
  readonly fontFamily?: FontReference;
  readonly fontSize?: FontSize;
  readonly fontWeight?: number | "normal" | "bold";
  readonly fontStyle?: "normal" | "italic";
  readonly color?: ColorValue;
  readonly lineHeight?: FontSize;
  readonly verticalAlign?: TableVerticalAlign;
  readonly letterSpacing?: TextStyle["letterSpacing"];
  readonly textTransform?: TextStyle["textTransform"];
  readonly textDecoration?: TextStyle["textDecoration"];
  readonly textShadow?: TextStyle["textShadow"];
  readonly lang?: string;
};

export type TableStyle = TableTextStyle & LayoutStyle;

export type TableRowStyle = TableTextStyle & {
  /** When unspecified, row height is determined from text measurement of unmerged cells in this row. */
  readonly height?: number;
  readonly backgroundColor?: ColorValue;
};

export type TableCellStyle = TableTextStyle & {
  readonly backgroundColor?: ColorValue;
  readonly borderColor?: ColorValue;
  readonly borderWidth?: number;
  readonly borderStyle?: "solid" | "dashed" | "dotted";
  readonly padding?: TableSpacing;
  readonly paddingTop?: number;
  readonly paddingRight?: number;
  readonly paddingBottom?: number;
  readonly paddingLeft?: number;
};

export type TableDeclarationProps = MotionProps &
  MorphProps & {
    readonly id?: string;
    readonly name?: string;
    readonly link?: LinkTarget;
    readonly x?: number;
    readonly y?: number;
    /** Provides allocatable width for fr columns; when omitted, all columns are sized to natural content width. */
    readonly width?: number;
    readonly columns: readonly TableColumn[];
    readonly style?: TableStyle;
  };

type TablePartDeclarationProps = {
  readonly id?: string;
};

export type RowDeclarationProps = TablePartDeclarationProps & {
  readonly style?: TableRowStyle;
};

export type CellDeclarationProps = TablePartDeclarationProps & {
  readonly colSpan?: number;
  readonly rowSpan?: number;
  readonly style?: TableCellStyle;
};

export type TableNode = CommonContainer<
  "table",
  TableDeclarationProps,
  NonNullable<TableDeclarationProps["style"]>
>;
export type RowNode = CommonContainer<
  "row",
  RowDeclarationProps,
  NonNullable<RowDeclarationProps["style"]>
>;
export type CellNode = RichTextContainer<
  "cell",
  CellDeclarationProps,
  NonNullable<CellDeclarationProps["style"]>
>;

/** Creates the editable native table shared by the command factory and the Table JSX element. */
export function createTableNode(props: TableDeclarationProps): TableNode {
  return new CommonContainer<
    "table",
    TableDeclarationProps,
    NonNullable<TableDeclarationProps["style"]>
  >("table", "rp:table", props);
}

/** Creates the editable row shared by command authoring and the Row JSX element. */
export function createRowNode(props: RowDeclarationProps): RowNode {
  return new CommonContainer<
    "row",
    RowDeclarationProps,
    NonNullable<RowDeclarationProps["style"]>
  >("row", "rp:row", props);
}

/** Creates the editable rich-text cell shared by command authoring and the Cell JSX element. */
export function createCellNode(props: CellDeclarationProps): CellNode {
  return new RichTextContainer<
    "cell",
    CellDeclarationProps,
    NonNullable<CellDeclarationProps["style"]>
  >("cell", "rp:cell", props);
}
