import { CommonContainer, RichTextContainer } from "@pptxdk/core/document";
import type { z } from "zod";

import type { MotionProps } from "../motion/motion-api.ts";
import type { MorphProps } from "../shared/morph-api.ts";
import type {
  CellDeclarationProps,
  RowDeclarationProps,
  TableCellStyle,
  TableColumn,
  tableDeclarationPropsSchema,
  TableRowStyle,
  TableSpacing,
  TableStyle,
  TableTextStyle,
  TableVerticalAlign,
} from "./table-schemas.ts";

export type {
  CellDeclarationProps,
  RowDeclarationProps,
  TableCellStyle,
  TableColumn,
  TableRowStyle,
  TableSpacing,
  TableStyle,
  TableTextStyle,
  TableVerticalAlign,
};

type TableDeclarationDataProps = Omit<
  z.infer<typeof tableDeclarationPropsSchema>,
  "enter" | "emphasis" | "motion" | "morphKey"
>;

export type TableDeclarationProps = MotionProps &
  MorphProps &
  TableDeclarationDataProps;

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
  >("table", "table", props);
}

/** Creates the editable row shared by command authoring and the Row JSX element. */
export function createRowNode(props: RowDeclarationProps): RowNode {
  return new CommonContainer<
    "row",
    RowDeclarationProps,
    NonNullable<RowDeclarationProps["style"]>
  >("row", "row", props);
}

/** Creates the editable rich-text cell shared by command authoring and the Cell JSX element. */
export function createCellNode(props: CellDeclarationProps): CellNode {
  return new RichTextContainer<
    "cell",
    CellDeclarationProps,
    NonNullable<CellDeclarationProps["style"]>
  >("cell", "cell", props);
}
