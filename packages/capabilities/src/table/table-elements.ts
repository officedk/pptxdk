import { hostNodeCreator } from "@pptxdk/core/host";
import { createElement, type ReactNode } from "react";

import { motionElement } from "../motion/motion-elements.ts";
import {
  createCellNode,
  createRowNode,
  createTableNode,
  type CellDeclarationProps,
  type RowDeclarationProps,
  type TableDeclarationProps,
} from "./table-declaration.ts";

export type TableProps = TableDeclarationProps & {
  readonly children?: ReactNode;
};

export type RowProps = RowDeclarationProps & {
  readonly children?: ReactNode;
};

export type CellProps = CellDeclarationProps & {
  readonly children?: ReactNode;
};

/** Declares an editable native table in the React root handled by the table module. */
export function Table(props: TableProps): ReactNode {
  return motionElement("rp:table", {
    ...props,
    [hostNodeCreator]: createTableNode,
  });
}

/** Declares a row of a native table; calculated from actual measurement of row content when height is omitted. */
export function Row(props: RowProps): ReactNode {
  return createElement(
    "rp:row",
    { ...props, [hostNodeCreator]: createRowNode },
    props.children,
  );
}

/** Declares table cell content and horizontal / vertical merge spans. */
export function Cell(props: CellProps): ReactNode {
  return createElement(
    "rp:cell",
    { ...props, [hostNodeCreator]: createCellNode },
    props.children,
  );
}
