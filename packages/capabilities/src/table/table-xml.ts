import type { ResolvedColor } from "@pptxdk/core/capability";
import { pxToEmu, resolvedColorXml, xmlEscape } from "@pptxdk/emit";

const tableUri = "http://schemas.openxmlformats.org/drawingml/2006/table";

/** Cell padding in pixels, already resolved by the table capability parser. */
export type TableCellPadding = {
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
};

/** Resolved cell border; the table capability owns selection and validation of these values. */
export type TableCellBorder = {
  readonly color: ResolvedColor;
  readonly width: number;
  readonly dash: "solid" | "dash" | "dot";
};

export type TableCellAlignment = "top" | "middle" | "bottom";

function fillXml(value: ResolvedColor | undefined): string {
  return value === undefined
    ? "<a:noFill/>"
    : `<a:solidFill>${resolvedColorXml(value)}</a:solidFill>`;
}

function lineXml(
  side: "L" | "R" | "T" | "B",
  border: TableCellBorder | undefined,
): string {
  if (border === undefined) return `<a:ln${side}><a:noFill/></a:ln${side}>`;
  const width = Math.max(1, pxToEmu(border.width));
  const dash =
    border.dash === "solid" ? "" : `<a:prstDash val="${border.dash}"/>`;
  return `<a:ln${side} w="${width}"><a:solidFill>${resolvedColorXml(border.color)}</a:solidFill>${dash}</a:ln${side}>`;
}

/** Writes the resolved cell margins, vertical anchor, borders, and fill into a:tcPr. */
export function tableCellPropertiesXml(
  padding: TableCellPadding,
  vertical: TableCellAlignment,
  border: TableCellBorder | undefined,
  fill: ResolvedColor | undefined,
): string {
  const lines = (["L", "R", "T", "B"] as const)
    .map((side) => lineXml(side, border))
    .join("");
  const anchor =
    vertical === "middle" ? "ctr" : vertical === "bottom" ? "b" : "t";
  return `<a:tcPr marL="${pxToEmu(padding.left)}" marR="${pxToEmu(padding.right)}" marT="${pxToEmu(padding.top)}" marB="${pxToEmu(padding.bottom)}" anchor="${anchor}">${lines}${fillXml(fill)}</a:tcPr>`;
}

/** Writes an empty continuation cell covered by a merged cell. */
export function emptyTableCellXml(lang?: string): string {
  const langAttr = lang !== undefined ? ` lang="${xmlEscape(lang)}"` : "";
  return `<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr${langAttr}/></a:p></a:txBody><a:tcPr/>`;
}

/** Writes the native graphicFrame and a:tbl once the capability has resolved tracks and cell bodies. */
export function tableFrameXml(
  id: number,
  escapedName: string,
  x: number,
  y: number,
  widths: readonly number[],
  heights: readonly number[],
  rows: readonly string[],
  linkXml: string,
): string {
  const columnXml = widths
    .map((width) => `<a:gridCol w="${pxToEmu(width)}"/>`)
    .join("");
  const rowXml = rows
    .map((cells, row) => `<a:tr h="${pxToEmu(heights[row]!)}">${cells}</a:tr>`)
    .join("");
  const width = widths.reduce((sum, value) => sum + pxToEmu(value), 0);
  const height = heights.reduce((sum, value) => sum + pxToEmu(value), 0);
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${escapedName}">${linkXml}</p:cNvPr><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${pxToEmu(x)}" y="${pxToEmu(y)}"/><a:ext cx="${width}" cy="${height}"/></p:xfrm><a:graphic><a:graphicData uri="${tableUri}"><a:tbl><a:tblPr/><a:tblGrid>${columnXml}</a:tblGrid>${rowXml}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}
