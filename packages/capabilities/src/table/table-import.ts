import {
  matchLayoutPlaceholderKey,
  type ColorValue,
  type NativePlaceholderRef,
} from "@pptxdk/core/capability";
import {
  A_NS,
  P_NS,
  TABLE_NS_URI,
  emuToPx,
  type ImportDiagnostic,
  type ImportLocation,
  type ShapeTreeImportContext,
  type ShapeTreeImportHandler,
  type ShapeTreeImportOutcome,
  type XmlElementCursor,
} from "@pptxdk/import";

import { readClickLink, type ClickLinkReadResult } from "../shared/identity.ts";
import {
  isPlaceholderGraphicFrameElement,
  readNvPrPlaceholder,
} from "../shared/placeholder-import.ts";
import {
  readDrawingColor,
  readLineBorder,
  readTransformXml,
  type ImportedLineBorder,
} from "../shared/visual.ts";
import { populateRichTextNode, readRichTextBody } from "../text/text-import.ts";
import {
  createCellNode,
  createRowNode,
  createTableNode,
  type CellDeclarationProps,
  type CellNode,
  type RowNode,
  type TableCellStyle,
  type TableDeclarationProps,
  type TableVerticalAlign,
} from "./table-declaration.ts";
import {
  graphicFrameLocksXmlSchema,
  nonVisualGraphicFrameIdSchema,
  tblCellPrXmlSchema,
  tblCellXmlSchema,
  tblGridColXmlSchema,
  tblPrXmlSchema,
  tblRowXmlSchema,
} from "./table-schemas.ts";

type XmlElement = XmlElementCursor["element"];

const anchorToVerticalAlign: Readonly<
  Record<"t" | "ctr" | "b", TableVerticalAlign>
> = {
  t: "top",
  ctr: "middle",
  b: "bottom",
};

function directElementChild(
  parent: XmlElement,
  namespaceUri: string,
  localName: string,
): XmlElement | undefined {
  for (
    let child = parent.firstChild;
    child !== null;
    child = child.nextSibling
  ) {
    if (child.nodeType !== 1) continue;
    const el = child as XmlElement;
    if (
      (el.namespaceURI ?? "") === namespaceUri &&
      (el.localName ?? el.tagName) === localName
    ) {
      return el;
    }
  }
  return undefined;
}

/** Side-effect-free predicate matching `<p:graphicFrame>` elements containing DrawingML tables. */
export function isTableGraphicFrame(cursor: XmlElementCursor): boolean {
  const graphic = directElementChild(cursor.element, A_NS, "graphic");
  if (graphic === undefined) return false;
  const graphicData = directElementChild(graphic, A_NS, "graphicData");
  if (graphicData === undefined) return false;
  return graphicData.getAttribute("uri") === TABLE_NS_URI;
}

type RawGridCell = {
  readonly cursor: XmlElementCursor;
  readonly gridSpan?: number;
  readonly rowSpan?: number;
  readonly hMerge: boolean;
  readonly vMerge: boolean;
};

function consumeContinuationCell(tcCursor: XmlElementCursor): void {
  const txBody = tcCursor.child(A_NS, "txBody");
  if (txBody !== undefined) {
    txBody.ignoreEmptyChild(A_NS, "bodyPr");
    txBody.ignoreEmptyChild(A_NS, "lstStyle");
    for (const p of txBody.children(A_NS, "p")) {
      const endParaRPr = p.child(A_NS, "endParaRPr");
      if (endParaRPr !== undefined) {
        endParaRPr.consumeAttr("lang");
        endParaRPr.consumeAttr("dirty");
        endParaRPr.consumeAttr("smtClean");
      }
      p.suppressUnconsumed();
    }
    txBody.suppressUnconsumed();
  }
  tcCursor.ignoreEmptyChild(A_NS, "tcPr");
  tcCursor.suppressUnconsumed();
}

function readCellProperties(
  tcPr: XmlElementCursor | undefined,
  cellLocation: ImportLocation,
  diagnostics: ImportDiagnostic[],
): Partial<TableCellStyle> {
  if (tcPr === undefined) return {};

  const rawAttrs = tcPr.attrsObject([
    "marL",
    "marR",
    "marT",
    "marB",
    "anchor",
    "vert",
    "horzOverflow",
    "anchorCtr",
  ]);
  const parsedAttrs = tblCellPrXmlSchema.safeParse(rawAttrs);
  let paddingTop: number | undefined;
  let paddingRight: number | undefined;
  let paddingBottom: number | undefined;
  let paddingLeft: number | undefined;
  let verticalAlign: TableVerticalAlign | undefined;

  if (!parsedAttrs.success) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "recovered",
      location: { ...cellLocation, xpath: tcPr.xpath },
      attributeOrElement: "a:tcPr",
      rawValue: JSON.stringify(rawAttrs),
    });
  } else {
    const attrs = parsedAttrs.data;
    if (attrs.marT !== undefined && attrs.marT !== 0) {
      paddingTop = emuToPx(attrs.marT);
    }
    if (attrs.marR !== undefined && attrs.marR !== 0) {
      paddingRight = emuToPx(attrs.marR);
    }
    if (attrs.marB !== undefined && attrs.marB !== 0) {
      paddingBottom = emuToPx(attrs.marB);
    }
    if (attrs.marL !== undefined && attrs.marL !== 0) {
      paddingLeft = emuToPx(attrs.marL);
    }
    if (attrs.anchor !== undefined) {
      verticalAlign = anchorToVerticalAlign[attrs.anchor];
    }
    if (attrs.vert !== undefined && attrs.vert !== "horz") {
      diagnostics.push({
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: { ...cellLocation, xpath: `${tcPr.xpath}/@vert` },
        namespaceUri: "",
        localName: "vert",
        mechanismHint: `table-cell-vert:${attrs.vert}`,
      });
    }
  }

  let border: ImportedLineBorder | undefined;
  for (const side of ["lnL", "lnR", "lnT", "lnB"] as const) {
    const ln = tcPr.child(A_NS, side);
    if (ln === undefined) continue;
    const parsedBorder = readLineBorder(ln, cellLocation, diagnostics);
    if (border === undefined && parsedBorder.borderColor !== undefined) {
      border = parsedBorder;
    }
  }

  tcPr.ignoreEmptyChild(A_NS, "noFill");
  let backgroundColor: ColorValue | undefined;
  const solidFill = tcPr.child(A_NS, "solidFill");
  if (solidFill !== undefined) {
    backgroundColor = readDrawingColor(solidFill, cellLocation, diagnostics);
  }

  return {
    ...(backgroundColor === undefined ? {} : { backgroundColor }),
    ...(border?.borderColor === undefined
      ? {}
      : { borderColor: border.borderColor }),
    ...(border?.borderWidth === undefined
      ? {}
      : { borderWidth: border.borderWidth }),
    ...(border?.borderStyle === undefined
      ? {}
      : { borderStyle: border.borderStyle }),
    ...(paddingTop === undefined ? {} : { paddingTop }),
    ...(paddingRight === undefined ? {} : { paddingRight }),
    ...(paddingBottom === undefined ? {} : { paddingBottom }),
    ...(paddingLeft === undefined ? {} : { paddingLeft }),
    ...(verticalAlign === undefined ? {} : { verticalAlign }),
  };
}

function readAnchorCell(
  tcCursor: XmlElementCursor,
  colSpan: number,
  rowSpan: number,
  cellLocation: ImportLocation,
  part: string,
  slideIndex: number,
  diagnostics: ImportDiagnostic[],
): CellNode {
  const txBody = tcCursor.child(A_NS, "txBody");
  const tcPr = tcCursor.child(A_NS, "tcPr");
  const cellPropsStyle = readCellProperties(tcPr, cellLocation, diagnostics);

  let hoistedRunStyle: Partial<TableCellStyle> = {};
  let singleTextContent: string | undefined;
  let richParagraphs:
    | ReturnType<typeof readRichTextBody>["importedParagraphs"]
    | undefined;

  if (txBody !== undefined) {
    const { importedParagraphs } = readRichTextBody(
      txBody,
      part,
      slideIndex,
      diagnostics,
    );
    const singleParagraph =
      importedParagraphs.length === 1 ? importedParagraphs[0] : undefined;
    const hasParagraphLevelStyle =
      singleParagraph !== undefined &&
      Object.keys(singleParagraph.style).length > 0;
    const singleTextRun =
      singleParagraph !== undefined &&
      !hasParagraphLevelStyle &&
      singleParagraph.runs.length === 1 &&
      singleParagraph.runs[0]!.kind === "text"
        ? singleParagraph.runs[0]
        : undefined;
    const isSingleEmptyParagraph =
      singleParagraph !== undefined &&
      !hasParagraphLevelStyle &&
      singleParagraph.runs.length === 0;

    if (singleTextRun !== undefined) {
      const {
        fontFamily,
        fontSize,
        fontWeight,
        fontStyle,
        color,
        letterSpacing,
        textTransform,
        textDecoration,
        textShadow,
        lang,
      } = singleTextRun.style;
      hoistedRunStyle = {
        ...(fontFamily === undefined ? {} : { fontFamily }),
        ...(fontSize === undefined ? {} : { fontSize }),
        ...(fontWeight === undefined ? {} : { fontWeight }),
        ...(fontStyle === undefined ? {} : { fontStyle }),
        ...(color === undefined ? {} : { color }),
        ...(letterSpacing === undefined ? {} : { letterSpacing }),
        ...(textTransform === undefined ? {} : { textTransform }),
        ...(textDecoration === undefined ? {} : { textDecoration }),
        ...(textShadow === undefined ? {} : { textShadow }),
        ...(lang === undefined ? {} : { lang }),
      };
      singleTextContent = singleTextRun.text;
    } else if (!isSingleEmptyParagraph && importedParagraphs.length > 0) {
      richParagraphs = importedParagraphs;
    }
  }

  const mergedStyle: TableCellStyle = {
    ...cellPropsStyle,
    ...hoistedRunStyle,
  };

  const cellProps: CellDeclarationProps = {
    ...(colSpan > 1 ? { colSpan } : {}),
    ...(rowSpan > 1 ? { rowSpan } : {}),
    ...(Object.keys(mergedStyle).length === 0 ? {} : { style: mergedStyle }),
  };
  const cellNode = createCellNode(cellProps);
  if (singleTextContent !== undefined) {
    cellNode.content(singleTextContent);
  } else if (richParagraphs !== undefined) {
    populateRichTextNode(cellNode, richParagraphs);
  }
  return cellNode;
}

/**
 * Reads a Transitional `<p:graphicFrame>` containing a DrawingML `<a:tbl>` into an editable `TableNode`
 * preserving solved column widths, row heights, 2D merge ranges, cell borders/fills/padding, and rich text.
 */
export function readTableGraphicFrame(
  context: ShapeTreeImportContext,
): ShapeTreeImportOutcome {
  const { cursor, part, slideIndex, layoutPlaceholders } = context;
  const diagnostics: ImportDiagnostic[] = [];

  const nvGraphicFramePr = cursor.child(P_NS, "nvGraphicFramePr");
  if (nvGraphicFramePr === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part, xpath: cursor.xpath, slideIndex },
    });
    return { diagnostics };
  }

  let objectId: number | undefined;
  let objectName: string | undefined;
  let clickResult: ClickLinkReadResult | undefined;

  const cNvPr = nvGraphicFramePr.child(P_NS, "cNvPr");
  if (cNvPr !== undefined) {
    const rawCNvPr = cNvPr.attrsObject(["id", "name"]);
    const parsedCNvPr = nonVisualGraphicFrameIdSchema.safeParse(rawCNvPr);
    if (!parsedCNvPr.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: { part, xpath: cNvPr.xpath, slideIndex },
        attributeOrElement: "p:cNvPr",
        rawValue: JSON.stringify(rawCNvPr),
      });
    } else {
      objectId = parsedCNvPr.data.id;
      if (
        parsedCNvPr.data.name !== undefined &&
        parsedCNvPr.data.name.trim() !== ""
      ) {
        objectName = parsedCNvPr.data.name;
      }
    }
    const hlinkClick = cNvPr.child(A_NS, "hlinkClick");
    if (hlinkClick !== undefined) {
      clickResult = readClickLink(hlinkClick, context, {
        ...(objectId === undefined ? {} : { objectId }),
        ...(objectName === undefined ? {} : { objectName }),
      });
      diagnostics.push(...clickResult.diagnostics);
    }
    cNvPr.ignoreEmptyChild(A_NS, "extLst");
  }

  const cNvGraphicFramePr = nvGraphicFramePr.child(P_NS, "cNvGraphicFramePr");
  if (cNvGraphicFramePr !== undefined) {
    const locks = cNvGraphicFramePr.child(A_NS, "graphicFrameLocks");
    if (locks !== undefined) {
      const rawLocks = locks.attrsObject([
        "noGrp",
        "noDrilldown",
        "noSelect",
        "noChangeAspect",
        "noMove",
        "noResize",
      ]);
      graphicFrameLocksXmlSchema.safeParse(rawLocks);
    }
  }
  const tableLocation: ImportLocation = {
    part,
    xpath: cursor.xpath,
    slideIndex,
    ...(objectId === undefined ? {} : { objectId }),
    ...(objectName === undefined ? {} : { objectName }),
  };

  const nvPr = nvGraphicFramePr.child(P_NS, "nvPr");
  const phResult = readNvPrPlaceholder(nvPr, tableLocation, diagnostics);
  if (phResult.kind === "error") {
    return { diagnostics };
  }
  const nativeRef: NativePlaceholderRef | undefined =
    phResult.kind === "ok" ? phResult.nativeRef : undefined;
  const matchedKey =
    nativeRef === undefined
      ? undefined
      : matchLayoutPlaceholderKey(layoutPlaceholders, {
          type: nativeRef.type,
          idx: nativeRef.idx,
        });

  const xfrmCursor = cursor.child(P_NS, "xfrm") ?? cursor.child(A_NS, "xfrm");
  if (xfrmCursor === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: tableLocation,
    });
    return { diagnostics };
  }
  const transform = readTransformXml(xfrmCursor, tableLocation, diagnostics);
  if (transform === undefined) {
    return { diagnostics };
  }

  const graphic = cursor.child(A_NS, "graphic");
  const graphicData = graphic?.child(A_NS, "graphicData");
  if (graphicData !== undefined) {
    graphicData.consumeAttr("uri");
  }
  const tbl = graphicData?.child(A_NS, "tbl");
  if (tbl === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: tableLocation,
    });
    return { diagnostics };
  }

  const tblPr = tbl.child(A_NS, "tblPr");
  if (tblPr !== undefined) {
    const rawTblPr = tblPr.attrsObject([
      "firstRow",
      "firstCol",
      "lastRow",
      "lastCol",
      "bandRow",
      "bandCol",
      "rtl",
    ]);
    const parsedTblPr = tblPrXmlSchema.safeParse(rawTblPr);
    if (!parsedTblPr.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: { ...tableLocation, xpath: tblPr.xpath },
        attributeOrElement: "a:tblPr",
        rawValue: JSON.stringify(rawTblPr),
      });
    }
    tblPr.ignoreEmptyChild(A_NS, "noFill");
    const tableStyleId = tblPr.child(A_NS, "tableStyleId");
    if (tableStyleId !== undefined) {
      const styleIdText = tableStyleId.textContent.trim();
      diagnostics.push({
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: { ...tableLocation, xpath: tableStyleId.xpath },
        namespaceUri: A_NS,
        localName: "tableStyleId",
        ...(styleIdText === "" ? {} : { mechanismHint: styleIdText }),
      });
      tableStyleId.suppressUnconsumed();
    }
  }

  const tblGrid = tbl.child(A_NS, "tblGrid");
  if (tblGrid === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...tableLocation, xpath: tbl.xpath },
    });
    return { diagnostics };
  }

  const gridCols = tblGrid.children(A_NS, "gridCol");
  if (gridCols.length === 0) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...tableLocation, xpath: tblGrid.xpath },
    });
    return { diagnostics };
  }

  const columnWidths: number[] = [];
  for (const gridCol of gridCols) {
    const rawCol = gridCol.attrsObject(["w"]);
    const parsedCol = tblGridColXmlSchema.safeParse(rawCol);
    if (!parsedCol.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "omitted",
        location: { ...tableLocation, xpath: gridCol.xpath },
        attributeOrElement: "a:gridCol/@w",
        rawValue: rawCol["w"] ?? "",
      });
      return { diagnostics };
    }
    columnWidths.push(emuToPx(parsedCol.data.w));
  }

  const trCursors = tbl.children(A_NS, "tr");
  if (trCursors.length === 0) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...tableLocation, xpath: tbl.xpath },
    });
    return { diagnostics };
  }

  const rowHeights: number[] = [];
  const rawGrid: RawGridCell[][] = [];
  for (const trCursor of trCursors) {
    const rawRow = trCursor.attrsObject(["h"]);
    const parsedRow = tblRowXmlSchema.safeParse(rawRow);
    if (!parsedRow.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "omitted",
        location: { ...tableLocation, xpath: trCursor.xpath },
        attributeOrElement: "a:tr/@h",
        rawValue: rawRow["h"] ?? "",
      });
      return { diagnostics };
    }
    rowHeights.push(emuToPx(parsedRow.data.h));

    const tcCursors = trCursor.children(A_NS, "tc");
    if (tcCursors.length !== columnWidths.length) {
      diagnostics.push({
        code: "corrupt-part",
        category: "corrupt-input",
        disposition: "omitted",
        location: { ...tableLocation, xpath: trCursor.xpath },
      });
      return { diagnostics };
    }

    const rowCells: RawGridCell[] = [];
    for (const tcCursor of tcCursors) {
      const rawCell = tcCursor.attrsObject([
        "gridSpan",
        "rowSpan",
        "hMerge",
        "vMerge",
      ]);
      const parsedCell = tblCellXmlSchema.safeParse(rawCell);
      if (!parsedCell.success) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "recovered",
          location: { ...tableLocation, xpath: tcCursor.xpath },
          attributeOrElement: "a:tc",
          rawValue: JSON.stringify(rawCell),
        });
        rowCells.push({ cursor: tcCursor, hMerge: false, vMerge: false });
      } else {
        rowCells.push({
          cursor: tcCursor,
          ...(parsedCell.data.gridSpan === undefined
            ? {}
            : { gridSpan: parsedCell.data.gridSpan }),
          ...(parsedCell.data.rowSpan === undefined
            ? {}
            : { rowSpan: parsedCell.data.rowSpan }),
          hMerge: parsedCell.data.hMerge ?? false,
          vMerge: parsedCell.data.vMerge ?? false,
        });
      }
    }
    rawGrid.push(rowCells);
  }

  const numRows = rawGrid.length;
  const numCols = columnWidths.length;
  const occupied = Array.from({ length: numRows }, () =>
    Array.from({ length: numCols }, () => false),
  );
  const rowNodes: RowNode[] = [];

  for (let r = 0; r < numRows; r += 1) {
    const rowNode = createRowNode({ style: { height: rowHeights[r]! } });
    for (let c = 0; c < numCols; c += 1) {
      const rawCell = rawGrid[r]![c]!;
      if (occupied[r]![c]) {
        consumeContinuationCell(rawCell.cursor);
        continue;
      }

      if (rawCell.hMerge || rawCell.vMerge) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "recovered",
          location: { ...tableLocation, xpath: rawCell.cursor.xpath },
          attributeOrElement: rawCell.hMerge ? "a:tc/@hMerge" : "a:tc/@vMerge",
          rawValue: "1",
        });
      }

      let contiguousH = 0;
      while (
        c + 1 + contiguousH < numCols &&
        rawGrid[r]![c + 1 + contiguousH]!.hMerge &&
        !rawGrid[r]![c + 1 + contiguousH]!.vMerge
      ) {
        contiguousH += 1;
      }
      let contiguousV = 0;
      while (
        r + 1 + contiguousV < numRows &&
        rawGrid[r + 1 + contiguousV]![c]!.vMerge &&
        !rawGrid[r + 1 + contiguousV]![c]!.hMerge
      ) {
        contiguousV += 1;
      }

      const maxColSpan = numCols - c;
      const maxRowSpan = numRows - r;
      if (
        (rawCell.gridSpan !== undefined && rawCell.gridSpan > maxColSpan) ||
        (rawCell.rowSpan !== undefined && rawCell.rowSpan > maxRowSpan)
      ) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "recovered",
          location: { ...tableLocation, xpath: rawCell.cursor.xpath },
          attributeOrElement: "a:tc",
        });
      }

      const colSpan = Math.min(
        maxColSpan,
        Math.max(rawCell.gridSpan ?? 1, contiguousH + 1),
      );
      const rowSpan = Math.min(
        maxRowSpan,
        Math.max(rawCell.rowSpan ?? 1, contiguousV + 1),
      );

      for (let rr = r; rr < r + rowSpan; rr += 1) {
        for (let cc = c; cc < c + colSpan; cc += 1) {
          occupied[rr]![cc] = true;
        }
      }

      const cellNode = readAnchorCell(
        rawCell.cursor,
        colSpan,
        rowSpan,
        { ...tableLocation, xpath: rawCell.cursor.xpath },
        part,
        slideIndex,
        diagnostics,
      );
      rowNode.append(cellNode);
    }
    rowNodes.push(rowNode);
  }

  const columns = columnWidths as unknown as readonly [number, ...number[]];
  const tableProps: TableDeclarationProps = {
    ...(objectName === undefined ? {} : { name: objectName }),
    ...(matchedKey === undefined ? {} : { placeholder: matchedKey }),
    ...(nativeRef === undefined ? {} : { nativePlaceholder: nativeRef }),
    ...(clickResult?.link === undefined ? {} : { link: clickResult.link }),
    x: transform.x,
    y: transform.y,
    columns,
  };
  const tableNode = createTableNode(tableProps);
  tableNode.append(...rowNodes);

  if (
    clickResult?.deferredSlideLink !== undefined &&
    context.registerDeferredSlideLink !== undefined
  ) {
    const deferred = clickResult.deferredSlideLink;
    context.registerDeferredSlideLink({
      ...deferred,
      apply(targetSlideId: string) {
        tableNode.set({ link: { id: targetSlideId } });
      },
    });
  }

  return { node: tableNode, diagnostics };
}

/** Shape-tree import handler wired into `tableMechanism` for free DrawingML `<a:tbl>` graphicFrames. */
export const tableImportHandler: ShapeTreeImportHandler = {
  target: "shape-tree",
  localName: "graphicFrame",
  matches: (cursor) =>
    isTableGraphicFrame(cursor) && !isPlaceholderGraphicFrameElement(cursor),
  read: readTableGraphicFrame,
};

/** Shape-tree import handler wired into `tablePlaceholderMechanism` for placeholder-bound `<a:tbl>` graphicFrames. */
export const tablePlaceholderImportHandler: ShapeTreeImportHandler = {
  target: "shape-tree",
  localName: "graphicFrame",
  matches: (cursor) =>
    isTableGraphicFrame(cursor) && isPlaceholderGraphicFrameElement(cursor),
  read: readTableGraphicFrame,
};
