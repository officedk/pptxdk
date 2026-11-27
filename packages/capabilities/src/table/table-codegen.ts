import { defineObjectCodegen, type CodegenContext } from "@pptxdk/codegen";
import { CommonNode } from "@pptxdk/core/document";

import {
  blockRichTextChildKinds,
  collectRichTextNodeProps,
  createRichTextCodegenHandler,
} from "../text/text-codegen.ts";
import {
  cellDeclarationPropsSchema,
  rowDeclarationPropsSchema,
  rowStyleKeys,
  tableCellStyleSchema,
  tableColumnSchema,
  tableDeclarationPropsSchema,
  tableRowStyleSchema,
  tableStyleKeys,
  tableStyleSchema,
  tableTextStyleKeys,
} from "./table-schemas.ts";

const tableAllowedPropKeys = new Set(
  Object.keys(tableDeclarationPropsSchema.unwrap().shape),
);
const rowAllowedPropKeys = new Set(
  Object.keys(rowDeclarationPropsSchema.unwrap().shape),
);
const cellAllowedPropKeys = new Set(
  Object.keys(cellDeclarationPropsSchema.unwrap().shape),
);

const tableCleanStyleKeys = new Set(tableStyleKeys);
const rowCleanStyleKeys = new Set([...tableTextStyleKeys, ...rowStyleKeys]);
const cellCleanStyleKeys = new Set([
  ...tableTextStyleKeys,
  ...Object.keys(tableCellStyleSchema.unwrap().shape),
]);

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return false;
  }
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

function validateTableNativePlaceholder(
  node: CommonNode,
  ctx: CodegenContext,
): void {
  const nativePh = node.props["nativePlaceholder"];
  if (
    isPlainRecord(nativePh) &&
    Object.hasOwn(nativePh, "sz") &&
    nativePh["sz"] !== undefined
  ) {
    ctx.fail(
      "unexpressible-value",
      "Table.nativePlaceholder.sz is no longer supported; use size",
      "@nativePlaceholder.sz",
    );
  }
}

function validateTableLayerStyle(
  style: unknown,
  layerLabel: "Table" | "Row" | "Cell",
  allowedKeys: ReadonlySet<string>,
  ctx: CodegenContext,
): void {
  if (style === undefined) return;
  if (!isPlainRecord(style)) {
    ctx.fail(
      "unexpressible-value",
      `${layerLabel}.style must be a plain object`,
      "@style",
    );
  }
  for (const [key, value] of Object.entries(style)) {
    if (value === undefined) continue;
    if (!allowedKeys.has(key)) {
      ctx.fail(
        "unexpressible-value",
        `${layerLabel}.style does not support property ${JSON.stringify(key)}`,
        `@style.${key}`,
      );
    }
  }
  const schema =
    layerLabel === "Table"
      ? tableStyleSchema
      : layerLabel === "Row"
        ? tableRowStyleSchema
        : tableCellStyleSchema;
  const parsed = schema.safeParse(style);
  if (!parsed.success) {
    const issue = parsed.error.issues[0];
    const issuePath =
      issue !== undefined && issue.path.length > 0
        ? `@style.${issue.path.join(".")}`
        : "@style";
    ctx.fail(
      "unexpressible-value",
      issue?.message ?? `${layerLabel}.style cannot be expressed`,
      issuePath,
    );
  }
}

function validateTableColumns(node: CommonNode, ctx: CodegenContext): void {
  const columns = node.props["columns"];
  if (!Array.isArray(columns) || columns.length === 0) {
    ctx.fail(
      "invalid-document",
      "table declaration requires a non-empty columns array",
      "@columns",
    );
  }
  for (const [index, column] of columns.entries()) {
    if (typeof column === "number" && !Number.isFinite(column)) {
      ctx.fail(
        "unexpressible-value",
        `Non-finite number ${String(column)} cannot be expressed in TypeScript`,
        `@columns[${index}]`,
      );
    }
    if (!tableColumnSchema.safeParse(column).success) {
      ctx.fail(
        "invalid-document",
        `Invalid table column track at index ${index}`,
        `@columns[${index}]`,
      );
    }
  }
}

function emitTableContainerChildren(
  node: CommonNode,
  ctx: CodegenContext,
  expectedChildType: "row" | "cell",
): readonly string[] {
  const children: string[] = [];
  let childIndex = 0;
  for (const child of node.children) {
    if (!(child instanceof CommonNode)) {
      if (
        child.type === "#text" &&
        "text" in child &&
        child.text.trim() !== ""
      ) {
        ctx.fail(
          "invalid-document",
          `${node.type} declaration cannot directly contain text content`,
        );
      }
      continue;
    }
    if (child.type !== expectedChildType) {
      const label =
        child.id === undefined
          ? `${child.type}[${childIndex + 1}]`
          : `${child.type}[${child.id}]`;
      ctx.fail(
        "invalid-document",
        `${node.type} declaration cannot contain ${JSON.stringify(child.type)} child`,
        `${ctx.path}/${label}`,
      );
    }
    children.push(ctx.emitChildNode(child, childIndex));
    childIndex += 1;
  }
  return children;
}

const cellRichTextHandler = createRichTextCodegenHandler({
  vanillaFactory: "cell",
  reactComponent: "Cell",
  allowedPropKeys: cellAllowedPropKeys,
  allowedChildKinds: blockRichTextChildKinds,
  supportObjectExtras: false,
});

/**
 * Source codegen contribution for native `table`, `row`, and `cell` declarations.
 *
 * Preserves column tracks (`px` / `"auto"` / `"Nfr"`), row order, `rowSpan`/`colSpan`
 * merges, layer-local styles and borders, placeholder bindings, and cell rich text.
 */
export const tableCodegen = defineObjectCodegen({
  declarations: {
    table: {
      emitVanilla(node, ctx) {
        validateTableColumns(node, ctx);
        const callee = ctx.useImport("table");
        const propsEntries = collectRichTextNodeProps(
          node,
          ctx,
          tableAllowedPropKeys,
          true,
        );
        validateTableNativePlaceholder(node, ctx);
        validateTableLayerStyle(
          node.props["style"],
          "Table",
          tableCleanStyleKeys,
          ctx,
        );
        const rows = emitTableContainerChildren(node, ctx, "row");
        return ctx.formatVanillaCall(
          callee,
          [ctx.formatProps(propsEntries)],
          rows,
        );
      },
      emitReact(node, ctx) {
        validateTableColumns(node, ctx);
        const tag = ctx.useImport("Table");
        const propsEntries = collectRichTextNodeProps(
          node,
          ctx,
          tableAllowedPropKeys,
          true,
        );
        validateTableNativePlaceholder(node, ctx);
        validateTableLayerStyle(
          node.props["style"],
          "Table",
          tableCleanStyleKeys,
          ctx,
        );
        const rows = emitTableContainerChildren(node, ctx, "row");
        return ctx.formatJsxElement(tag, propsEntries, rows);
      },
    },
    row: {
      emitVanilla(node, ctx) {
        const callee = ctx.useImport("row");
        const propsEntries = collectRichTextNodeProps(
          node,
          ctx,
          rowAllowedPropKeys,
          false,
        );
        validateTableLayerStyle(
          node.props["style"],
          "Row",
          rowCleanStyleKeys,
          ctx,
        );
        const cells = emitTableContainerChildren(node, ctx, "cell");
        const args =
          propsEntries.length === 0 ? [] : [ctx.formatProps(propsEntries)];
        return ctx.formatVanillaCall(callee, args, cells);
      },
      emitReact(node, ctx) {
        const tag = ctx.useImport("Row");
        const propsEntries = collectRichTextNodeProps(
          node,
          ctx,
          rowAllowedPropKeys,
          false,
        );
        validateTableLayerStyle(
          node.props["style"],
          "Row",
          rowCleanStyleKeys,
          ctx,
        );
        const cells = emitTableContainerChildren(node, ctx, "cell");
        return ctx.formatJsxElement(tag, propsEntries, cells);
      },
    },
    cell: {
      emitVanilla(node, ctx) {
        validateTableLayerStyle(
          node.props["style"],
          "Cell",
          cellCleanStyleKeys,
          ctx,
        );
        return cellRichTextHandler.emitVanilla(node, ctx);
      },
      emitReact(node, ctx) {
        validateTableLayerStyle(
          node.props["style"],
          "Cell",
          cellCleanStyleKeys,
          ctx,
        );
        return cellRichTextHandler.emitReact(node, ctx);
      },
    },
  },
});
