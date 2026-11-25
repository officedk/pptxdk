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
  tableColumnSchema,
  tableDeclarationPropsSchema,
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
        const cells = emitTableContainerChildren(node, ctx, "cell");
        return ctx.formatJsxElement(tag, propsEntries, cells);
      },
    },
    cell: createRichTextCodegenHandler({
      vanillaFactory: "cell",
      reactComponent: "Cell",
      allowedPropKeys: cellAllowedPropKeys,
      allowedChildKinds: blockRichTextChildKinds,
      supportObjectExtras: false,
    }),
  },
});
