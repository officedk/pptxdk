import {
  defineObjectCodegen,
  expr,
  type CodegenContext,
  type CodegenExpr,
} from "@pptxdk/codegen";
import type { DeclarationText } from "@pptxdk/core/build";
import { CommonNode } from "@pptxdk/core/document";

const objectMotionPropKeys = new Set(["enter", "emphasis", "motion"]);

function singleTextLeaf(node: CommonNode): string | undefined {
  if (node.children.length !== 1) return undefined;
  const only = node.children[0];
  if (only === undefined || only instanceof CommonNode) return undefined;
  const leaf = only as DeclarationText;
  return leaf.type === "#text" ? leaf.text : undefined;
}

function collectNodeProps(
  node: CommonNode,
  ctx: CodegenContext,
  supportObjectExtras: boolean,
): readonly (readonly [string, unknown])[] {
  const extra: Readonly<Record<string, CodegenExpr>> = supportObjectExtras
    ? ctx.emitObjectExtraProps(node.props)
    : {};
  const entries: [string, unknown][] = [];
  for (const [key, value] of Object.entries(node.props)) {
    if (value === undefined) continue;
    if (supportObjectExtras && objectMotionPropKeys.has(key)) {
      if (Object.hasOwn(extra, key)) {
        entries.push([key, extra[key]]);
      } else {
        ctx.fail(
          "unexpressible-value",
          `Object property ${JSON.stringify(key)} requires motion codegen support`,
          `@${key}`,
        );
      }
      continue;
    }
    entries.push([key, value]);
  }
  return entries;
}

function emitRichTextVanilla(
  node: CommonNode,
  ctx: CodegenContext,
  factoryName: string,
  supportObjectExtras: boolean,
): string {
  const callee = ctx.useImport(factoryName);
  const propsEntries = collectNodeProps(node, ctx, supportObjectExtras);
  const leafText = singleTextLeaf(node);
  if (leafText !== undefined) {
    const args =
      propsEntries.length === 0
        ? [JSON.stringify(leafText)]
        : [JSON.stringify(leafText), ctx.formatProps(propsEntries)];
    return ctx.formatVanillaCall(callee, args);
  }
  if (node.children.length === 0) {
    const args =
      propsEntries.length === 0 ? [] : ["[]", ctx.formatProps(propsEntries)];
    return ctx.formatVanillaCall(callee, args);
  }
  const children = ctx.emitRichTextChildren(node.children);
  if (propsEntries.length === 0) {
    return ctx.formatVanillaCall(callee, [], children);
  }
  const contentArg = ctx.emitValueExpr(children.map((code) => expr(code)));
  return ctx.formatVanillaCall(callee, [
    contentArg,
    ctx.formatProps(propsEntries),
  ]);
}

function emitRichTextReact(
  node: CommonNode,
  ctx: CodegenContext,
  componentName: string,
  supportObjectExtras: boolean,
): string {
  const tag = ctx.useImport(componentName);
  const propsEntries = collectNodeProps(node, ctx, supportObjectExtras);
  const children = ctx.emitRichTextChildren(node.children);
  return ctx.formatJsxElement(tag, propsEntries, children);
}

/** Basic source codegen contribution for `text`, `paragraph`, and `span` declarations. */
export const textCodegen = defineObjectCodegen({
  declarations: {
    text: {
      emitVanilla: (node, ctx) => emitRichTextVanilla(node, ctx, "text", true),
      emitReact: (node, ctx) => emitRichTextReact(node, ctx, "Text", true),
    },
    paragraph: {
      emitVanilla: (node, ctx) =>
        emitRichTextVanilla(node, ctx, "paragraph", false),
      emitReact: (node, ctx) => emitRichTextReact(node, ctx, "P", false),
    },
    span: {
      emitVanilla: (node, ctx) => emitRichTextVanilla(node, ctx, "span", false),
      emitReact: (node, ctx) => emitRichTextReact(node, ctx, "Span", false),
    },
  },
});
