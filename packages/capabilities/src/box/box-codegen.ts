import { defineObjectCodegen, type CodegenContext } from "@pptxdk/codegen";
import { CommonNode } from "@pptxdk/core/document";

const objectMotionPropKeys = new Set(["enter", "emphasis", "motion"]);

function collectBoxProps(
  node: CommonNode,
  ctx: CodegenContext,
): readonly (readonly [string, unknown])[] {
  const extra = ctx.emitObjectExtraProps(node.props);
  const entries: [string, unknown][] = [];
  for (const [key, value] of Object.entries(node.props)) {
    if (value === undefined) continue;
    if (objectMotionPropKeys.has(key)) {
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

function emitBoxChildren(
  node: CommonNode,
  ctx: CodegenContext,
): readonly string[] {
  const children: string[] = [];
  let childIndex = 0;
  for (const child of node.children) {
    if (child instanceof CommonNode) {
      children.push(ctx.emitChildNode(child, childIndex));
      childIndex += 1;
    } else if (
      "text" in child &&
      typeof child.text === "string" &&
      child.text.trim() !== ""
    ) {
      ctx.fail("invalid-document", "Box cannot directly contain text content");
    }
  }
  return children;
}

function emitBoxVanilla(node: CommonNode, ctx: CodegenContext): string {
  const callee = ctx.useImport("box");
  const propsEntries = collectBoxProps(node, ctx);
  const args = propsEntries.length === 0 ? [] : [ctx.formatProps(propsEntries)];
  const children = emitBoxChildren(node, ctx);
  return ctx.formatVanillaCall(callee, args, children);
}

function emitBoxReact(node: CommonNode, ctx: CodegenContext): string {
  const tag = ctx.useImport("Box");
  const propsEntries = collectBoxProps(node, ctx);
  const children = emitBoxChildren(node, ctx);
  return ctx.formatJsxElement(tag, propsEntries, children);
}

/** Source codegen contribution for `box` layout container declarations. */
export const boxCodegen = defineObjectCodegen({
  declarations: {
    box: {
      emitVanilla: emitBoxVanilla,
      emitReact: emitBoxReact,
    },
  },
});
