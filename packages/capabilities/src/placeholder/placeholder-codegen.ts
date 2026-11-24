import { defineObjectCodegen, type CodegenContext } from "@pptxdk/codegen";
import type { CommonNode } from "@pptxdk/core/document";

function collectPlaceholderProps(
  node: CommonNode,
  ctx: CodegenContext,
): readonly (readonly [string, unknown])[] {
  if (node.children.length > 0) {
    return ctx.fail(
      "unexpressible-value",
      "Unpopulated placeholder cannot contain children",
    );
  }
  const entries: [string, unknown][] = [];
  for (const [key, value] of Object.entries(node.props)) {
    if (value !== undefined) {
      entries.push([key, value]);
    }
  }
  return entries;
}

/** Source codegen contribution for unpopulated `placeholder` declarations. */
export const placeholderCodegen = defineObjectCodegen({
  declarations: {
    placeholder: {
      emitVanilla: (node, ctx) => {
        const callee = ctx.useImport("placeholder");
        const propsEntries = collectPlaceholderProps(node, ctx);
        return propsEntries.length === 0
          ? ctx.formatVanillaCall(callee, [])
          : ctx.formatVanillaCall(callee, [ctx.formatProps(propsEntries)]);
      },
      emitReact: (node, ctx) => {
        const tag = ctx.useImport("Placeholder");
        const propsEntries = collectPlaceholderProps(node, ctx);
        return ctx.formatJsxElement(tag, propsEntries);
      },
    },
  },
});
