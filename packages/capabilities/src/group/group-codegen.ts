import {
  defineObjectCodegen,
  expr,
  type CodegenContext,
  type CodegenExpr,
} from "@pptxdk/codegen";
import { readInternalDeclarationProps } from "@pptxdk/core/capability";
import { CommonNode } from "@pptxdk/core/document";

import {
  groupDeclarationPropsKeys,
  groupStyleSchema,
} from "./group-declaration.ts";

const objectMotionPropKeys = new Set(["enter", "emphasis", "motion"]);

const groupAllowedPropKeys = new Set(groupDeclarationPropsKeys);

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return false;
  }
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

function serializeGroupStyleProp(
  style: unknown,
  ctx: CodegenContext,
): CodegenExpr {
  if (!isPlainRecord(style)) {
    ctx.fail(
      "unexpressible-value",
      "Group.style must be a plain object",
      "@style",
    );
  }
  const parsed = groupStyleSchema.safeParse(style);
  if (!parsed.success) {
    const issue = parsed.error.issues[0];
    const issuePath =
      issue !== undefined && issue.path.length > 0
        ? `@style.${issue.path.join(".")}`
        : "@style";
    ctx.fail(
      "unexpressible-value",
      issue?.message ?? "Group.style cannot be expressed",
      issuePath,
    );
  }
  const entries = Object.entries(style).filter(
    ([, entryValue]) => entryValue !== undefined,
  );
  return expr(ctx.formatProps(entries));
}

function collectGroupProps(
  node: CommonNode,
  ctx: CodegenContext,
): readonly (readonly [string, unknown])[] {
  const props = readInternalDeclarationProps(node);
  const extra = ctx.emitObjectExtraProps(props);
  const entries: [string, unknown][] = [];
  for (const [key, value] of Object.entries(props)) {
    if (value === undefined) continue;
    if (!groupAllowedPropKeys.has(key)) {
      ctx.fail(
        "invalid-document",
        `group declaration does not support property ${JSON.stringify(key)}`,
        `@${key}`,
      );
    }
    if (key === "style") {
      entries.push([key, serializeGroupStyleProp(value, ctx)]);
      continue;
    }
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

function emitGroupChildren(
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
      ctx.fail(
        "invalid-document",
        "Group cannot directly contain text content",
      );
    }
  }
  return children;
}

function emitGroupVanilla(node: CommonNode, ctx: CodegenContext): string {
  const callee = ctx.useImport("group");
  const propsEntries = collectGroupProps(node, ctx);
  const args = propsEntries.length === 0 ? [] : [ctx.formatProps(propsEntries)];
  const children = emitGroupChildren(node, ctx);
  return ctx.formatVanillaCall(callee, args, children);
}

function emitGroupReact(node: CommonNode, ctx: CodegenContext): string {
  const tag = ctx.useImport("Group");
  const propsEntries = collectGroupProps(node, ctx);
  const children = emitGroupChildren(node, ctx);
  return ctx.formatJsxElement(tag, propsEntries, children);
}

/** Source codegen contribution for native `group` container declarations. */
export const groupCodegen = defineObjectCodegen({
  declarations: {
    group: {
      emitVanilla: emitGroupVanilla,
      emitReact: emitGroupReact,
    },
  },
});
