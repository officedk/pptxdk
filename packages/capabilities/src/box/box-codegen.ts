import {
  defineObjectCodegen,
  expr,
  type CodegenContext,
  type CodegenExpr,
} from "@pptxdk/codegen";
import { readInternalDeclarationProps } from "@pptxdk/core/capability";
import { CommonNode } from "@pptxdk/core/document";

import { serializeImageDescriptorValue } from "../image/image-codegen.ts";
import { motionPropKeySet } from "../motion/motion-schemas.ts";
import { boxDeclarationPropsKeys, boxStyleSchema } from "./box-declaration.ts";

const objectMotionPropKeys = motionPropKeySet;

const boxAllowedPropKeys = new Set<string>(boxDeclarationPropsKeys);

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return false;
  }
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

function serializeBoxStyleProp(
  style: unknown,
  ctx: CodegenContext,
): CodegenExpr {
  if (!isPlainRecord(style)) {
    ctx.fail(
      "unexpressible-value",
      "Box.style must be a plain object",
      "@style",
    );
  }
  const parsed = boxStyleSchema.safeParse(style);
  if (!parsed.success) {
    const issue = parsed.error.issues[0];
    const issuePath =
      issue !== undefined && issue.path.length > 0
        ? `@style.${issue.path.join(".")}`
        : "@style";
    ctx.fail(
      "unexpressible-value",
      issue?.message ?? "Box.style cannot be expressed",
      issuePath,
    );
  }
  const entries: [string, unknown][] = [];
  for (const [key, entryValue] of Object.entries(style)) {
    if (entryValue === undefined) continue;
    if (key === "background") {
      const serializedImage = serializeImageDescriptorValue(entryValue, ctx);
      entries.push([key, serializedImage ?? entryValue]);
      continue;
    }
    entries.push([key, entryValue]);
  }
  return expr(ctx.formatProps(entries));
}

function collectBoxProps(
  node: CommonNode,
  ctx: CodegenContext,
): readonly (readonly [string, unknown])[] {
  const props = readInternalDeclarationProps(node);
  const extra = ctx.emitObjectExtraProps(props);
  const entries: [string, unknown][] = [];
  for (const [key, value] of Object.entries(props)) {
    if (value === undefined) continue;
    if (!boxAllowedPropKeys.has(key)) {
      ctx.fail(
        "invalid-document",
        `box declaration does not support property ${JSON.stringify(key)}`,
        `@${key}`,
      );
    }
    if (key === "style") {
      entries.push([key, serializeBoxStyleProp(value, ctx)]);
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
