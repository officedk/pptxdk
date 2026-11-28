import {
  defineObjectCodegen,
  expr,
  type CodegenContext,
  type CodegenExpr,
} from "@pptxdk/codegen";
import type { DeclarationText } from "@pptxdk/core/build";
import { readInternalDeclarationProps } from "@pptxdk/core/capability";
import { CommonNode } from "@pptxdk/core/document";

import { motionPropsShape } from "../motion/motion-schemas.ts";
import {
  lineAdjustSchema,
  lineAllowedPropKeyList,
  lineEndpointSchema,
  lineRoutingSchema,
  lineStrokeValueSchema,
  lineStyleSchema,
} from "./line-declaration.ts";

const objectMotionPropKeys = new Set(Object.keys(motionPropsShape));
const lineAllowedPropKeys = new Set(lineAllowedPropKeyList);

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  if (!isRecord(value)) return false;
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

function assertNoLineChildren(node: CommonNode, ctx: CodegenContext): void {
  for (const child of node.children) {
    if (child instanceof CommonNode) {
      ctx.fail("invalid-document", "Line does not accept child declarations");
    }
    const leaf = child as DeclarationText;
    if (leaf.type === "#text" && leaf.text.trim() !== "") {
      ctx.fail("invalid-document", "Line does not accept text children");
    }
  }
}

function serializeLineStyleProp(
  style: unknown,
  ctx: CodegenContext,
): CodegenExpr | undefined {
  if (!isPlainRecord(style)) {
    ctx.fail(
      "unexpressible-value",
      "Line.style must be a plain object",
      "@style",
    );
  }
  const parsed = lineStyleSchema.safeParse(style);
  if (!parsed.success) {
    const issue = parsed.error.issues[0];
    const issuePath =
      issue !== undefined && issue.path.length > 0
        ? `@style.${issue.path.join(".")}`
        : "@style";
    ctx.fail(
      "unexpressible-value",
      issue?.message ?? "Line.style cannot be expressed",
      issuePath,
    );
  }
  const entries: [string, unknown][] = [];
  for (const [key, entryValue] of Object.entries(style)) {
    if (entryValue === undefined) continue;
    entries.push([key, entryValue]);
  }
  if (entries.length === 0) return undefined;
  return expr(ctx.formatProps(entries));
}

function validateLinePropValue(
  key: string,
  value: unknown,
  ctx: CodegenContext,
): void {
  if (key === "routing" && !lineRoutingSchema.safeParse(value).success) {
    ctx.fail(
      "unexpressible-value",
      "Line.routing cannot be expressed",
      `@${key}`,
    );
  }
  if (key === "adjust" && !lineAdjustSchema.safeParse(value).success) {
    ctx.fail(
      "unexpressible-value",
      "Line.adjust cannot be expressed",
      `@${key}`,
    );
  }
  if (key === "stroke" && !lineStrokeValueSchema.safeParse(value).success) {
    ctx.fail(
      "unexpressible-value",
      "Line.stroke cannot be expressed",
      `@${key}`,
    );
  }
  if (
    (key === "from" || key === "to") &&
    !lineEndpointSchema.safeParse(value).success
  ) {
    ctx.fail(
      "unexpressible-value",
      `Line.${key} cannot be expressed`,
      `@${key}`,
    );
  }
}

function collectLineProps(
  node: CommonNode,
  ctx: CodegenContext,
): readonly (readonly [string, unknown])[] {
  const props = readInternalDeclarationProps(node);
  const extra = ctx.emitObjectExtraProps(props);
  const entries: [string, unknown][] = [];
  for (const [key, value] of Object.entries(props)) {
    if (value === undefined) continue;
    if (!lineAllowedPropKeys.has(key)) {
      ctx.fail(
        "invalid-document",
        `line declaration does not support property ${JSON.stringify(key)}`,
        `@${key}`,
      );
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
    if (key === "style") {
      const styleExpr = serializeLineStyleProp(value, ctx);
      if (styleExpr !== undefined) {
        entries.push([key, styleExpr]);
      }
      continue;
    }
    validateLinePropValue(key, value, ctx);
    entries.push([key, value]);
  }
  return entries;
}

function emitLineVanilla(node: CommonNode, ctx: CodegenContext): string {
  assertNoLineChildren(node, ctx);
  const callee = ctx.useImport("line");
  const propsEntries = collectLineProps(node, ctx);
  const args = propsEntries.length === 0 ? [] : [ctx.formatProps(propsEntries)];
  return ctx.formatVanillaCall(callee, args);
}

function emitLineReact(node: CommonNode, ctx: CodegenContext): string {
  assertNoLineChildren(node, ctx);
  const tag = ctx.useImport("Line");
  const propsEntries = collectLineProps(node, ctx);
  return ctx.formatJsxElement(tag, propsEntries);
}

/** Source codegen contribution for `line` declarations. */
export const lineCodegen = defineObjectCodegen({
  declarations: {
    line: {
      emitVanilla: emitLineVanilla,
      emitReact: emitLineReact,
    },
  },
});
