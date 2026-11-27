import { defineObjectCodegen, type CodegenContext } from "@pptxdk/codegen";
import {
  layoutStyleSchema,
  nativePlaceholderRefSchema,
  visualStyleSchema,
} from "@pptxdk/core/capability";
import type { CommonNode } from "@pptxdk/core/document";
import { z } from "zod";

const placeholderAllowedPropKeys = new Set([
  "id",
  "name",
  "morphKey",
  "placeholder",
  "nativePlaceholder",
  "link",
  "style",
]);

const placeholderCleanStyleSchema = z
  .strictObject({
    ...layoutStyleSchema.unwrap().shape,
    ...visualStyleSchema.unwrap().shape,
  })
  .readonly();

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return false;
  }
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

function validatePlaceholderNativeRef(
  value: unknown,
  ctx: CodegenContext,
): void {
  const parsed = nativePlaceholderRefSchema.safeParse(value);
  if (!parsed.success) {
    const issue = parsed.error.issues[0];
    ctx.fail(
      "unexpressible-value",
      issue?.message ?? "Placeholder.nativePlaceholder cannot be expressed",
      "@nativePlaceholder",
    );
  }
}

function validatePlaceholderStyle(value: unknown, ctx: CodegenContext): void {
  if (!isPlainRecord(value)) {
    ctx.fail(
      "unexpressible-value",
      "Placeholder.style must be a plain object",
      "@style",
    );
  }
  const parsed = placeholderCleanStyleSchema.safeParse(value);
  if (!parsed.success) {
    const issue = parsed.error.issues[0];
    const issuePath =
      issue !== undefined && issue.path.length > 0
        ? `@style.${issue.path.join(".")}`
        : "@style";
    ctx.fail(
      "unexpressible-value",
      issue?.message ?? "Placeholder.style cannot be expressed",
      issuePath,
    );
  }
}

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
    if (value === undefined) continue;
    if (!placeholderAllowedPropKeys.has(key)) {
      ctx.fail(
        "invalid-document",
        `placeholder declaration does not support property ${JSON.stringify(key)}`,
        `@${key}`,
      );
    }
    if (key === "nativePlaceholder") {
      validatePlaceholderNativeRef(value, ctx);
    } else if (key === "style") {
      validatePlaceholderStyle(value, ctx);
    }
    entries.push([key, value]);
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
