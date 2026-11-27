import {
  defineObjectCodegen,
  expr,
  type CodegenContext,
  type CodegenExpr,
  type DeclarationCodegenHandler,
} from "@pptxdk/codegen";
import type { DeclarationText } from "@pptxdk/core/build";
import { nativePlaceholderRefSchema } from "@pptxdk/core/capability";
import { CommonNode } from "@pptxdk/core/document";

import { serializeImageDescriptorValue } from "../image/image-codegen.ts";
import {
  paragraphAuthorPropsKeys,
  paragraphStyleSchema,
  runStyleSchema,
  slideNumberAuthorPropsKeys,
  spanAuthorPropsKeys,
  textAuthorPropsKeys,
  textBoxStyleSchema,
} from "../shared/text-schemas.ts";

const objectMotionPropKeys = new Set(["enter", "emphasis", "motion"]);

const textAllowedPropKeys = new Set(textAuthorPropsKeys);
const paragraphAllowedPropKeys = new Set(paragraphAuthorPropsKeys);
const spanAllowedPropKeys = new Set(spanAuthorPropsKeys);
const slideNumberAllowedPropKeys = new Set(slideNumberAuthorPropsKeys);

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return false;
  }
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

function validateRichTextNativePlaceholder(
  nodeType: string,
  value: unknown,
  ctx: CodegenContext,
): void {
  const parsed = nativePlaceholderRefSchema.safeParse(value);
  if (!parsed.success) {
    ctx.fail(
      "unexpressible-value",
      parsed.error.issues[0]?.message ??
        `${nodeType}.nativePlaceholder cannot be expressed`,
      "@nativePlaceholder",
    );
  }
}

function serializeRichTextStyleProp(
  nodeType: string,
  style: unknown,
  ctx: CodegenContext,
): CodegenExpr {
  if (!isPlainRecord(style)) {
    ctx.fail(
      "unexpressible-value",
      `${nodeType}.style must be a plain object`,
      "@style",
    );
  }
  const schema =
    nodeType === "text"
      ? textBoxStyleSchema
      : nodeType === "paragraph"
        ? paragraphStyleSchema
        : nodeType === "span" || nodeType === "slide-number"
          ? runStyleSchema
          : undefined;
  if (schema !== undefined) {
    const parsed = schema.safeParse(style);
    if (!parsed.success) {
      const issue = parsed.error.issues[0];
      const issuePath =
        issue !== undefined && issue.path.length > 0
          ? `@style.${issue.path.join(".")}`
          : "@style";
      ctx.fail(
        "unexpressible-value",
        issue?.message ?? `${nodeType}.style cannot be expressed`,
        issuePath,
      );
    }
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

/** Container-level rich-text child declaration types allowed directly inside `Text` and `Notes`. */
export const blockRichTextChildKinds: ReadonlySet<string> = new Set([
  "paragraph",
  "span",
  "slide-number",
]);

/** Inline-level rich-text child declaration types allowed inside `P` and `Span`. */
export const inlineRichTextChildKinds: ReadonlySet<string> = new Set([
  "span",
  "slide-number",
]);

function singleTextLeaf(node: CommonNode): string | undefined {
  if (node.children.length !== 1) return undefined;
  const only = node.children[0];
  if (only === undefined || only instanceof CommonNode) return undefined;
  const leaf = only as DeclarationText;
  return leaf.type === "#text" ? leaf.text : undefined;
}

/** Collects and validates own defined props on a rich-text or slide-number node. */
export function collectRichTextNodeProps(
  node: CommonNode,
  ctx: CodegenContext,
  allowedPropKeys: ReadonlySet<string>,
  supportObjectExtras: boolean,
): readonly (readonly [string, unknown])[] {
  const extra: Readonly<Record<string, CodegenExpr>> = supportObjectExtras
    ? ctx.emitObjectExtraProps(node.props)
    : {};
  const entries: [string, unknown][] = [];
  for (const [key, value] of Object.entries(node.props)) {
    if (value === undefined) continue;
    if (!allowedPropKeys.has(key)) {
      ctx.fail(
        "invalid-document",
        `${node.type} declaration does not support property ${JSON.stringify(key)}`,
        `@${key}`,
      );
    }
    if (key === "nativePlaceholder") {
      validateRichTextNativePlaceholder(node.type, value, ctx);
      entries.push([key, value]);
      continue;
    }
    if (key === "style") {
      entries.push([key, serializeRichTextStyleProp(node.type, value, ctx)]);
      continue;
    }
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

function assertAllowedRichTextChildren(
  node: CommonNode,
  ctx: CodegenContext,
  allowedChildKinds: ReadonlySet<string>,
): void {
  let childIndex = 0;
  for (const child of node.children) {
    if (!(child instanceof CommonNode)) continue;
    const label =
      child.id === undefined
        ? `${child.type}[${childIndex + 1}]`
        : `${child.type}[${child.id}]`;
    if (!allowedChildKinds.has(child.type)) {
      ctx.fail(
        "invalid-document",
        `${node.type} declaration cannot contain ${JSON.stringify(child.type)} child`,
        `${ctx.path}/${label}`,
      );
    }
    childIndex += 1;
  }
}

export type RichTextHandlerOptions = {
  readonly vanillaFactory: string;
  readonly reactComponent: string;
  readonly allowedPropKeys: ReadonlySet<string>;
  readonly allowedChildKinds: ReadonlySet<string>;
  readonly supportObjectExtras: boolean;
};

/**
 * Creates a dual-target (`vanilla` / `react`) codegen handler for a rich-text container node
 * (`Text`, `P`, `Span`, or `Notes`), validating props and child declaration types.
 */
export function createRichTextCodegenHandler(
  options: RichTextHandlerOptions,
): DeclarationCodegenHandler {
  return {
    emitVanilla(node, ctx) {
      assertAllowedRichTextChildren(node, ctx, options.allowedChildKinds);
      const callee = ctx.useImport(options.vanillaFactory);
      const propsEntries = collectRichTextNodeProps(
        node,
        ctx,
        options.allowedPropKeys,
        options.supportObjectExtras,
      );
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
          propsEntries.length === 0
            ? []
            : ["[]", ctx.formatProps(propsEntries)];
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
    },
    emitReact(node, ctx) {
      assertAllowedRichTextChildren(node, ctx, options.allowedChildKinds);
      const tag = ctx.useImport(options.reactComponent);
      const propsEntries = collectRichTextNodeProps(
        node,
        ctx,
        options.allowedPropKeys,
        options.supportObjectExtras,
      );
      const children = ctx.emitRichTextChildren(node.children);
      return ctx.formatJsxElement(tag, propsEntries, children);
    },
  };
}

function emitSlideNumberProps(
  node: CommonNode,
  ctx: CodegenContext,
): readonly (readonly [string, unknown])[] {
  if (node.children.length > 0) {
    ctx.fail(
      "invalid-document",
      "SlideNumber declaration cannot have children",
    );
  }
  return collectRichTextNodeProps(node, ctx, slideNumberAllowedPropKeys, false);
}

/** Source codegen contribution for `text`, `paragraph`, `span`, and `slide-number` declarations. */
export const textCodegen = defineObjectCodegen({
  declarations: {
    text: createRichTextCodegenHandler({
      vanillaFactory: "text",
      reactComponent: "Text",
      allowedPropKeys: textAllowedPropKeys,
      allowedChildKinds: blockRichTextChildKinds,
      supportObjectExtras: true,
    }),
    paragraph: createRichTextCodegenHandler({
      vanillaFactory: "paragraph",
      reactComponent: "P",
      allowedPropKeys: paragraphAllowedPropKeys,
      allowedChildKinds: inlineRichTextChildKinds,
      supportObjectExtras: false,
    }),
    span: createRichTextCodegenHandler({
      vanillaFactory: "span",
      reactComponent: "Span",
      allowedPropKeys: spanAllowedPropKeys,
      allowedChildKinds: inlineRichTextChildKinds,
      supportObjectExtras: false,
    }),
    "slide-number": {
      emitVanilla(node, ctx) {
        const callee = ctx.useImport("slideNumber");
        const propsEntries = emitSlideNumberProps(node, ctx);
        const args =
          propsEntries.length === 0 ? [] : [ctx.formatProps(propsEntries)];
        return ctx.formatVanillaCall(callee, args);
      },
      emitReact(node, ctx) {
        const tag = ctx.useImport("SlideNumber");
        const propsEntries = emitSlideNumberProps(node, ctx);
        return ctx.formatJsxElement(tag, propsEntries, []);
      },
    },
  },
});
