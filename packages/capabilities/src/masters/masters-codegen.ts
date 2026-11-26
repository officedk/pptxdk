import {
  defineSlideCodegen,
  expr,
  type CodegenContext,
  type CodegenExpr,
} from "@pptxdk/codegen";
import {
  footerPlaceholdersSchema,
  headerFooterSchema,
  layoutPlaceholderSpecSchema,
  placeholderRoleSchema,
  slideFooterValueSchema,
  slideNumberValueSchema,
} from "@pptxdk/core/capability";

import { serializeImageDescriptorValue } from "../image/image-codegen.ts";

const legacyPlaceholderSpecKeys = new Set([
  "x",
  "y",
  "width",
  "height",
  "flipH",
  "flipV",
  "sz",
]);

const legacyPlaceholderVisualKeys = new Set([
  "backgroundColor",
  "backgroundImage",
  "borderColor",
  "borderWidth",
  "borderStyle",
  "startArrow",
  "endArrow",
]);

const legacyFooterPlaceholderKeys = new Set([
  "x",
  "y",
  "width",
  "height",
  "sz",
]);

const masterOrLayoutPathPattern =
  /^deck\/slide-master\[[^\]]+\](?:\/slide-layout\[[^\]]+\])?$/u;

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return false;
  }
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

function serializePlaceholderStyleValue(
  style: Record<string, unknown>,
  ctx: CodegenContext,
): CodegenExpr {
  const entries: [string, unknown][] = [];
  for (const [key, value] of Object.entries(style)) {
    if (value === undefined) continue;
    if (legacyPlaceholderVisualKeys.has(key)) {
      ctx.fail(
        "unexpressible-value",
        `Placeholder.style does not support legacy property ${JSON.stringify(key)}`,
        `@placeholders.style.${key}`,
      );
    }
    if (key === "background") {
      const serializedImage = serializeImageDescriptorValue(value, ctx);
      entries.push([key, serializedImage ?? value]);
      continue;
    }
    entries.push([key, value]);
  }
  return expr(ctx.formatProps(entries));
}

function serializeMasterOrLayoutPlaceholderSpec(
  value: unknown,
  ctx: CodegenContext,
): CodegenExpr | undefined {
  if (
    !masterOrLayoutPathPattern.test(ctx.path) ||
    !isPlainRecord(value) ||
    !placeholderRoleSchema.safeParse(value["type"]).success
  ) {
    return undefined;
  }

  for (const [key, entryValue] of Object.entries(value)) {
    if (entryValue === undefined) continue;
    if (legacyPlaceholderSpecKeys.has(key)) {
      ctx.fail(
        "unexpressible-value",
        `Placeholder spec does not support legacy property ${JSON.stringify(key)}; use size and style`,
        `@placeholders.${key}`,
      );
    }
  }

  const parsed = layoutPlaceholderSpecSchema.safeParse(value);
  if (!parsed.success) {
    const issue = parsed.error.issues[0];
    ctx.fail(
      "unexpressible-value",
      `Invalid placeholder declaration: ${issue?.message ?? "invalid value"}`,
      "@placeholders",
    );
  }

  const specEntries: [string, unknown][] = [];
  for (const [key, entryValue] of Object.entries(value)) {
    if (entryValue === undefined) continue;
    if (key === "style") {
      if (!isPlainRecord(entryValue)) {
        ctx.fail(
          "unexpressible-value",
          "Placeholder.style must be a plain object",
          "@placeholders.style",
        );
      }
      specEntries.push([
        "style",
        serializePlaceholderStyleValue(entryValue, ctx),
      ]);
      continue;
    }
    specEntries.push([key, entryValue]);
  }

  return expr(ctx.formatProps(specEntries));
}

function serializeFooterPlaceholdersValue(
  value: unknown,
  ctx: CodegenContext,
): CodegenExpr {
  const parsed = footerPlaceholdersSchema.safeParse(value);
  if (!parsed.success) {
    return ctx.fail(
      "unexpressible-value",
      `Invalid footerPlaceholders declaration: ${parsed.error.issues[0]?.message ?? "invalid value"}`,
    );
  }
  if (isPlainRecord(value)) {
    for (const [slotKey, slotSpec] of Object.entries(value)) {
      if (!isPlainRecord(slotSpec)) continue;
      for (const [propKey, propValue] of Object.entries(slotSpec)) {
        if (
          propValue !== undefined &&
          legacyFooterPlaceholderKeys.has(propKey)
        ) {
          ctx.fail(
            "unexpressible-value",
            `footerPlaceholders.${slotKey} does not support legacy property ${JSON.stringify(propKey)}; use size and style`,
            `${slotKey}.${propKey}`,
          );
        }
      }
    }
  }
  return expr(ctx.emitValueExpr(value));
}

function serializeValidatedValue(
  label: string,
  validate: (value: unknown) => {
    readonly success: boolean;
    readonly error?: {
      readonly issues: readonly { readonly message: string }[];
    };
  },
  value: unknown,
  ctx: CodegenContext,
): CodegenExpr {
  const parsed = validate(value);
  if (!parsed.success) {
    return ctx.fail(
      "unexpressible-value",
      `Invalid ${label} declaration: ${parsed.error?.issues[0]?.message ?? "invalid value"}`,
    );
  }
  return expr(ctx.emitValueExpr(value));
}

/** Source codegen contribution for `SlideMaster` and `SlideLayout` `headerFooter` and placeholder specs. */
export const mastersCodegen = defineSlideCodegen({
  slideProps: {
    headerFooter: (value, ctx) =>
      serializeValidatedValue(
        "headerFooter",
        (candidate) => headerFooterSchema.safeParse(candidate),
        value,
        ctx,
      ),
  },
  valueSerializers: [serializeMasterOrLayoutPlaceholderSpec],
});

/** Source codegen contribution for `footerPlaceholders` and slide `footer` / `dateTime` / `slideNumber`. */
export const footerPlaceholderCodegen = defineSlideCodegen({
  slideProps: {
    footerPlaceholders: serializeFooterPlaceholdersValue,
    footer: (value, ctx) =>
      serializeValidatedValue(
        "footer",
        (candidate) => slideFooterValueSchema.safeParse(candidate),
        value,
        ctx,
      ),
    dateTime: (value, ctx) =>
      serializeValidatedValue(
        "dateTime",
        (candidate) => slideFooterValueSchema.safeParse(candidate),
        value,
        ctx,
      ),
    slideNumber: (value, ctx) =>
      serializeValidatedValue(
        "slideNumber",
        (candidate) => slideNumberValueSchema.safeParse(candidate),
        value,
        ctx,
      ),
  },
});
