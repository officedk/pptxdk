import {
  defineObjectCodegen,
  expr,
  type CodegenContext,
  type CodegenExpr,
} from "@pptxdk/codegen";
import type { DeclarationText } from "@pptxdk/core/build";
import {
  imageBackgroundSchema,
  nativePlaceholderRefSchema,
} from "@pptxdk/core/capability";
import { CommonNode } from "@pptxdk/core/document";

import { serializeGeometryProp } from "../shape/shape-codegen.ts";
import { readImage } from "./image-decode.ts";
import {
  externalImageSourceSchema,
  imageDeclarationPropsSchema,
  imageStyleSchema,
} from "./image-schemas.ts";

const objectMotionPropKeys = new Set(["enter", "emphasis", "motion"]);

const imageDescriptorAllowedKeys = new Set([
  "image",
  "fit",
  "position",
  "opacity",
]);

const imageAllowedPropKeys = new Set(
  Object.keys(imageDeclarationPropsSchema.unwrap().shape),
);

const dummyImageBytes = new Uint8Array([0]);

function isPlainObject(value: unknown): value is Record<string, unknown> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return false;
  }
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

function assertNoImageChildren(node: CommonNode, ctx: CodegenContext): void {
  for (const child of node.children) {
    if (child instanceof CommonNode) {
      ctx.fail("invalid-document", "Image does not accept child declarations");
    }
    const leaf = child as DeclarationText;
    if (leaf.type === "#text" && leaf.text.trim() !== "") {
      ctx.fail("invalid-document", "Image does not accept text children");
    }
  }
}

function serializeImageSrcProp(
  value: unknown,
  ctx: CodegenContext,
): CodegenExpr {
  if (isPlainObject(value) && Object.hasOwn(value, "externalUrl")) {
    const parsedExternal = externalImageSourceSchema.safeParse(value);
    if (!parsedExternal.success) {
      ctx.fail(
        "unexpressible-value",
        "Image.src externalUrl must be a non-empty string",
        "@src",
      );
    }
    return expr(
      ctx.emitValueExpr(
        { externalUrl: parsedExternal.data.externalUrl },
        "@src",
      ),
    );
  }
  try {
    const loaded = readImage(value);
    const assetVar = ctx.registerImageAsset(loaded.bytes, loaded.extension);
    return expr(assetVar);
  } catch (error) {
    ctx.fail(
      "unexpressible-value",
      error instanceof Error
        ? error.message
        : "Image.src could not be read or decoded",
      "@src",
    );
  }
}

function validateImageNativePlaceholder(
  value: unknown,
  ctx: CodegenContext,
): void {
  const parsed = nativePlaceholderRefSchema.safeParse(value);
  if (!parsed.success) {
    ctx.fail(
      "unexpressible-value",
      parsed.error.issues[0]?.message ??
        "Image.nativePlaceholder cannot be expressed",
      "@nativePlaceholder",
    );
  }
}

function serializeImageStyleProp(
  style: unknown,
  ctx: CodegenContext,
): CodegenExpr {
  if (!isPlainObject(style)) {
    ctx.fail(
      "unexpressible-value",
      "Image.style must be a plain object",
      "@style",
    );
  }
  const parsed = imageStyleSchema.safeParse(style);
  if (!parsed.success) {
    const issue = parsed.error.issues[0];
    const issuePath =
      issue !== undefined && issue.path.length > 0
        ? `@style.${issue.path.join(".")}`
        : "@style";
    ctx.fail(
      "unexpressible-value",
      issue?.message ?? "Image.style cannot be expressed",
      issuePath,
    );
  }
  const styleEntries: [string, unknown][] = [];
  for (const [key, entryValue] of Object.entries(style)) {
    if (entryValue === undefined) continue;
    if (key === "background") {
      const serializedBg = serializeImageDescriptorValue(entryValue, ctx);
      styleEntries.push([key, serializedBg ?? entryValue]);
      continue;
    }
    styleEntries.push([key, entryValue]);
  }
  return expr(ctx.formatProps(styleEntries));
}

/**
 * Serializes an `{ image, fit?, position?, opacity? }` descriptor (used in `style.background`
 * and `Slide.background`) by reading/validating embedded image bytes or `{ externalUrl }`
 * and emitting a formatted descriptor expression.
 */
export function serializeImageDescriptorValue(
  value: unknown,
  ctx: CodegenContext,
): CodegenExpr | undefined {
  if (!isPlainObject(value) || value["image"] === undefined) {
    return undefined;
  }
  const definedKeys = Object.keys(value).filter(
    (key) => value[key] !== undefined,
  );
  if (
    definedKeys.length === 0 ||
    !definedKeys.every((key) => imageDescriptorAllowedKeys.has(key))
  ) {
    return undefined;
  }

  const subPath = ctx.path.includes("@") ? undefined : "@style.background";
  const rawSource = value["image"];
  let imageExpr: CodegenExpr;

  if (isPlainObject(rawSource) && Object.hasOwn(rawSource, "externalUrl")) {
    const parsedExternal = externalImageSourceSchema.safeParse(rawSource);
    if (!parsedExternal.success) {
      ctx.fail(
        "unexpressible-value",
        "Background image externalUrl must be a non-empty string",
        subPath,
      );
    }
    imageExpr = expr(
      ctx.emitValueExpr(
        { externalUrl: parsedExternal.data.externalUrl },
        subPath,
      ),
    );
  } else {
    try {
      const loaded = readImage(rawSource);
      const assetVar = ctx.registerImageAsset(loaded.bytes, loaded.extension);
      imageExpr = expr(assetVar);
    } catch (error) {
      ctx.fail(
        "unexpressible-value",
        error instanceof Error
          ? error.message
          : "Background image source could not be read or decoded",
        subPath,
      );
    }
  }

  const parsedOptions = imageBackgroundSchema.safeParse({
    ...value,
    image: dummyImageBytes,
  });
  if (!parsedOptions.success) {
    ctx.fail(
      "unexpressible-value",
      parsedOptions.error.issues[0]?.message ??
        "Background image descriptor options are invalid",
      subPath,
    );
  }

  const descriptorEntries: [string, unknown][] = [["image", imageExpr]];
  for (const optionKey of ["fit", "position", "opacity"] as const) {
    const optionValue = value[optionKey];
    if (optionValue !== undefined) {
      descriptorEntries.push([
        optionKey,
        expr(ctx.emitValueExpr(optionValue, subPath)),
      ]);
    }
  }
  return expr(ctx.formatProps(descriptorEntries));
}

function collectImageProps(
  node: CommonNode,
  ctx: CodegenContext,
): readonly (readonly [string, unknown])[] {
  assertNoImageChildren(node, ctx);
  const extra: Readonly<Record<string, CodegenExpr>> = ctx.emitObjectExtraProps(
    node.props,
  );
  const entries: [string, unknown][] = [];
  let sawSrc = false;

  for (const [key, value] of Object.entries(node.props)) {
    if (value === undefined) continue;
    if (!imageAllowedPropKeys.has(key)) {
      ctx.fail(
        "invalid-document",
        `image declaration does not support property ${JSON.stringify(key)}`,
        `@${key}`,
      );
    }
    if (key === "src") {
      sawSrc = true;
      entries.push(["src", serializeImageSrcProp(value, ctx)]);
      continue;
    }
    if (key === "nativePlaceholder") {
      validateImageNativePlaceholder(value, ctx);
      entries.push([key, value]);
      continue;
    }
    if (key === "geometry") {
      entries.push([key, serializeGeometryProp(value, ctx)]);
      continue;
    }
    if (key === "style") {
      entries.push([key, serializeImageStyleProp(value, ctx)]);
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

  if (!sawSrc) {
    ctx.fail("unexpressible-value", "Image declaration is missing src", "@src");
  }
  return entries;
}

/** Source codegen contribution for `image` declarations and `{ image, fit?, position?, opacity? }` descriptors. */
export const imageCodegen = defineObjectCodegen({
  declarations: {
    image: {
      emitVanilla: (node, ctx) => {
        const callee = ctx.useImport("image");
        const props = collectImageProps(node, ctx);
        return ctx.formatVanillaCall(callee, [ctx.formatProps(props)]);
      },
      emitReact: (node, ctx) => {
        const tag = ctx.useImport("Image");
        const props = collectImageProps(node, ctx);
        return ctx.formatJsxElement(tag, props);
      },
    },
  },
  valueSerializers: [serializeImageDescriptorValue],
});
