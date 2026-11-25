import {
  defineObjectCodegen,
  expr,
  type CodegenContext,
  type CodegenExpr,
} from "@pptxdk/codegen";
import type { DeclarationText } from "@pptxdk/core/build";
import { CommonNode } from "@pptxdk/core/document";

import { readImage } from "./image-decode.ts";
import { externalImageSourceSchema } from "./image-schemas.ts";

const objectMotionPropKeys = new Set(["enter", "emphasis", "motion"]);

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

function serializeImageSrcProp(value: unknown, ctx: CodegenContext): unknown {
  if (isPlainObject(value) && Object.hasOwn(value, "externalUrl")) {
    const parsedExternal = externalImageSourceSchema.safeParse(value);
    if (!parsedExternal.success) {
      ctx.fail(
        "unexpressible-value",
        "Image.src externalUrl must be a non-empty string",
        "@src",
      );
    }
    return { externalUrl: parsedExternal.data.externalUrl };
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

/**
 * Serializes an `{ image: string | URL | Uint8Array }` descriptor (used in `VisualStyle.backgroundImage`
 * and `SlideBackground`) by reading/validating the embedded image and registering a relative asset URL.
 */
export function serializeImageDescriptorValue(
  value: unknown,
  ctx: CodegenContext,
): CodegenExpr | undefined {
  if (!isPlainObject(value)) return undefined;
  const definedKeys = Object.keys(value).filter(
    (key) => value[key] !== undefined,
  );
  if (definedKeys.length !== 1 || definedKeys[0] !== "image") {
    return undefined;
  }
  const rawSource = value["image"];
  const subPath = ctx.path.includes("@") ? undefined : "@style.backgroundImage";
  try {
    const loaded = readImage(rawSource);
    const assetVar = ctx.registerImageAsset(loaded.bytes, loaded.extension);
    return expr(ctx.formatProps([["image", expr(assetVar)]]));
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
    if (key === "src") {
      sawSrc = true;
      entries.push(["src", serializeImageSrcProp(value, ctx)]);
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

/** Source codegen contribution for `image` declarations and `{ image: ... }` fill descriptors. */
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
