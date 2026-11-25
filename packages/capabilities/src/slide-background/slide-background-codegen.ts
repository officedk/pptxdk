import {
  defineSlideCodegen,
  expr,
  type CodegenContext,
  type CodegenExpr,
} from "@pptxdk/codegen";
import { slideBackgroundSchema } from "@pptxdk/core/capability";

import { serializeImageDescriptorValue } from "../image/image-codegen.ts";

function isPlainObject(value: unknown): value is Record<string, unknown> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return false;
  }
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

/**
 * Serializes `LinearGradient` and `RadialGradient` author values back into `linearGradient(...)`
 * and `radialGradient(...)` calls so gradients in `Slide.background` and `VisualStyle.backgroundImage`
 * round-trip idiomatically.
 */
export function serializeGradientValue(
  value: unknown,
  ctx: CodegenContext,
): CodegenExpr | undefined {
  if (!isPlainObject(value)) return undefined;
  const type = value["type"];
  if (type !== "linear-gradient" && type !== "radial-gradient") {
    return undefined;
  }

  const definedKeys = Object.keys(value).filter(
    (key) => value[key] !== undefined,
  );
  const subPath = ctx.path.includes("@") ? undefined : "@style.backgroundImage";

  if (type === "linear-gradient") {
    if (
      !definedKeys.every(
        (key) => key === "type" || key === "angle" || key === "stops",
      ) ||
      typeof value["angle"] !== "number" ||
      !Array.isArray(value["stops"])
    ) {
      return undefined;
    }
    const fn = ctx.useImport("linearGradient");
    const angleExpr = ctx.emitValueExpr(value["angle"], subPath);
    const stopExprs = value["stops"].map((stop) =>
      ctx.emitValueExpr(stop, subPath),
    );
    return expr(ctx.formatVanillaCall(fn, [angleExpr, ...stopExprs]));
  }

  if (
    !definedKeys.every((key) => key === "type" || key === "stops") ||
    !Array.isArray(value["stops"])
  ) {
    return undefined;
  }
  const fn = ctx.useImport("radialGradient");
  const stopExprs = value["stops"].map((stop) =>
    ctx.emitValueExpr(stop, subPath),
  );
  return expr(ctx.formatVanillaCall(fn, stopExprs));
}

/** Source codegen contribution for `Slide.background` / `SlideMaster.background` / `SlideLayout.background` and gradients. */
export const slideBackgroundCodegen = defineSlideCodegen({
  slideProps: {
    background: (value, ctx) => {
      const parsed = slideBackgroundSchema.safeParse(value);
      if (!parsed.success) {
        ctx.fail(
          "unexpressible-value",
          parsed.error.issues[0]?.message ??
            "Slide.background declaration is invalid",
        );
      }
      return expr(ctx.emitValueExpr(parsed.data));
    },
  },
  valueSerializers: [serializeGradientValue, serializeImageDescriptorValue],
});
