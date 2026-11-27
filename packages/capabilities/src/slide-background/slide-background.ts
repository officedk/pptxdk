import {
  CapabilityError,
  applyFormatColorTransforms,
  slideBackgroundSchema,
  type BoundSlideCapability,
  type CapabilityInput,
  type ColorValue,
  type FormatColorTransform,
  type ResolvedColor,
  type SealedCanvasMechanism,
  type SlideResolution,
  type SlideBackground,
  type SlideCapabilityModule,
  type SlideEmitContext,
  type SlideEmitOutput,
  type SlidePaintContext,
  type SlidePaintOutput,
  type SlideResolveContext,
  type ThemeContext,
} from "@pptxdk/core/capability";
import { defineCanvasMechanism } from "@pptxdk/import";

import { readImage } from "../image/image-decode.ts";
import {
  gradientDefSvg,
  paintImageFillSvg,
  resolveGradientSpec,
  resolveImageBackgroundSpec,
  type ResolvedGradientStop,
  type ResolvedImageFill,
} from "../shared/visual.ts";
import { slideBackgroundCodegen } from "./slide-background-codegen.ts";
import { readSlideBackgroundSlot } from "./slide-background-import.ts";
import {
  slideBackgroundGradientXml,
  slideBackgroundImageXml,
  slideBackgroundNoneXml,
  slideBackgroundRefXml,
  slideBackgroundSolidXml,
} from "./slide-background-xml.ts";

type ResolvedBgRefFill =
  | { readonly kind: "none" }
  | { readonly kind: "solid"; readonly color: ResolvedColor }
  | {
      readonly kind: "linear-gradient";
      readonly angle: number;
      readonly stops: readonly ResolvedGradientStop[];
    }
  | {
      readonly kind: "radial-gradient";
      readonly stops: readonly ResolvedGradientStop[];
    };

/** Closed resolved representation of Slide.background consumed exhaustively by emit and paint. */
export type ResolvedSlideBackground =
  | { readonly kind: "default" }
  | { readonly kind: "none" }
  | { readonly kind: "solid"; readonly color: ResolvedColor }
  | {
      readonly kind: "linear-gradient";
      readonly angle: number;
      readonly stops: readonly ResolvedGradientStop[];
    }
  | {
      readonly kind: "radial-gradient";
      readonly stops: readonly ResolvedGradientStop[];
    }
  | ResolvedImageFill
  | {
      readonly kind: "ref";
      readonly idx: 1001 | 1002 | 1003;
      readonly refColor: ResolvedColor;
      readonly fill: ResolvedBgRefFill;
    };

function resolveFormatColor(
  theme: ThemeContext,
  refColor: ResolvedColor,
  transform: FormatColorTransform,
): ResolvedColor {
  const base =
    transform.color === "phClr"
      ? refColor
      : theme.resolveColor(transform.color);
  return applyFormatColorTransforms(base, transform);
}

function resolveBgRefFill(
  theme: ThemeContext,
  idx: 1001 | 1002 | 1003,
  refColor: ResolvedColor,
): ResolvedBgRefFill {
  const bgStyle = theme.formatScheme?.bgFillStyles?.[idx - 1001];
  if (bgStyle === undefined) {
    return { kind: "solid", color: refColor };
  }
  if (bgStyle.kind === "none") {
    return { kind: "none" };
  }
  if (bgStyle.kind === "solid") {
    return {
      kind: "solid",
      color: resolveFormatColor(theme, refColor, bgStyle.color),
    };
  }
  const stops: ResolvedGradientStop[] = bgStyle.stops.map((stop) => ({
    position: stop.position,
    color: resolveFormatColor(theme, refColor, stop.color),
  }));
  if (bgStyle.linear !== undefined) {
    const authorAngle = (((bgStyle.linear.angle + 90) % 360) + 360) % 360;
    return { kind: "linear-gradient", angle: authorAngle, stops };
  }
  return { kind: "radial-gradient", stops };
}

function resolveSlideBackground(
  theme: ThemeContext,
  background: unknown,
): ResolvedSlideBackground {
  if (background === undefined) {
    return { kind: "default" };
  }
  if (background === "none") {
    return { kind: "none" };
  }

  const parsed = slideBackgroundSchema.safeParse(background);
  if (!parsed.success) {
    // Preserve clear diagnostic codes for gradient/image subfields when author passed a malformed variant.
    if (
      typeof background === "object" &&
      background !== null &&
      !Array.isArray(background)
    ) {
      const candidate = background as Readonly<Record<string, unknown>>;
      if (
        candidate["type"] === "linear-gradient" ||
        candidate["type"] === "radial-gradient"
      ) {
        return resolveGradientSpec(theme, background, undefined);
      }
      if (
        Object.keys(candidate).length === 1 &&
        Object.hasOwn(candidate, "image")
      ) {
        readImage(candidate["image"]);
      }
      if ("color" in candidate && !("ref" in candidate)) {
        return {
          kind: "solid",
          color: theme.resolveColor(background as ColorValue),
        };
      }
    }
    throw new CapabilityError(
      "invalid-slide-background",
      parsed.error.issues[0]?.message ??
        "Slide.background declaration is invalid",
    );
  }

  const value: SlideBackground = parsed.data;
  if (value === "none") {
    return { kind: "none" };
  }
  if (typeof value === "object" && value !== null) {
    if ("image" in value) {
      return resolveImageBackgroundSpec(value, undefined);
    }
    if ("ref" in value) {
      const refColor = theme.resolveColor(value.color ?? "bg1");
      return {
        kind: "ref",
        idx: value.ref,
        refColor,
        fill: resolveBgRefFill(theme, value.ref, refColor),
      };
    }
    if ("type" in value) {
      return resolveGradientSpec(theme, value, undefined);
    }
  }

  return {
    kind: "solid",
    color: theme.resolveColor(value),
  };
}

function emitSlideBackground(
  resolved: ResolvedSlideBackground,
  emitContext: SlideEmitContext,
): SlideEmitOutput {
  switch (resolved.kind) {
    case "default":
      return { xml: "" };
    case "none":
      return { xml: slideBackgroundNoneXml() };
    case "solid":
      return { xml: slideBackgroundSolidXml(resolved.color) };
    case "linear-gradient":
    case "radial-gradient":
      return { xml: slideBackgroundGradientXml(resolved) };
    case "image":
      return {
        xml: slideBackgroundImageXml(
          resolved,
          (resource) => emitContext.registerResource(resource),
          {
            x: 0,
            y: 0,
            width: emitContext.dimensions.width,
            height: emitContext.dimensions.height,
          },
        ),
      };
    case "ref":
      return { xml: slideBackgroundRefXml(resolved.idx, resolved.refColor) };
    default: {
      const exhaustive: never = resolved;
      return exhaustive;
    }
  }
}

function paintSlideBackground(
  resolved: ResolvedSlideBackground,
  paintContext: SlidePaintContext,
  field = "style.background",
): SlidePaintOutput {
  const { width, height } = paintContext.dimensions;
  switch (resolved.kind) {
    case "default":
    case "none":
      return { svg: "" };
    case "solid":
      return {
        svg: `<rect width="100%" height="100%" fill="#${resolved.color.hex}" fill-opacity="${resolved.color.alpha ?? 1}"/>`,
      };
    case "linear-gradient":
    case "radial-gradient": {
      const gradId = `bg-s${paintContext.slideNumber}-grad`;
      const def = gradientDefSvg(gradId, resolved);
      return {
        svg: `<defs>${def}</defs><rect x="0" y="0" width="${width}" height="${height}" fill="url(#${gradId})"/>`,
      };
    }
    case "image": {
      const patId = `bg-s${paintContext.slideNumber}-img`;
      let def: string;
      try {
        def = paintImageFillSvg(resolved, patId, {
          x: 0,
          y: 0,
          width,
          height,
        });
      } catch (error) {
        const message =
          error instanceof Error
            ? error.message
            : "Failed to decode slide background image for preview";
        throw new CapabilityError(
          "preview-corrupt-image",
          message,
          `${paintContext.path}/background`,
          {
            ...(error instanceof CapabilityError ? error.details : undefined),
            field,
          },
        );
      }
      return {
        svg: `<defs>${def}</defs><rect x="0" y="0" width="${width}" height="${height}" fill="url(#${patId})"/>`,
        ...(resolved.source.kind === "external"
          ? {
              placeholders: [
                {
                  path: `${paintContext.path}/background`,
                  url: resolved.source.externalUrl,
                },
              ],
            }
          : {}),
      };
    }
    case "ref": {
      switch (resolved.fill.kind) {
        case "none":
          return { svg: '<rect width="100%" height="100%" fill="white"/>' };
        case "solid":
          return {
            svg: `<rect width="100%" height="100%" fill="#${resolved.fill.color.hex}" fill-opacity="${resolved.fill.color.alpha ?? 1}"/>`,
          };
        case "linear-gradient":
        case "radial-gradient": {
          const gradId = `bg-s${paintContext.slideNumber}-grad`;
          const def = gradientDefSvg(gradId, resolved.fill);
          return {
            svg: `<defs>${def}</defs><rect x="0" y="0" width="${width}" height="${height}" fill="url(#${gradId})"/>`,
          };
        }
        default: {
          const exhaustiveFill: never = resolved.fill;
          return exhaustiveFill;
        }
      }
    }
    default: {
      const exhaustive: never = resolved;
      return exhaustive;
    }
  }
}

type SolidBackgroundModel = Extract<
  ResolvedSlideBackground,
  { readonly kind: "solid" }
>;
type GradientBackgroundModel = Extract<
  ResolvedSlideBackground,
  { readonly kind: "linear-gradient" | "radial-gradient" }
>;
type ImageBackgroundModel = Extract<
  ResolvedSlideBackground,
  { readonly kind: "image" }
>;
type RefBackgroundModel = Extract<
  ResolvedSlideBackground,
  { readonly kind: "ref" }
>;

/** Native canvas mechanism for `<p:bg><p:bgPr><a:solidFill>`. */
export const slideBackgroundSolidMechanism = defineCanvasMechanism({
  key: "slide-background-solid",
  role: "canvas",
  emit: (model: SolidBackgroundModel, context: SlideEmitContext) =>
    emitSlideBackground(model, context),
  paint: (model: SolidBackgroundModel, context: SlidePaintContext) =>
    paintSlideBackground(model, context),
  import: {
    target: "slide",
    slot: "background",
    read: readSlideBackgroundSlot,
  },
});

/** Native canvas mechanism for `<p:bg><p:bgPr><a:gradFill>`. */
export const slideBackgroundGradientMechanism = defineCanvasMechanism({
  key: "slide-background-gradient",
  role: "canvas",
  emit: (model: GradientBackgroundModel, context: SlideEmitContext) =>
    emitSlideBackground(model, context),
  paint: (model: GradientBackgroundModel, context: SlidePaintContext) =>
    paintSlideBackground(model, context),
  import: {
    target: "slide",
    slot: "background",
    read: readSlideBackgroundSlot,
  },
});

/** Native canvas mechanism for `<p:bg><p:bgPr><a:blipFill>`. */
export const slideBackgroundImageMechanism = defineCanvasMechanism({
  key: "slide-background-image",
  role: "canvas",
  emit: (model: ImageBackgroundModel, context: SlideEmitContext) =>
    emitSlideBackground(model, context),
  paint: (model: ImageBackgroundModel, context: SlidePaintContext) =>
    paintSlideBackground(model, context),
  import: {
    target: "slide",
    slot: "background",
    read: readSlideBackgroundSlot,
  },
});

/** Native canvas mechanism for `<p:bg><p:bgRef idx="1001..1003">`. */
export const slideBackgroundRefMechanism = defineCanvasMechanism({
  key: "slide-background-ref",
  role: "canvas",
  emit: (model: RefBackgroundModel, context: SlideEmitContext) =>
    emitSlideBackground(model, context),
  paint: (model: RefBackgroundModel, context: SlidePaintContext) =>
    paintSlideBackground(model, context),
  import: {
    target: "slide",
    slot: "background",
    read: readSlideBackgroundSlot,
  },
});

/**
 * Bound slide-background module: validates and resolves background declarations once per build
 * and owns both PresentationML `<p:bg>` emission and static preview canvas painting.
 */
export const slideBackgroundCapability: SlideCapabilityModule<
  BoundSlideCapability,
  readonly [
    SealedCanvasMechanism<"slide-background-solid">,
    SealedCanvasMechanism<"slide-background-gradient">,
    SealedCanvasMechanism<"slide-background-image">,
    SealedCanvasMechanism<"slide-background-ref">,
  ]
> = {
  declaration: "slide-background",
  slideProp: "background",
  consumesOwnerBackgroundStyle: true,
  slot: "background",
  always: true,
  codegen: slideBackgroundCodegen,
  mechanisms: [
    slideBackgroundSolidMechanism,
    slideBackgroundGradientMechanism,
    slideBackgroundImageMechanism,
    slideBackgroundRefMechanism,
  ] as const,
  bindSlide(input: CapabilityInput, context) {
    const rawStyle = input.props["style"];
    const styleBackground =
      typeof rawStyle === "object" &&
      rawStyle !== null &&
      !Array.isArray(rawStyle)
        ? (rawStyle as Readonly<Record<string, unknown>>)["background"]
        : undefined;
    const backgroundField =
      styleBackground !== undefined ? "style.background" : "background";
    const background = styleBackground ?? input.props["background"];
    try {
      // Validate background declaration eagerly at bind time so invalid tokens/resources surface early.
      resolveSlideBackground(context.theme, background);
    } catch (error) {
      if (error instanceof CapabilityError) {
        throw new CapabilityError(
          error.code,
          error.message,
          `${context.path}/background`,
        );
      }
      throw error;
    }

    return {
      resolveSlide(resolveContext: SlideResolveContext): SlideResolution {
        const resolved = resolveSlideBackground(
          resolveContext.theme,
          background,
        );
        return {
          extensions: [],
          emit(emitContext: SlideEmitContext): SlideEmitOutput {
            return emitSlideBackground(resolved, emitContext);
          },
          paint(paintContext: SlidePaintContext): SlidePaintOutput {
            return paintSlideBackground(
              resolved,
              paintContext,
              backgroundField,
            );
          },
        };
      },
    };
  },
} satisfies SlideCapabilityModule;
