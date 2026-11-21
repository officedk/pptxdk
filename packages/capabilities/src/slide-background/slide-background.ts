import {
  CapabilityError,
  type CapabilityInput,
  type ColorValue,
  type ResolvedColor,
  type ResolvedSlide,
  type SlideBackground,
  type SlideCapabilityModule,
  type SlideEmitContext,
  type SlideEmitOutput,
  type SlidePaintContext,
  type SlidePaintOutput,
  type SlideResolveContext,
  type ThemeContext,
} from "@pptxdk/core/capability";

import {
  decodeImageToPngBytes,
  readImage,
  type LoadedImage,
} from "../image/image-decode.ts";
import {
  gradientDefSvg,
  resolveGradientSpec,
  type ResolvedGradientStop,
} from "../shared/visual.ts";
import {
  slideBackgroundGradientXml,
  slideBackgroundImageXml,
  slideBackgroundSolidXml,
} from "./slide-background-xml.ts";

const imageRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/** Closed resolved representation of Slide.background consumed exhaustively by emit and paint. */
export type ResolvedSlideBackground =
  | { readonly kind: "default" }
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
  | { readonly kind: "image"; readonly image: LoadedImage };

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function resolveSlideBackground(
  theme: ThemeContext,
  background: SlideBackground | undefined,
): ResolvedSlideBackground {
  if (background === undefined) {
    return { kind: "default" };
  }

  const backgroundRecord: Record<string, unknown> | undefined = isRecord(
    background,
  )
    ? background
    : undefined;

  if (
    backgroundRecord !== undefined &&
    Object.hasOwn(backgroundRecord, "image")
  ) {
    const unknown = Object.keys(backgroundRecord).find(
      (key) => key !== "image",
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "invalid-slide-background",
        `Slide.background contains unsupported parameter ${unknown}`,
      );
    }
    return {
      kind: "image",
      image: readImage(backgroundRecord["image"]),
    };
  }

  if (
    backgroundRecord !== undefined &&
    (backgroundRecord["type"] === "linear-gradient" ||
      backgroundRecord["type"] === "radial-gradient")
  ) {
    return resolveGradientSpec(theme, backgroundRecord, undefined);
  }

  return {
    kind: "solid",
    color: theme.resolveColor(background as ColorValue),
  };
}

function emitSlideBackground(
  resolved: ResolvedSlideBackground,
  emitContext: SlideEmitContext,
): SlideEmitOutput {
  switch (resolved.kind) {
    case "default":
      return { xml: "" };
    case "solid":
      return { xml: slideBackgroundSolidXml(resolved.color) };
    case "linear-gradient":
    case "radial-gradient":
      return { xml: slideBackgroundGradientXml(resolved) };
    case "image": {
      const relationshipId = emitContext.registerResource({
        bytes: resolved.image.bytes,
        extension: resolved.image.extension,
        contentType: resolved.image.contentType,
        relationshipType: imageRelationship,
      });
      return {
        xml: slideBackgroundImageXml(relationshipId),
      };
    }
    default: {
      const exhaustive: never = resolved;
      return exhaustive;
    }
  }
}

function paintSlideBackground(
  resolved: ResolvedSlideBackground,
  paintContext: SlidePaintContext,
): SlidePaintOutput {
  const { width, height } = paintContext.dimensions;
  switch (resolved.kind) {
    case "default":
      return {
        svg: '<rect width="100%" height="100%" fill="white"/>',
      };
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
      let pngBytes: Uint8Array;
      try {
        pngBytes = decodeImageToPngBytes(resolved.image);
      } catch (error) {
        const message =
          error instanceof Error
            ? error.message
            : "Failed to decode slide background image for preview";
        throw new CapabilityError(
          "preview-corrupt-image",
          message,
          `${paintContext.path}/background`,
        );
      }
      const base64 = Buffer.from(pngBytes).toString("base64");
      return {
        svg: `<image x="0" y="0" width="${width}" height="${height}" preserveAspectRatio="none" href="data:image/png;base64,${base64}"/>`,
      };
    }
    default: {
      const exhaustive: never = resolved;
      return exhaustive;
    }
  }
}

/**
 * Bound slide-background module: validates and resolves background declarations once per build
 * and owns both PresentationML `<p:bg>` emission and static preview canvas painting.
 */
export const slideBackgroundCapability: SlideCapabilityModule = {
  declaration: "slide-background",
  slideProp: "background",
  slot: "background",
  always: true,
  bindSlide(input: CapabilityInput, context) {
    const background = input.props["background"] as SlideBackground | undefined;
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
      resolveSlide(resolveContext: SlideResolveContext): ResolvedSlide {
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
            return paintSlideBackground(resolved, paintContext);
          },
        };
      },
    };
  },
};
