import {
  CapabilityError,
  type BoundSlideCapabilityModule,
  type CapabilityInput,
  type ColorValue,
  type SlideBackground,
} from "@react-pptx/core/capability";

import { readImage } from "./image.ts";
import {
  slideBackgroundFillXml,
  slideBackgroundImageXml,
} from "./slide-background-xml.ts";

const imageRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * Bound slide-background module: validates and resolves background declarations once per build
 * and owns both PresentationML `<p:bg>` compilation and static preview canvas rendering.
 */
export const slideBackgroundCapability: BoundSlideCapabilityModule = {
  declaration: "slide-background",
  slideProp: "background",
  slot: "background",
  always: true,
  bindSlide(input: CapabilityInput, context) {
    const background = input.props["background"] as SlideBackground | undefined;
    if (background === undefined) {
      return {
        compileSlide() {
          return { xml: "", extensions: [] };
        },
        previewSlide() {
          return { svg: '<rect width="100%" height="100%" fill="white"/>' };
        },
      };
    }

    const backgroundRecord: Record<string, unknown> | undefined = isRecord(
      background,
    )
      ? background
      : undefined;

    try {
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
        const image = readImage(backgroundRecord["image"]);
        return {
          compileSlide(compileContext) {
            const relationshipId = compileContext.registerResource({
              bytes: image.bytes,
              extension: image.extension,
              contentType: image.contentType,
              relationshipType: imageRelationship,
            });
            return {
              xml: slideBackgroundImageXml(relationshipId),
              extensions: [],
            };
          },
          previewSlide(previewContext) {
            throw new CapabilityError(
              "preview-unsupported-background",
              "Preview does not yet support image or gradient backgrounds",
              `${previewContext.path}/background`,
            );
          },
        };
      }

      const isGradient =
        backgroundRecord !== undefined &&
        (backgroundRecord["type"] === "linear-gradient" ||
          backgroundRecord["type"] === "radial-gradient");

      if (isGradient) {
        const xml = slideBackgroundFillXml(context.theme, {
          backgroundImage: backgroundRecord,
        });
        return {
          compileSlide() {
            return { xml, extensions: [] };
          },
          previewSlide(previewContext) {
            throw new CapabilityError(
              "preview-unsupported-background",
              "Preview does not yet support image or gradient backgrounds",
              `${previewContext.path}/background`,
            );
          },
        };
      }

      const xml = slideBackgroundFillXml(context.theme, {
        backgroundColor: background,
      });
      const color = context.theme.resolveColor(background as ColorValue);
      return {
        compileSlide() {
          return { xml, extensions: [] };
        },
        previewSlide() {
          return {
            svg: `<rect width="100%" height="100%" fill="#${color.hex}" fill-opacity="${color.alpha ?? 1}"/>`,
          };
        },
      };
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
  },
};
