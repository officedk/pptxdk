import {
  CapabilityError,
  type CapabilityInput,
  type SlideBackground,
  type SlideCapabilityModule,
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

/** Slide backgrounds share theme resolution, gradient writing, and image resource registration services. */
export const slideBackgroundCapability: SlideCapabilityModule = {
  declaration: "slide-background",
  slideProp: "background",
  slot: "background",
  compileSlide(input: CapabilityInput, context) {
    const background = input.props["background"] as SlideBackground | undefined;
    if (background === undefined) return { xml: "", extensions: [] };
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
        const relationshipId = context.registerResource({
          bytes: image.bytes,
          extension: image.extension,
          contentType: image.contentType,
          relationshipType: imageRelationship,
        });
        return {
          xml: slideBackgroundImageXml(relationshipId),
          extensions: [],
        };
      }
      const style =
        backgroundRecord !== undefined &&
        (backgroundRecord["type"] === "linear-gradient" ||
          backgroundRecord["type"] === "radial-gradient")
          ? { backgroundImage: backgroundRecord }
          : { backgroundColor: background };
      return {
        xml: slideBackgroundFillXml(context.theme, style),
        extensions: [],
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
