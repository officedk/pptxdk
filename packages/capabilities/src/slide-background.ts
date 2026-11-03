import {
  CapabilityError,
  xmlEscape,
  type CapabilityInput,
  type SlideBackground,
  type SlideCapabilityModule,
} from "@react-pptx/core";

import { readImage } from "./image.ts";
import { fillXml } from "./visual.ts";

const imageRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** 页面背景共用主题解析、渐变 writer 与图片资源登记服务。 */
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
            `Slide.background 含未实现参数 ${unknown}`,
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
          xml: `<p:bg><p:bgPr><a:blipFill><a:blip r:embed="${xmlEscape(relationshipId)}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>`,
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
        xml: `<p:bg><p:bgPr>${fillXml(context.theme, style, undefined, undefined)}<a:effectLst/></p:bgPr></p:bg>`,
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
