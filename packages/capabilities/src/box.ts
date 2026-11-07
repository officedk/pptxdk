import {
  CapabilityError,
  xmlEscape,
  type CapabilityInput,
  type CapabilityModule,
  type GeometryBox,
} from "@react-pptx/core";

import { nativeLinkXml } from "./identity.ts";
import { motionAnnotations } from "./motion-api.ts";
import { shapePropertiesXml } from "./visual.ts";

function styleRecord(value: unknown): Record<string, unknown> {
  if (value === undefined) return {};
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new CapabilityError("invalid-style", "Box.style must be an object");
  }
  return value as Record<string, unknown>;
}

function objectXml(
  id: number,
  name: string,
  box: GeometryBox,
  properties: string,
  linkXml: string,
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${xmlEscape(name)}">${linkXml}</p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr>${properties}</p:sp>`;
}

/** Visible Box backgrounds use the same shape property writer; boxes without fill, stroke, or shadow retain layout semantics only. */
export const boxCapability: CapabilityModule = {
  declaration: "rp:box",
  compile(input: CapabilityInput, context) {
    const unknown = Object.keys(input.props).find(
      (key) =>
        ![
          "testId",
          "part",
          "style",
          "theme",
          "morphKey",
          "link",
          "__motionContext",
        ].includes(key),
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-parameter",
        `Box contains unsupported parameter ${unknown}`,
      );
    }
    const box = context.layoutBox;
    if (box === undefined)
      throw new CapabilityError(
        "layout-unavailable",
        "Box is missing layout results",
      );
    const testId = input.props["testId"];
    if (
      testId !== undefined &&
      (typeof testId !== "string" || testId.trim() === "")
    ) {
      throw new CapabilityError(
        "invalid-parameter",
        "testId must be a non-empty string",
      );
    }
    const part = input.props["part"];
    if (
      part !== undefined &&
      (typeof part !== "string" || part.trim() === "")
    ) {
      throw new CapabilityError(
        "invalid-parameter",
        "part must be a non-empty string",
      );
    }
    const name =
      context.morphName ??
      (typeof testId === "string" ? testId : `Box ${context.id - 1}`);
    const style = styleRecord(input.props["style"]);
    const properties = shapePropertiesXml({
      theme: context.theme,
      style,
      box,
      layoutOnlyWhenInvisible: true,
    });
    const motionTarget = {
      kind: "shape",
      name,
      ...(typeof testId === "string" ? { testId } : {}),
      textBody: false,
      filled:
        style["backgroundColor"] !== undefined ||
        style["backgroundImage"] !== undefined,
    } as const;
    const annotations = properties.hasVisibleStyle
      ? motionAnnotations(input.props, motionTarget)
      : motionAnnotations(
          { motion: false, __motionContext: input.props["__motionContext"] },
          motionTarget,
        );
    return {
      xml: properties.hasVisibleStyle
        ? objectXml(
            context.id,
            name,
            box,
            properties.xml,
            nativeLinkXml(context.nativeLink),
          )
        : "",
      observation: {
        id: context.id,
        kind: "shape",
        name,
        ...(typeof testId === "string" ? { testId } : {}),
        ...(typeof part === "string" ? { part } : {}),
        geometry: properties.geometryName,
        box,
      },
      extensions: [],
      ...(annotations.length === 0 ? {} : { annotations }),
      ...(properties.warnings.length === 0
        ? {}
        : { warnings: properties.warnings }),
    };
  },
};
