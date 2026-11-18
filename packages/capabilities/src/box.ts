import {
  CapabilityError,
  layoutStyleKeys,
  paddingStyleKeys,
  parseCapabilityLayoutContract,
  textStyleFrom,
  textStyleKeys,
  xmlEscape,
  type BoundContainerCollaboration,
  type BoundObjectCapability,
  type BoundObjectCapabilityModule,
  type BoundObjectNativeCollaboration,
  type CapabilityContext,
  type CapabilityInput,
  type CapabilityParseContext,
  type GeometryBox,
  type LayoutStyle,
  type LinkTarget,
  type ObjectPreviewContext,
  type ObjectPreviewOutput,
  type TextStyle,
  type ThemeScope,
  type VisualStyle,
} from "@react-pptx/core/capability";
import { CommonContainer } from "@react-pptx/core/document";

import { nativeLinkXml, parseLinkTarget, parseMorphKey } from "./identity.ts";
import type { MorphProps } from "./morph-api.ts";
import {
  motionAnnotations,
  validateMotionFlag,
  type MotionProps,
} from "./motion-api.ts";
import {
  hasVisibleBoxStyle,
  resolveBoxPreviewVisuals,
  shapePropertiesXml,
  visualStyleKeys,
} from "./visual.ts";

/** Layout container appearance with visual styling that optionally generates a native background shape. */
export type BoxDeclarationProps = MorphProps &
  Pick<MotionProps, "enter" | "motion"> & {
    readonly id?: string;
    readonly link?: LinkTarget;
    readonly style?: LayoutStyle & TextStyle & VisualStyle;
    readonly theme?: ThemeScope;
  };

/** Creates the editable layout container shared by the command factory and the Box JSX element. */
export function createBoxNode(
  props: BoxDeclarationProps,
): CommonContainer<
  "box",
  BoxDeclarationProps,
  NonNullable<BoxDeclarationProps["style"]>
> {
  return new CommonContainer<
    "box",
    BoxDeclarationProps,
    NonNullable<BoxDeclarationProps["style"]>
  >("box", "rp:box", props);
}

const boxAllowedParameters = [
  "id",
  "style",
  "theme",
  "morphKey",
  "link",
  "enter",
  "motion",
] as const;

const boxAllowedStyleKeys = new Set<string>([
  ...layoutStyleKeys,
  ...paddingStyleKeys,
  ...textStyleKeys,
  ...visualStyleKeys,
]);

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
  properties: string,
  linkXml: string,
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${xmlEscape(name)}">${linkXml}</p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr>${properties}</p:sp>`;
}

/** Visible Box backgrounds use the same shape property writer; boxes without fill, stroke, or shadow retain layout semantics only. */
export const boxCapability: BoundObjectCapabilityModule = {
  declaration: "rp:box",

  bind(
    input: CapabilityInput,
    context: CapabilityParseContext,
  ): BoundObjectCapability {
    const unknown = Object.keys(input.props).find(
      (key) => !(boxAllowedParameters as readonly string[]).includes(key),
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-parameter",
        `Box contains unknown parameter ${unknown}`,
        context.path,
      );
    }
    validateMotionFlag(input.props["motion"]);
    const morphKey = parseMorphKey(input.props["morphKey"], context.path);
    const link = parseLinkTarget(input.props["link"], context.path);
    const style = { ...styleRecord(input.props["style"]) };
    const hasVisibleStyle = hasVisibleBoxStyle(style);
    if (morphKey !== undefined && !hasVisibleStyle) {
      throw new CapabilityError(
        "morph-needs-object",
        "Only visible Box elements generating native background shapes can participate in Morph",
        context.path,
      );
    }

    const contract = parseCapabilityLayoutContract({
      nativeKind: "sp",
      props: input.props,
      allowedStyleKeys: boxAllowedStyleKeys,
      allowPadding: true,
      isContainer: true,
      path: context.path,
    });

    const theme = context.theme.with(
      input.props["theme"] as ThemeScope | undefined,
      textStyleFrom(style),
    );

    shapePropertiesXml({
      theme,
      style,
      box: { x: 0, y: 0, width: 100, height: 100 },
      layoutOnlyWhenInvisible: true,
    });

    const authorId =
      typeof input.props["id"] === "string" ? input.props["id"] : undefined;
    const motionProps: Readonly<Record<string, unknown>> = {
      ...(input.props["enter"] !== undefined
        ? { enter: input.props["enter"] }
        : {}),
      ...(input.props["motion"] !== undefined
        ? { motion: input.props["motion"] }
        : {}),
    };

    const container: BoundContainerCollaboration = {
      children: input.children,
      theme,
      emitsObject: hasVisibleStyle,
      observesLayout: true,
      propagatesBaseline: true,
      checksCanvasOverflow: true,
    };

    const native: BoundObjectNativeCollaboration = {
      ...(morphKey === undefined ? {} : { morphKey }),
      ...(link === undefined ? {} : { link }),
    };

    let compiledBox: GeometryBox | undefined;

    return {
      contract,
      container,
      ...(morphKey !== undefined || link !== undefined ? { native } : {}),
      compile(compileContext: CapabilityContext) {
        const box = compileContext.layoutBox;
        if (box === undefined) {
          throw new CapabilityError(
            "layout-unavailable",
            "Box is missing layout results",
            compileContext.path,
          );
        }
        compiledBox = box;
        const name = compileContext.morphName ?? `Box ${compileContext.id - 1}`;
        const properties = shapePropertiesXml({
          theme: compileContext.theme,
          style,
          box,
          layoutOnlyWhenInvisible: true,
        });
        const motionTarget = {
          kind: "shape",
          name,
          ...(authorId === undefined ? {} : { authorId }),
          textBody: false,
          filled:
            style["backgroundColor"] !== undefined ||
            style["backgroundImage"] !== undefined,
        } as const;
        const annotations = properties.hasVisibleStyle
          ? motionAnnotations(motionProps, motionTarget)
          : [];
        return {
          xml: properties.hasVisibleStyle
            ? objectXml(
                compileContext.id,
                name,
                properties.xml,
                nativeLinkXml(compileContext.nativeLink),
              )
            : "",
          observation: {
            id: compileContext.id,
            kind: "shape",
            name,
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
      preview(previewContext: ObjectPreviewContext): ObjectPreviewOutput {
        const childrenSvg = previewContext.childrenSvg ?? "";
        if (!hasVisibleStyle) {
          return { svg: childrenSvg };
        }
        const box = previewContext.box ?? compiledBox;
        if (box === undefined) {
          throw new CapabilityError(
            "layout-unavailable",
            "Preview requires compiled Box geometry",
            previewContext.path,
          );
        }
        const resolved = resolveBoxPreviewVisuals({
          label: "box",
          visualStyle: style,
          box,
          theme: previewContext.theme,
        });
        const backgroundSvg = `<g transform="${resolved.svgTransform}"><rect x="${box.x}" y="${box.y}" width="${box.width}" height="${box.height}" ${resolved.attributes}/></g>`;
        return {
          svg: `${backgroundSvg}${childrenSvg}`,
          transform: resolved.transform,
        };
      },
    };
  },
};
