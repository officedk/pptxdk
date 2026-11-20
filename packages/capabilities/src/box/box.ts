import {
  CapabilityError,
  layoutStyleKeys,
  paddingStyleKeys,
  parseCapabilityLayoutContract,
  textStyleFrom,
  textStyleKeys,
  type BoundContainerCollaboration,
  type BoundObjectCapability,
  type BoundObjectNativeCollaboration,
  type CapabilityInput,
  type CapabilityParseContext,
  type EmitContext,
  type EmitOutput,
  type LayoutStyle,
  type LinkTarget,
  type ObjectCapabilityModule,
  type ObjectResolveContext,
  type PaintContext,
  type PaintOutput,
  type PreviewTransform,
  type ResolvedObject,
  type TextStyle,
  type ThemeScope,
  type VisualStyle,
} from "@react-pptx/core/capability";
import { CommonContainer } from "@react-pptx/core/document";
import { xmlEscape } from "@react-pptx/emit";

import {
  motionContainerAnnotations,
  validateMotionFlag,
  type MotionProps,
} from "../motion/motion-api.ts";
import {
  nativeLinkXml,
  parseLinkTarget,
  parseMorphKey,
} from "../shared/identity.ts";
import type { MorphProps } from "../shared/morph-api.ts";
import {
  emitVisualModel,
  hasVisibleBoxStyle,
  paintVisualModel,
  resolveVisualModel,
  visualStyleKeys,
} from "../shared/visual.ts";

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
  escapedName: string,
  properties: string,
  linkXml: string,
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapedName}">${linkXml}</p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr>${properties}</p:sp>`;
}

/** Visible Box backgrounds use the same shape property writer; boxes without fill, stroke, or shadow retain layout semantics only. */
export const boxCapability: ObjectCapabilityModule = {
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
      );
    }
    validateMotionFlag(input.props["motion"]);
    const morphKey = parseMorphKey(input.props["morphKey"]);
    const link = parseLinkTarget(input.props["link"]);
    const style = { ...styleRecord(input.props["style"]) };
    const hasVisibleStyle = hasVisibleBoxStyle(style);
    if (morphKey !== undefined && !hasVisibleStyle) {
      throw new CapabilityError(
        "morph-needs-object",
        "Only visible Box elements generating native background shapes can participate in Morph",
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

    resolveVisualModel({
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

    return {
      contract,
      container,
      ...(morphKey !== undefined || link !== undefined ? { native } : {}),
      resolve(resolveContext: ObjectResolveContext): ResolvedObject {
        const box = resolveContext.layoutBox;
        if (box === undefined) {
          throw new CapabilityError(
            "layout-unavailable",
            "Box is missing layout results",
          );
        }
        const name = resolveContext.morphName ?? `Box ${resolveContext.id - 1}`;
        let escapedName: string;
        try {
          escapedName = xmlEscape(name);
        } catch {
          throw new CapabilityError(
            "invalid-xml-character",
            "Box.name contains characters not allowed in XML 1.0",
          );
        }
        const visual = resolveVisualModel({
          theme: resolveContext.theme,
          style,
          box,
          layoutOnlyWhenInvisible: true,
        });
        const geometryName =
          visual.geometry.kind === "custom" ? "custom" : visual.geometry.name;
        const motionTarget = {
          kind: "shape",
          name,
          ...(authorId === undefined ? {} : { authorId }),
          textBody: false,
          filled:
            style["backgroundColor"] !== undefined ||
            style["backgroundImage"] !== undefined,
        } as const;
        const annotations = motionContainerAnnotations(
          motionProps,
          visual.hasVisibleStyle ? motionTarget : undefined,
          { propagatesDefaultEnter: true },
        );
        const transform: PreviewTransform | undefined = visual.hasVisibleStyle
          ? {
              box,
              angle: visual.rotate,
              flipH: visual.flipH,
              flipV: visual.flipV,
            }
          : undefined;
        return {
          observation: {
            id: resolveContext.id,
            kind: "shape",
            name,
            geometry: geometryName,
            box,
          },
          ...(transform === undefined ? {} : { transform }),
          extensions: [],
          ...(annotations.length === 0 ? {} : { annotations }),
          ...(visual.warnings.length === 0
            ? {}
            : { warnings: visual.warnings }),
          emit(emitContext: EmitContext): EmitOutput {
            return {
              xml: visual.hasVisibleStyle
                ? objectXml(
                    emitContext.id,
                    escapedName,
                    emitVisualModel(visual, emitContext.registerResource).xml,
                    nativeLinkXml(emitContext.nativeLink),
                  )
                : "",
            };
          },
          paint(paintContext: PaintContext): PaintOutput {
            const childrenSvg = paintContext.childrenSvg ?? "";
            if (!visual.hasVisibleStyle) {
              return { svg: childrenSvg };
            }
            const background = paintVisualModel({
              model: visual,
              slideNumber: paintContext.slideNumber,
              objectId: paintContext.objectId,
            });
            return {
              svg: `${background.svg}${childrenSvg}`,
            };
          },
        };
      },
    };
  },
};
