import {
  CapabilityError,
  isMorphKey,
  layoutStyleKeys,
  parseCapabilityLayoutContract,
  type BoundObjectCapability,
  type BoundObjectCapabilityModule,
  type BoundObjectNativeCollaboration,
  type CapabilityInput,
  type CapabilityLayoutMeasurement,
  type CapabilityLayoutRequest,
  type CapabilityParseContext,
  type EmitContext,
  type EmitOutput,
  type GeometryBox,
  type MorphKey,
  type ObjectResolveContext,
  type PaintContext,
  type PaintOutput,
  type ResolvedObject,
} from "@pptxdk/core/capability";
import { assertXml10Characters, pxToEmu, xmlEscape } from "@pptxdk/emit";

import { motionAnnotations } from "../motion/motion-api.ts";
import { evaluateGeometry, type Geometry } from "../shared/geometry.ts";
import { nativeLinkXml, parseLinkTarget } from "../shared/identity.ts";
import {
  emitVisualModel,
  resolveVisualModel,
  visualSvgTransform,
  type ResolvedBorder,
  type ResolvedVisualModel,
} from "../shared/visual.ts";
import { imageVisualStyleKeys } from "./image-declaration.ts";
import {
  decodeImageToPngBytes,
  readImage,
  type LoadedImage,
} from "./image-decode.ts";
import { imageXml } from "./image-xml.ts";

export { readImage, type LoadedImage } from "./image-decode.ts";
export {
  createImageNode,
  imageVisualStyleKeys,
  type ImageDeclarationProps,
  type ImageSource,
  type ImageVisualStyle,
} from "./image-declaration.ts";

type ImageFit = "fill" | "cover" | "contain";

type ImageCrop = {
  readonly l: number;
  readonly t: number;
  readonly r: number;
  readonly b: number;
};

type ImagePlacement = {
  readonly box: GeometryBox;
  readonly sourceCrop: ImageCrop;
};

/** Closed resolved image model produced in resolve() and exhaustively consumed by emit and paint. */
type ResolvedImageModel = {
  readonly image: LoadedImage;
  readonly placement: ImagePlacement;
  readonly visual: ResolvedVisualModel;
  readonly name: string;
  readonly alt?: string;
};

function roundSvgNumber(value: number): string {
  const rounded = Math.round(value * 1_000_000) / 1_000_000;
  return Object.is(rounded, -0) ? "0" : String(rounded);
}

function borderStrokeAttributes(border: ResolvedBorder): string {
  const w = Math.max(1, border.width);
  let dashAttr = "";
  switch (border.style) {
    case "solid":
      dashAttr = "";
      break;
    case "dashed":
      dashAttr = ` stroke-dasharray="${roundSvgNumber(w * 4)} ${roundSvgNumber(w * 3)}"`;
      break;
    case "dotted":
      dashAttr = ` stroke-dasharray="${roundSvgNumber(w)} ${roundSvgNumber(w * 2)}"`;
      break;
    default: {
      const exhaustive: never = border.style;
      return exhaustive;
    }
  }
  return `stroke="#${border.color.hex}" stroke-opacity="${roundSvgNumber(border.color.alpha ?? 1)}" stroke-width="${roundSvgNumber(border.width)}"${dashAttr}`;
}

function emitImageModel(
  model: ResolvedImageModel,
  emitContext: EmitContext,
): EmitOutput {
  const properties = emitVisualModel(model.visual, undefined, {
    omitFill: true,
  });
  const escapedName = xmlEscape(model.name);
  const escapedAlt = model.alt === undefined ? undefined : xmlEscape(model.alt);
  let relationshipId: string;
  try {
    relationshipId = emitContext.registerResource({
      bytes: model.image.bytes,
      extension: model.image.extension,
      contentType: model.image.contentType,
      relationshipType:
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    });
  } catch (error) {
    if (error instanceof CapabilityError) throw error;
    throw new CapabilityError(
      "image-resource-failure",
      error instanceof Error ? error.message : String(error),
    );
  }
  const { sourceCrop } = model.placement;
  const serializedCrop = Object.fromEntries(
    (
      [
        ["l", sourceCrop.l],
        ["t", sourceCrop.t],
        ["r", sourceCrop.r],
        ["b", sourceCrop.b],
      ] as const
    ).filter(([, fraction]) => fraction !== 0),
  );
  return {
    xml: imageXml(
      emitContext.id,
      escapedName,
      escapedAlt,
      relationshipId,
      model.placement.box,
      serializedCrop,
      properties.xml,
      properties.opacity,
      nativeLinkXml(emitContext.nativeLink),
    ),
  };
}

function roundCropFraction(fraction: number): number {
  return Math.round(fraction * 100000) / 100000;
}

function paintImageModel(
  model: ResolvedImageModel,
  paintContext: PaintContext,
): PaintOutput {
  const { image, placement, visual } = model;
  const pngBytes = decodeImageToPngBytes(image);

  const { box, sourceCrop } = placement;
  const leftCrop = roundCropFraction(sourceCrop.l);
  const topCrop = roundCropFraction(sourceCrop.t);
  const rightCrop = roundCropFraction(sourceCrop.r);
  const bottomCrop = roundCropFraction(sourceCrop.b);
  const widthFactor = 1 - leftCrop - rightCrop;
  const heightFactor = 1 - topCrop - bottomCrop;

  const prefix = `img-s${paintContext.slideNumber}-o${paintContext.objectId}`;
  const defs: string[] = [];
  const evaluated = evaluateGeometry(visual.geometry, box);
  const fillablePaths = evaluated.paths.filter(
    (pathItem) => pathItem.fillMode !== "none" && pathItem.d !== "",
  );

  const opacity = visual.opacity ?? 1;
  let imageElement = "";
  if (widthFactor > 0 && heightFactor > 0 && fillablePaths.length > 0) {
    const imgWidth = box.width / widthFactor;
    const imgHeight = box.height / heightFactor;
    const imgX = box.x - leftCrop * imgWidth;
    const imgY = box.y - topCrop * imgHeight;
    const clipId = `${prefix}-clip`;
    const clipMarkup = fillablePaths
      .map((pathItem) => `<path d="${pathItem.d}"/>`)
      .join("");
    defs.push(
      `<clipPath id="${clipId}" clipPathUnits="userSpaceOnUse">${clipMarkup}</clipPath>`,
    );
    const opacityAttr =
      opacity === 1 ? "" : ` opacity="${roundSvgNumber(opacity)}"`;
    imageElement = `<image clip-path="url(#${clipId})" x="${roundSvgNumber(imgX)}" y="${roundSvgNumber(imgY)}" width="${roundSvgNumber(imgWidth)}" height="${roundSvgNumber(imgHeight)}" preserveAspectRatio="none"${opacityAttr} href="data:image/png;base64,${Buffer.from(pngBytes).toString("base64")}"/>`;

    if (visual.softEdge !== undefined && visual.softEdge > 0) {
      const r = visual.softEdge;
      const minSide = Math.min(box.width, box.height);
      const borderPad = visual.border === undefined ? 0 : visual.border.width;
      const pad = Math.ceil(r * 2 + borderPad + 8);
      const maskX = roundSvgNumber(box.x - pad);
      const maskY = roundSvgNumber(box.y - pad);
      const maskW = roundSvgNumber(box.width + pad * 2);
      const maskH = roundSvgNumber(box.height + pad * 2);
      const blurId = `${prefix}-soft-blur`;
      const maskId = `${prefix}-soft-mask`;
      const innerStrokeWidth =
        r >= minSide / 2 ? 0 : Math.min(r, minSide * 0.45);
      const maskAttrs =
        innerStrokeWidth > 0
          ? `fill="white" stroke="black" stroke-width="${roundSvgNumber(innerStrokeWidth)}" filter="url(#${blurId})"`
          : `fill="white" stroke="none" filter="url(#${blurId})"`;
      const maskShapes = fillablePaths
        .map((pathItem) => `<path d="${pathItem.d}" ${maskAttrs}/>`)
        .join("");
      defs.push(
        `<filter id="${blurId}" filterUnits="userSpaceOnUse" x="${maskX}" y="${maskY}" width="${maskW}" height="${maskH}"><feGaussianBlur stdDeviation="${roundSvgNumber(r / 2)}"/></filter>`,
      );
      defs.push(
        `<mask id="${maskId}" maskUnits="userSpaceOnUse" x="${maskX}" y="${maskY}" width="${maskW}" height="${maskH}">${maskShapes}</mask>`,
      );
      imageElement = `<g mask="url(#${maskId})">${imageElement}</g>`;
    }
  }

  let borderElements = "";
  if (visual.border !== undefined) {
    const strokeAttr = borderStrokeAttributes(visual.border);
    borderElements = evaluated.paths
      .filter((pathItem) => pathItem.stroke && pathItem.d !== "")
      .map((pathItem) => `<path d="${pathItem.d}" fill="none" ${strokeAttr}/>`)
      .join("");
  }

  let layer = `<g transform="${visualSvgTransform(visual)}">${imageElement}${borderElements}</g>`;

  if (visual.boxShadow !== undefined) {
    const shadow = visual.boxShadow;
    const cx = box.x + box.width / 2;
    const cy = box.y + box.height / 2;
    const halfDiag = Math.hypot(box.width, box.height) / 2;
    const rad = (shadow.angle * Math.PI) / 180;
    const dx = shadow.distance * Math.cos(rad);
    const dy = shadow.distance * Math.sin(rad);
    const sigma = shadow.blur / 2;
    const borderPad = visual.border === undefined ? 0 : visual.border.width;
    const pad = shadow.blur * 3 + Math.abs(dx) + Math.abs(dy) + borderPad + 16;
    const minX = roundSvgNumber(cx - halfDiag - pad);
    const minY = roundSvgNumber(cy - halfDiag - pad);
    const filterSize = roundSvgNumber((halfDiag + pad) * 2);
    const shadowId = `${prefix}-shadow`;
    defs.push(
      `<filter id="${shadowId}" filterUnits="userSpaceOnUse" x="${minX}" y="${minY}" width="${filterSize}" height="${filterSize}"><feDropShadow dx="${roundSvgNumber(dx)}" dy="${roundSvgNumber(dy)}" stdDeviation="${roundSvgNumber(sigma)}" flood-color="#${shadow.color.hex}" flood-opacity="${roundSvgNumber(shadow.color.alpha ?? 1)}"/></filter>`,
    );
    layer = `<g filter="url(#${shadowId})">${layer}</g>`;
  }

  const defsXml = defs.length > 0 ? `<defs>${defs.join("")}</defs>` : "";
  return {
    svg: `${defsXml}${layer}`,
  };
}

/** Shared image placement calculation for native OOXML output and static preview. */
function imagePlacement(
  image: { readonly width: number; readonly height: number },
  fit: ImageFit,
  crop: Partial<ImageCrop> | undefined,
  allocated: GeometryBox,
): ImagePlacement {
  const { width, height } = allocated;
  const box = { ...allocated };
  let l = crop?.l ?? 0;
  let t = crop?.t ?? 0;
  let r = crop?.r ?? 0;
  let b = crop?.b ?? 0;
  if (crop === undefined && fit === "cover") {
    const imageRatio = image.width / image.height;
    const boxRatio = width / height;
    if (imageRatio > boxRatio) {
      const excess = (1 - boxRatio / imageRatio) / 2;
      l = excess;
      r = excess;
    } else if (imageRatio < boxRatio) {
      const excess = (1 - imageRatio / boxRatio) / 2;
      t = excess;
      b = excess;
    }
  } else if (crop === undefined && fit === "contain") {
    const imageRatio = image.width / image.height;
    const boxRatio = width / height;
    if (imageRatio > boxRatio) {
      const fittedHeight = width / imageRatio;
      box.y += (height - fittedHeight) / 2;
      box.height = fittedHeight;
    } else if (imageRatio < boxRatio) {
      const fittedWidth = height * imageRatio;
      box.x += (width - fittedWidth) / 2;
      box.width = fittedWidth;
    }
  }
  // Dimensions derived by contain must also be representable as non-zero EMU.
  box.width = dimension(box.width, "width");
  box.height = dimension(box.height, "height");
  return { box, sourceCrop: { l, t, r, b } };
}

const allowedParameters = [
  "id",
  "name",
  "morphKey",
  "link",
  "src",
  "alt",
  "x",
  "y",
  "width",
  "height",
  "geometry",
  "flipH",
  "flipV",
  "style",
  "fit",
  "crop",
  "enter",
  "emphasis",
  "motion",
] as const;
const maximumShapeExtentEmu = 2_147_483_647;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function optionalText(
  props: Record<string, unknown>,
  name: "name",
  path?: string,
): string | undefined {
  const value = props[name];
  if (value === undefined) return undefined;
  if (typeof value !== "string" || value.trim() === "") {
    throw new CapabilityError(
      "invalid-parameter",
      `${name} must be a non-empty string`,
      path,
    );
  }
  return value;
}

function dimension(
  value: unknown,
  name: "x" | "y" | "width" | "height",
  fallback?: number,
): number {
  const actual = value === undefined ? fallback : value;
  const isExtent = name === "width" || name === "height";
  if (
    typeof actual !== "number" ||
    !Number.isFinite(actual) ||
    (isExtent ? actual <= 0 : actual < -1_000_000 || actual > 1_000_000)
  ) {
    throw new CapabilityError(
      "invalid-coordinate",
      `${name} must be a finite pixel number, width and height must be greater than 0, and coordinates must be between -1000000 and 1000000`,
    );
  }
  if (isExtent) {
    const extentEmu = pxToEmu(actual);
    if (extentEmu < 1 || extentEmu > maximumShapeExtentEmu) {
      throw new CapabilityError(
        "invalid-coordinate",
        `${name} must convert to an EMU value between 1 and ${maximumShapeExtentEmu} inclusive`,
      );
    }
  }
  return actual;
}

function requestedDimension(
  style: Readonly<Record<string, unknown>>,
  props: Readonly<Record<string, unknown>>,
  name: "width" | "height",
): unknown {
  const styled = style[name];
  return styled === undefined
    ? props[name]
    : typeof styled === "number"
      ? styled
      : undefined;
}

function cropValues(
  value: unknown,
  path?: string,
): Partial<ImageCrop> | undefined {
  if (value === undefined) return undefined;
  if (!isRecord(value)) {
    throw new CapabilityError(
      "invalid-crop",
      "crop must be a four-sided fraction object",
      path,
    );
  }
  const crop = value;
  const allowed = ["left", "top", "right", "bottom"] as const;
  const unknown = Object.keys(crop).find(
    (key) => !(allowed as readonly string[]).includes(key),
  );
  if (unknown !== undefined) {
    throw new CapabilityError(
      "unknown-parameter",
      `crop contains unknown parameter ${unknown}`,
      path,
    );
  }
  const result: { l?: number; t?: number; r?: number; b?: number } = {};
  const sideMap = {
    left: "l",
    top: "t",
    right: "r",
    bottom: "b",
  } as const;
  for (const side of allowed) {
    const fraction = crop[side];
    if (fraction === undefined) continue;
    if (
      typeof fraction !== "number" ||
      !Number.isFinite(fraction) ||
      fraction < 0 ||
      fraction > 1
    ) {
      throw new CapabilityError(
        "invalid-crop",
        `crop.${side} must be a finite fraction between 0 and 1`,
        path,
      );
    }
    result[sideMap[side]] = fraction;
  }
  if ((result.l ?? 0) + (result.r ?? 0) > 1) {
    throw new CapabilityError(
      "invalid-crop",
      "Sum of crop left and right cannot exceed 1",
      path,
    );
  }
  if ((result.t ?? 0) + (result.b ?? 0) > 1) {
    throw new CapabilityError(
      "invalid-crop",
      "Sum of crop top and bottom cannot exceed 1",
      path,
    );
  }
  return result;
}

/** Image module reads and validates formal resources, encapsulates bound state, and emits native picture objects. */
export const imageCapability: BoundObjectCapabilityModule = {
  declaration: "image",

  bind(
    input: CapabilityInput,
    context: CapabilityParseContext,
  ): BoundObjectCapability {
    const unknown = Object.keys(input.props).find(
      (key) => !(allowedParameters as readonly string[]).includes(key),
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-parameter",
        `Image contains unsupported parameter ${unknown}`,
      );
    }
    if (
      input.children.some(
        (child) => "props" in child || child.text.trim() !== "",
      )
    ) {
      throw new CapabilityError(
        "unsupported-children",
        "Image does not accept children",
      );
    }

    const requestedName = optionalText(input.props, "name");
    const rawMorphKey = input.props["morphKey"];
    let morphKey: MorphKey | undefined;
    if (rawMorphKey !== undefined) {
      if (!isMorphKey(rawMorphKey)) {
        throw new CapabilityError(
          "invalid-morph-key",
          "morphKey must be created by createMorphScope().handle(name)",
        );
      }
      morphKey = rawMorphKey;
    }
    const link = parseLinkTarget(input.props["link"]);

    const altValue = input.props["alt"];
    if (altValue !== undefined && typeof altValue !== "string") {
      throw new CapabilityError("invalid-parameter", "alt must be a string");
    }

    const allowedStyleKeys = new Set<string>([
      ...layoutStyleKeys,
      ...imageVisualStyleKeys,
    ]);

    const contract = parseCapabilityLayoutContract({
      nativeKind: "pic",
      props: input.props,
      allowedStyleKeys,
      allowPadding: false,
      path: context.path,
    });

    const imageStyle = isRecord(input.props["style"])
      ? { ...input.props["style"] }
      : {};

    const requestedFit = input.props["fit"];
    const fit = requestedFit === undefined ? "fill" : requestedFit;
    if (fit !== "fill" && fit !== "cover" && fit !== "contain") {
      throw new CapabilityError(
        "unsupported-fit",
        "fit must be fill, cover, or contain",
      );
    }

    const crop = cropValues(input.props["crop"]);
    const image = readImage(input.props["src"]);
    const geometry = input.props["geometry"] as Geometry | undefined;
    const flipH = input.props["flipH"] as boolean | undefined;
    const flipV = input.props["flipV"] as boolean | undefined;

    // Validate geometry and visual properties during bind before layout runs.
    resolveVisualModel({
      theme: context.theme,
      style: imageStyle,
      box: { x: 0, y: 0, width: 100, height: 100 },
      ...(geometry !== undefined ? { geometry } : {}),
      ...(flipH !== undefined ? { flipH } : {}),
      ...(flipV !== undefined ? { flipV } : {}),
      image: true,
    });

    const declaredX = input.props["x"];
    const declaredY = input.props["y"];
    const requestedWidth = requestedDimension(imageStyle, input.props, "width");
    const requestedHeight = requestedDimension(
      imageStyle,
      input.props,
      "height",
    );
    const authorId =
      typeof input.props["id"] === "string" ? input.props["id"] : undefined;
    const motionProps: Readonly<Record<string, unknown>> = {
      ...(input.props["enter"] !== undefined
        ? { enter: input.props["enter"] }
        : {}),
      ...(input.props["emphasis"] !== undefined
        ? { emphasis: input.props["emphasis"] }
        : {}),
      ...(input.props["motion"] !== undefined
        ? { motion: input.props["motion"] }
        : {}),
    };

    const native: BoundObjectNativeCollaboration = {
      ...(morphKey === undefined ? {} : { morphKey }),
      ...(requestedName === undefined ? {} : { hasExplicitName: true }),
      ...(link === undefined ? {} : { link }),
    };

    return {
      contract,
      ...(morphKey !== undefined ||
      requestedName !== undefined ||
      link !== undefined
        ? { native }
        : {}),

      measureLayout(
        request: CapabilityLayoutRequest,
      ): CapabilityLayoutMeasurement {
        const declaredWidth =
          typeof requestedWidth === "number" ? requestedWidth : undefined;
        const declaredHeight =
          typeof requestedHeight === "number" ? requestedHeight : undefined;
        const width =
          declaredWidth ??
          Math.min(image.width, request.availableWidth ?? image.width);
        const height = declaredHeight ?? (width * image.height) / image.width;
        return {
          width,
          height,
          minContentWidth: image.width,
          minContentHeight: image.height,
        };
      },

      resolve(resolveContext: ObjectResolveContext): ResolvedObject {
        const x = dimension(resolveContext.layoutBox?.x ?? declaredX, "x", 0);
        const y = dimension(resolveContext.layoutBox?.y ?? declaredY, "y", 0);
        let width: number;
        let height: number;
        if (resolveContext.layoutBox !== undefined) {
          width = dimension(resolveContext.layoutBox.width, "width");
          height = dimension(resolveContext.layoutBox.height, "height");
        } else if (
          requestedWidth === undefined &&
          requestedHeight === undefined
        ) {
          width = image.width;
          height = image.height;
        } else if (requestedWidth === undefined) {
          height = dimension(requestedHeight, "height");
          width = (height * image.width) / image.height;
        } else if (requestedHeight === undefined) {
          width = dimension(requestedWidth, "width");
          height = (width * image.height) / image.width;
        } else {
          width = dimension(requestedWidth, "width");
          height = dimension(requestedHeight, "height");
        }
        width = dimension(width, "width");
        height = dimension(height, "height");

        const placement = imagePlacement(image, fit, crop, {
          x,
          y,
          width,
          height,
        });
        const { box } = placement;
        const resolvedVisual = resolveVisualModel({
          theme: resolveContext.theme,
          style: imageStyle,
          box,
          ...(geometry === undefined ? {} : { geometry }),
          ...(flipH === undefined ? {} : { flipH }),
          ...(flipV === undefined ? {} : { flipV }),
          image: true,
        });
        const name =
          resolveContext.morphName ??
          requestedName ??
          `Image ${resolveContext.id - 1}`;
        try {
          assertXml10Characters(name);
          if (altValue !== undefined) assertXml10Characters(altValue);
        } catch {
          throw new CapabilityError(
            "invalid-xml-character",
            "Image.name or alt contains characters not allowed in XML 1.0",
          );
        }
        const angle = resolvedVisual.rotate;
        const annotations = motionAnnotations(motionProps, {
          kind: "image",
          name,
          ...(authorId === undefined ? {} : { authorId }),
          textBody: false,
          filled: false,
          box,
          rotationDegrees: angle,
          hasBackground:
            imageStyle["borderColor"] !== undefined ||
            (typeof imageStyle["borderWidth"] === "number" &&
              imageStyle["borderWidth"] > 0),
        });
        const model: ResolvedImageModel = {
          image,
          placement,
          visual: resolvedVisual,
          name,
          ...(altValue === undefined ? {} : { alt: altValue }),
        };

        return {
          observation: {
            id: resolveContext.id,
            kind: "image",
            name,
            box,
          },
          transform: {
            box,
            angle,
            flipH: resolvedVisual.flipH,
            flipV: resolvedVisual.flipV,
          },
          extensions: [],
          ...(resolvedVisual.warnings.length === 0
            ? {}
            : { warnings: resolvedVisual.warnings }),
          ...(annotations.length === 0 ? {} : { annotations }),

          emit(emitContext: EmitContext): EmitOutput {
            return emitImageModel(model, emitContext);
          },

          paint(paintContext: PaintContext): PaintOutput {
            return paintImageModel(model, paintContext);
          },
        };
      },
    };
  },
};
