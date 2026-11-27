import {
  CapabilityError,
  matchLayoutPlaceholderKey,
  matchMasterPlaceholderSpec,
  nativePlaceholderRefSchema,
  parseCapabilityLayoutContract,
  type BoundObjectCapability,
  type BoundObjectNativeCollaboration,
  type CapabilityInput,
  type CapabilityParseContext,
  type EmitContext,
  type EmitOutput,
  type Rect,
  type NativePlaceholderRef,
  type ObjectCapabilityModule,
  type ObjectResolveContext,
  type PaintOutput,
  type ObjectResolution,
} from "@pptxdk/core/capability";
import {
  assertXml10Characters,
  placeholderRefXml,
  xmlEscape,
} from "@pptxdk/emit";
import {
  A_NS,
  EMU_PER_PX,
  extentEmuFragmentSchema,
  offsetEmuFragmentSchema,
  ooxmlBoolSchema,
  ooxmlNonNegativeIntSchema,
  P_NS,
  transformAttributesFragmentSchema,
  defineCanvasMechanism,
  emuToPx,
  type ImportDiagnostic,
  type ImportLocation,
  type ShapeTreeImportContext,
  type ShapeTreeImportHandler,
  type ShapeTreeImportOutcome,
} from "@pptxdk/import";
import { z } from "zod";

import {
  nativeLinkXml,
  parseLinkTarget,
  parseMorphKey,
  parseOptionalName,
  readClickLink,
  type ClickLinkReadResult,
} from "../shared/identity.ts";
import {
  directElementChild,
  isPlaceholderSpElement,
  readNvPrPlaceholder,
} from "../shared/placeholder-import.ts";
import {
  hasLocalDeclaration,
  resolvePlaceholderGeometry,
  resolvePlaceholderVisualStyle,
} from "../shared/style-inheritance.ts";
import {
  assertVisualBoxExtent,
  emitPlaceholderVisualSpPr,
  readSpPrAppearance,
  resolveVisualModel,
  type PlaceholderSpPrFlags,
  type ResolvedVisualModel,
} from "../shared/visual.ts";
import { placeholderCodegen } from "./placeholder-codegen.ts";
import {
  createPlaceholderNode,
  placeholderAllowedPropKeys,
  placeholderDeclarationPropsSchema,
  placeholderObjectStyleSchema,
  type PlaceholderDeclarationProps,
  type PlaceholderNode,
} from "./placeholder-schemas.ts";

export {
  createPlaceholderNode,
  placeholderDeclarationPropsSchema,
  placeholderObjectStyleSchema,
  type PlaceholderDeclarationProps,
  type PlaceholderNode,
};

const NonVisualSpIdSchema = z.object({
  id: ooxmlNonNegativeIntSchema,
  name: z.string().optional(),
});

const NonVisualSpPrFlagsSchema = z.object({
  txBox: ooxmlBoolSchema.optional(),
});

const placeholderAllowedStyleKeys = new Set<string>(
  Object.keys(placeholderObjectStyleSchema.unwrap().shape),
);

type ResolvedPlaceholderModel = {
  readonly box: Rect;
  readonly name: string;
  readonly nativeRef: NativePlaceholderRef;
  readonly visual: ResolvedVisualModel;
  readonly spPrFlags: PlaceholderSpPrFlags;
};

function emitPlaceholderModel(
  model: ResolvedPlaceholderModel,
  emitContext: EmitContext,
): EmitOutput {
  const escapedName = xmlEscape(model.name);
  const spPrXml = emitPlaceholderVisualSpPr(
    model.visual,
    model.spPrFlags,
    (resource) => emitContext.registerResource(resource),
  );
  return {
    xml: `<p:sp><p:nvSpPr><p:cNvPr id="${emitContext.id}" name="${escapedName}">${nativeLinkXml(emitContext.nativeLink)}</p:cNvPr><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${placeholderRefXml(model.nativeRef)}</p:nvPr></p:nvSpPr>${spPrXml}</p:sp>`,
  };
}

function paintPlaceholderModel(): PaintOutput {
  // Unpopulated content placeholders show prompt chrome only inside interactive editors and draw nothing on static preview.
  return { svg: "" };
}

/**
 * Reads an unpopulated `<p:sp>` content/object placeholder (no `<p:txBody>`)
 * into a layout-bound `PlaceholderNode` while preserving `<p:ph>` binding.
 */
export function readContentPlaceholderSp(
  context: ShapeTreeImportContext,
): ShapeTreeImportOutcome {
  const { cursor, part, slideIndex, layoutPlaceholders } = context;
  const diagnostics: ImportDiagnostic[] = [];

  const nvSpPr = cursor.child(P_NS, "nvSpPr");
  if (nvSpPr === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part, xpath: cursor.xpath, slideIndex },
    });
    return { diagnostics };
  }

  let objectId: number | undefined;
  let objectName: string | undefined;
  let clickResult: ClickLinkReadResult | undefined;
  const cNvPr = nvSpPr.child(P_NS, "cNvPr");
  if (cNvPr !== undefined) {
    const rawCNvPr = cNvPr.attrsObject(["id", "name"]);
    const parsedCNvPr = NonVisualSpIdSchema.safeParse(rawCNvPr);
    if (!parsedCNvPr.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: { part, xpath: cNvPr.xpath, slideIndex },
        attributeOrElement: "p:cNvPr",
        rawValue: JSON.stringify(rawCNvPr),
      });
    } else {
      objectId = parsedCNvPr.data.id;
      if (
        parsedCNvPr.data.name !== undefined &&
        parsedCNvPr.data.name.trim() !== ""
      ) {
        objectName = parsedCNvPr.data.name;
      }
    }
    const hlinkClick = cNvPr.child(A_NS, "hlinkClick");
    if (hlinkClick !== undefined) {
      clickResult = readClickLink(hlinkClick, context, {
        ...(objectId === undefined ? {} : { objectId }),
        ...(objectName === undefined ? {} : { objectName }),
      });
      diagnostics.push(...clickResult.diagnostics);
    }
    cNvPr.ignoreEmptyChild(A_NS, "extLst");
  }

  const cNvSpPr = nvSpPr.child(P_NS, "cNvSpPr");
  if (cNvSpPr !== undefined) {
    NonVisualSpPrFlagsSchema.safeParse(cNvSpPr.attrsObject(["txBox"]));
    const spLocks = cNvSpPr.child(A_NS, "spLocks");
    spLocks?.attrsObject([
      "noGrp",
      "noRot",
      "noChangeAspect",
      "noMove",
      "noResize",
      "noSelect",
    ]);
  }

  const phLocation: ImportLocation = {
    part,
    xpath: cursor.xpath,
    slideIndex,
    ...(objectId === undefined ? {} : { objectId }),
    ...(objectName === undefined ? {} : { objectName }),
  };
  const nvPr = nvSpPr.child(P_NS, "nvPr");
  const phResult = readNvPrPlaceholder(nvPr, phLocation, diagnostics);
  if (phResult.kind === "none") {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: phLocation,
    });
    return { diagnostics };
  }
  if (phResult.kind === "error") {
    return { diagnostics };
  }
  const nativeRef = phResult.nativeRef;

  const spPr = cursor.child(P_NS, "spPr");
  if (spPr === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: phLocation,
    });
    return { diagnostics };
  }

  let localBox:
    | {
        readonly x: number;
        readonly y: number;
        readonly width: number;
        readonly height: number;
      }
    | undefined;
  let rotate: number | undefined;
  let flipH: boolean | undefined;
  let flipV: boolean | undefined;
  let fallbackUnits = { w: 0, h: 0 };
  const xfrm = spPr.child(A_NS, "xfrm");
  if (xfrm !== undefined) {
    const hasExplicitFlipH = xfrm.element.hasAttribute("flipH");
    const hasExplicitFlipV = xfrm.element.hasAttribute("flipV");
    const rawXfrm = xfrm.attrsObject(["rot", "flipH", "flipV"]);
    const parsedXfrm = transformAttributesFragmentSchema.safeParse(rawXfrm);
    if (parsedXfrm.success) {
      if (parsedXfrm.data.rot !== undefined && parsedXfrm.data.rot !== 0) {
        rotate = parsedXfrm.data.rot / 60_000;
      }
      if (hasExplicitFlipH && parsedXfrm.data.flipH !== undefined) {
        flipH = parsedXfrm.data.flipH;
      }
      if (hasExplicitFlipV && parsedXfrm.data.flipV !== undefined) {
        flipV = parsedXfrm.data.flipV;
      }
    }
    const off = xfrm.child(A_NS, "off");
    const ext = xfrm.child(A_NS, "ext");
    if (off !== undefined && ext !== undefined) {
      const parsedOff = offsetEmuFragmentSchema.safeParse(
        off.attrsObject(["x", "y"]),
      );
      const parsedExt = extentEmuFragmentSchema.safeParse(
        ext.attrsObject(["cx", "cy"]),
      );
      if (parsedOff.success && parsedExt.success) {
        fallbackUnits = { w: parsedExt.data.cx, h: parsedExt.data.cy };
        localBox = {
          x: emuToPx(parsedOff.data.x),
          y: emuToPx(parsedOff.data.y),
          width: emuToPx(parsedExt.data.cx),
          height: emuToPx(parsedExt.data.cy),
        };
      }
    }
  }

  const matchedKey = matchLayoutPlaceholderKey(layoutPlaceholders, {
    type: nativeRef.type,
    idx: nativeRef.idx,
  });
  const layoutPlaceholder =
    matchedKey === undefined ? undefined : layoutPlaceholders?.[matchedKey];
  const masterPlaceholder = matchMasterPlaceholderSpec(
    context.masterPlaceholders,
    nativeRef.type,
  );
  const inheritedWidth =
    (typeof layoutPlaceholder?.style?.width === "number"
      ? layoutPlaceholder.style.width
      : undefined) ??
    (typeof masterPlaceholder?.style?.width === "number"
      ? masterPlaceholder.style.width
      : undefined);
  const inheritedHeight =
    (typeof layoutPlaceholder?.style?.height === "number"
      ? layoutPlaceholder.style.height
      : undefined) ??
    (typeof masterPlaceholder?.style?.height === "number"
      ? masterPlaceholder.style.height
      : undefined);
  const inheritedBox =
    inheritedWidth !== undefined && inheritedHeight !== undefined
      ? { width: inheritedWidth, height: inheritedHeight }
      : undefined;
  const effectiveFallbackUnits =
    localBox === undefined && inheritedBox !== undefined
      ? {
          w: Math.round(inheritedBox.width * EMU_PER_PX),
          h: Math.round(inheritedBox.height * EMU_PER_PX),
        }
      : fallbackUnits;

  const appearance = readSpPrAppearance(
    spPr,
    context.package,
    effectiveFallbackUnits,
    phLocation,
    diagnostics,
    { preservePlaceholderClears: true },
  );

  const { border: rawBorder, ...restVisualStyle } = appearance.style;
  const border: NonNullable<PlaceholderDeclarationProps["style"]>["border"] =
    rawBorder === "none"
      ? "none"
      : rawBorder === undefined
        ? undefined
        : (() => {
            const {
              startArrow: _startArrow,
              endArrow: _endArrow,
              ...plainBorder
            } = rawBorder;
            return plainBorder;
          })();
  const boxSize = localBox ?? inheritedBox;
  let borderRadius: number | undefined;
  if (appearance.geometry === "roundRect") {
    if (boxSize !== undefined) {
      borderRadius =
        (16667 / 100_000) * Math.min(boxSize.width, boxSize.height);
    }
  } else if (
    typeof appearance.geometry === "object" &&
    appearance.geometry.type === "preset" &&
    appearance.geometry.name === "roundRect" &&
    boxSize !== undefined
  ) {
    const adj = appearance.geometry.adjust?.["adj"] ?? 16667;
    if (adj > 0) {
      borderRadius = (adj / 100_000) * Math.min(boxSize.width, boxSize.height);
    }
  }

  const combinedStyle: Record<string, unknown> = {
    ...(localBox === undefined
      ? {}
      : {
          position: "absolute",
          left: localBox.x,
          top: localBox.y,
          width: localBox.width,
          height: localBox.height,
        }),
    ...restVisualStyle,
    ...(border === undefined ? {} : { border }),
    ...(borderRadius === undefined ? {} : { borderRadius }),
    ...(rotate === undefined ? {} : { rotate }),
    ...(flipH === undefined ? {} : { flipH }),
    ...(flipV === undefined ? {} : { flipV }),
  };

  let validatedStyle: PlaceholderDeclarationProps["style"] | undefined;
  if (Object.keys(combinedStyle).length > 0) {
    const parsedStyle = placeholderObjectStyleSchema.safeParse(combinedStyle);
    if (parsedStyle.success) {
      validatedStyle = parsedStyle.data;
    }
  }

  const node = createPlaceholderNode({
    ...(objectName === undefined ? {} : { name: objectName }),
    ...(matchedKey === undefined ? {} : { placeholder: matchedKey }),
    nativePlaceholder: nativeRef,
    ...(clickResult?.link === undefined ? {} : { link: clickResult.link }),
    ...(validatedStyle === undefined ? {} : { style: validatedStyle }),
  });

  if (
    clickResult?.deferredSlideLink !== undefined &&
    context.registerDeferredSlideLink !== undefined
  ) {
    const deferred = clickResult.deferredSlideLink;
    context.registerDeferredSlideLink({
      ...deferred,
      apply(targetSlideId: string) {
        node.set({ link: { id: targetSlideId } });
      },
    });
  }

  return { node, diagnostics };
}

/** Shape-tree import handler wired into `contentPlaceholderMechanism` for unpopulated `<p:sp><p:ph .../></p:sp>` placeholders without `<p:txBody>`. */
export const contentPlaceholderImportHandler: ShapeTreeImportHandler = {
  target: "shape-tree",
  localName: "sp",
  matches: (cursor) =>
    isPlaceholderSpElement(cursor) &&
    directElementChild(cursor.element, P_NS, "txBody") === undefined,
  read: readContentPlaceholderSp,
};

/** Typed native canvas mechanism connecting emit, paint, and import for unpopulated `content-placeholder`. */
export const contentPlaceholderMechanism = defineCanvasMechanism({
  key: "content-placeholder",
  role: "canvas",
  emit: emitPlaceholderModel,
  paint: paintPlaceholderModel,
  import: contentPlaceholderImportHandler,
});

/** Capability preserving unpopulated non-text content placeholder slots across import/emit without synthesizing empty text bodies. */
export const placeholderCapability = {
  declaration: "placeholder",
  mechanisms: [contentPlaceholderMechanism] as const,
  codegen: placeholderCodegen,

  bind(
    input: CapabilityInput,
    context: CapabilityParseContext,
  ): BoundObjectCapability {
    const unknown = Object.keys(input.props).find(
      (key) => !placeholderAllowedPropKeys.has(key),
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-parameter",
        `Placeholder contains unsupported parameter ${unknown}`,
      );
    }
    if (input.children.length > 0) {
      throw new CapabilityError(
        "unsupported-children",
        "Unpopulated content placeholder does not accept children",
      );
    }

    const placeholder = context.placeholder;
    const rawNative = nativePlaceholderRefSchema.safeParse(
      input.props["nativePlaceholder"],
    );
    const nativeRef =
      placeholder?.nativeRef ??
      (rawNative.success ? rawNative.data : undefined);
    if (nativeRef === undefined) {
      throw new CapabilityError(
        "missing-placeholder-binding",
        "Content placeholder requires a bound layout placeholder or nativePlaceholder descriptor",
      );
    }

    const requestedName = parseOptionalName(input.props["name"]);
    const morphKey = parseMorphKey(input.props["morphKey"]);
    const link = parseLinkTarget(input.props["link"]);

    const contract = parseCapabilityLayoutContract({
      nativeKind: "sp",
      props: input.props,
      allowedStyleKeys: placeholderAllowedStyleKeys,
      allowPadding: false,
      path: context.path,
    });

    const rawSlideStyle =
      typeof input.props["style"] === "object" &&
      input.props["style"] !== null &&
      !Array.isArray(input.props["style"])
        ? (input.props["style"] as Readonly<Record<string, unknown>>)
        : undefined;
    const inheritedPlaceholderGeometry =
      placeholder?.layoutPlaceholder?.geometry ??
      placeholder?.masterPlaceholder?.geometry;
    const visualStyle =
      placeholder !== undefined
        ? resolvePlaceholderVisualStyle(placeholder, rawSlideStyle)
        : rawSlideStyle !== undefined
          ? { ...rawSlideStyle }
          : {};
    resolveVisualModel({
      theme: placeholder?.masterTheme ?? context.theme,
      style: visualStyle,
      box: { x: 0, y: 0, width: 100, height: 100 },
      ...(inheritedPlaceholderGeometry === undefined
        ? {}
        : { geometry: inheritedPlaceholderGeometry }),
      layoutOnlyWhenInvisible: true,
    });

    const placeholderGeometry =
      placeholder === undefined
        ? undefined
        : resolvePlaceholderGeometry(placeholder, input.props);
    const rawX = placeholderGeometry?.x;
    const rawY = placeholderGeometry?.y;
    const rawWidth = placeholderGeometry?.width;
    const rawHeight = placeholderGeometry?.height;

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

      resolve(resolveContext: ObjectResolveContext): ObjectResolution {
        const box: Rect = {
          x:
            resolveContext.layoutBox?.x ??
            (typeof rawX === "number" ? rawX : 0),
          y:
            resolveContext.layoutBox?.y ??
            (typeof rawY === "number" ? rawY : 0),
          width:
            resolveContext.layoutBox?.width ??
            (typeof rawWidth === "number" ? rawWidth : 100),
          height:
            resolveContext.layoutBox?.height ??
            (typeof rawHeight === "number" ? rawHeight : 100),
        };
        const name =
          resolveContext.morphName ??
          requestedName ??
          placeholder?.layoutPlaceholder?.name ??
          `Placeholder ${resolveContext.id - 1}`;
        try {
          assertXml10Characters(name);
        } catch {
          throw new CapabilityError(
            "invalid-xml-character",
            "Placeholder.name contains characters not allowed in XML 1.0",
          );
        }
        const hasBox =
          placeholder === undefined ||
          placeholderGeometry?.hasLocalBoxOverride === true;
        if (hasBox) {
          assertVisualBoxExtent(box);
        }
        const visual = resolveVisualModel({
          theme: placeholder?.masterTheme ?? resolveContext.theme,
          style: visualStyle,
          box,
          ...(inheritedPlaceholderGeometry === undefined
            ? {}
            : { geometry: inheritedPlaceholderGeometry }),
          layoutOnlyWhenInvisible: !hasBox,
        });
        const spPrFlags: PlaceholderSpPrFlags = {
          hasBox,
          hasRotate: hasLocalDeclaration(rawSlideStyle, "rotate"),
          hasFlipH: hasLocalDeclaration(rawSlideStyle, "flipH"),
          hasFlipV: hasLocalDeclaration(rawSlideStyle, "flipV"),
          hasGeometryOrRadius: hasLocalDeclaration(
            rawSlideStyle,
            "borderRadius",
          ),
          hasFill: hasLocalDeclaration(rawSlideStyle, "background"),
          hasBorder: hasLocalDeclaration(rawSlideStyle, "border"),
          hasEffects:
            hasLocalDeclaration(rawSlideStyle, "boxShadow") ||
            hasLocalDeclaration(rawSlideStyle, "softEdge"),
          includeDefaultRectWhenHasBox: true,
        };
        const model: ResolvedPlaceholderModel = {
          box,
          name,
          nativeRef,
          visual,
          spPrFlags,
        };
        return {
          observation: {
            id: resolveContext.id,
            kind: "placeholder",
            name,
            box,
          },
          extensions: [],
          emit(emitContext: EmitContext): EmitOutput {
            return contentPlaceholderMechanism.emit(model, emitContext);
          },
          paint(): PaintOutput {
            return contentPlaceholderMechanism.paint(model, undefined as never);
          },
        };
      },
    };
  },
} satisfies ObjectCapabilityModule;
