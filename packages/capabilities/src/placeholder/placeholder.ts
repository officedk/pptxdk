import {
  CapabilityError,
  authorIdSchema,
  isMorphKey,
  layoutStyleKeys,
  layoutStyleSchema,
  linkTargetSchema,
  matchLayoutPlaceholderKey,
  nativePlaceholderRefSchema,
  parseCapabilityLayoutContract,
  visualStyleSchema,
  type BoundObjectCapability,
  type BoundObjectNativeCollaboration,
  type CapabilityInput,
  type CapabilityParseContext,
  type EmitContext,
  type EmitOutput,
  type Rect,
  type MorphKey,
  type NativePlaceholderRef,
  type ObjectCapabilityModule,
  type ObjectResolveContext,
  type PaintOutput,
  type ObjectResolution,
} from "@pptxdk/core/capability";
import { CommonNode } from "@pptxdk/core/document";
import {
  assertXml10Characters,
  placeholderRefXml,
  xmlEscape,
} from "@pptxdk/emit";
import {
  A_NS,
  extentEmuFragmentSchema,
  offsetEmuFragmentSchema,
  ooxmlBoolSchema,
  ooxmlNonNegativeIntSchema,
  P_NS,
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
  optionalNameSchema,
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
  emitPlaceholderVisualSpPr,
  resolveVisualModel,
  visualStyleKeys,
  type PlaceholderSpPrFlags,
  type ResolvedVisualModel,
} from "../shared/visual.ts";
import { placeholderCodegen } from "./placeholder-codegen.ts";

const NonVisualSpIdSchema = z.object({
  id: ooxmlNonNegativeIntSchema,
  name: z.string().optional(),
});

const NonVisualSpPrFlagsSchema = z.object({
  txBox: ooxmlBoolSchema.optional(),
});

export const placeholderObjectStyleSchema = z
  .strictObject({
    ...layoutStyleSchema.unwrap().shape,
    ...visualStyleSchema.unwrap().shape,
  })
  .readonly();

export const placeholderDeclarationPropsSchema = z.strictObject({
  id: authorIdSchema.optional(),
  name: optionalNameSchema.optional(),
  morphKey: z.custom<MorphKey>(isMorphKey).optional(),
  placeholder: z.string().min(1).optional(),
  nativePlaceholder: nativePlaceholderRefSchema.optional(),
  link: linkTargetSchema.optional(),
  style: placeholderObjectStyleSchema.optional(),
  x: z.number().finite().optional(),
  y: z.number().finite().optional(),
  width: z.number().finite().positive().optional(),
  height: z.number().finite().positive().optional(),
});

export type PlaceholderDeclarationProps = z.infer<
  typeof placeholderDeclarationPropsSchema
>;

export type PlaceholderNode = CommonNode<
  "placeholder",
  PlaceholderDeclarationProps
>;

const placeholderAllowedParameters = Object.keys(
  placeholderDeclarationPropsSchema.shape,
);

/** Creates an editable AST node representing an unpopulated native content placeholder (`<p:sp><p:ph .../></p:sp>` without `<p:txBody>`). */
export function createPlaceholderNode(
  props: PlaceholderDeclarationProps,
): PlaceholderNode {
  return new CommonNode("placeholder", "placeholder", props);
}

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
  const xfrm = spPr.child(A_NS, "xfrm");
  if (xfrm !== undefined) {
    xfrm.attrsObject(["rot", "flipH", "flipV"]);
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
        localBox = {
          x: emuToPx(parsedOff.data.x),
          y: emuToPx(parsedOff.data.y),
          width: emuToPx(parsedExt.data.cx),
          height: emuToPx(parsedExt.data.cy),
        };
      }
    }
  }
  const prstGeomEl = directElementChild(spPr.element, A_NS, "prstGeom");
  if (
    prstGeomEl !== undefined &&
    prstGeomEl.getAttribute("prst") === "rect" &&
    prstGeomEl.attributes.length === 1
  ) {
    const prstGeom = spPr.child(A_NS, "prstGeom")!;
    prstGeom.consumeAttr("prst");
    const avLstEl = directElementChild(prstGeom.element, A_NS, "avLst");
    if (
      avLstEl !== undefined &&
      avLstEl.children.length === 0 &&
      avLstEl.attributes.length === 0
    ) {
      prstGeom.child(A_NS, "avLst");
    }
  }
  const noFillEl = directElementChild(spPr.element, A_NS, "noFill");
  if (
    noFillEl !== undefined &&
    noFillEl.children.length === 0 &&
    noFillEl.attributes.length === 0
  ) {
    spPr.child(A_NS, "noFill");
  }

  const matchedKey = matchLayoutPlaceholderKey(layoutPlaceholders, {
    type: nativeRef.type,
    idx: nativeRef.idx,
  });

  const node = createPlaceholderNode({
    ...(objectName === undefined ? {} : { name: objectName }),
    ...(matchedKey === undefined ? {} : { placeholder: matchedKey }),
    nativePlaceholder: nativeRef,
    ...(clickResult?.link === undefined ? {} : { link: clickResult.link }),
    ...(localBox === undefined
      ? {}
      : {
          style: {
            position: "absolute",
            left: localBox.x,
            top: localBox.y,
            width: localBox.width,
            height: localBox.height,
          },
        }),
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
      (key) => !placeholderAllowedParameters.includes(key),
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
      allowedStyleKeys: new Set<string>([
        ...layoutStyleKeys,
        ...visualStyleKeys,
      ]),
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
    const inheritedPlaceholderFlipH =
      placeholder?.layoutPlaceholder?.flipH ??
      placeholder?.masterPlaceholder?.flipH;
    const inheritedPlaceholderFlipV =
      placeholder?.layoutPlaceholder?.flipV ??
      placeholder?.masterPlaceholder?.flipV;
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
      ...(inheritedPlaceholderFlipH === undefined
        ? {}
        : { flipH: inheritedPlaceholderFlipH }),
      ...(inheritedPlaceholderFlipV === undefined
        ? {}
        : { flipV: inheritedPlaceholderFlipV }),
      layoutOnlyWhenInvisible: true,
    });

    const placeholderGeometry =
      placeholder === undefined
        ? undefined
        : resolvePlaceholderGeometry(placeholder, input.props);
    const rawX = placeholderGeometry?.x ?? input.props["x"];
    const rawY = placeholderGeometry?.y ?? input.props["y"];
    const rawWidth = placeholderGeometry?.width ?? input.props["width"];
    const rawHeight = placeholderGeometry?.height ?? input.props["height"];

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
        const visual = resolveVisualModel({
          theme: placeholder?.masterTheme ?? resolveContext.theme,
          style: visualStyle,
          box,
          ...(inheritedPlaceholderGeometry === undefined
            ? {}
            : { geometry: inheritedPlaceholderGeometry }),
          ...(inheritedPlaceholderFlipH === undefined
            ? {}
            : { flipH: inheritedPlaceholderFlipH }),
          ...(inheritedPlaceholderFlipV === undefined
            ? {}
            : { flipV: inheritedPlaceholderFlipV }),
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
          hasFill:
            hasLocalDeclaration(rawSlideStyle, "background") ||
            hasLocalDeclaration(rawSlideStyle, "backgroundColor") ||
            hasLocalDeclaration(rawSlideStyle, "backgroundImage"),
          hasBorder:
            hasLocalDeclaration(rawSlideStyle, "border") ||
            hasLocalDeclaration(rawSlideStyle, "borderColor") ||
            hasLocalDeclaration(rawSlideStyle, "borderWidth") ||
            hasLocalDeclaration(rawSlideStyle, "borderStyle"),
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
