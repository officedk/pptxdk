import {
  CapabilityError,
  authorIdSchema,
  isMorphKey,
  layoutStyleKeys,
  linkTargetSchema,
  matchLayoutPlaceholderKey,
  nativePlaceholderRefSchema,
  parseCapabilityLayoutContract,
  type BoundObjectCapability,
  type BoundObjectNativeCollaboration,
  type CapabilityInput,
  type CapabilityParseContext,
  type EmitContext,
  type EmitOutput,
  type GeometryBox,
  type MorphKey,
  type NativePlaceholderRef,
  type ObjectCapabilityModule,
  type ObjectResolveContext,
  type PaintOutput,
  type ResolvedObject,
} from "@pptxdk/core/capability";
import { CommonNode } from "@pptxdk/core/document";
import {
  assertXml10Characters,
  placeholderRefXml,
  pxToEmu,
  xmlEscape,
} from "@pptxdk/emit";
import {
  A_NS,
  ExtentEmuFragmentSchema,
  OffsetEmuFragmentSchema,
  OoxmlBoolSchema,
  OoxmlNonNegativeIntSchema,
  P_NS,
  defineCanvasMechanism,
  emuToPx,
  type ImportDiagnostic,
  type ImportLocation,
  type ShapeTreeImportContext,
  type ShapeTreeImportHandler,
  type ShapeTreeImportOutcome,
  type XmlElementCursor,
} from "@pptxdk/import";
import { z } from "zod";

import {
  OptionalNameSchema,
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
import { resolvePlaceholderGeometry } from "../shared/style-inheritance.ts";

const NonVisualSpIdSchema = z.object({
  id: OoxmlNonNegativeIntSchema,
  name: z.string().optional(),
});

const NonVisualSpPrFlagsSchema = z.object({
  txBox: OoxmlBoolSchema.optional(),
});

export const placeholderDeclarationPropsSchema = z.strictObject({
  id: authorIdSchema.optional(),
  name: OptionalNameSchema.optional(),
  morphKey: z.custom<MorphKey>(isMorphKey).optional(),
  placeholder: z.string().min(1).optional(),
  nativePlaceholder: nativePlaceholderRefSchema.optional(),
  link: linkTargetSchema.optional(),
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
  readonly box: GeometryBox;
  readonly name: string;
  readonly nativeRef: NativePlaceholderRef;
  readonly omitSpPrTransform: boolean;
};

function emitPlaceholderModel(
  model: ResolvedPlaceholderModel,
  emitContext: EmitContext,
): EmitOutput {
  const escapedName = xmlEscape(model.name);
  const spPrXml = model.omitSpPrTransform
    ? "<p:spPr/>"
    : `<p:spPr><a:xfrm><a:off x="${pxToEmu(model.box.x)}" y="${pxToEmu(model.box.y)}"/><a:ext cx="${pxToEmu(model.box.width)}" cy="${pxToEmu(model.box.height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`;
  return {
    xml: `<p:sp><p:nvSpPr><p:cNvPr id="${emitContext.id}" name="${escapedName}">${nativeLinkXml(emitContext.nativeLink)}</p:cNvPr><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${placeholderRefXml(model.nativeRef)}</p:nvPr></p:nvSpPr>${spPrXml}</p:sp>`,
  };
}

function paintPlaceholderModel(): PaintOutput {
  // Unpopulated content placeholders show prompt chrome only inside interactive editors and draw nothing on static preview.
  return { svg: "" };
}

function consumeUnknownSubtree(cursor: XmlElementCursor): void {
  const attrs = cursor.element.attributes;
  for (let i = 0; i < attrs.length; i += 1) {
    const attr = attrs.item(i);
    if (attr === null) continue;
    const ns = attr.namespaceURI ?? "";
    const local = attr.localName ?? attr.name;
    if (
      ns === "http://www.w3.org/2000/xmlns/" ||
      attr.name === "xmlns" ||
      attr.prefix === "xmlns"
    ) {
      continue;
    }
    cursor.consumeAttr(local, ns);
  }
  for (const child of cursor.remainingElementChildren()) {
    consumeUnknownSubtree(child);
  }
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
      const parsedOff = OffsetEmuFragmentSchema.safeParse(
        off.attrsObject(["x", "y"]),
      );
      const parsedExt = ExtentEmuFragmentSchema.safeParse(
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
  consumeUnknownSubtree(spPr);

  const matchedKey = matchLayoutPlaceholderKey(layoutPlaceholders, {
    type: nativeRef.type,
    idx: nativeRef.idx,
  });

  const node = createPlaceholderNode({
    ...(objectName === undefined ? {} : { name: objectName }),
    ...(matchedKey === undefined ? {} : { placeholder: matchedKey }),
    nativePlaceholder: nativeRef,
    ...(clickResult?.link === undefined ? {} : { link: clickResult.link }),
    ...localBox,
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
      allowedStyleKeys: new Set<string>(layoutStyleKeys),
      allowPadding: false,
      path: context.path,
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

      resolve(resolveContext: ObjectResolveContext): ResolvedObject {
        const box: GeometryBox = {
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
        const omitSpPrTransform =
          placeholder !== undefined &&
          placeholderGeometry?.hasLocalBoxOverride !== true;
        const model: ResolvedPlaceholderModel = {
          box,
          name,
          nativeRef,
          omitSpPrTransform,
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
