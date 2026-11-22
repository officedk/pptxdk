import {
  CapabilityError,
  isMorphKey,
  linkTargetSchema,
  parseLinkTarget,
  type LinkTarget,
  type MorphKey,
  type NativeLink,
} from "@pptxdk/core/capability";
import type { CommonNode } from "@pptxdk/core/document";
import { xmlEscape } from "@pptxdk/emit";
import {
  ooxmlBoolSchema,
  P_NS,
  REL_HYPERLINK,
  REL_SLIDE,
  R_NS,
  resolveExternalRelationshipTarget,
  resolvePartPath,
  type ImportDiagnostic,
  type ImportLocation,
  type ShapeTreeImportContext,
  type XmlElementCursor,
} from "@pptxdk/import";
import { z } from "zod";

const NvPrUserDrawnSchema = z.object({
  userDrawn: ooxmlBoolSchema.optional(),
});

function directChildDom(
  parent: XmlElementCursor["element"],
  namespaceUri: string,
  localName: string,
): XmlElementCursor["element"] | undefined {
  for (
    let child = parent.firstChild;
    child !== null;
    child = child.nextSibling
  ) {
    if (child.nodeType !== 1) continue;
    const el = child as XmlElementCursor["element"];
    if (
      (el.namespaceURI ?? "") === namespaceUri &&
      (el.localName ?? el.tagName) === localName
    ) {
      return el;
    }
  }
  return undefined;
}

/** Side-effect-free check for `<p:ph>` inside an object's `<p:nvPr>`. */
export function hasPlaceholderElement(cursor: XmlElementCursor): boolean {
  for (const wrapperName of [
    "nvSpPr",
    "nvPicPr",
    "nvGrpSpPr",
    "nvGraphicFramePr",
    "nvCxnSpPr",
  ] as const) {
    const wrapper = directChildDom(cursor.element, P_NS, wrapperName);
    if (wrapper === undefined) continue;
    const nvPr = directChildDom(wrapper, P_NS, "nvPr");
    if (nvPr === undefined) continue;
    if (directChildDom(nvPr, P_NS, "ph") !== undefined) return true;
  }
  return false;
}

/** Consumes `<p:nvPr>` (and optional `userDrawn` flag) for non-placeholder objects. */
export function consumeNonPlaceholderNvPr(
  nvParentCursor: XmlElementCursor,
): void {
  const nvPr = nvParentCursor.child(P_NS, "nvPr");
  if (nvPr === undefined) return;
  const raw = nvPr.attrsObject(["userDrawn"]);
  NvPrUserDrawnSchema.safeParse(raw);
}

export const optionalNameSchema = z
  .string()
  .refine((value) => value.trim() !== "", {
    message: "name must be a non-empty string",
  });

const HlinkClickFragmentSchema = z.object({
  id: z.string().optional(),
  action: z.string().optional(),
});

export { linkTargetSchema, parseLinkTarget };
export type { LinkTarget };

/** Validates an optional non-empty object name parameter. */
export function parseOptionalName(
  value: unknown,
  path?: string,
): string | undefined {
  if (value === undefined) return undefined;
  const parsed = optionalNameSchema.safeParse(value);
  if (!parsed.success) {
    throw new CapabilityError(
      "invalid-parameter",
      "name must be a non-empty string",
      path,
    );
  }
  return parsed.data;
}

/** Validates an optional Morph handle created by createMorphScope().handle(name). */
export function parseMorphKey(
  value: unknown,
  path?: string,
): MorphKey | undefined {
  if (value === undefined) return undefined;
  if (!isMorphKey(value)) {
    throw new CapabilityError(
      "invalid-morph-key",
      "morphKey must be created by createMorphScope().handle(name)",
      path,
    );
  }
  return value;
}

/** Native object capabilities share the same click link representation, with relationships and actions pre-resolved by core. */
export function nativeLinkXml(link: NativeLink | undefined): string {
  if (link === undefined) return "";
  const attributes = [
    ...(link.relationshipId === undefined
      ? []
      : [`r:id="${xmlEscape(link.relationshipId)}"`]),
    ...(link.action === undefined
      ? []
      : [`action="${xmlEscape(link.action)}"`]),
  ];
  return `<a:hlinkClick ${attributes.join(" ")}/>`;
}

export type ClickLinkReadResult = {
  readonly link?: LinkTarget;
  readonly deferredSlideLink?: {
    readonly targetPart: string;
    readonly relationshipId: string;
    readonly location: ImportLocation;
  };
  readonly relationshipId?: string;
  readonly clickLocation?: ImportLocation;
  readonly diagnostics: readonly ImportDiagnostic[];
};

const homeAction = "ppaction://hlinkshowjump?jump=firstslide";
const slideJumpAction = "ppaction://hlinksldjump";

/**
 * Reads an `<a:hlinkClick>` child under `<p:cNvPr>` into a `LinkTarget` or deferred internal slide jump,
 * keeping external URLs offline and diagnosing broken relationships without dropping the host object.
 */
export function readClickLink(
  hlinkCursor: XmlElementCursor,
  context: ShapeTreeImportContext,
  objectIdentity?: {
    readonly objectId?: number;
    readonly objectName?: string;
  },
): ClickLinkReadResult {
  const clickLocation: ImportLocation = {
    part: context.part,
    xpath: hlinkCursor.xpath,
    slideIndex: context.slideIndex,
    ...(objectIdentity?.objectId === undefined
      ? {}
      : { objectId: objectIdentity.objectId }),
    ...(objectIdentity?.objectName === undefined
      ? {}
      : { objectName: objectIdentity.objectName }),
  };

  const rawId = hlinkCursor.attr("id", R_NS);
  const rawAction = hlinkCursor.attr("action");
  const parsed = HlinkClickFragmentSchema.safeParse({
    ...(rawId === undefined ? {} : { id: rawId }),
    ...(rawAction === undefined ? {} : { action: rawAction }),
  });
  if (!parsed.success) {
    return {
      clickLocation,
      diagnostics: [
        {
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "omitted",
          location: clickLocation,
          attributeOrElement: "a:hlinkClick",
        },
      ],
    };
  }

  const rId = parsed.data.id ?? "";
  const action = parsed.data.action ?? "";

  if (rId === "" && action === homeAction) {
    return {
      link: { action: "home" },
      clickLocation,
      diagnostics: [],
    };
  }

  if (rId === "") {
    return {
      clickLocation,
      diagnostics: [
        {
          code: "broken-relationship",
          category: "corrupt-input",
          disposition: "omitted",
          location: clickLocation,
          relationshipId: "",
        },
      ],
    };
  }

  const relationships = context.relationships ?? new Map();
  if (action === slideJumpAction) {
    const rel = relationships.get(rId);
    if (
      rel === undefined ||
      rel.Type !== REL_SLIDE ||
      rel.TargetMode === "External"
    ) {
      return {
        relationshipId: rId,
        clickLocation,
        diagnostics: [
          {
            code: "broken-relationship",
            category: "corrupt-input",
            disposition: "omitted",
            location: clickLocation,
            relationshipId: rId,
          },
        ],
      };
    }
    const targetPart = resolvePartPath(context.part, rel.Target);
    if (targetPart === undefined || !context.package.hasPart(targetPart)) {
      return {
        relationshipId: rId,
        clickLocation,
        diagnostics: [
          {
            code: "missing-target-part",
            category: "corrupt-input",
            disposition: "omitted",
            location: clickLocation,
            relationshipId: rId,
            targetPart: targetPart ?? rel.Target,
          },
        ],
      };
    }
    return {
      deferredSlideLink: {
        targetPart,
        relationshipId: rId,
        location: clickLocation,
      },
      relationshipId: rId,
      clickLocation,
      diagnostics: [],
    };
  }

  if (action !== "") {
    return {
      relationshipId: rId,
      clickLocation,
      diagnostics: [
        {
          code: "unsupported-mechanism",
          category: "unsupported",
          disposition: "omitted",
          location: clickLocation,
          namespaceUri: hlinkCursor.namespaceUri,
          localName: hlinkCursor.localName,
          mechanismHint: action,
        },
      ],
    };
  }

  const resolvedExternal = resolveExternalRelationshipTarget(
    relationships,
    rId,
    clickLocation,
    REL_HYPERLINK,
  );
  if (!resolvedExternal.ok) {
    return {
      relationshipId: rId,
      clickLocation,
      diagnostics: [resolvedExternal.diagnostic],
    };
  }

  try {
    new URL(resolvedExternal.target);
  } catch {
    return {
      relationshipId: rId,
      clickLocation,
      diagnostics: [
        {
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "omitted",
          location: clickLocation,
          attributeOrElement: "Relationship/@Target",
          rawValue: resolvedExternal.target,
        },
      ],
    };
  }

  return {
    link: { url: resolvedExternal.target },
    relationshipId: rId,
    clickLocation,
    diagnostics: [],
  };
}

/** Registers a deferred internal slide-jump callback so links resolve via slidePartPath -> allocated Slide.id. */
export function attachDeferredSlideLink(
  node: CommonNode,
  deferredSlideLink: ClickLinkReadResult["deferredSlideLink"],
  context: ShapeTreeImportContext,
): void {
  if (
    deferredSlideLink !== undefined &&
    context.registerDeferredSlideLink !== undefined
  ) {
    context.registerDeferredSlideLink({
      ...deferredSlideLink,
      apply(targetSlideId: string) {
        node.set({ link: { id: targetSlideId } });
      },
    });
  }
}
