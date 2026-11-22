import {
  nativePlaceholderRefSchema,
  placeholderRoleSchema,
  type NativePlaceholderRef,
} from "@pptxdk/core/capability";
import {
  OoxmlBoolSchema,
  OoxmlNonNegativeIntSchema,
  P_NS,
  type ImportDiagnostic,
  type ImportLocation,
  type XmlElementCursor,
} from "@pptxdk/import";

const P14_MOD_ID_EXT_URI = "{D42A27DB-BD31-4B8C-83A1-F6EECF244321}";

export type XmlElement = XmlElementCursor["element"];

/** Finds a direct child XML element matching namespace URI and local name without mutating cursor tracking. */
export function directElementChild(
  parent: XmlElement,
  namespaceUri: string,
  localName: string,
): XmlElement | undefined {
  for (
    let child = parent.firstChild;
    child !== null;
    child = child.nextSibling
  ) {
    if (child.nodeType !== 1) continue;
    const el = child as XmlElement;
    if (
      (el.namespaceURI ?? "") === namespaceUri &&
      (el.localName ?? el.tagName) === localName
    ) {
      return el;
    }
  }
  return undefined;
}

/** Side-effect-free lookup of `<p:ph>` inside `<p:{nvContainer}>/<p:nvPr>`. */
export function directPlaceholderElement(
  cursor: XmlElementCursor,
  nvContainerLocalName: "nvSpPr" | "nvPicPr" | "nvGraphicFramePr",
): XmlElement | undefined {
  const nvContainer = directElementChild(
    cursor.element,
    P_NS,
    nvContainerLocalName,
  );
  if (nvContainer === undefined) return undefined;
  const nvPr = directElementChild(nvContainer, P_NS, "nvPr");
  if (nvPr === undefined) return undefined;
  return directElementChild(nvPr, P_NS, "ph");
}

/** Side-effect-free predicate identifying `<p:sp>` elements bound to a placeholder (`<p:ph>`). */
export function isPlaceholderSpElement(cursor: XmlElementCursor): boolean {
  return directPlaceholderElement(cursor, "nvSpPr") !== undefined;
}

/** Side-effect-free predicate identifying `<p:pic>` elements bound to a placeholder (`<p:ph>`). */
export function isPlaceholderPicElement(cursor: XmlElementCursor): boolean {
  return directPlaceholderElement(cursor, "nvPicPr") !== undefined;
}

/** Side-effect-free predicate identifying `<p:graphicFrame>` elements bound to a placeholder (`<p:ph>`). */
export function isPlaceholderGraphicFrameElement(
  cursor: XmlElementCursor,
): boolean {
  return directPlaceholderElement(cursor, "nvGraphicFramePr") !== undefined;
}

function consumeSubtree(cursor: XmlElementCursor): void {
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
    consumeSubtree(child);
  }
}

export type NvPrPlaceholderReadResult =
  | { readonly kind: "none" }
  | { readonly kind: "ok"; readonly nativeRef: NativePlaceholderRef }
  | { readonly kind: "error" };

/**
 * Reads optional `<p:ph>` and benign `p14:modId` extensions inside `<p:nvPr>`.
 * Shared across text, image, table, chart, and unpopulated placeholder importers.
 */
export function readNvPrPlaceholder(
  nvPr: XmlElementCursor | undefined,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): NvPrPlaceholderReadResult {
  if (nvPr === undefined) return { kind: "none" };

  const extLst = nvPr.child(P_NS, "extLst");
  if (extLst !== undefined) {
    for (const ext of extLst.children(P_NS, "ext")) {
      if (ext.attr("uri") === P14_MOD_ID_EXT_URI) {
        consumeSubtree(ext);
      }
    }
  }

  const ph = nvPr.child(P_NS, "ph");
  if (ph === undefined) {
    return { kind: "none" };
  }

  const rawPh = ph.attrsObject([
    "type",
    "idx",
    "sz",
    "orient",
    "hasCustomPrompt",
  ]);
  const explicitType = rawPh["type"] !== undefined;
  const parsedRole = placeholderRoleSchema.safeParse(rawPh["type"] ?? "obj");
  if (!parsedRole.success) {
    diagnostics.push({
      code: "unsupported-mechanism",
      category: "unsupported",
      disposition: "omitted",
      location: { ...location, xpath: ph.xpath },
      namespaceUri: P_NS,
      localName: "ph",
      mechanismHint: `ph:${rawPh["type"] ?? ""}`,
    });
    return { kind: "error" };
  }

  const idxParsed = OoxmlNonNegativeIntSchema.safeParse(rawPh["idx"]);
  const hasPromptParsed = OoxmlBoolSchema.safeParse(rawPh["hasCustomPrompt"]);
  const parsedNativeRef = nativePlaceholderRefSchema.safeParse({
    type: parsedRole.data,
    explicitType,
    ...(idxParsed.success ? { idx: idxParsed.data } : {}),
    ...(rawPh["sz"] !== undefined ? { sz: rawPh["sz"] } : {}),
    ...(rawPh["orient"] !== undefined ? { orient: rawPh["orient"] } : {}),
    ...(hasPromptParsed.success && hasPromptParsed.data
      ? { hasCustomPrompt: true }
      : {}),
  });
  if (!parsedNativeRef.success) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: ph.xpath },
      attributeOrElement: "p:ph",
      rawValue: JSON.stringify(rawPh),
    });
    return { kind: "error" };
  }

  return { kind: "ok", nativeRef: parsedNativeRef.data };
}
