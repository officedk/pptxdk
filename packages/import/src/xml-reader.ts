import {
  DOMParser,
  type Document as XmlDocument,
  type Element as XmlElement,
  type Node as XmlNode,
} from "@xmldom/xmldom";

import type { UnconsumedXmlItem } from "#/diagnostics.ts";
import { MC_NS, XMLNS_NS, canonicalQName } from "#/namespaces.ts";

/** Error raised when an XML part fails well-formedness or lacks a root element. */
export class MalformedXmlPartError extends Error {
  readonly partPath: string;
  readonly detail: string;

  constructor(partPath: string, detail: string) {
    super(`Malformed XML in ${partPath}: ${detail}`);
    this.name = "MalformedXmlPartError";
    this.partPath = partPath;
    this.detail = detail;
  }
}

/** Parses an XML part with `@xmldom/xmldom` without emitting console noise. */
export function parseXmlPart(partPath: string, xmlText: string): XmlElement {
  let parseFailure: string | undefined;
  const parser = new DOMParser({
    onError(level, message) {
      if (level === "error" || level === "fatalError") {
        parseFailure ??= message;
      }
    },
  });
  let doc: XmlDocument;
  try {
    doc = parser.parseFromString(xmlText, "application/xml");
  } catch (error) {
    throw new MalformedXmlPartError(
      partPath,
      error instanceof Error ? error.message : String(error),
    );
  }
  if (parseFailure !== undefined || doc.documentElement === null) {
    throw new MalformedXmlPartError(
      partPath,
      parseFailure ?? "Missing root element",
    );
  }
  return doc.documentElement;
}

function attrKey(namespaceUri: string, localName: string): string {
  return `${namespaceUri}\u0000${localName}`;
}

function isIgnorableDeclarationAttr(
  namespaceUri: string,
  name: string,
  prefix: string | null,
  localName: string,
): boolean {
  if (namespaceUri === XMLNS_NS || name === "xmlns" || prefix === "xmlns") {
    return true;
  }
  return namespaceUri === MC_NS && localName === "Ignorable";
}

function isXmlElement(node: XmlNode): node is XmlElement {
  return node.nodeType === 1;
}

function elementChildrenOf(element: XmlElement): XmlElement[] {
  const result: XmlElement[] = [];
  for (
    let child = element.firstChild;
    child !== null;
    child = child.nextSibling
  ) {
    if (isXmlElement(child)) {
      result.push(child);
    }
  }
  return result;
}

function childXpathFor(
  parentXpath: string,
  siblings: readonly XmlElement[],
  target: XmlElement,
): string {
  const ns = target.namespaceURI ?? "";
  const local = target.localName ?? target.tagName;
  const qname = canonicalQName(ns, local);
  let total = 0;
  let position = 0;
  for (const sibling of siblings) {
    const siblingNs = sibling.namespaceURI ?? "";
    const siblingLocal = sibling.localName ?? sibling.tagName;
    if (siblingNs === ns && siblingLocal === local) {
      total += 1;
      if (sibling === target) {
        position = total;
      }
    }
  }
  return total > 1
    ? `${parentXpath}/${qname}[${position}]`
    : `${parentXpath}/${qname}`;
}

type CursorRegistry = {
  readonly byElement: Map<XmlElement, XmlElementCursor>;
  readonly ordered: XmlElementCursor[];
};

/**
 * Namespace-aware XML cursor that tracks consumed attributes and child elements so
 * unconsumed OOXML constructs are surfaced with exact package part and XPath locations.
 */
export class XmlElementCursor {
  readonly element: XmlElement;
  readonly part: string;
  readonly xpath: string;
  readonly namespaceUri: string;
  readonly localName: string;
  readonly #registry: CursorRegistry;
  readonly #consumedAttrs = new Set<string>();
  readonly #consumedChildren = new Set<XmlElement>();
  #suppressUnconsumed = false;
  #cachedChildren: readonly XmlElement[] | undefined;

  private constructor(
    element: XmlElement,
    part: string,
    xpath: string,
    registry: CursorRegistry,
  ) {
    this.element = element;
    this.part = part;
    this.xpath = xpath;
    this.namespaceUri = element.namespaceURI ?? "";
    this.localName = element.localName ?? element.tagName;
    this.#registry = registry;
    this.#registry.byElement.set(element, this);
    this.#registry.ordered.push(this);
  }

  static createRoot(
    element: XmlElement,
    part: string,
    rootXpath?: string,
  ): XmlElementCursor {
    const registry: CursorRegistry = {
      byElement: new Map(),
      ordered: [],
    };
    const ns = element.namespaceURI ?? "";
    const local = element.localName ?? element.tagName;
    return new XmlElementCursor(
      element,
      part,
      rootXpath ?? `/${canonicalQName(ns, local)}`,
      registry,
    );
  }

  get textContent(): string {
    return this.element.textContent ?? "";
  }

  #elementChildren(): readonly XmlElement[] {
    this.#cachedChildren ??= elementChildrenOf(this.element);
    return this.#cachedChildren;
  }

  #cursorFor(child: XmlElement): XmlElementCursor {
    const existing = this.#registry.byElement.get(child);
    if (existing !== undefined) return existing;
    const xpath = childXpathFor(this.xpath, this.#elementChildren(), child);
    return new XmlElementCursor(child, this.part, xpath, this.#registry);
  }

  /** Reads an attribute by `(namespaceUri, localName)` and marks it consumed if present. */
  attr(localName: string, namespaceUri = ""): string | undefined {
    const attributes = this.element.attributes;
    for (let i = 0; i < attributes.length; i += 1) {
      const attribute = attributes.item(i);
      if (attribute === null) continue;
      const attrNs = attribute.namespaceURI ?? "";
      const attrLocal = attribute.localName ?? attribute.name;
      if (attrNs === namespaceUri && attrLocal === localName) {
        this.#consumedAttrs.add(attrKey(attrNs, attrLocal));
        return attribute.value;
      }
    }
    return undefined;
  }

  /** Reads a set of allowed attributes into a plain record for Zod fragment parsing. */
  attrsObject(
    allowed: readonly string[],
    namespaceUri = "",
  ): Record<string, string> {
    const result: Record<string, string> = {};
    for (const localName of allowed) {
      const value = this.attr(localName, namespaceUri);
      if (value !== undefined) {
        result[localName] = value;
      }
    }
    return result;
  }

  /** Marks an attribute consumed even if its value is not retained. */
  consumeAttr(localName: string, namespaceUri = ""): boolean {
    return this.attr(localName, namespaceUri) !== undefined;
  }

  /** Returns the first matching child cursor and marks that child element consumed. */
  child(namespaceUri: string, localName: string): XmlElementCursor | undefined {
    for (const child of this.#elementChildren()) {
      const childNs = child.namespaceURI ?? "";
      const childLocal = child.localName ?? child.tagName;
      if (childNs === namespaceUri && childLocal === localName) {
        this.#consumedChildren.add(child);
        return this.#cursorFor(child);
      }
    }
    return undefined;
  }

  /** Returns all matching child cursors in document order and marks them consumed. */
  children(
    namespaceUri: string,
    localName: string,
  ): readonly XmlElementCursor[] {
    const matches: XmlElementCursor[] = [];
    for (const child of this.#elementChildren()) {
      const childNs = child.namespaceURI ?? "";
      const childLocal = child.localName ?? child.tagName;
      if (childNs === namespaceUri && childLocal === localName) {
        this.#consumedChildren.add(child);
        matches.push(this.#cursorFor(child));
      }
    }
    return matches;
  }

  /** Returns matching direct element children without marking non-matching siblings consumed. */
  matchingChildren(
    predicate: (element: XmlElement) => boolean,
  ): readonly XmlElementCursor[] {
    const matches: XmlElementCursor[] = [];
    for (const child of this.#elementChildren()) {
      if (predicate(child)) {
        this.#consumedChildren.add(child);
        matches.push(this.#cursorFor(child));
      }
    }
    return matches;
  }

  /** Returns all direct element children as cursors and marks them consumed by the caller. */
  allElementChildren(): readonly XmlElementCursor[] {
    return this.#elementChildren().map((child) => {
      this.#consumedChildren.add(child);
      return this.#cursorFor(child);
    });
  }

  /** Returns direct element children not yet consumed in document order and marks them consumed. */
  remainingElementChildren(): readonly XmlElementCursor[] {
    const remaining: XmlElementCursor[] = [];
    for (const child of this.#elementChildren()) {
      if (this.#consumedChildren.has(child)) continue;
      this.#consumedChildren.add(child);
      remaining.push(this.#cursorFor(child));
    }
    return remaining;
  }

  /** Marks a specific child cursor consumed (useful after inspecting an element child). */
  markChildConsumed(child: XmlElementCursor): void {
    this.#consumedChildren.add(child.element);
  }

  /** Marks all direct element children consumed without descending into them. */
  consumeChildrenWithoutVisiting(): void {
    for (const child of this.#elementChildren()) {
      this.#consumedChildren.add(child);
    }
  }

  /** Suppresses unconsumed-xml diagnostics for this cursor and its remaining children (used when the subtree is already reported as omitted). */
  suppressUnconsumed(): void {
    this.#suppressUnconsumed = true;
    this.consumeChildrenWithoutVisiting();
  }

  /** Marks an empty schema-boilerplate child element consumed only when it has no attributes or element children. */
  ignoreEmptyChild(namespaceUri: string, localName: string): boolean {
    for (const child of this.#elementChildren()) {
      const childNs = child.namespaceURI ?? "";
      const childLocal = child.localName ?? child.tagName;
      if (childNs !== namespaceUri || childLocal !== localName) continue;
      if (
        hasNonIgnorableAttributes(child) ||
        elementChildrenOf(child).length > 0
      ) {
        // Visit the child so its own non-empty contents are reported with exact XPath.
        this.#consumedChildren.add(child);
        this.#cursorFor(child);
        return false;
      }
      this.#consumedChildren.add(child);
      return true;
    }
    return false;
  }

  /** Collects unconsumed attributes and direct unconsumed child elements across all visited cursors. */
  collectUnconsumed(): readonly UnconsumedXmlItem[] {
    const items: UnconsumedXmlItem[] = [];
    for (const cursor of this.#registry.ordered) {
      if (cursor.#suppressUnconsumed) continue;
      const attributes = cursor.element.attributes;
      for (let i = 0; i < attributes.length; i += 1) {
        const attribute = attributes.item(i);
        if (attribute === null) continue;
        const attrNs = attribute.namespaceURI ?? "";
        const attrLocal = attribute.localName ?? attribute.name;
        if (
          isIgnorableDeclarationAttr(
            attrNs,
            attribute.name,
            attribute.prefix,
            attrLocal,
          )
        ) {
          continue;
        }
        if (!cursor.#consumedAttrs.has(attrKey(attrNs, attrLocal))) {
          items.push({
            kind: "attribute",
            location: {
              part: cursor.part,
              xpath: `${cursor.xpath}/@${canonicalQName(attrNs, attrLocal)}`,
            },
            namespaceUri: attrNs,
            localName: attrLocal,
          });
        }
      }
      const siblings = cursor.#elementChildren();
      for (const child of siblings) {
        if (!cursor.#consumedChildren.has(child)) {
          const childNs = child.namespaceURI ?? "";
          const childLocal = child.localName ?? child.tagName;
          items.push({
            kind: "element",
            location: {
              part: cursor.part,
              xpath: childXpathFor(cursor.xpath, siblings, child),
            },
            namespaceUri: childNs,
            localName: childLocal,
          });
        }
      }
    }
    return items;
  }
}

function hasNonIgnorableAttributes(element: XmlElement): boolean {
  const attributes = element.attributes;
  for (let i = 0; i < attributes.length; i += 1) {
    const attribute = attributes.item(i);
    if (attribute === null) continue;
    const attrNs = attribute.namespaceURI ?? "";
    const attrLocal = attribute.localName ?? attribute.name;
    if (
      !isIgnorableDeclarationAttr(
        attrNs,
        attribute.name,
        attribute.prefix,
        attrLocal,
      )
    ) {
      return true;
    }
  }
  return false;
}
