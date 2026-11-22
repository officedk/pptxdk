import {
  CommonNode,
  createSectionNode,
  type DeckContainer,
} from "@pptxdk/core/document";
import { z } from "zod";

import type { ImportDiagnostic } from "#/diagnostics.ts";
import { allocateScopedAuthorId } from "#/masters.ts";
import { P14_NS, P14_SECTION_LIST_EXT_URI, P_NS } from "#/namespaces.ts";
import { OoxmlNonNegativeIntSchema } from "#/xml-fragments.ts";
import type { XmlElementCursor } from "#/xml-reader.ts";

const SectionAttrsXmlSchema = z.object({
  name: z.string().refine((value) => value.trim() !== ""),
  id: z.string().optional(),
});

const SectionSlideIdXmlSchema = z.object({
  id: OoxmlNonNegativeIntSchema,
});

function hasMatchingSectionExtension(extLst: XmlElementCursor): boolean {
  for (
    let child = extLst.element.firstChild;
    child !== null;
    child = child.nextSibling
  ) {
    if (child.nodeType !== 1) continue;
    const el = child as XmlElementCursor["element"];
    if (
      (el.namespaceURI ?? "") === P_NS &&
      (el.localName ?? el.tagName) === "ext" &&
      el.getAttribute("uri") === P14_SECTION_LIST_EXT_URI
    ) {
      return true;
    }
  }
  return false;
}

/**
 * Reads PowerPoint `<p14:sectionLst>` from `ppt/presentation.xml` and groups imported
 * `SlideNode` children into editable `SectionNode` containers while preserving slide order.
 */
export function importPresentationSections(
  presCursor: XmlElementCursor,
  deckNode: DeckContainer,
  slideNodeBySldId: ReadonlyMap<number, CommonNode>,
  presentationPath: string,
  diagnostics: ImportDiagnostic[],
): void {
  const existingSections = deckNode.children.some(
    (child) => child instanceof CommonNode && child.type === "section",
  );
  if (existingSections) return;

  const rawExtLst = (() => {
    for (
      let child = presCursor.element.firstChild;
      child !== null;
      child = child.nextSibling
    ) {
      if (child.nodeType !== 1) continue;
      const el = child as XmlElementCursor["element"];
      if (
        (el.namespaceURI ?? "") === P_NS &&
        (el.localName ?? el.tagName) === "extLst"
      ) {
        return el;
      }
    }
    return undefined;
  })();
  if (rawExtLst === undefined) return;

  const extLst = presCursor.child(P_NS, "extLst");
  if (extLst === undefined) return;
  if (!hasMatchingSectionExtension(extLst)) {
    return;
  }

  const parsedSections: {
    readonly name: string;
    readonly slides: CommonNode[];
  }[] = [];
  const assignedSlides = new Set<CommonNode>();
  const usedSectionNames = new Set<string>();

  for (const ext of extLst.children(P_NS, "ext")) {
    if (ext.element.getAttribute("uri") !== P14_SECTION_LIST_EXT_URI) {
      continue;
    }
    ext.consumeAttr("uri");
    const sectionLst = ext.child(P14_NS, "sectionLst");
    if (sectionLst === undefined) continue;

    const sectionCursors = sectionLst.children(P14_NS, "section");
    for (let index = 0; index < sectionCursors.length; index += 1) {
      const sectionCursor = sectionCursors[index]!;
      const rawAttrs = sectionCursor.attrsObject(["name", "id"]);
      const parsedAttrs = SectionAttrsXmlSchema.safeParse(rawAttrs);
      let sectionName: string;
      if (!parsedAttrs.success) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "recovered",
          location: {
            part: presentationPath,
            xpath: `${sectionCursor.xpath}/@name`,
          },
          attributeOrElement: "p14:section/@name",
          rawValue: rawAttrs["name"] ?? "",
        });
        sectionName = allocateScopedAuthorId(
          undefined,
          usedSectionNames,
          `Section ${index + 1}`,
        );
      } else {
        sectionName = parsedAttrs.data.name;
        usedSectionNames.add(sectionName);
      }

      const sectionSlides: CommonNode[] = [];
      const sldIdLst = sectionCursor.child(P14_NS, "sldIdLst");
      if (sldIdLst !== undefined) {
        for (const sldIdCursor of sldIdLst.children(P14_NS, "sldId")) {
          const rawSlideId = sldIdCursor.attrsObject(["id"]);
          const parsedSlideId = SectionSlideIdXmlSchema.safeParse(rawSlideId);
          if (!parsedSlideId.success) {
            diagnostics.push({
              code: "invalid-xml-value",
              category: "corrupt-input",
              disposition: "omitted",
              location: {
                part: presentationPath,
                xpath: `${sldIdCursor.xpath}/@id`,
              },
              attributeOrElement: "p14:sldId/@id",
              rawValue: rawSlideId["id"] ?? "",
            });
            continue;
          }
          const targetSlide = slideNodeBySldId.get(parsedSlideId.data.id);
          if (targetSlide === undefined) {
            diagnostics.push({
              code: "broken-relationship",
              category: "corrupt-input",
              disposition: "omitted",
              location: {
                part: presentationPath,
                xpath: sldIdCursor.xpath,
              },
              relationshipId: String(parsedSlideId.data.id),
            });
            continue;
          }
          if (assignedSlides.has(targetSlide)) {
            diagnostics.push({
              code: "invalid-xml-value",
              category: "corrupt-input",
              disposition: "omitted",
              location: {
                part: presentationPath,
                xpath: sldIdCursor.xpath,
              },
              attributeOrElement: "p14:sldId/@id",
              rawValue: String(parsedSlideId.data.id),
            });
            continue;
          }
          assignedSlides.add(targetSlide);
          sectionSlides.push(targetSlide);
        }
      }

      parsedSections.push({ name: sectionName, slides: sectionSlides });
    }
  }

  if (parsedSections.length === 0) return;

  const directSlides = deckNode.children.filter(
    (child): child is CommonNode =>
      child instanceof CommonNode && child.type === "slide",
  );
  const unassignedSlides = directSlides.filter(
    (slide) => !assignedSlides.has(slide),
  );

  for (const slide of directSlides) {
    slide.remove();
  }

  if (unassignedSlides.length > 0) {
    const fallbackName = allocateScopedAuthorId(
      undefined,
      usedSectionNames,
      "Default Section",
    );
    const defaultSection = createSectionNode({ name: fallbackName });
    for (const slide of unassignedSlides) {
      defaultSection.append(slide);
    }
    deckNode.append(defaultSection);
  }

  for (const entry of parsedSections) {
    const sectionNode = createSectionNode({ name: entry.name });
    for (const slide of entry.slides) {
      sectionNode.append(slide);
    }
    deckNode.append(sectionNode);
  }
}
