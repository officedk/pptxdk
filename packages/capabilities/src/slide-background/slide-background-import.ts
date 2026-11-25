import type { BackgroundRef, SlideBackground } from "@pptxdk/core/capability";
import {
  A_NS,
  ooxmlBoolSchema,
  P_NS,
  type ImportDiagnostic,
  type ImportLocation,
  type PptxPackage,
  type SlideImportContext,
  type SlideImportOutcome,
  type XmlElementCursor,
} from "@pptxdk/import";
import { z } from "zod";

import {
  readBlipFill,
  readDrawingColor,
  readGradientFill,
} from "../shared/visual.ts";

const BgAttrsSchema = z.object({
  shadeToTitle: ooxmlBoolSchema.optional(),
});

const BgPrAttrsSchema = z.object({
  rotWithShape: ooxmlBoolSchema.optional(),
});

const bgRefAttrsSchema = z.object({
  idx: z.enum(["1001", "1002", "1003"]),
});

const processedCommonSlideElements = new WeakSet<object>();

function readBackgroundElement(
  bgCursor: XmlElementCursor,
  pkg: PptxPackage,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): SlideBackground | undefined {
  const rawBgAttrs = bgCursor.attrsObject(["shadeToTitle"]);
  const parsedBgAttrs = BgAttrsSchema.safeParse(rawBgAttrs);
  if (!parsedBgAttrs.success) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "recovered",
      location: { ...location, xpath: bgCursor.xpath },
      attributeOrElement: "p:bg",
      rawValue: JSON.stringify(rawBgAttrs),
    });
  }

  let result: SlideBackground | undefined;
  for (const child of bgCursor.remainingElementChildren()) {
    if (child.namespaceUri !== P_NS) {
      diagnostics.push({
        code: "unconsumed-xml",
        category: "unsupported",
        disposition: "omitted",
        kind: "element",
        location: { ...location, xpath: child.xpath },
        namespaceUri: child.namespaceUri,
        localName: child.localName,
      });
      continue;
    }

    if (child.localName === "bgPr") {
      const rawPrAttrs = child.attrsObject(["rotWithShape"]);
      const parsedPrAttrs = BgPrAttrsSchema.safeParse(rawPrAttrs);
      if (!parsedPrAttrs.success) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "recovered",
          location: { ...location, xpath: child.xpath },
          attributeOrElement: "p:bgPr",
          rawValue: JSON.stringify(rawPrAttrs),
        });
      }
      child.ignoreEmptyChild(A_NS, "effectLst");

      for (const fillChild of child.remainingElementChildren()) {
        if (fillChild.namespaceUri !== A_NS) {
          diagnostics.push({
            code: "unconsumed-xml",
            category: "unsupported",
            disposition: "omitted",
            kind: "element",
            location: { ...location, xpath: fillChild.xpath },
            namespaceUri: fillChild.namespaceUri,
            localName: fillChild.localName,
          });
          continue;
        }
        switch (fillChild.localName) {
          case "noFill":
            break;
          case "solidFill":
            result = readDrawingColor(fillChild, location, diagnostics);
            break;
          case "gradFill":
            result = readGradientFill(fillChild, location, diagnostics);
            break;
          case "blipFill": {
            const blip = readBlipFill(fillChild, pkg, location, diagnostics);
            if (blip !== undefined) {
              result = { image: blip.image };
            }
            break;
          }
          default:
            diagnostics.push({
              code: "unsupported-mechanism",
              category: "unsupported",
              disposition: "omitted",
              location: { ...location, xpath: fillChild.xpath },
              namespaceUri: fillChild.namespaceUri,
              localName: fillChild.localName,
              mechanismHint: `background:${fillChild.localName}`,
            });
            break;
        }
      }
    } else if (child.localName === "bgRef") {
      const rawRefAttrs = child.attrsObject(["idx"]);
      const parsedRefAttrs = bgRefAttrsSchema.safeParse(rawRefAttrs);
      if (!parsedRefAttrs.success) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "omitted",
          location: { ...location, xpath: child.xpath },
          attributeOrElement: "p:bgRef/@idx",
          rawValue: rawRefAttrs["idx"] ?? "",
        });
        child.suppressUnconsumed();
        continue;
      }
      const color = readDrawingColor(child, location, diagnostics);
      const ref: BackgroundRef = {
        ref: Number(parsedRefAttrs.data.idx) as 1001 | 1002 | 1003,
        ...(color === undefined ? {} : { color }),
      };
      result = ref;
    } else {
      diagnostics.push({
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: { ...location, xpath: child.xpath },
        namespaceUri: child.namespaceUri,
        localName: child.localName,
        mechanismHint: `background:${child.localName}`,
      });
    }
  }
  return result;
}

/**
 * Shared PresentationML `<p:cSld>/<p:bg>` import handler for slides, slide layouts, and slide masters.
 * Deduplicates by `<p:cSld>` element identity so multiple background mechanisms share one read pass.
 */
export function readSlideBackgroundSlot(
  context: SlideImportContext,
): SlideImportOutcome {
  const cSld = context.slideCursor.child(P_NS, "cSld");
  if (cSld === undefined) return {};
  if (processedCommonSlideElements.has(cSld.element)) return {};
  processedCommonSlideElements.add(cSld.element);

  const bgCursor = cSld.child(P_NS, "bg");
  if (bgCursor === undefined) return {};

  const diagnostics: ImportDiagnostic[] = [];
  const location: ImportLocation = {
    part: context.part,
    slideIndex: context.slideIndex,
  };
  const background = readBackgroundElement(
    bgCursor,
    context.package,
    location,
    diagnostics,
  );
  if (background !== undefined) {
    context.slideNode.set({ background });
  }
  return diagnostics.length === 0 ? {} : { diagnostics };
}
