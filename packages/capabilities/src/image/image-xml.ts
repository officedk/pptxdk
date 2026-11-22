import type { NativePlaceholderRef } from "@pptxdk/core/capability";
import { placeholderRefXml, xmlEscape } from "@pptxdk/emit";

/** Internal PresentationML wrapper for native picture objects and their crop/opacity values. */
export function imageXml(
  id: number,
  escapedName: string,
  escapedAlt: string | undefined,
  blipRef: { readonly kind: "embed" | "link"; readonly relationshipId: string },
  _box: { x: number; y: number; width: number; height: number },
  crop: Record<string, number>,
  properties: string,
  opacity: number | undefined,
  linkXml: string,
  placeholder?: {
    readonly nativeRef: NativePlaceholderRef;
    readonly spPrXml?: string;
  },
): string {
  const description = escapedAlt === undefined ? "" : ` descr="${escapedAlt}"`;
  const cropXml =
    Object.keys(crop).length === 0
      ? ""
      : `<a:srcRect${Object.entries(crop)
          .map(
            ([side, fraction]) => ` ${side}="${Math.round(fraction * 100000)}"`,
          )
          .join("")}/>`;
  const alpha =
    opacity === undefined
      ? ""
      : `<a:alphaModFix amt="${Math.round(opacity * 100000)}"/>`;
  const blipAttr =
    blipRef.kind === "link"
      ? `r:link="${xmlEscape(blipRef.relationshipId)}"`
      : `r:embed="${xmlEscape(blipRef.relationshipId)}"`;
  const cNvPicPr =
    placeholder === undefined
      ? '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
      : '<p:cNvPicPr><a:picLocks noGrp="1" noChangeAspect="1"/></p:cNvPicPr>';
  const nvPr =
    placeholder === undefined
      ? "<p:nvPr/>"
      : `<p:nvPr>${placeholderRefXml(placeholder.nativeRef)}</p:nvPr>`;
  const spPr = placeholder?.spPrXml ?? properties;
  return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${escapedName}"${description}>${linkXml}</p:cNvPr>${cNvPicPr}${nvPr}</p:nvPicPr><p:blipFill><a:blip ${blipAttr}>${alpha}</a:blip>${cropXml}<a:stretch><a:fillRect/></a:stretch></p:blipFill>${spPr}</p:pic>`;
}
