import { xmlEscape } from "@react-pptx/core/capability";

/** Internal PresentationML wrapper for native picture objects and their crop/opacity values. */
export function imageXml(
  id: number,
  escapedName: string,
  escapedAlt: string | undefined,
  relationshipId: string,
  box: { x: number; y: number; width: number; height: number },
  crop: Record<string, number>,
  properties: string,
  opacity: number | undefined,
  linkXml: string,
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
  return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${escapedName}"${description}>${linkXml}</p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${xmlEscape(relationshipId)}">${alpha}</a:blip>${cropXml}<a:stretch><a:fillRect/></a:stretch></p:blipFill>${properties}</p:pic>`;
}
