import { xmlDocument } from "@react-pptx/core/capability";

const drawingNamespace =
  "http://schemas.openxmlformats.org/drawingml/2006/main";
const presentationNamespace =
  "http://schemas.openxmlformats.org/presentationml/2006/main";
const relationshipNamespace =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export const notesMasterContentType =
  "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml";
export const notesSlideContentType =
  "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
export const notesMasterRelationshipType = `${relationshipNamespace}/notesMaster`;
export const notesSlideRelationshipType = `${relationshipNamespace}/notesSlide`;
export const themeRelationshipType = `${relationshipNamespace}/theme`;
export const slideRelationshipType = `${relationshipNamespace}/slide`;
export const notesMasterPathTemplate = "ppt/notesMasters/notesMaster{n}.xml";
export const notesSlidePathTemplate = "ppt/notesSlides/notesSlide{n}.xml";
export const themePath = "ppt/theme/theme1.xml";

const colorMap =
  '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';

function groupShapeTree(shapes: string): string {
  return `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapes}`;
}

/**
 * Shared notes master for one deck.
 *
 * The master is intentionally fixed rather than author-composed: it supplies the slide-image and
 * body placeholders plus the notes text defaults, while the page's notes body carries the actual
 * rich-text content. The notes theme relationship points at the deck theme so page theme tokens and
 * embedded font resources keep their normal ownership.
 */
export function notesMasterXml(): string {
  const slideImage =
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="1143000"/><a:ext cx="5486400" cy="3086100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr></p:sp>';
  const body =
    '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="4400550"/><a:ext cx="5486400" cy="3600450"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US" sz="1200"/></a:p></p:txBody></p:sp>';
  const notesStyle =
    '<p:notesStyle><a:lvl1pPr><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:notesStyle>';
  return xmlDocument(
    `<p:notesMaster xmlns:a="${drawingNamespace}" xmlns:r="${relationshipNamespace}" xmlns:p="${presentationNamespace}"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${groupShapeTree(slideImage + body)}</p:spTree></p:cSld>${colorMap}${notesStyle}</p:notesMaster>`,
  );
}

/** One notes slide body placeholder containing the parsed notes rich text. */
export function notesSlideXml(textBody: string): string {
  const slideImage =
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>';
  const body = `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>${textBody}</p:sp>`;
  return xmlDocument(
    `<p:notes xmlns:a="${drawingNamespace}" xmlns:r="${relationshipNamespace}" xmlns:p="${presentationNamespace}"><p:cSld><p:spTree>${groupShapeTree(slideImage + body)}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`,
  );
}
