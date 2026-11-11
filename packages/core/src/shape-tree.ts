const emptyShapeTree =
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

/**
 * Internal PresentationML shape tree wrapper shared by slides, masters, and layouts.
 *
 * Capability writers contribute complete native object XML; this module owns only the surrounding
 * shape-tree envelope and its required non-visual group properties.
 */
export function shapeTreeXml(shapeXml = ""): string {
  return `<p:spTree>${emptyShapeTree}${shapeXml}</p:spTree>`;
}
