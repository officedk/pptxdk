/** Internal PresentationML wrapper for native shape objects. */
export function shapeXml(
  id: number,
  escapedName: string,
  properties: string,
  textBody: string,
  linkXml: string,
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapedName}">${linkXml}</p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr>${properties}${textBody}</p:sp>`;
}
