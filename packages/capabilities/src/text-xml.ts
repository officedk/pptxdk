/** Internal PresentationML text box wrapper for the text capability. */
export function textBoxXml(
  id: number,
  escapedName: string,
  box: {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
  },
  textBody: string,
  properties: string,
  linkXml: string,
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapedName}">${linkXml}</p:cNvPr><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>${properties}${textBody}</p:sp>`;
}
