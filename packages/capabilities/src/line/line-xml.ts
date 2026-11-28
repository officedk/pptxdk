export type LineConnectionEndpointXml = {
  readonly id: number;
  readonly idx: number;
};

export type LineConnectionXml = {
  readonly start?: LineConnectionEndpointXml;
  readonly end?: LineConnectionEndpointXml;
};

/** Internal PresentationML wrapper for native `<p:cxnSp>` connector objects. */
export function lineXml(
  id: number,
  escapedName: string,
  properties: string,
  linkXml: string,
  connection?: LineConnectionXml,
): string {
  const stCxn =
    connection?.start === undefined
      ? ""
      : `<a:stCxn id="${connection.start.id}" idx="${connection.start.idx}"/>`;
  const endCxn =
    connection?.end === undefined
      ? ""
      : `<a:endCxn id="${connection.end.id}" idx="${connection.end.idx}"/>`;
  const cxnBody = stCxn === "" && endCxn === "" ? "" : `${stCxn}${endCxn}`;
  const cxnProps =
    cxnBody === ""
      ? "<p:cNvCxnSpPr/>"
      : `<p:cNvCxnSpPr>${cxnBody}</p:cNvCxnSpPr>`;
  return `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="${escapedName}">${linkXml}</p:cNvPr>${cxnProps}<p:nvPr/></p:nvCxnSpPr>${properties}</p:cxnSp>`;
}
