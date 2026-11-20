type EmuBox = {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
};

/** Internal PresentationML wrapper for native groups; member baselines stay in slide coordinates. */
export function groupXml(
  id: number,
  escapedName: string,
  box: EmuBox,
  members: string,
  rotation: number | undefined,
  linkXml: string,
): string {
  const angle =
    rotation === undefined
      ? undefined
      : Math.round((((rotation % 360) + 360) % 360) * 60000);
  return `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="${escapedName}">${linkXml}</p:cNvPr><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm${angle === undefined || angle === 0 ? "" : ` rot="${angle}"`}><a:off x="${box.x}" y="${box.y}"/><a:ext cx="${box.width}" cy="${box.height}"/><a:chOff x="${box.x}" y="${box.y}"/><a:chExt cx="${box.width}" cy="${box.height}"/></a:xfrm></p:grpSpPr>${members}</p:grpSp>`;
}
