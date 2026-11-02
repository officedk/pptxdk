export function pxToEmu(value: number): number {
  return Math.round(value * 9525);
}

export function assertXml10Characters(value: string): void {
  for (const character of value) {
    const codePoint = character.codePointAt(0);
    if (
      codePoint !== undefined &&
      codePoint !== 0x9 &&
      codePoint !== 0xa &&
      codePoint !== 0xd &&
      (codePoint < 0x20 ||
        (codePoint >= 0xd800 && codePoint <= 0xdfff) ||
        codePoint === 0xfffe ||
        codePoint === 0xffff)
    ) {
      throw new Error("字符串包含 XML 1.0 不允许的字符");
    }
  }
}

export function xmlEscape(value: string): string {
  assertXml10Characters(value);
  // XML 会规范化属性中的原始空白，字符引用可以保留身份值。
  return value
    .replace(/&/gu, "&amp;")
    .replace(/</gu, "&lt;")
    .replace(/>/gu, "&gt;")
    .replace(/"/gu, "&quot;")
    .replace(/'/gu, "&apos;")
    .replace(/\r/gu, "&#xD;")
    .replace(/\n/gu, "&#xA;")
    .replace(/\t/gu, "&#x9;");
}

export function xmlDocument(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${body}`;
}

const namespaces = {
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  relationships: "http://schemas.openxmlformats.org/package/2006/relationships",
  contentTypes: "http://schemas.openxmlformats.org/package/2006/content-types",
} as const;

const emptyShapeTree =
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
const minimumSlideLayoutId = 2_147_483_648;

function relationships(
  entries: readonly { id: string; type: string; target: string }[],
): string {
  return xmlDocument(
    `<Relationships xmlns="${namespaces.relationships}">${entries
      .map(
        ({ id, type, target }) =>
          `<Relationship Id="${xmlEscape(id)}" Type="${xmlEscape(type)}" Target="${xmlEscape(target)}"/>`,
      )
      .join("")}</Relationships>`,
  );
}

export type PptxPackageResource = {
  readonly name: string;
  readonly bytes: Uint8Array;
  readonly extension: string;
  readonly contentType: string;
};

export type PptxRelationship = {
  readonly id: string;
  readonly type: string;
  readonly target: string;
};

function contentTypes(
  slideCount: number,
  resources: readonly PptxPackageResource[],
): string {
  const overrides = [
    [
      "/ppt/presentation.xml",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
    ],
    [
      "/ppt/slideMasters/slideMaster1.xml",
      "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
    ],
    [
      "/ppt/slideLayouts/slideLayout1.xml",
      "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
    ],
    [
      "/ppt/theme/theme1.xml",
      "application/vnd.openxmlformats-officedocument.theme+xml",
    ],
    ...Array.from({ length: slideCount }, (_, index) => [
      `/ppt/slides/slide${index + 1}.xml`,
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    ]),
  ];
  const defaults = new Map<string, string>([
    ["rels", "application/vnd.openxmlformats-package.relationships+xml"],
    ["xml", "application/xml"],
  ]);
  for (const resource of resources) {
    defaults.set(resource.extension, resource.contentType);
  }
  return xmlDocument(
    `<Types xmlns="${namespaces.contentTypes}">${[...defaults]
      .map(
        ([extension, contentType]) =>
          `<Default Extension="${xmlEscape(extension)}" ContentType="${xmlEscape(contentType)}"/>`,
      )
      .join("")}${overrides
      .map(
        ([partName, contentType]) =>
          `<Override PartName="${xmlEscape(partName!)}" ContentType="${xmlEscape(contentType!)}"/>`,
      )
      .join("")}</Types>`,
  );
}

function theme(): string {
  const colors = `<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>`;
  const fonts = `<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>`;
  const fill = `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>`;
  const line = `<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>`;
  const effect = `<a:effectStyle><a:effectLst/></a:effectStyle>`;
  const format = `<a:fmtScheme name="Office"><a:fillStyleLst>${fill}${fill}${fill}</a:fillStyleLst><a:lnStyleLst>${line}${line}${line}</a:lnStyleLst><a:effectStyleLst>${effect}${effect}${effect}</a:effectStyleLst><a:bgFillStyleLst>${fill}${fill}${fill}</a:bgFillStyleLst></a:fmtScheme>`;
  return xmlDocument(
    `<a:theme xmlns:a="${namespaces.a}" name="Office Theme"><a:themeElements>${colors}${fonts}${format}</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`,
  );
}

function master(): string {
  const colorMap =
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
  return xmlDocument(
    `<p:sldMaster xmlns:a="${namespaces.a}" xmlns:r="${namespaces.r}" xmlns:p="${namespaces.p}"><p:cSld name="Office Theme"><p:spTree>${emptyShapeTree}</p:spTree></p:cSld>${colorMap}<p:sldLayoutIdLst><p:sldLayoutId id="${minimumSlideLayoutId}" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`,
  );
}

function layout(): string {
  return xmlDocument(
    `<p:sldLayout xmlns:a="${namespaces.a}" xmlns:r="${namespaces.r}" xmlns:p="${namespaces.p}" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree>${emptyShapeTree}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
  );
}

function slide(slideId: string, shapeXml: string, slideXml: string): string {
  return xmlDocument(
    `<p:sld xmlns:a="${namespaces.a}" xmlns:r="${namespaces.r}" xmlns:p="${namespaces.p}"><p:cSld name="${xmlEscape(slideId)}"><p:spTree>${emptyShapeTree}${shapeXml}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${slideXml}</p:sld>`,
  );
}

export function pptxParts(
  size: { readonly width: number; readonly height: number },
  slides: readonly {
    readonly id: string;
    readonly shapeXml: string;
    readonly slideXml?: string;
    readonly relationships: readonly PptxRelationship[];
  }[],
  resources: readonly PptxPackageResource[] = [],
): Record<string, string | Uint8Array> {
  const slideRelationships = slides.map((_, index) => ({
    id: `rId${index + 2}`,
    type: `${namespaces.r}/slide`,
    target: `slides/slide${index + 1}.xml`,
  }));
  const slideIds = slides
    .map((_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 2}"/>`)
    .join("");
  const presentation = xmlDocument(
    `<p:presentation xmlns:a="${namespaces.a}" xmlns:r="${namespaces.r}" xmlns:p="${namespaces.p}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${slideIds}</p:sldIdLst><p:sldSz cx="${pxToEmu(size.width)}" cy="${pxToEmu(size.height)}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
  );
  const parts: Record<string, string | Uint8Array> = {
    "[Content_Types].xml": contentTypes(slides.length, resources),
    "_rels/.rels": relationships([
      {
        id: "rId1",
        type: `${namespaces.r}/officeDocument`,
        target: "ppt/presentation.xml",
      },
    ]),
    "ppt/presentation.xml": presentation,
    "ppt/_rels/presentation.xml.rels": relationships([
      {
        id: "rId1",
        type: `${namespaces.r}/slideMaster`,
        target: "slideMasters/slideMaster1.xml",
      },
      ...slideRelationships,
    ]),
    "ppt/slideMasters/slideMaster1.xml": master(),
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": relationships([
      {
        id: "rId1",
        type: `${namespaces.r}/slideLayout`,
        target: "../slideLayouts/slideLayout1.xml",
      },
      {
        id: "rId2",
        type: `${namespaces.r}/theme`,
        target: "../theme/theme1.xml",
      },
    ]),
    "ppt/slideLayouts/slideLayout1.xml": layout(),
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": relationships([
      {
        id: "rId1",
        type: `${namespaces.r}/slideMaster`,
        target: "../slideMasters/slideMaster1.xml",
      },
    ]),
    "ppt/theme/theme1.xml": theme(),
  };
  slides.forEach(({ shapeXml, slideXml }, index) => {
    parts[`ppt/slides/slide${index + 1}.xml`] = slide(
      slides[index]!.id,
      shapeXml,
      slideXml ?? "",
    );
    parts[`ppt/slides/_rels/slide${index + 1}.xml.rels`] = relationships([
      {
        id: "rId1",
        type: `${namespaces.r}/slideLayout`,
        target: "../slideLayouts/slideLayout1.xml",
      },
      ...slides[index]!.relationships,
    ]);
  });
  for (const resource of resources) {
    parts[`ppt/media/${resource.name}`] = resource.bytes;
  }
  return parts;
}
