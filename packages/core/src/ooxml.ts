import { createHash } from "node:crypto";

import { defaultTheme, isFont, nativeColorSlots, type Theme } from "#/theme.ts";

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
  entries: readonly {
    id: string;
    type: string;
    target: string;
    targetMode?: "External";
  }[],
): string {
  return xmlDocument(
    `<Relationships xmlns="${namespaces.relationships}">${entries
      .map(
        ({ id, type, target, targetMode }) =>
          `<Relationship Id="${xmlEscape(id)}" Type="${xmlEscape(type)}" Target="${xmlEscape(target)}"${targetMode === undefined ? "" : ` TargetMode="${xmlEscape(targetMode)}"`}/>`,
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
  readonly targetMode?: "External";
};

export type PptxPackagePart = {
  readonly path: string;
  readonly content: string | Uint8Array;
  readonly contentType: string;
};

export type PptxPackageRelationshipSet = {
  readonly sourcePath: string;
  readonly relationships: readonly PptxRelationship[];
};

function contentTypes(
  slideCount: number,
  resources: readonly PptxPackageResource[],
  packageParts: readonly PptxPackagePart[],
  hasCoreProperties: boolean,
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
    ...(hasCoreProperties
      ? [
          [
            "/docProps/core.xml",
            "application/vnd.openxmlformats-package.core-properties+xml",
          ],
        ]
      : []),
    ...packageParts.map((part) => [`/${part.path}`, part.contentType]),
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

function themePart(value: Theme): string {
  const name = xmlEscape(value.name);
  const colors = `<a:clrScheme name="${name}">${nativeColorSlots
    .map((slot) => {
      const color = value.colors[slot];
      if (!/^#[0-9a-f]{6}$/iu.test(color)) {
        throw new Error(`Theme.colors.${slot} 必须是 #RRGGBB 颜色`);
      }
      const schemeName =
        slot === "hlink" ? "hlink" : slot === "folHlink" ? "folHlink" : slot;
      return `<a:${schemeName}><a:srgbClr val="${color.slice(1).toUpperCase()}"/></a:${schemeName}>`;
    })
    .join("")}</a:clrScheme>`;
  const major = isFont(value.fonts.major)
    ? value.fonts.major.family
    : value.fonts.major;
  const minor = isFont(value.fonts.minor)
    ? value.fonts.minor.family
    : value.fonts.minor;
  const majorEscaped = xmlEscape(major);
  const minorEscaped = xmlEscape(minor);
  const fonts = `<a:fontScheme name="${name}"><a:majorFont><a:latin typeface="${majorEscaped}"/><a:ea typeface="${majorEscaped}"/><a:cs typeface="${majorEscaped}"/></a:majorFont><a:minorFont><a:latin typeface="${minorEscaped}"/><a:ea typeface="${minorEscaped}"/><a:cs typeface="${minorEscaped}"/></a:minorFont></a:fontScheme>`;
  const fill = `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>`;
  const line = `<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>`;
  const effect = `<a:effectStyle><a:effectLst/></a:effectStyle>`;
  const format = `<a:fmtScheme name="Office"><a:fillStyleLst>${fill}${fill}${fill}</a:fillStyleLst><a:lnStyleLst>${line}${line}${line}</a:lnStyleLst><a:effectStyleLst>${effect}${effect}${effect}</a:effectStyleLst><a:bgFillStyleLst>${fill}${fill}${fill}</a:bgFillStyleLst></a:fmtScheme>`;
  return xmlDocument(
    `<a:theme xmlns:a="${namespaces.a}" name="${name}"><a:themeElements>${colors}${fonts}${format}</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`,
  );
}

function master(): string {
  const colorMap =
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
  // 单母版、单版式骨架必须避开 ID 冲突，因此版式取下一个合法值。
  return xmlDocument(
    `<p:sldMaster xmlns:a="${namespaces.a}" xmlns:r="${namespaces.r}" xmlns:p="${namespaces.p}"><p:cSld name="Office Theme"><p:spTree>${emptyShapeTree}</p:spTree></p:cSld>${colorMap}<p:sldLayoutIdLst><p:sldLayoutId id="${minimumSlideLayoutId + 1}" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`,
  );
}

function layout(): string {
  return xmlDocument(
    `<p:sldLayout xmlns:a="${namespaces.a}" xmlns:r="${namespaces.r}" xmlns:p="${namespaces.p}" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree>${emptyShapeTree}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
  );
}

function slide(
  slideId: string,
  shapeXml: string,
  slideXml: string,
  backgroundXml: string,
  hidden: boolean,
): string {
  return xmlDocument(
    `<p:sld xmlns:a="${namespaces.a}" xmlns:r="${namespaces.r}" xmlns:p="${namespaces.p}"${hidden ? ' show="false"' : ""}><p:cSld name="${xmlEscape(slideId)}">${backgroundXml}<p:spTree>${emptyShapeTree}${shapeXml}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${slideXml}</p:sld>`,
  );
}

function sectionId(index: number, name: string): string {
  const hex = createHash("sha1")
    .update(`section:${index}:${name}`)
    .digest("hex")
    .toUpperCase();
  return `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20, 32)}}`;
}

function sectionListXml(
  sections: readonly {
    readonly name: string;
    readonly slideIndexes: readonly number[];
  }[],
): string {
  if (sections.length === 0) return "";
  const entries = sections
    .map(
      ({ name, slideIndexes }, index) =>
        `<p14:section name="${xmlEscape(name)}" id="${sectionId(index, name)}"><p14:sldIdLst>${slideIndexes
          .map((slideIndex) => `<p14:sldId id="${255 + slideIndex}"/>`)
          .join("")}</p14:sldIdLst></p14:section>`,
    )
    .join("");
  return `<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">${entries}</p14:sectionLst></p:ext></p:extLst>`;
}

function coreProperties(title: string): string {
  return xmlDocument(
    `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xmlEscape(title)}</dc:title></cp:coreProperties>`,
  );
}

export function pptxParts(
  size: { readonly width: number; readonly height: number },
  slides: readonly {
    readonly id: string;
    readonly hidden?: boolean;
    readonly shapeXml: string;
    readonly backgroundXml?: string;
    readonly slideXml?: string;
    readonly relationships: readonly PptxRelationship[];
  }[],
  resources: readonly PptxPackageResource[] = [],
  packageParts: readonly PptxPackagePart[] = [],
  packageRelationships: readonly PptxPackageRelationshipSet[] = [],
  presentationContent = "",
  presentationAttributes: readonly { name: string; value: string }[] = [],
  themeValue: Theme = defaultTheme,
  title?: string,
  sections: readonly {
    readonly name: string;
    readonly slideIndexes: readonly number[];
  }[] = [],
): Record<string, string | Uint8Array> {
  const slideRelationships = slides.map((_, index) => ({
    id: `rId${index + 2}`,
    type: `${namespaces.r}/slide`,
    target: `slides/slide${index + 1}.xml`,
  }));
  const slideIds = slides
    .map((_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 2}"/>`)
    .join("");
  const presentationAttributeXml = presentationAttributes
    .map(({ name, value }) => ` ${name}="${xmlEscape(value)}"`)
    .join("");
  const extraPresentationRelationships =
    packageRelationships.find(
      ({ sourcePath }) => sourcePath === "ppt/presentation.xml",
    )?.relationships ?? [];
  const presentation = xmlDocument(
    `<p:presentation xmlns:a="${namespaces.a}" xmlns:r="${namespaces.r}" xmlns:p="${namespaces.p}"${presentationAttributeXml}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${slideIds}</p:sldIdLst><p:sldSz cx="${pxToEmu(size.width)}" cy="${pxToEmu(size.height)}"/><p:notesSz cx="6858000" cy="9144000"/>${presentationContent}${sectionListXml(sections)}</p:presentation>`,
  );
  const parts: Record<string, string | Uint8Array> = {
    "[Content_Types].xml": contentTypes(
      slides.length,
      resources,
      packageParts,
      title !== undefined,
    ),
    "_rels/.rels": relationships([
      {
        id: "rId1",
        type: `${namespaces.r}/officeDocument`,
        target: "ppt/presentation.xml",
      },
      ...(title === undefined
        ? []
        : [
            {
              id: "rId2",
              type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
              target: "docProps/core.xml",
            },
          ]),
    ]),
    "ppt/presentation.xml": presentation,
    "ppt/_rels/presentation.xml.rels": relationships([
      {
        id: "rId1",
        type: `${namespaces.r}/slideMaster`,
        target: "slideMasters/slideMaster1.xml",
      },
      ...slideRelationships,
      ...extraPresentationRelationships,
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
    "ppt/theme/theme1.xml": themePart(themeValue),
  };
  slides.forEach(({ shapeXml, slideXml, backgroundXml, hidden }, index) => {
    parts[`ppt/slides/slide${index + 1}.xml`] = slide(
      slides[index]!.id,
      shapeXml,
      slideXml ?? "",
      backgroundXml ?? "",
      hidden ?? false,
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
  if (title !== undefined) parts["docProps/core.xml"] = coreProperties(title);
  for (const resource of resources) {
    parts[`ppt/media/${resource.name}`] = resource.bytes;
  }
  for (const part of packageParts) {
    parts[part.path] = part.content;
  }
  for (const relationshipSet of packageRelationships) {
    if (relationshipSet.sourcePath === "ppt/presentation.xml") continue;
    const marker = relationshipSet.sourcePath.lastIndexOf("/");
    const directory =
      marker === -1 ? "" : relationshipSet.sourcePath.slice(0, marker + 1);
    const name = relationshipSet.sourcePath.slice(marker + 1);
    parts[`${directory}_rels/${name}.rels`] = relationships(
      relationshipSet.relationships,
    );
  }
  return parts;
}
