import { posix } from "node:path";

import validateOpenXml from "@ooxml-tools/validate";
import { DOMParser } from "@xmldom/xmldom";
import { strFromU8, unzipSync } from "fflate";

const presentationNamespace =
  "http://schemas.openxmlformats.org/presentationml/2006/main";
const drawingNamespace =
  "http://schemas.openxmlformats.org/drawingml/2006/main";
const relationshipNamespace =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const packageRelationshipNamespace =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const contentTypeNamespace =
  "http://schemas.openxmlformats.org/package/2006/content-types";
const transitionalStandard = "ECMA-376 5th edition Transitional";
const schemaValidatorPackage = "@ooxml-tools/validate";
const schemaValidatorPackageVersion = "0.4.6";
const openXmlSdkVersion = "3.0.2";
const schemaValidatorOfficeVersion = "Microsoft365";
const schemaValidatorProfile = {
  package: schemaValidatorPackage,
  packageVersion: schemaValidatorPackageVersion,
  sdkVersion: openXmlSdkVersion,
  officeVersion: schemaValidatorOfficeVersion,
} as const;
const validationScope =
  "Open XML SDK 3.0.2 在 Microsoft365 配置下实现的 Transitional 包、Schema 与约束校验；另由本地读取器独立核对 OPC 关系、内容类型及页面、形状、图片、组合和文字结构。MC 仅按 SDK 支持选择 Choice 或 Fallback；p14／p159 的所有兼容分支及未知扩展不保证均被验证。校验不等于完整标准认证或应用证据。";
const applicationEvidence = "not-validated";
const minimumSlideLayoutId = 2_147_483_648;
const maximumSlideLayoutId = 4_294_967_295;
const minimumSlideSizeEmu = 914_400;
const maximumSlideSizeEmu = 51_206_400;
const maximumShapeExtentEmu = 2_147_483_647;
const minimumFontSizeCentipoints = 100;
const maximumFontSizeCentipoints = 400_000;
const maximumTextSpacingCentipoints = 158_400;

export type ValidationIssue = {
  readonly part: string;
  readonly code: string;
  readonly message: string;
  readonly xpath?: string;
};

export type FileObject = {
  readonly id: number;
  readonly name: string;
  readonly kind: "shape" | "image" | "group";
  readonly geometry?: string;
  readonly boxEmu: {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
  };
  readonly image?: {
    readonly relationshipId: string;
    readonly part: string;
    readonly contentType: string;
  };
  readonly crop?: {
    readonly left?: number;
    readonly top?: number;
    readonly right?: number;
    readonly bottom?: number;
  };
  readonly groupTransform?: {
    readonly childOffset: { readonly x: number; readonly y: number };
    readonly childExtent: { readonly width: number; readonly height: number };
  };
  readonly members?: readonly FileObject[];
  readonly text?: FileText;
};

export type FileTextRun = {
  readonly text: string;
  readonly fontFamily?: string;
  readonly fontSizePt?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly color?: string;
};

export type FileTextList =
  | { readonly type: "bullet"; readonly character: string }
  | {
      readonly type: "number";
      readonly style: string;
      readonly startAt?: number;
    };

export type FileTextParagraph = {
  readonly text: string;
  readonly runs: readonly FileTextRun[];
  readonly alignment?: string;
  readonly lineSpacingPoints?: number;
  readonly spacingAfterPoints?: number;
  readonly list?: FileTextList;
};

export type FileText = {
  readonly text: string;
  readonly paragraphs: readonly FileTextParagraph[];
};

export type FileSlide = {
  readonly id: string;
  readonly objects: readonly FileObject[];
};

export type ValidationResult = {
  readonly profile: {
    readonly standard: typeof transitionalStandard;
    readonly schemaValidator: typeof schemaValidatorProfile;
    readonly scope: string;
    readonly applicationEvidence: typeof applicationEvidence;
  };
  readonly errors: readonly ValidationIssue[];
  readonly presentation: {
    readonly slideCount: number;
    readonly widthEmu?: number;
    readonly heightEmu?: number;
  };
  readonly slides: readonly FileSlide[];
};

type XmlFiles = Record<string, Uint8Array>;
type XmlDocument = ReturnType<DOMParser["parseFromString"]>;
type XmlNode = NonNullable<XmlDocument["firstChild"]>;
type XmlElement = NonNullable<XmlDocument["documentElement"]>;
type Documents = Map<string, XmlDocument>;

function issue(
  errors: ValidationIssue[],
  part: string,
  code: string,
  message: string,
): void {
  errors.push({ part, code, message });
}

async function validateStandardStructure(
  input: Uint8Array,
): Promise<ValidationIssue[]> {
  try {
    const results = await validateOpenXml(
      input,
      "pptx",
      schemaValidatorOfficeVersion,
    );
    return results.map(({ id, description, path }) => ({
      part: path.partUri,
      code: id,
      message: description,
      ...(path.xpath === "" ? {} : { xpath: path.xpath }),
    }));
  } catch (error) {
    return [
      {
        part: "",
        code: "schema-validator-failed",
        message: `Open XML SDK 结构校验器执行失败：${error instanceof Error ? error.message : String(error)}`,
      },
    ];
  }
}

function elements(
  parent: XmlNode,
  namespace: string,
  localName: string,
): XmlElement[] {
  const result: XmlElement[] = [];
  for (
    let child = parent.firstChild;
    child !== null;
    child = child.nextSibling
  ) {
    if (
      child.nodeType === 1 &&
      child.namespaceURI === namespace &&
      child.localName === localName
    ) {
      result.push(child as XmlElement);
    }
  }
  return result;
}

function first(
  parent: XmlNode,
  namespace: string,
  localName: string,
): XmlElement | undefined {
  return elements(parent, namespace, localName)[0];
}

function attribute(element: XmlElement, name: string): string {
  return element.getAttribute(name) ?? "";
}

function integerAttribute(
  element: XmlElement | undefined,
  name: string,
): number | undefined {
  if (element === undefined) return undefined;
  const value = attribute(element, name);
  if (value === "" || !/^-?\d+$/u.test(value)) return undefined;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) ? parsed : undefined;
}

function decimalAttribute(
  element: XmlElement | undefined,
  name: string,
): number | undefined {
  if (element === undefined) return undefined;
  const value = attribute(element, name);
  if (value === "") return undefined;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
}

function readTextRun(run: XmlElement): FileTextRun {
  const text = first(run, drawingNamespace, "t")?.textContent ?? "";
  const properties = first(run, drawingNamespace, "rPr");
  if (properties === undefined) return { text };

  const latin = first(properties, drawingNamespace, "latin");
  const fill = first(properties, drawingNamespace, "solidFill");
  const color =
    fill === undefined ? undefined : first(fill, drawingNamespace, "srgbClr");
  const fontFamily = latin === undefined ? "" : attribute(latin, "typeface");
  const fontSize = decimalAttribute(properties, "sz");
  const bold = attribute(properties, "b");
  const italic = attribute(properties, "i");
  return {
    text,
    ...(fontFamily === "" ? {} : { fontFamily }),
    ...(fontSize === undefined ? {} : { fontSizePt: fontSize / 100 }),
    ...(bold === "" ? {} : { bold: bold === "1" || bold === "true" }),
    ...(italic === "" ? {} : { italic: italic === "1" || italic === "true" }),
    ...(color === undefined || attribute(color, "val") === ""
      ? {}
      : { color: `#${attribute(color, "val").toUpperCase()}` }),
  };
}

function readParagraph(paragraph: XmlElement): FileTextParagraph {
  const runs = elements(paragraph, drawingNamespace, "r").map(readTextRun);
  const paragraphProperties = first(paragraph, drawingNamespace, "pPr");
  const lineSpacing =
    paragraphProperties === undefined
      ? undefined
      : first(paragraphProperties, drawingNamespace, "lnSpc");
  const exactSpacing =
    lineSpacing === undefined
      ? undefined
      : first(lineSpacing, drawingNamespace, "spcPts");
  const lineSpacingValue = decimalAttribute(exactSpacing, "val");
  const spacingAfter =
    paragraphProperties === undefined
      ? undefined
      : first(paragraphProperties, drawingNamespace, "spcAft");
  const spacingAfterValue = decimalAttribute(
    spacingAfter === undefined
      ? undefined
      : first(spacingAfter, drawingNamespace, "spcPts"),
    "val",
  );
  const bullet =
    paragraphProperties === undefined
      ? undefined
      : first(paragraphProperties, drawingNamespace, "buChar");
  const autoNumber =
    paragraphProperties === undefined
      ? undefined
      : first(paragraphProperties, drawingNamespace, "buAutoNum");
  const alignment =
    paragraphProperties === undefined
      ? ""
      : attribute(paragraphProperties, "algn");
  let list: FileTextList | undefined;
  if (bullet !== undefined) {
    list = { type: "bullet", character: attribute(bullet, "char") };
  } else if (autoNumber !== undefined) {
    const startAt = integerAttribute(autoNumber, "startAt");
    list = {
      type: "number",
      style: attribute(autoNumber, "type"),
      ...(startAt === undefined ? {} : { startAt }),
    };
  }
  return {
    text: runs.map(({ text }) => text).join(""),
    runs,
    ...(alignment === "" ? {} : { alignment }),
    ...(lineSpacingValue === undefined
      ? {}
      : { lineSpacingPoints: lineSpacingValue / 100 }),
    ...(spacingAfterValue === undefined
      ? {}
      : { spacingAfterPoints: spacingAfterValue / 100 }),
    ...(list === undefined ? {} : { list }),
  };
}

function readTextBody(shape: XmlElement): FileText | undefined {
  const body = first(shape, presentationNamespace, "txBody");
  if (body === undefined) return undefined;
  const paragraphs = elements(body, drawingNamespace, "p").map(readParagraph);
  return {
    text: paragraphs.map(({ text }) => text).join("\n"),
    paragraphs,
  };
}

function parseXmlFiles(files: XmlFiles, errors: ValidationIssue[]): Documents {
  const documents: Documents = new Map();
  const parser = new DOMParser({
    onError(level, message) {
      if (level !== "warning") throw new Error(message);
    },
  });
  for (const [part, bytes] of Object.entries(files)) {
    if (!part.endsWith(".xml") && !part.endsWith(".rels")) continue;
    try {
      const document = parser.parseFromString(
        strFromU8(bytes),
        "application/xml",
      );
      if (document.documentElement === null) throw new Error("XML 缺少根元素");
      documents.set(part, document);
    } catch (error) {
      issue(
        errors,
        part,
        "invalid-xml",
        error instanceof Error ? error.message : String(error),
      );
    }
  }
  return documents;
}

function sourcePartForRelationships(path: string): string | undefined {
  if (path === "_rels/.rels") return "";
  const marker = "/_rels/";
  const index = path.lastIndexOf(marker);
  if (index === -1 || !path.endsWith(".rels")) return undefined;
  const directory = path.slice(0, index);
  const file = path.slice(index + marker.length, -".rels".length);
  return directory === "" ? file : `${directory}/${file}`;
}

function relationshipTargetPath(source: string, target: string): string {
  const joined = posix.normalize(
    posix.join(posix.dirname(source || "."), target),
  );
  return joined === "." ? "" : joined;
}

function checkRelationships(
  files: XmlFiles,
  documents: Documents,
  errors: ValidationIssue[],
): Map<string, Map<string, XmlElement>> {
  const bySource = new Map<string, Map<string, XmlElement>>();
  for (const [path, document] of documents) {
    if (!path.endsWith(".rels")) continue;
    const source = sourcePartForRelationships(path);
    if (source === undefined) continue;
    const root = document.documentElement;
    if (
      root === null ||
      root.namespaceURI !== packageRelationshipNamespace ||
      root.localName !== "Relationships"
    ) {
      issue(
        errors,
        path,
        "invalid-relationships-root",
        "关系部件根元素不符合 OPC 命名空间",
      );
      continue;
    }
    const entries = new Map<string, XmlElement>();
    for (const relationship of elements(
      root,
      packageRelationshipNamespace,
      "Relationship",
    )) {
      const id = attribute(relationship, "Id");
      const target = attribute(relationship, "Target");
      const mode = attribute(relationship, "TargetMode");
      if (
        id === "" ||
        target === "" ||
        attribute(relationship, "Type") === ""
      ) {
        issue(
          errors,
          path,
          "invalid-relationship",
          "关系缺少 Id、Type 或 Target",
        );
        continue;
      }
      if (entries.has(id))
        issue(errors, path, "duplicate-relationship-id", `关系标识 ${id} 重复`);
      entries.set(id, relationship);
      if (mode !== "External") {
        const resolved = relationshipTargetPath(source, target);
        if (resolved.startsWith("../") || files[resolved] === undefined) {
          issue(
            errors,
            path,
            "missing-relationship-target",
            `关系目标 ${target} 不存在`,
          );
        }
      }
    }
    bySource.set(source, entries);
  }
  return bySource;
}

function validateContentTypes(
  files: XmlFiles,
  documents: Documents,
  errors: ValidationIssue[],
): { defaults: Map<string, string>; overrides: Map<string, string> } {
  const part = "[Content_Types].xml";
  const document = documents.get(part);
  if (document === undefined) {
    issue(errors, part, "missing-content-types", "缺少内容类型部件");
    return { defaults: new Map(), overrides: new Map() };
  }
  const root = document.documentElement;
  if (
    root === null ||
    root.namespaceURI !== contentTypeNamespace ||
    root.localName !== "Types"
  ) {
    issue(
      errors,
      part,
      "invalid-content-types-root",
      "内容类型部件根元素不符合 OPC 命名空间",
    );
    return { defaults: new Map(), overrides: new Map() };
  }
  const defaults = new Map(
    elements(root, contentTypeNamespace, "Default").map((entry) => [
      attribute(entry, "Extension").toLowerCase(),
      attribute(entry, "ContentType"),
    ]),
  );
  const overrides = new Map(
    elements(root, contentTypeNamespace, "Override").map((entry) => [
      attribute(entry, "PartName"),
      attribute(entry, "ContentType"),
    ]),
  );
  const required: readonly (readonly [string, string])[] = [
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
  ];
  for (const [path, contentType] of required) {
    if (overrides.get(path) !== contentType) {
      issue(errors, part, "missing-content-type", `${path} 缺少预期内容类型`);
    }
  }
  for (const path of Object.keys(files).filter((entry) =>
    /^ppt\/slides\/slide\d+\.xml$/u.test(entry),
  )) {
    if (
      overrides.get(`/${path}`) !==
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
    ) {
      issue(errors, part, "missing-content-type", `/${path} 缺少页面内容类型`);
    }
  }
  return { defaults, overrides };
}

function contentTypeForPart(
  path: string,
  contentTypes: {
    defaults: Map<string, string>;
    overrides: Map<string, string>;
  },
): string | undefined {
  const override = contentTypes.overrides.get(`/${path}`);
  if (override !== undefined) return override;
  const extension = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
  return contentTypes.defaults.get(extension);
}

function imageSignatureMatches(
  bytes: Uint8Array,
  contentType: string,
): boolean {
  if (contentType === "image/png") {
    return (
      bytes.length >= 8 &&
      bytes[0] === 0x89 &&
      bytes[1] === 0x50 &&
      bytes[2] === 0x4e &&
      bytes[3] === 0x47 &&
      bytes[4] === 0x0d &&
      bytes[5] === 0x0a &&
      bytes[6] === 0x1a &&
      bytes[7] === 0x0a
    );
  }
  if (contentType === "image/jpeg") {
    return bytes.length >= 4 && bytes[0] === 0xff && bytes[1] === 0xd8;
  }
  if (contentType === "image/gif") {
    const signature = String.fromCharCode(...bytes.subarray(0, 6));
    return signature === "GIF87a" || signature === "GIF89a";
  }
  return false;
}

function requireRelationship(
  source: string,
  type: string,
  targetPath: string,
  relationships: Map<string, Map<string, XmlElement>>,
  errors: ValidationIssue[],
): void {
  const found = [...(relationships.get(source)?.values() ?? [])].some(
    (entry) =>
      attribute(entry, "Type") === type &&
      relationshipTargetPath(source, attribute(entry, "Target")) === targetPath,
  );
  if (!found) {
    issue(
      errors,
      source === "" ? "_rels/.rels" : source,
      "missing-required-relationship",
      `缺少指向 ${targetPath} 的关系`,
    );
  }
}

function requireRoot(
  documents: Documents,
  errors: ValidationIssue[],
  path: string,
  namespace: string,
  localName: string,
): void {
  const root = documents.get(path)?.documentElement;
  if (
    root == null ||
    root.namespaceURI !== namespace ||
    root.localName !== localName
  ) {
    issue(errors, path, "invalid-part-root", `部件根元素应为 ${localName}`);
  }
}

function validateSlideLayoutIds(
  documents: Documents,
  errors: ValidationIssue[],
): void {
  const part = "ppt/slideMasters/slideMaster1.xml";
  const root = documents.get(part)?.documentElement;
  if (root == null) return;
  const layoutIds = first(root, presentationNamespace, "sldLayoutIdLst");
  if (layoutIds === undefined) return;
  for (const layoutId of elements(
    layoutIds,
    presentationNamespace,
    "sldLayoutId",
  )) {
    const id = integerAttribute(layoutId, "id");
    if (
      id === undefined ||
      id < minimumSlideLayoutId ||
      id > maximumSlideLayoutId
    ) {
      issue(
        errors,
        part,
        "invalid-slide-layout-id",
        `母版版式标识 ${attribute(layoutId, "id")} 必须是 ${minimumSlideLayoutId} 至 ${maximumSlideLayoutId} 的整数`,
      );
    }
  }
}

function validateTextValueRange(
  element: XmlElement | undefined,
  name: string,
  minimum: number,
  maximum: number,
  part: string,
  errors: ValidationIssue[],
  code: "invalid-font-size" | "invalid-text-spacing",
  label: string,
): void {
  if (element === undefined || !element.hasAttribute(name)) return;
  const value = integerAttribute(element, name);
  if (value === undefined || value < minimum || value > maximum) {
    issue(
      errors,
      part,
      code,
      `${label}字段必须是 ${minimum} 到 ${maximum} 间的整数`,
    );
  }
}

function validateTextValueRanges(
  shape: XmlElement,
  part: string,
  errors: ValidationIssue[],
): void {
  const body = first(shape, presentationNamespace, "txBody");
  if (body === undefined) return;
  for (const paragraph of elements(body, drawingNamespace, "p")) {
    for (const run of elements(paragraph, drawingNamespace, "r")) {
      validateTextValueRange(
        first(run, drawingNamespace, "rPr"),
        "sz",
        minimumFontSizeCentipoints,
        maximumFontSizeCentipoints,
        part,
        errors,
        "invalid-font-size",
        "字号",
      );
    }
    validateTextValueRange(
      first(paragraph, drawingNamespace, "endParaRPr"),
      "sz",
      minimumFontSizeCentipoints,
      maximumFontSizeCentipoints,
      part,
      errors,
      "invalid-font-size",
      "字号",
    );
    const properties = first(paragraph, drawingNamespace, "pPr");
    for (const spacingName of ["lnSpc", "spcAft"] as const) {
      const spacing =
        properties === undefined
          ? undefined
          : first(properties, drawingNamespace, spacingName);
      validateTextValueRange(
        spacing === undefined
          ? undefined
          : first(spacing, drawingNamespace, "spcPts"),
        "val",
        0,
        maximumTextSpacingCentipoints,
        part,
        errors,
        "invalid-text-spacing",
        "段落间距",
      );
    }
  }
}

function validateShape(
  shape: XmlElement,
  part: string,
  errors: ValidationIssue[],
): FileObject | undefined {
  const shapeProperties = first(shape, presentationNamespace, "nvSpPr");
  const drawingProperties =
    shapeProperties === undefined
      ? undefined
      : first(shapeProperties, presentationNamespace, "cNvPr");
  const id = integerAttribute(drawingProperties, "id");
  const name =
    drawingProperties === undefined ? "" : attribute(drawingProperties, "name");
  if (id === undefined || name === "") {
    issue(errors, part, "invalid-shape-identity", "形状缺少有效对象标识或名称");
    return undefined;
  }
  const properties = first(shape, presentationNamespace, "spPr");
  const geometry =
    properties === undefined
      ? undefined
      : first(properties, drawingNamespace, "prstGeom");
  const geometryName =
    geometry === undefined ? "" : attribute(geometry, "prst");
  if (geometryName === "") {
    issue(
      errors,
      part,
      "missing-shape-geometry",
      `形状 ${name} 缺少 preset geometry`,
    );
    return undefined;
  }
  const transform =
    properties === undefined
      ? undefined
      : first(properties, drawingNamespace, "xfrm");
  const offset =
    transform === undefined
      ? undefined
      : first(transform, drawingNamespace, "off");
  const extent =
    transform === undefined
      ? undefined
      : first(transform, drawingNamespace, "ext");
  const x = integerAttribute(offset, "x");
  const y = integerAttribute(offset, "y");
  const width = integerAttribute(extent, "cx");
  const height = integerAttribute(extent, "cy");
  if (
    x === undefined ||
    y === undefined ||
    width === undefined ||
    height === undefined ||
    width < 1 ||
    width > maximumShapeExtentEmu ||
    height < 1 ||
    height > maximumShapeExtentEmu
  ) {
    issue(
      errors,
      part,
      "invalid-shape-transform",
      `形状 ${name} 的坐标或尺寸无效`,
    );
    return undefined;
  }
  validateTextValueRanges(shape, part, errors);
  const text = readTextBody(shape);
  return {
    id,
    name,
    kind: "shape",
    geometry: geometryName,
    boxEmu: { x, y, width, height },
    ...(text === undefined ? {} : { text }),
  };
}

function objectIdentity(
  object: XmlElement,
  propertiesName: string,
  part: string,
  kind: string,
  errors: ValidationIssue[],
): { id: number; name: string } | undefined {
  const properties = first(object, presentationNamespace, propertiesName);
  const drawingProperties =
    properties === undefined
      ? undefined
      : first(properties, presentationNamespace, "cNvPr");
  const id = integerAttribute(drawingProperties, "id");
  const name =
    drawingProperties === undefined ? "" : attribute(drawingProperties, "name");
  if (id === undefined || id < 1 || name === "") {
    issue(
      errors,
      part,
      `invalid-${kind}-identity`,
      `${kind} 缺少有效对象标识或名称`,
    );
    return undefined;
  }
  return { id, name };
}

function transformBox(
  transform: XmlElement | undefined,
  part: string,
  name: string,
  kind: string,
  errors: ValidationIssue[],
): FileObject["boxEmu"] | undefined {
  const offset =
    transform === undefined
      ? undefined
      : first(transform, drawingNamespace, "off");
  const extent =
    transform === undefined
      ? undefined
      : first(transform, drawingNamespace, "ext");
  const x = integerAttribute(offset, "x");
  const y = integerAttribute(offset, "y");
  const width = integerAttribute(extent, "cx");
  const height = integerAttribute(extent, "cy");
  if (
    x === undefined ||
    y === undefined ||
    width === undefined ||
    height === undefined ||
    width < 1 ||
    width > maximumShapeExtentEmu ||
    height < 1 ||
    height > maximumShapeExtentEmu
  ) {
    issue(
      errors,
      part,
      `invalid-${kind}-transform`,
      `${kind} ${name} 的坐标或尺寸无效`,
    );
    return undefined;
  }
  return { x, y, width, height };
}

function validatePicture(
  picture: XmlElement,
  part: string,
  files: XmlFiles,
  relationships: Map<string, Map<string, XmlElement>>,
  contentTypes: {
    defaults: Map<string, string>;
    overrides: Map<string, string>;
  },
  errors: ValidationIssue[],
): FileObject | undefined {
  const identity = objectIdentity(picture, "nvPicPr", part, "picture", errors);
  if (identity === undefined) return undefined;
  const properties = first(picture, presentationNamespace, "spPr");
  const geometry =
    properties === undefined
      ? undefined
      : first(properties, drawingNamespace, "prstGeom");
  const geometryName =
    geometry === undefined ? "" : attribute(geometry, "prst");
  if (geometryName === "") {
    issue(
      errors,
      part,
      "missing-picture-geometry",
      `图片 ${identity.name} 缺少几何`,
    );
    return undefined;
  }
  const boxEmu = transformBox(
    properties === undefined
      ? undefined
      : first(properties, drawingNamespace, "xfrm"),
    part,
    identity.name,
    "picture",
    errors,
  );
  const fill = first(picture, presentationNamespace, "blipFill");
  const blip =
    fill === undefined ? undefined : first(fill, drawingNamespace, "blip");
  const sourceRect =
    fill === undefined ? undefined : first(fill, drawingNamespace, "srcRect");
  const crop: NonNullable<FileObject["crop"]> = {};
  for (const [field, attributeName] of [
    ["left", "l"],
    ["top", "t"],
    ["right", "r"],
    ["bottom", "b"],
  ] as const) {
    if (sourceRect?.hasAttribute(attributeName) !== true) continue;
    const value = integerAttribute(sourceRect, attributeName);
    if (value === undefined || value < 0 || value > 100000) {
      issue(
        errors,
        part,
        "invalid-picture-crop",
        `图片 ${identity.name} 的裁切值无效`,
      );
      continue;
    }
    Object.assign(crop, { [field]: value });
  }
  if ((crop.left ?? 0) + (crop.right ?? 0) > 100000) {
    issue(
      errors,
      part,
      "invalid-picture-crop",
      `图片 ${identity.name} 左右裁切越界`,
    );
  }
  if ((crop.top ?? 0) + (crop.bottom ?? 0) > 100000) {
    issue(
      errors,
      part,
      "invalid-picture-crop",
      `图片 ${identity.name} 上下裁切越界`,
    );
  }
  const relationshipId =
    blip?.getAttributeNS(relationshipNamespace, "embed") ?? "";
  const relationship = relationships.get(part)?.get(relationshipId);
  let image: FileObject["image"];
  if (relationship === undefined) {
    issue(
      errors,
      part,
      "missing-image-relationship",
      `图片 ${identity.name} 引用的关系 ${relationshipId} 不存在`,
    );
  } else if (
    attribute(relationship, "Type") !== `${relationshipNamespace}/image` ||
    attribute(relationship, "TargetMode") === "External"
  ) {
    issue(
      errors,
      part,
      "invalid-image-relationship",
      `图片 ${identity.name} 必须引用包内图片部件`,
    );
  } else {
    const imagePart = relationshipTargetPath(
      part,
      attribute(relationship, "Target"),
    );
    const bytes = files[imagePart];
    const contentType = contentTypeForPart(imagePart, contentTypes);
    const expectedType = imagePart.endsWith(".png")
      ? "image/png"
      : imagePart.endsWith(".jpeg") || imagePart.endsWith(".jpg")
        ? "image/jpeg"
        : imagePart.endsWith(".gif")
          ? "image/gif"
          : undefined;
    if (
      !imagePart.startsWith("ppt/media/") ||
      bytes === undefined ||
      contentType === undefined ||
      contentType !== expectedType ||
      expectedType === undefined ||
      !imageSignatureMatches(bytes, contentType)
    ) {
      issue(
        errors,
        part,
        "invalid-image-part",
        `图片 ${identity.name} 的媒体部件缺失或内容类型无效`,
      );
    } else {
      image = { relationshipId, part: imagePart, contentType };
    }
  }
  if (boxEmu === undefined) return undefined;
  return {
    id: identity.id,
    name: identity.name,
    kind: "image",
    geometry: geometryName,
    boxEmu,
    ...(image === undefined ? {} : { image }),
    ...(Object.keys(crop).length === 0 ? {} : { crop }),
  };
}

function validateGroup(
  group: XmlElement,
  part: string,
  files: XmlFiles,
  relationships: Map<string, Map<string, XmlElement>>,
  contentTypes: {
    defaults: Map<string, string>;
    overrides: Map<string, string>;
  },
  errors: ValidationIssue[],
): FileObject | undefined {
  const identity = objectIdentity(group, "nvGrpSpPr", part, "group", errors);
  if (identity === undefined) return undefined;
  const properties = first(group, presentationNamespace, "grpSpPr");
  const transform =
    properties === undefined
      ? undefined
      : first(properties, drawingNamespace, "xfrm");
  const boxEmu = transformBox(transform, part, identity.name, "group", errors);
  const childOffset =
    transform === undefined
      ? undefined
      : first(transform, drawingNamespace, "chOff");
  const childExtent =
    transform === undefined
      ? undefined
      : first(transform, drawingNamespace, "chExt");
  const childX = integerAttribute(childOffset, "x");
  const childY = integerAttribute(childOffset, "y");
  const childWidth = integerAttribute(childExtent, "cx");
  const childHeight = integerAttribute(childExtent, "cy");
  if (
    childX === undefined ||
    childY === undefined ||
    childWidth === undefined ||
    childHeight === undefined ||
    childWidth < 1 ||
    childHeight < 1 ||
    childWidth > maximumShapeExtentEmu ||
    childHeight > maximumShapeExtentEmu
  ) {
    issue(
      errors,
      part,
      "invalid-group-transform",
      `组合 ${identity.name} 缺少有效的成员坐标范围`,
    );
  }
  const members = readObjects(
    group,
    part,
    files,
    relationships,
    contentTypes,
    errors,
  );
  if (members.length === 0) {
    issue(errors, part, "empty-group", `组合 ${identity.name} 没有原生成员`);
  }
  if (boxEmu === undefined) return undefined;
  return {
    id: identity.id,
    name: identity.name,
    kind: "group",
    boxEmu,
    ...(childX === undefined || childY === undefined
      ? {}
      : {
          groupTransform: {
            childOffset: { x: childX, y: childY },
            childExtent: {
              width: childWidth ?? 0,
              height: childHeight ?? 0,
            },
          },
        }),
    members,
  };
}

function readObjects(
  parent: XmlNode,
  part: string,
  files: XmlFiles,
  relationships: Map<string, Map<string, XmlElement>>,
  contentTypes: {
    defaults: Map<string, string>;
    overrides: Map<string, string>;
  },
  errors: ValidationIssue[],
): FileObject[] {
  const objects: FileObject[] = [];
  for (
    let child = parent.firstChild;
    child !== null;
    child = child.nextSibling
  ) {
    if (child.nodeType !== 1 || child.namespaceURI !== presentationNamespace) {
      continue;
    }
    let object: FileObject | undefined;
    switch (child.localName) {
      case "sp":
        object = validateShape(child as XmlElement, part, errors);
        break;
      case "pic":
        object = validatePicture(
          child as XmlElement,
          part,
          files,
          relationships,
          contentTypes,
          errors,
        );
        break;
      case "grpSp":
        object = validateGroup(
          child as XmlElement,
          part,
          files,
          relationships,
          contentTypes,
          errors,
        );
        break;
    }
    if (object !== undefined) objects.push(object);
  }
  return objects;
}

function readPresentation(
  files: XmlFiles,
  documents: Documents,
  relationships: Map<string, Map<string, XmlElement>>,
  contentTypes: {
    defaults: Map<string, string>;
    overrides: Map<string, string>;
  },
  errors: ValidationIssue[],
): { presentation: ValidationResult["presentation"]; slides: FileSlide[] } {
  const part = "ppt/presentation.xml";
  const document = documents.get(part);
  const root = document?.documentElement;
  if (
    root == null ||
    root.localName !== "presentation" ||
    root.namespaceURI !== presentationNamespace
  ) {
    issue(
      errors,
      part,
      "invalid-presentation",
      "缺少 Transitional PresentationML 根部件",
    );
    return { presentation: { slideCount: 0 }, slides: [] };
  }
  const size = first(root, presentationNamespace, "sldSz");
  const widthEmu = integerAttribute(size, "cx");
  const heightEmu = integerAttribute(size, "cy");
  if (
    widthEmu === undefined ||
    heightEmu === undefined ||
    widthEmu < minimumSlideSizeEmu ||
    widthEmu > maximumSlideSizeEmu ||
    heightEmu < minimumSlideSizeEmu ||
    heightEmu > maximumSlideSizeEmu
  ) {
    issue(
      errors,
      part,
      "invalid-slide-size",
      `页面尺寸必须在 ${minimumSlideSizeEmu} 到 ${maximumSlideSizeEmu} EMU 的闭区间内`,
    );
  }
  const ids = first(root, presentationNamespace, "sldIdLst");
  const slideIds =
    ids === undefined ? [] : elements(ids, presentationNamespace, "sldId");
  const presentationRelationships = relationships.get(part) ?? new Map();
  const slides: FileSlide[] = [];
  slideIds.forEach((slideId, index) => {
    const relationshipId =
      slideId.getAttributeNS(relationshipNamespace, "id") ?? "";
    const relationship = presentationRelationships.get(relationshipId);
    if (relationship === undefined) {
      issue(
        errors,
        part,
        "missing-slide-relationship",
        `页面 ${index + 1} 缺少对应关系`,
      );
      return;
    }
    const slidePath = relationshipTargetPath(
      part,
      attribute(relationship, "Target"),
    );
    const slideDocument = documents.get(slidePath);
    const slideRoot = slideDocument?.documentElement;
    if (
      slideRoot == null ||
      slideRoot.localName !== "sld" ||
      slideRoot.namespaceURI !== presentationNamespace
    ) {
      issue(
        errors,
        slidePath,
        "invalid-slide",
        "页面关系目标不是有效的 Transitional 页面部件",
      );
      return;
    }
    const commonSlideData = first(slideRoot, presentationNamespace, "cSld");
    const slideTree =
      commonSlideData === undefined
        ? undefined
        : first(commonSlideData, presentationNamespace, "spTree");
    if (slideTree === undefined) {
      issue(errors, slidePath, "missing-shape-tree", "页面缺少 shape tree");
      return;
    }
    const objects = readObjects(
      slideTree,
      slidePath,
      files,
      relationships,
      contentTypes,
      errors,
    );
    const seenIds = new Set<number>();
    const pending = [...objects];
    while (pending.length > 0) {
      const object = pending.pop();
      if (object === undefined) continue;
      if (seenIds.has(object.id)) {
        issue(
          errors,
          slidePath,
          "duplicate-object-id",
          `对象标识 ${object.id} 重复`,
        );
      }
      seenIds.add(object.id);
      pending.push(...(object.members ?? []));
    }
    slides.push({
      id:
        commonSlideData === undefined ? "" : attribute(commonSlideData, "name"),
      objects,
    });
  });
  if (slideIds.length === 0)
    issue(errors, part, "missing-slides", "Presentation 不含页面");
  if (files["ppt/theme/theme1.xml"] === undefined) {
    issue(
      errors,
      "ppt/theme/theme1.xml",
      "missing-theme",
      "缺少必要的主题部件",
    );
  }
  return {
    presentation: {
      slideCount: slideIds.length,
      ...(widthEmu === undefined ? {} : { widthEmu }),
      ...(heightEmu === undefined ? {} : { heightEmu }),
    },
    slides,
  };
}

/**
 * 对同一份 PPTX 字节执行 SDK 结构校验并独立读取页面对象、几何和原生文字，
 * 不读取构建报告。SDK 的 MC 分支选择和扩展覆盖有限；此结果不代表应用证据。
 */
export async function validatePptx(
  input: Uint8Array,
): Promise<ValidationResult> {
  const errors = await validateStandardStructure(input);
  let files: XmlFiles;
  try {
    files = unzipSync(input);
  } catch (error) {
    issue(
      errors,
      "",
      "invalid-package",
      error instanceof Error ? error.message : String(error),
    );
    return {
      profile: {
        standard: transitionalStandard,
        schemaValidator: schemaValidatorProfile,
        scope: validationScope,
        applicationEvidence,
      },
      errors,
      presentation: { slideCount: 0 },
      slides: [],
    };
  }
  const documents = parseXmlFiles(files, errors);
  for (const requiredPart of [
    "[Content_Types].xml",
    "_rels/.rels",
    "ppt/presentation.xml",
    "ppt/_rels/presentation.xml.rels",
    "ppt/slideMasters/slideMaster1.xml",
    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    "ppt/slideLayouts/slideLayout1.xml",
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    "ppt/theme/theme1.xml",
  ]) {
    if (files[requiredPart] === undefined)
      issue(errors, requiredPart, "missing-part", "PPTX 缺少必要部件");
  }
  const relationships = checkRelationships(files, documents, errors);
  const contentTypes = validateContentTypes(files, documents, errors);
  requireRoot(
    documents,
    errors,
    "ppt/slideMasters/slideMaster1.xml",
    presentationNamespace,
    "sldMaster",
  );
  validateSlideLayoutIds(documents, errors);
  requireRoot(
    documents,
    errors,
    "ppt/slideLayouts/slideLayout1.xml",
    presentationNamespace,
    "sldLayout",
  );
  requireRoot(
    documents,
    errors,
    "ppt/theme/theme1.xml",
    drawingNamespace,
    "theme",
  );
  requireRelationship(
    "",
    `${relationshipNamespace}/officeDocument`,
    "ppt/presentation.xml",
    relationships,
    errors,
  );
  requireRelationship(
    "ppt/presentation.xml",
    `${relationshipNamespace}/slideMaster`,
    "ppt/slideMasters/slideMaster1.xml",
    relationships,
    errors,
  );
  requireRelationship(
    "ppt/slideMasters/slideMaster1.xml",
    `${relationshipNamespace}/slideLayout`,
    "ppt/slideLayouts/slideLayout1.xml",
    relationships,
    errors,
  );
  requireRelationship(
    "ppt/slideMasters/slideMaster1.xml",
    `${relationshipNamespace}/theme`,
    "ppt/theme/theme1.xml",
    relationships,
    errors,
  );
  requireRelationship(
    "ppt/slideLayouts/slideLayout1.xml",
    `${relationshipNamespace}/slideMaster`,
    "ppt/slideMasters/slideMaster1.xml",
    relationships,
    errors,
  );
  const { presentation, slides } = readPresentation(
    files,
    documents,
    relationships,
    contentTypes,
    errors,
  );
  for (const [index] of slides.entries()) {
    requireRelationship(
      `ppt/slides/slide${index + 1}.xml`,
      `${relationshipNamespace}/slideLayout`,
      "ppt/slideLayouts/slideLayout1.xml",
      relationships,
      errors,
    );
  }
  return {
    profile: {
      standard: transitionalStandard,
      schemaValidator: schemaValidatorProfile,
      scope: validationScope,
      applicationEvidence,
    },
    errors,
    presentation,
    slides,
  };
}
