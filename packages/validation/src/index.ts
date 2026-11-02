import { posix } from "node:path";

import validateOpenXml from "@ooxml-tools/validate";
import { DOMParser } from "@xmldom/xmldom";
import { strFromU8, unzipSync } from "fflate";

const presentationNamespace =
  "http://schemas.openxmlformats.org/presentationml/2006/main";
const drawingNamespace =
  "http://schemas.openxmlformats.org/drawingml/2006/main";
const tableUri = "http://schemas.openxmlformats.org/drawingml/2006/table";
const relationshipNamespace =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const packageRelationshipNamespace =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const contentTypeNamespace =
  "http://schemas.openxmlformats.org/package/2006/content-types";
const p14Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const p159Namespace =
  "http://schemas.microsoft.com/office/powerpoint/2015/09/main";
const markupCompatibilityNamespace =
  "http://schemas.openxmlformats.org/markup-compatibility/2006";
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
  "Open XML SDK 3.0.2 在 Microsoft365 配置下实现的 Transitional 包、Schema 与约束校验；另由本地读取器独立核对 OPC 关系、内容类型及页面、形状、图片、组合、文字与转场结构，并逐支核对转场 Choice／Fallback 的 Requires 和目标扩展限制。SDK 的 MC 仅按其支持选择 Choice 或 Fallback；p14／p159 的所有兼容分支及未知扩展不保证均被 Schema 验证。校验不等于完整标准认证或应用证据。";
const applicationEvidence = "not-validated";
const minimumSlideLayoutId = 2_147_483_648;
const maximumSlideLayoutId = 4_294_967_295;
const minimumSlideSizeEmu = 914_400;
const maximumSlideSizeEmu = 51_206_400;
const maximumShapeExtentEmu = 2_147_483_647;
const minimumFontSizeCentipoints = 100;
const maximumFontSizeCentipoints = 400_000;
const maximumTextSpacingCentipoints = 158_400;
// 外部 SDK 用共享文件编号关联输入；串行调用可避免并发覆盖，失败后队列仍会释放。
let schemaValidationQueue: Promise<void> = Promise.resolve();

export type ValidationIssue = {
  readonly part: string;
  readonly code: string;
  readonly message: string;
  readonly xpath?: string;
};

export type FileObject = {
  readonly id: number;
  readonly name: string;
  readonly kind: "shape" | "image" | "group" | "table";
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
  readonly table?: {
    readonly columnWidthsEmu: readonly number[];
    readonly rowHeightsEmu: readonly number[];
    readonly cells: readonly (readonly string[])[];
    readonly merges: readonly {
      readonly row: number;
      readonly column: number;
      readonly rowSpan: number;
      readonly colSpan: number;
    }[];
    readonly cellTexts: readonly {
      readonly row: number;
      readonly column: number;
      readonly text: FileText;
    }[];
  };
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
  readonly transition?: FileTransition;
  readonly animations?: readonly FileAnimationStep[];
};

export type FileAnimationStep = {
  readonly trigger: "auto" | "click";
  readonly effects: readonly FileAnimationEffect[];
};

export type FileAnimationEffect = {
  readonly targetId: number;
  readonly kind: "entrance" | "emphasis";
  readonly preset:
    | "fade"
    | "wipe"
    | "zoom"
    | "fly"
    | "floatUp"
    | "spinIn"
    | "stretch"
    | "spin";
  readonly durationMs: number;
  readonly delayMs: number;
  readonly from?: "left" | "right" | "top" | "bottom" | number;
  readonly easing?: {
    readonly accelerate?: number;
    readonly decelerate?: number;
  };
};

export type FileTransition = {
  readonly kind: "fade" | "strips" | "push" | "flip" | "morph";
  readonly speed?: "fast" | "med" | "slow";
  readonly direction?: string;
  readonly durationMs?: number;
  readonly by?: "object" | "word" | "character";
  readonly fallback?: FileTransition;
};

export type ValidationExtension = "p14" | "p159";

export type ValidationOptions = {
  /** 拒绝实际文件使用的目标未允许扩展；省略时只做结构读取。 */
  readonly allowedExtensions?: readonly ValidationExtension[];
};

export type ValidationResult = {
  readonly profile: {
    readonly standard: typeof transitionalStandard;
    readonly schemaValidator: typeof schemaValidatorProfile;
    readonly scope: string;
    readonly applicationEvidence: typeof applicationEvidence;
  };
  readonly errors: readonly ValidationIssue[];
  readonly extensions: readonly ValidationExtension[];
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

type TransitionRead = {
  readonly transition?: FileTransition;
  readonly extensions: ReadonlySet<ValidationExtension>;
};

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
  const validation = schemaValidationQueue.then(async () => {
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
  });
  schemaValidationQueue = validation.then(
    () => undefined,
    () => undefined,
  );

  try {
    return await validation;
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
  validateTextBodyValueRanges(body, part, errors);
}

function validateTextBodyValueRanges(
  body: XmlElement,
  part: string,
  errors: ValidationIssue[],
): void {
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

function tableSpan(
  cell: XmlElement,
  name: "gridSpan" | "rowSpan",
  part: string,
  errors: ValidationIssue[],
): number {
  if (!cell.hasAttribute(name)) return 1;
  const value = integerAttribute(cell, name);
  if (value === undefined || value < 1) {
    issue(errors, part, "invalid-table-merge", `${name} 必须是正整数`);
    return 1;
  }
  return value;
}

function tableMergeFlag(
  cell: XmlElement,
  name: "hMerge" | "vMerge",
  part: string,
  errors: ValidationIssue[],
): boolean {
  const value = attribute(cell, name);
  if (value === "" || value === "0" || value === "false") return false;
  if (value === "1" || value === "true") return true;
  issue(errors, part, "invalid-table-merge", `${name} 必须是布尔值`);
  return false;
}

function tableCellText(
  cell: XmlElement,
  part: string,
  errors: ValidationIssue[],
): FileText {
  const body = first(cell, drawingNamespace, "txBody");
  if (body === undefined) {
    issue(errors, part, "missing-table-cell-text", "表格单元格缺少 a:txBody");
    return { text: "", paragraphs: [] };
  }
  validateTextBodyValueRanges(body, part, errors);
  const paragraphs = elements(body, drawingNamespace, "p").map(readParagraph);
  return { text: paragraphs.map(({ text }) => text).join("\n"), paragraphs };
}

/** 独立读取表格网格与合并覆盖，不调用生产表格模块的计算结果。 */
function validateTable(
  frame: XmlElement,
  part: string,
  errors: ValidationIssue[],
): FileObject | undefined {
  const identity = objectIdentity(
    frame,
    "nvGraphicFramePr",
    part,
    "table",
    errors,
  );
  if (identity === undefined) return undefined;
  const transform = first(frame, presentationNamespace, "xfrm");
  const boxEmu = transformBox(transform, part, identity.name, "table", errors);
  const graphic = first(frame, drawingNamespace, "graphic");
  const graphicData =
    graphic === undefined
      ? undefined
      : first(graphic, drawingNamespace, "graphicData");
  const table =
    graphicData === undefined
      ? undefined
      : first(graphicData, drawingNamespace, "tbl");
  if (
    graphicData === undefined ||
    attribute(graphicData, "uri") !== tableUri ||
    table === undefined
  ) {
    issue(
      errors,
      part,
      "invalid-table-graphic",
      `表格 ${identity.name} 缺少 DrawingML 表格内容`,
    );
    return undefined;
  }
  const grid = first(table, drawingNamespace, "tblGrid");
  const columns =
    grid === undefined ? [] : elements(grid, drawingNamespace, "gridCol");
  if (columns.length === 0) {
    issue(
      errors,
      part,
      "invalid-table-grid",
      `表格 ${identity.name} 缺少列网格`,
    );
  }
  const columnWidthsEmu = columns.map((column, index) => {
    const width = integerAttribute(column, "w");
    if (width === undefined || width < 1 || width > maximumShapeExtentEmu) {
      issue(
        errors,
        part,
        "invalid-table-grid",
        `表格 ${identity.name} 第 ${index + 1} 列宽无效`,
      );
      return 0;
    }
    return width;
  });
  const rowNodes = elements(table, drawingNamespace, "tr");
  if (rowNodes.length === 0) {
    issue(
      errors,
      part,
      "invalid-table-grid",
      `表格 ${identity.name} 至少需要一行`,
    );
  }
  const rowHeightsEmu: number[] = [];
  const rowCells: XmlElement[][] = [];
  for (const [rowIndex, row] of rowNodes.entries()) {
    const height = integerAttribute(row, "h");
    if (height === undefined || height < 1 || height > maximumShapeExtentEmu) {
      issue(
        errors,
        part,
        "invalid-table-grid",
        `表格 ${identity.name} 第 ${rowIndex + 1} 行高无效`,
      );
      rowHeightsEmu.push(0);
    } else rowHeightsEmu.push(height);
    const cells = elements(row, drawingNamespace, "tc");
    if (cells.length !== columns.length) {
      issue(
        errors,
        part,
        "invalid-table-grid",
        `表格 ${identity.name} 第 ${rowIndex + 1} 行含 ${cells.length} 个格子，列网格含 ${columns.length} 列`,
      );
    }
    rowCells.push(cells);
  }

  type Owner = {
    readonly row: number;
    readonly column: number;
    readonly rowSpan: number;
    readonly colSpan: number;
  };
  const owners = new Map<string, Owner>();
  const continuations = new Map<
    XmlElement,
    { readonly horizontal: boolean; readonly vertical: boolean }
  >();
  const merges: {
    row: number;
    column: number;
    rowSpan: number;
    colSpan: number;
  }[] = [];
  const cells: string[][] = rowNodes.map(() =>
    Array.from({ length: columns.length }, () => ""),
  );
  const cellTexts: { row: number; column: number; text: FileText }[] = [];
  for (const [rowIndex, currentRow] of rowCells.entries()) {
    for (
      let column = 0;
      column < Math.min(currentRow.length, columns.length);
      column += 1
    ) {
      const cell = currentRow[column]!;
      const horizontalContinuation = tableMergeFlag(
        cell,
        "hMerge",
        part,
        errors,
      );
      const verticalContinuation = tableMergeFlag(cell, "vMerge", part, errors);
      continuations.set(cell, {
        horizontal: horizontalContinuation,
        vertical: verticalContinuation,
      });
      if (horizontalContinuation || verticalContinuation) continue;
      const colSpan = tableSpan(cell, "gridSpan", part, errors);
      const rowSpan = tableSpan(cell, "rowSpan", part, errors);
      if (
        column + colSpan > columns.length ||
        rowIndex + rowSpan > rowNodes.length
      ) {
        issue(
          errors,
          part,
          "invalid-table-merge",
          `表格 ${identity.name} 的合并范围越界`,
        );
      }
      const owner = { row: rowIndex, column, rowSpan, colSpan };
      for (
        let row = rowIndex;
        row < Math.min(rowNodes.length, rowIndex + rowSpan);
        row += 1
      ) {
        for (
          let col = column;
          col < Math.min(columns.length, column + colSpan);
          col += 1
        ) {
          const key = `${row}:${col}`;
          if (owners.has(key))
            issue(
              errors,
              part,
              "conflicting-table-merge",
              `表格 ${identity.name} 的合并范围重叠`,
            );
          else owners.set(key, owner);
        }
      }
      if (rowSpan > 1 || colSpan > 1) merges.push(owner);
      const text = tableCellText(cell, part, errors);
      cells[rowIndex]![column] = text.text;
      cellTexts.push({ row: rowIndex, column, text });
    }
  }
  for (const [rowIndex, currentRow] of rowCells.entries()) {
    for (let column = 0; column < columns.length; column += 1) {
      const cell = currentRow[column];
      const owner = owners.get(`${rowIndex}:${column}`);
      if (owner === undefined || cell === undefined) {
        issue(
          errors,
          part,
          "invalid-table-grid",
          `表格 ${identity.name} 的第 ${rowIndex + 1} 行第 ${column + 1} 列没有有效单元格`,
        );
        continue;
      }
      const isAnchor = rowIndex === owner.row && column === owner.column;
      const continuation = continuations.get(cell);
      const horizontalContinuation = continuation?.horizontal ?? false;
      const verticalContinuation = continuation?.vertical ?? false;
      if (
        horizontalContinuation !== (!isAnchor && column > owner.column) ||
        verticalContinuation !== (!isAnchor && rowIndex > owner.row)
      ) {
        issue(
          errors,
          part,
          "invalid-table-merge",
          `表格 ${identity.name} 的合并续格标记与锚点范围不一致`,
        );
      }
    }
  }
  const widthEmu = columnWidthsEmu.reduce((sum, width) => sum + width, 0);
  const heightEmu = rowHeightsEmu.reduce((sum, height) => sum + height, 0);
  if (
    boxEmu !== undefined &&
    (boxEmu.width !== widthEmu || boxEmu.height !== heightEmu)
  ) {
    issue(
      errors,
      part,
      "invalid-table-transform",
      `表格 ${identity.name} 的框架尺寸与行列网格合计不一致`,
    );
  }
  if (boxEmu === undefined) return undefined;
  return {
    id: identity.id,
    name: identity.name,
    kind: "table",
    boxEmu,
    table: {
      columnWidthsEmu,
      rowHeightsEmu,
      cells,
      merges,
      cellTexts,
    },
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
      case "graphicFrame":
        object = validateTable(child as XmlElement, part, errors);
        break;
    }
    if (object !== undefined) objects.push(object);
  }
  return objects;
}

function childElements(parent: XmlNode): XmlElement[] {
  const result: XmlElement[] = [];
  for (
    let child = parent.firstChild;
    child !== null;
    child = child.nextSibling
  ) {
    if (child.nodeType === 1) result.push(child as XmlElement);
  }
  return result;
}

function extensionForNamespace(
  namespace: string | null,
): ValidationExtension | undefined {
  if (namespace === p14Namespace) return "p14";
  if (namespace === p159Namespace) return "p159";
  return undefined;
}

function transitionDuration(value: string): number | undefined {
  const match = /^(\d+(?:\.\d)?)(h|min|s|ms|µs|ns)?$/u.exec(value);
  if (match === null) return undefined;
  const amount = Number(match[1]);
  const multiplier =
    match[2] === "h"
      ? 3_600_000
      : match[2] === "min"
        ? 60_000
        : match[2] === "s"
          ? 1_000
          : match[2] === "µs"
            ? 0.001
            : match[2] === "ns"
              ? 0.000001
              : 1;
  const milliseconds = amount * multiplier;
  return Number.isFinite(milliseconds) ? milliseconds : undefined;
}

function readTransitionElement(
  transition: XmlElement,
  part: string,
  errors: ValidationIssue[],
  insideChoice: boolean,
): {
  readonly transition: FileTransition;
  readonly extensions: ReadonlySet<ValidationExtension>;
} {
  const extensions = new Set<ValidationExtension>();
  const speedValue = attribute(transition, "spd");
  const speed =
    speedValue === "fast" || speedValue === "med" || speedValue === "slow"
      ? speedValue
      : undefined;
  if (speedValue !== "" && speed === undefined) {
    issue(
      errors,
      part,
      "invalid-transition-speed",
      `未知转场速度 ${speedValue}`,
    );
  }
  const durationValue = transition.getAttributeNS(p14Namespace, "dur");
  let durationMs: number | undefined;
  if (transition.hasAttributeNS(p14Namespace, "dur")) {
    extensions.add("p14");
    durationMs = transitionDuration(durationValue ?? "");
    if (durationMs === undefined) {
      issue(
        errors,
        part,
        "invalid-transition-duration",
        "p14:dur 不是有效时长",
      );
    }
  }
  const effects = childElements(transition);
  if (effects.length !== 1) {
    issue(
      errors,
      part,
      "invalid-transition-effect",
      "转场必须恰好包含一个效果元素",
    );
  }
  const effect = effects[0];
  let kind: FileTransition["kind"] = "fade";
  let direction: string | undefined;
  let by: FileTransition["by"];
  if (effect?.namespaceURI === presentationNamespace) {
    if (effect.localName === "fade") kind = "fade";
    else if (effect.localName === "strips") {
      kind = "strips";
      direction = attribute(effect, "dir");
      if (!["lu", "ru", "ld", "rd"].includes(direction)) {
        issue(errors, part, "invalid-transition-direction", "strips.dir 无效");
      }
    } else if (effect.localName === "push") {
      kind = "push";
      direction = attribute(effect, "dir");
      if (!["l", "r", "u", "d"].includes(direction)) {
        issue(errors, part, "invalid-transition-direction", "push.dir 无效");
      }
    } else {
      issue(
        errors,
        part,
        "unsupported-transition-effect",
        `不支持转场效果 p:${effect.localName}`,
      );
    }
  } else if (
    effect?.namespaceURI === p14Namespace &&
    effect.localName === "flip"
  ) {
    kind = "flip";
    extensions.add("p14");
    direction = attribute(effect, "dir");
    if (direction !== "l" && direction !== "r") {
      issue(
        errors,
        part,
        "invalid-transition-direction",
        "flip.dir 必须是 l 或 r",
      );
    }
  } else if (
    effect?.namespaceURI === p159Namespace &&
    effect.localName === "morph"
  ) {
    kind = "morph";
    extensions.add("p159");
    const option = attribute(effect, "option");
    if (option === "byObject") by = "object";
    else if (option === "byWord") by = "word";
    else if (option === "byChar") by = "character";
    else
      issue(errors, part, "invalid-morph-option", `未知 Morph 选项 ${option}`);
  } else if (effect !== undefined) {
    issue(
      errors,
      part,
      "unsupported-transition-effect",
      "转场效果使用未知命名空间",
    );
  }
  if (!insideChoice && extensions.size > 0) {
    issue(
      errors,
      part,
      "extension-outside-choice",
      "p14／p159 转场机制必须放在 mc:Choice 内",
    );
  }
  const result: FileTransition = {
    kind,
    ...(speed === undefined ? {} : { speed }),
    ...(direction === undefined ? {} : { direction }),
    ...(durationMs === undefined ? {} : { durationMs }),
    ...(by === undefined ? {} : { by }),
  };
  return { transition: result, extensions };
}

function readAlternateTransition(
  alternate: XmlElement,
  part: string,
  errors: ValidationIssue[],
): TransitionRead {
  const choices = elements(alternate, markupCompatibilityNamespace, "Choice");
  const fallbacks = elements(
    alternate,
    markupCompatibilityNamespace,
    "Fallback",
  );
  if (choices.length !== 1 || fallbacks.length !== 1) {
    issue(
      errors,
      part,
      "invalid-transition-alternate",
      "转场兼容分支必须包含一个 Choice 和一个 Fallback",
    );
  }
  const choice = choices[0];
  const fallback = fallbacks[0];
  const choiceTransition =
    choice === undefined
      ? undefined
      : elements(choice, presentationNamespace, "transition")[0];
  const fallbackTransition =
    fallback === undefined
      ? undefined
      : elements(fallback, presentationNamespace, "transition")[0];
  if (
    choice === undefined ||
    fallback === undefined ||
    choiceTransition === undefined ||
    fallbackTransition === undefined
  ) {
    issue(
      errors,
      part,
      "invalid-transition-alternate",
      "转场兼容分支的 Choice 和 Fallback 都必须包含 p:transition",
    );
    return { extensions: new Set() };
  }
  const choiceResult = readTransitionElement(
    choiceTransition,
    part,
    errors,
    true,
  );
  const fallbackResult = readTransitionElement(
    fallbackTransition,
    part,
    errors,
    false,
  );
  if (
    fallbackResult.transition.kind !== "fade" &&
    fallbackResult.transition.kind !== "strips" &&
    fallbackResult.transition.kind !== "push"
  ) {
    issue(
      errors,
      part,
      "invalid-transition-fallback",
      "转场 Fallback 必须使用 Transitional 基础效果",
    );
  }
  const required = new Set<ValidationExtension>();
  const requiredPrefixes = attribute(choice, "Requires")
    .trim()
    .split(/\s+/u)
    .filter((prefix) => prefix !== "");
  for (const prefix of requiredPrefixes) {
    const extension = extensionForNamespace(choice.lookupNamespaceURI(prefix));
    if (extension === undefined) {
      issue(
        errors,
        part,
        "unsupported-extension",
        `转场 Choice 要求不支持的命名空间 ${prefix}`,
      );
    } else {
      required.add(extension);
    }
  }
  for (const extension of choiceResult.extensions) {
    if (!required.has(extension)) {
      issue(
        errors,
        part,
        "missing-extension-requirement",
        `转场 Choice 使用 ${extension}，但 Requires 未声明该机制`,
      );
    }
  }
  for (const extension of required) {
    if (!choiceResult.extensions.has(extension)) {
      issue(
        errors,
        part,
        "unused-extension-requirement",
        `转场 Choice 声明了未使用的 ${extension} 机制`,
      );
    }
  }
  const extensions = new Set<ValidationExtension>([
    ...choiceResult.extensions,
    ...required,
    ...fallbackResult.extensions,
  ]);
  return {
    transition: {
      ...choiceResult.transition,
      fallback: fallbackResult.transition,
    },
    extensions,
  };
}

function readSlideTransition(
  slide: XmlElement,
  part: string,
  errors: ValidationIssue[],
): TransitionRead {
  const candidates = childElements(slide).filter(
    (child) =>
      (child.namespaceURI === presentationNamespace &&
        child.localName === "transition") ||
      (child.namespaceURI === markupCompatibilityNamespace &&
        child.localName === "AlternateContent"),
  );
  if (candidates.length > 1) {
    issue(errors, part, "duplicate-transition", "页面包含多个转场声明");
  }
  const candidate = candidates[0];
  if (candidate === undefined) return { extensions: new Set() };
  if (candidate.namespaceURI === markupCompatibilityNamespace) {
    return readAlternateTransition(candidate, part, errors);
  }
  const result = readTransitionElement(candidate, part, errors, false);
  return { transition: result.transition, extensions: result.extensions };
}

function descendantsByName(
  parent: XmlNode,
  namespace: string,
  localName: string,
): XmlElement[] {
  const result: XmlElement[] = [];
  const pending = [parent];
  while (pending.length > 0) {
    const current = pending.pop();
    if (current === undefined) continue;
    for (
      let child = current.firstChild;
      child !== null;
      child = child.nextSibling
    ) {
      if (child.nodeType !== 1) continue;
      const element = child as XmlElement;
      if (
        element.namespaceURI === namespace &&
        element.localName === localName
      ) {
        result.push(element);
      }
      pending.push(element);
    }
  }
  return result;
}

function positiveIntegerValue(value: string): number | undefined {
  if (!/^\d+$/u.test(value)) return undefined;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed > 0 ? parsed : undefined;
}

function nonNegativeIntegerValue(value: string): number | undefined {
  if (!/^\d+$/u.test(value)) return undefined;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed >= 0 ? parsed : undefined;
}

function collectObjectIds(objects: readonly FileObject[]): Set<number> {
  const result = new Set<number>();
  const pending = [...objects];
  while (pending.length > 0) {
    const object = pending.pop();
    if (object === undefined) continue;
    result.add(object.id);
    pending.push(...(object.members ?? []));
  }
  return result;
}

function readSlideAnimations(
  slide: XmlElement,
  part: string,
  objects: readonly FileObject[],
  errors: ValidationIssue[],
): readonly FileAnimationStep[] | undefined {
  const timingNodes = elements(slide, presentationNamespace, "timing");
  if (timingNodes.length === 0) return undefined;
  if (timingNodes.length > 1) {
    issue(errors, part, "duplicate-timing", "页面包含多个动画时序声明");
  }
  const timing = timingNodes[0];
  if (timing === undefined) return undefined;

  const timeIds = new Set<number>();
  for (const timeNode of descendantsByName(
    timing,
    presentationNamespace,
    "cTn",
  )) {
    const id = positiveIntegerValue(attribute(timeNode, "id"));
    if (id === undefined) {
      issue(errors, part, "invalid-animation-id", "动画 cTn 缺少正整数 id");
      continue;
    }
    if (timeIds.has(id)) {
      issue(
        errors,
        part,
        "duplicate-animation-id",
        `动画时间节点 id ${id} 重复`,
      );
    }
    timeIds.add(id);
  }

  const knownObjects = collectObjectIds(objects);
  const mainSequence = descendantsByName(
    timing,
    presentationNamespace,
    "cTn",
  ).find((timeNode) => attribute(timeNode, "nodeType") === "mainSeq");
  const sequenceChildren =
    mainSequence === undefined
      ? undefined
      : first(mainSequence, presentationNamespace, "childTnLst");
  if (sequenceChildren === undefined) {
    issue(
      errors,
      part,
      "invalid-animation-tree",
      "动画时序缺少 mainSeq 点击步列表",
    );
    return [];
  }

  const result: FileAnimationStep[] = [];
  for (const stepNode of elements(
    sequenceChildren,
    presentationNamespace,
    "par",
  )) {
    const stepTimeNode = first(stepNode, presentationNamespace, "cTn");
    if (stepTimeNode === undefined) {
      issue(errors, part, "invalid-animation-step", "点击步缺少 cTn");
      continue;
    }
    const stepConditions = first(
      stepTimeNode,
      presentationNamespace,
      "stCondLst",
    );
    const isAutomatic =
      stepConditions !== undefined &&
      descendantsByName(stepConditions, presentationNamespace, "cond").some(
        (condition) => attribute(condition, "evt") === "onBegin",
      );
    const trigger: FileAnimationStep["trigger"] = isAutomatic
      ? "auto"
      : "click";
    const stepChildren = first(
      stepTimeNode,
      presentationNamespace,
      "childTnLst",
    );
    const effects: FileAnimationEffect[] = [];
    for (const group of elements(
      stepChildren ?? stepTimeNode,
      presentationNamespace,
      "par",
    )) {
      const groupTimeNode = first(group, presentationNamespace, "cTn");
      if (groupTimeNode === undefined) continue;
      const groupCondition = first(
        first(groupTimeNode, presentationNamespace, "stCondLst") ??
          groupTimeNode,
        presentationNamespace,
        "cond",
      );
      const groupDelayText =
        groupCondition === undefined ? "0" : attribute(groupCondition, "delay");
      const groupDelay = nonNegativeIntegerValue(groupDelayText || "0");
      if (groupDelay === undefined) {
        issue(
          errors,
          part,
          "invalid-animation-delay",
          "动画组延迟必须是非负整数毫秒",
        );
        continue;
      }
      const groupChildren = first(
        groupTimeNode,
        presentationNamespace,
        "childTnLst",
      );
      for (const effectNode of elements(
        groupChildren ?? groupTimeNode,
        presentationNamespace,
        "par",
      )) {
        const effectTimeNode = first(effectNode, presentationNamespace, "cTn");
        if (effectTimeNode === undefined) continue;
        const presetId = integerAttribute(effectTimeNode, "presetID");
        if (presetId === undefined) continue;
        const presetClass = attribute(effectTimeNode, "presetClass");
        const presetById: Readonly<
          Record<number, FileAnimationEffect["preset"]>
        > = {
          2: "fly",
          8: "spin",
          10: "fade",
          17: "stretch",
          22: "wipe",
          42: "floatUp",
          49: "spinIn",
          53: "zoom",
        };
        const preset = presetById[presetId];
        const kind =
          presetClass === "entr"
            ? "entrance"
            : presetClass === "emph"
              ? "emphasis"
              : undefined;
        if (preset === undefined || kind === undefined) {
          issue(
            errors,
            part,
            "unsupported-animation",
            `不支持动画预设 ${presetClass}#${presetId}`,
          );
          continue;
        }
        const expectedKind = preset === "spin" ? "emphasis" : "entrance";
        if (kind !== expectedKind) {
          issue(
            errors,
            part,
            "invalid-animation-kind",
            `${preset} 的 presetClass 必须是 ${expectedKind === "emphasis" ? "emph" : "entr"}`,
          );
          continue;
        }
        const effectConditionList = first(
          effectTimeNode,
          presentationNamespace,
          "stCondLst",
        );
        const effectCondition =
          effectConditionList === undefined
            ? undefined
            : first(effectConditionList, presentationNamespace, "cond");
        const effectDelay = nonNegativeIntegerValue(
          effectCondition === undefined
            ? "0"
            : attribute(effectCondition, "delay") || "0",
        );
        if (effectDelay === undefined) {
          issue(
            errors,
            part,
            "invalid-animation-delay",
            "动画效果延迟必须是非负整数毫秒",
          );
          continue;
        }
        const behaviorNodes = descendantsByName(
          effectTimeNode,
          presentationNamespace,
          "cBhvr",
        );
        const durations = behaviorNodes
          .map((behavior) =>
            integerAttribute(
              first(behavior, presentationNamespace, "cTn"),
              "dur",
            ),
          )
          .filter((duration): duration is number => duration !== undefined);
        const durationMs =
          durations.length === 0 ? undefined : Math.max(...durations);
        if (durationMs === undefined || durationMs <= 0) {
          issue(
            errors,
            part,
            "invalid-animation-duration",
            "动画行为缺少正整数时长",
          );
          continue;
        }
        const targetNodes = descendantsByName(
          effectTimeNode,
          presentationNamespace,
          "spTgt",
        );
        const targetIds = targetNodes.map((target) =>
          positiveIntegerValue(attribute(target, "spid")),
        );
        const validTargetIds = targetIds.filter(
          (targetId): targetId is number => targetId !== undefined,
        );
        if (
          targetNodes.length === 0 ||
          validTargetIds.length !== targetNodes.length ||
          new Set(validTargetIds).size !== 1
        ) {
          issue(
            errors,
            part,
            "invalid-animation-target",
            "一个动画效果必须指向唯一原生对象",
          );
          continue;
        }
        const targetId = validTargetIds[0];
        if (targetId === undefined || !knownObjects.has(targetId)) {
          issue(
            errors,
            part,
            "missing-animation-target",
            `动画目标 ${String(targetId)} 不存在`,
          );
          continue;
        }
        const easing: { accelerate?: number; decelerate?: number } = {};
        const accelerate = integerAttribute(effectTimeNode, "accel");
        const decelerate = integerAttribute(effectTimeNode, "decel");
        if (accelerate !== undefined && accelerate !== 0) {
          easing.accelerate = accelerate / 100000;
        }
        if (decelerate !== undefined && decelerate !== 0) {
          easing.decelerate = decelerate / 100000;
        }
        const directionSubtype = integerAttribute(
          effectTimeNode,
          "presetSubtype",
        );
        const direction =
          preset === "wipe" || preset === "fly"
            ? ({ 1: "top", 2: "right", 4: "bottom", 8: "left" } as const)[
                directionSubtype ?? -1
              ]
            : undefined;
        if (
          (preset === "wipe" || preset === "fly") &&
          direction === undefined
        ) {
          issue(
            errors,
            part,
            "invalid-animation-direction",
            `${preset} 缺少有效方向`,
          );
          continue;
        }
        const rotation = descendantsByName(
          effectTimeNode,
          presentationNamespace,
          "animRot",
        )[0];
        const rotationFrom =
          rotation === undefined
            ? undefined
            : decimalAttribute(rotation, "from");
        if (preset === "spin" && rotationFrom === undefined) {
          issue(
            errors,
            part,
            "invalid-animation-rotation",
            "spin 缺少有效起始角度",
          );
          continue;
        }
        const from =
          direction === undefined
            ? rotationFrom === undefined
              ? undefined
              : rotationFrom / 60000
            : direction;
        effects.push({
          targetId,
          kind,
          preset,
          durationMs,
          delayMs: groupDelay + effectDelay,
          ...(from === undefined ? {} : { from }),
          ...(Object.keys(easing).length === 0 ? {} : { easing }),
        });
      }
    }
    if (effects.length > 0) result.push({ trigger, effects });
  }
  return result;
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
): {
  presentation: ValidationResult["presentation"];
  slides: FileSlide[];
  extensions: ValidationExtension[];
} {
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
    return {
      presentation: { slideCount: 0 },
      slides: [],
      extensions: [],
    };
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
  const usedExtensions = new Set<ValidationExtension>();
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
    const transition = readSlideTransition(slideRoot, slidePath, errors);
    const animations = readSlideAnimations(
      slideRoot,
      slidePath,
      objects,
      errors,
    );
    for (const extension of transition.extensions) {
      usedExtensions.add(extension);
    }
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
      ...(transition.transition === undefined
        ? {}
        : { transition: transition.transition }),
      ...(animations === undefined ? {} : { animations }),
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
    extensions: [...usedExtensions],
  };
}

/**
 * 对同一份 PPTX 字节执行 SDK 结构校验并独立读取页面对象、几何和原生文字，
 * 不读取构建报告。allowedExtensions 会约束本地读取器发现的所有转场兼容分支；
 * SDK 的 MC 分支选择和扩展覆盖有限，校验结果不代表应用证据。
 */
export async function validatePptx(
  input: Uint8Array,
  options: ValidationOptions = {},
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
      extensions: [],
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
  const { presentation, slides, extensions } = readPresentation(
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
  for (const extension of extensions) {
    if (
      options.allowedExtensions !== undefined &&
      !options.allowedExtensions.includes(extension)
    ) {
      issue(
        errors,
        "ppt/slides",
        "forbidden-extension",
        `输出目标禁止实际文件使用 ${extension}`,
      );
    }
  }
  return {
    profile: {
      standard: transitionalStandard,
      schemaValidator: schemaValidatorProfile,
      scope: validationScope,
      applicationEvidence,
    },
    errors,
    extensions,
    presentation,
    slides,
  };
}
