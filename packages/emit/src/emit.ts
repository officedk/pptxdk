import { createHash } from "node:crypto";

import {
  BuildError,
  CapabilityError,
  buildOptionsSchema,
  capabilityErrorToDiagnostic,
  formatBaseline,
  targets,
  toBuildDiagnostic,
  type BuildDiagnostic,
  type BuildOptions,
  type BuildReport,
  type ResolvedResource,
} from "@pptxdk/core/build";
import type {
  CapabilityPackagePart,
  CapabilityPartRelationship,
  CapabilityPresentationContribution,
  CapabilityResource,
  CapabilityResourceClaim,
  CapabilitySharedPackagePart,
  CapabilitySlidePackagePart,
  EmitContext,
  EmitOptions,
  EmitOutput,
  EmitResult,
  Extension,
  NativeLink,
  Resolution,
  ResolvedLink,
  SlideCapabilitySlot,
  SlideEmitContext,
  Target,
} from "@pptxdk/core/capability";
import { createBuildQuery, type SlideLocatorIndex } from "@pptxdk/core/query";
import {
  getInternalResolution,
  type InternalResolution,
  type ResolvedObjectLayer,
  type ResolvedSlideCapabilityRecord,
} from "@pptxdk/core/resolution";
import { strToU8, zipSync } from "fflate";
import type { z } from "zod";

import {
  assertXml10Characters,
  countAllocatedMasterThemeParts,
  hasThemeOverrideContent,
  pptxParts,
  type EmittedSlideMasterSpec,
  type PptxPackagePart,
  type PptxPackageRelationshipSet,
  type PptxPackageResource,
  type PptxRelationship,
} from "#/ooxml.ts";

type StoredPresentationItem = {
  readonly identity: string;
  readonly order: number;
  readonly preview: string;
  readonly content: string;
  readonly partPath: string;
  readonly relationshipId: string;
  readonly relationshipType: string;
  readonly contentType: string;
  readonly path: string;
  readonly resourceOwners: Map<PartPackageState, Set<number>>;
  readonly bytes?: Uint8Array;
  readonly resourceClaim?: Pick<CapabilityResourceClaim, "kind" | "key">;
  readonly resolveBytes?: (characters: string) => Uint8Array;
};

type SharedPartRecord = {
  readonly path: string;
  readonly relationshipId?: string;
  readonly descriptor: CapabilitySharedPackagePart;
};

type PresentationElementRecord = {
  readonly position: "before-slides" | "after-notes-size";
  readonly content: string;
};

type DocumentPackageState = {
  readonly resources: Map<string, PptxPackageResource>;
  readonly resourceTypes: Map<string, string>;
  readonly packageParts: Map<string, PptxPackagePart>;
  readonly packageRelationships: Map<string, PptxRelationship[]>;
  readonly presentationPartPaths: Set<string>;
  readonly sharedParts: Map<string, SharedPartRecord>;
  readonly presentationElements: PresentationElementRecord[];
  nextPackagePartId: number;
  nextPresentationRelationshipId: number;
  nextTextFieldId: number;
};

type PartPackageState = {
  readonly partPath: string;
  readonly baseRelationshipCount: number;
  readonly index: number;
  readonly relationships: PptxRelationship[];
  readonly relationshipIds: Map<string, string>;
  readonly resourceObservations: (ResolvedResource & { objectId: number })[];
};

type PackageContext = {
  readonly document: DocumentPackageState;
  readonly slide: PartPackageState;
};

type DocumentEmitState = DocumentPackageState & {
  readonly usedExtensions: Set<Extension>;
  readonly resourceClaims: ReadonlyMap<
    string,
    {
      readonly identity: string;
      readonly path: string;
      readonly characters: string;
    }
  >;
  readonly presentationAttributes: Map<string, string>;
  readonly presentationCollections: Map<
    string,
    {
      readonly tagName: string;
      readonly records: Map<
        string,
        {
          readonly identity: string;
          readonly tagName: string;
          readonly content: string;
          readonly items: Map<string, StoredPresentationItem>;
        }
      >;
    }
  >;
};

type SlideEmitState = PartPackageState & {
  readonly diagnostics: BuildDiagnostic[];
};

function diagnostic(
  path: string,
  code: string,
  message: string,
  details?: Readonly<Record<string, unknown>>,
): BuildDiagnostic {
  return toBuildDiagnostic(path, code, message, details);
}

function allocateTextFieldId(document: DocumentPackageState): string {
  const serial = document.nextTextFieldId++;
  const suffix = serial.toString(16).padStart(12, "0").toUpperCase();
  return `{00000000-0000-4000-8000-${suffix}}`;
}

function readValueAtPath(root: unknown, path: readonly PropertyKey[]): unknown {
  let current: unknown = root;
  for (const key of path) {
    if (typeof current !== "object" || current === null) return undefined;
    current = Reflect.get(current, key);
  }
  return current;
}

function mapBuildOptionsIssue(
  rawOptions: unknown,
  issue: z.core.$ZodIssue,
): BuildDiagnostic {
  const [head, second, third] = issue.path;
  if (head === undefined) {
    if (issue.code === "unrecognized_keys") {
      return diagnostic(
        "build",
        "unknown-parameter",
        `Build options contain unknown parameter ${issue.keys[0]}`,
        { prop: String(issue.keys[0]), scope: "options" },
      );
    }
    return diagnostic(
      "build",
      "invalid-options",
      "Build options must be an object",
    );
  }
  if (head === "embedFonts") {
    return diagnostic(
      "build.embedFonts",
      "unsupported-font-embedding",
      "embedFonts currently only supports 'full' or 'subset'",
      { field: "embedFonts", mechanism: "embedFonts" },
    );
  }
  if (head === "signal") {
    return diagnostic(
      "build.signal",
      "invalid-options",
      "Build options signal must be an AbortSignal",
      { field: "signal" },
    );
  }
  if (head === "target") {
    if (second === undefined) {
      if (issue.code === "unrecognized_keys") {
        return diagnostic(
          "build.target",
          "unknown-parameter",
          `Target contains unknown parameter ${issue.keys[0]}`,
          { prop: String(issue.keys[0]), scope: "options", field: "target" },
        );
      }
      return diagnostic(
        "build.target",
        "invalid-target",
        "Target must be an object",
        { field: "target" },
      );
    }
    if (second === "name") {
      return diagnostic(
        "build.target",
        "invalid-target",
        "Target requires a non-empty name",
        { field: "target.name" },
      );
    }
    if (second === "format") {
      return diagnostic(
        "build.target",
        "unsupported-format",
        `This build only emits ${formatBaseline}`,
        { field: "target.format", expected: formatBaseline },
      );
    }
    if (second === "extensions") {
      if (third === undefined) {
        return diagnostic(
          "build.target",
          "invalid-target",
          "Target extensions must be an array",
          { field: "target.extensions" },
        );
      }
      if (issue.code === "custom") {
        return diagnostic(
          "build.target",
          "duplicate-extension",
          issue.message,
          { field: "target.extensions" },
        );
      }
      const badExtension = readValueAtPath(rawOptions, issue.path);
      return diagnostic(
        "build.target",
        "unsupported-extension",
        `Unsupported target extension ${String(badExtension)}`,
        { field: "target.extensions", extension: String(badExtension) },
      );
    }
  }
  return diagnostic("build", "invalid-options", issue.message);
}

/** Validates emit/build options (`target`, `embedFonts`, `signal`) and normalizes the target. */
export function validateOptions(options: BuildOptions): Target {
  const parsed = buildOptionsSchema.safeParse(options);
  if (!parsed.success) {
    throw new BuildError(
      parsed.error.issues.map((issue) => mapBuildOptionsIssue(options, issue)),
    );
  }
  return parsed.data.target ?? targets.transitional;
}

function addResource(
  resource: CapabilityResource,
  context: PackageContext,
  objectId?: number,
): string {
  if ("externalTarget" in resource) {
    if (
      typeof resource.externalTarget !== "string" ||
      resource.externalTarget.trim() === "" ||
      typeof resource.relationshipType !== "string" ||
      resource.relationshipType.trim() === ""
    ) {
      throw new CapabilityError(
        "invalid-resource",
        "External resource target or relationship type is invalid",
      );
    }
    try {
      assertXml10Characters(resource.externalTarget);
      assertXml10Characters(resource.relationshipType);
    } catch {
      throw new CapabilityError(
        "invalid-resource",
        "External resource contains characters not allowed in XML",
      );
    }
    const relationshipKey = `external\u0000${resource.relationshipType}\u0000${resource.externalTarget}`;
    const existingRelationship =
      context.slide.relationshipIds.get(relationshipKey);
    if (existingRelationship !== undefined) return existingRelationship;
    const id = `rId${context.slide.relationships.length + context.slide.baseRelationshipCount + 1}`;
    context.slide.relationshipIds.set(relationshipKey, id);
    context.slide.relationships.push({
      id,
      type: resource.relationshipType,
      target: resource.externalTarget,
      targetMode: "External",
    });
    return id;
  }
  if (
    !(resource.bytes instanceof Uint8Array) ||
    resource.bytes.byteLength === 0
  ) {
    throw new CapabilityError(
      "invalid-resource",
      "Resource must contain non-empty binary content",
    );
  }
  if (
    typeof resource.extension !== "string" ||
    !/^[a-z0-9]+$/iu.test(resource.extension) ||
    typeof resource.contentType !== "string" ||
    resource.contentType.trim() === "" ||
    typeof resource.relationshipType !== "string" ||
    resource.relationshipType.trim() === ""
  ) {
    throw new CapabilityError(
      "invalid-resource",
      "Resource format or relationship type is invalid",
    );
  }
  const extension = resource.extension.toLowerCase();
  if (extension === "xml" || extension === "rels") {
    throw new CapabilityError(
      "invalid-resource",
      `Resource extension .${extension} conflicts with core PPTX parts`,
    );
  }
  const knownType = context.document.resourceTypes.get(extension);
  if (knownType !== undefined && knownType !== resource.contentType) {
    throw new CapabilityError(
      "conflicting-resource-type",
      `Resource extension .${extension} corresponds to multiple content types`,
    );
  }
  context.document.resourceTypes.set(extension, resource.contentType);
  try {
    assertXml10Characters(resource.contentType);
    assertXml10Characters(resource.relationshipType);
  } catch {
    throw new CapabilityError(
      "invalid-resource",
      "Resource relationship contains characters not allowed in XML",
    );
  }

  const digest = createHash("sha256").update(resource.bytes).digest("hex");
  const assetKey = `${extension}\u0000${resource.contentType}\u0000${digest}`;
  let asset = context.document.resources.get(assetKey);
  if (asset === undefined) {
    asset = {
      name: `resource${context.document.resources.size + 1}.${extension}`,
      bytes: new Uint8Array(resource.bytes),
      extension,
      contentType: resource.contentType,
    };
    context.document.resources.set(assetKey, asset);
  }

  if (objectId !== undefined) {
    recordResourceObservation(context.slide, objectId, {
      part: `ppt/media/${asset.name}`,
      contentType: asset.contentType,
      byteLength: asset.bytes.byteLength,
    });
  }

  const relationshipKey = `${assetKey}\u0000${resource.relationshipType}`;
  const existingRelationship =
    context.slide.relationshipIds.get(relationshipKey);
  if (existingRelationship !== undefined) return existingRelationship;
  const id = `rId${context.slide.relationships.length + context.slide.baseRelationshipCount + 1}`;
  context.slide.relationshipIds.set(relationshipKey, id);
  context.slide.relationships.push({
    id,
    type: resource.relationshipType,
    target: relationshipTarget(
      context.slide.partPath,
      `ppt/media/${asset.name}`,
    ),
  });
  return id;
}

function relationshipTarget(sourcePath: string, targetPath: string): string {
  const sourceSegments = sourcePath.split("/").slice(0, -1);
  const targetSegments = targetPath.split("/");
  while (
    sourceSegments.length > 0 &&
    targetSegments.length > 0 &&
    sourceSegments[0] === targetSegments[0]
  ) {
    sourceSegments.shift();
    targetSegments.shift();
  }
  return [...sourceSegments.map(() => ".."), ...targetSegments].join("/");
}

function recordResourceObservation(
  slide: PartPackageState,
  objectId: number,
  resource: ResolvedResource,
): void {
  if (
    slide.resourceObservations.some(
      (existing) =>
        existing.objectId === objectId && existing.part === resource.part,
    )
  ) {
    return;
  }
  slide.resourceObservations.push({ objectId, ...resource });
}

function packagePartPath(template: string, sequence: number): string {
  if (
    (template.match(/\{n\}/gu) ?? []).length !== 1 ||
    !/^(?:[a-z0-9_-]+\/)*[a-z0-9_-]+\.[a-z0-9]+$/iu.test(
      template.replace("{n}", String(sequence)),
    )
  ) {
    throw new CapabilityError(
      "invalid-part",
      "Native part path must be a safe package-relative path containing a single part number",
    );
  }
  return template.replace("{n}", String(sequence));
}

/**
 * Allocates a shared part path without colliding with generated deck parts.
 *
 * Shared part templates normally start at a free sequence number. A template such as
 * `ppt/theme/theme{n}.xml` must nevertheless skip the theme part generated from the deck theme,
 * which is written outside shared part registration but already reserved for the package.
 */
function allocateSharedPackagePartPath(
  template: string,
  document: DocumentPackageState,
): string {
  for (;;) {
    const path = packagePartPath(template, document.nextPackagePartId++);
    if (
      !document.packageParts.has(path) &&
      !document.presentationPartPaths.has(path)
    ) {
      return path;
    }
  }
}

function addPackagePart(
  packagePart: CapabilityPackagePart,
  context: PackageContext,
  objectId?: number,
): string {
  const { part, relatedParts } = packagePart;
  if (
    typeof part.pathTemplate !== "string" ||
    typeof part.contentType !== "string" ||
    part.contentType.trim() === "" ||
    typeof packagePart.relationshipType !== "string" ||
    packagePart.relationshipType.trim() === "" ||
    typeof packagePart.content !== "function" ||
    !Array.isArray(relatedParts) ||
    relatedParts.some(
      (related) =>
        typeof related.pathTemplate !== "string" ||
        typeof related.contentType !== "string" ||
        related.contentType.trim() === "" ||
        typeof related.relationshipType !== "string" ||
        related.relationshipType.trim() === "" ||
        !(related.bytes instanceof Uint8Array) ||
        related.bytes.byteLength === 0,
    )
  ) {
    throw new CapabilityError(
      "invalid-part",
      "Native part or its related parts format and relationships are invalid",
    );
  }
  try {
    for (const value of [part.contentType, packagePart.relationshipType]) {
      assertXml10Characters(value);
    }
    for (const related of relatedParts) {
      for (const value of [related.contentType, related.relationshipType]) {
        assertXml10Characters(value);
      }
    }
  } catch {
    throw new CapabilityError(
      "invalid-part",
      "Native part content type or relationship contains characters not allowed in XML 1.0",
    );
  }

  const sequence = context.document.nextPackagePartId++;
  const partPath = packagePartPath(part.pathTemplate, sequence);
  const related = relatedParts.map((item, index) => ({
    path: packagePartPath(item.pathTemplate, sequence),
    relationship: {
      id: `rId${index + 1}`,
      type: item.relationshipType,
      target: "",
    },
    descriptor: item,
  }));
  for (const item of related) {
    item.relationship.target = relationshipTarget(partPath, item.path);
  }
  const relatedIds = related.map(({ relationship }) => relationship.id);
  const content = packagePart.content(relatedIds);
  if (
    !(typeof content === "string"
      ? content.length > 0
      : content instanceof Uint8Array && content.byteLength > 0)
  ) {
    throw new CapabilityError(
      "invalid-part",
      "Native part must contain content",
    );
  }
  const relatedPaths = new Set(related.map(({ path }) => path));
  if (
    context.document.packageParts.has(partPath) ||
    relatedPaths.size !== related.length ||
    relatedPaths.has(partPath) ||
    related.some(({ path }) => context.document.packageParts.has(path))
  ) {
    throw new CapabilityError("duplicate-part", "Duplicate PPTX part path");
  }
  if (objectId !== undefined) {
    recordResourceObservation(context.slide, objectId, {
      part: partPath,
      contentType: part.contentType,
      byteLength:
        typeof content === "string"
          ? strToU8(content).byteLength
          : content.byteLength,
    });
    for (const item of related) {
      recordResourceObservation(context.slide, objectId, {
        part: item.path,
        contentType: item.descriptor.contentType,
        byteLength: item.descriptor.bytes.byteLength,
      });
    }
  }
  context.document.packageParts.set(partPath, {
    path: partPath,
    content,
    contentType: part.contentType,
  });
  for (const item of related) {
    context.document.packageParts.set(item.path, {
      path: item.path,
      content: new Uint8Array(item.descriptor.bytes),
      contentType: item.descriptor.contentType,
    });
  }
  if (related.length > 0) {
    context.document.packageRelationships.set(
      partPath,
      related.map(({ relationship }) => relationship),
    );
  }

  const relationshipId = `rId${context.slide.relationships.length + context.slide.baseRelationshipCount + 1}`;
  context.slide.relationships.push({
    id: relationshipId,
    type: packagePart.relationshipType,
    target: relationshipTarget(context.slide.partPath, partPath),
  });
  return relationshipId;
}

function packagePath(value: unknown, label: string): string {
  if (
    typeof value !== "string" ||
    !/^(?:[a-z0-9_-]+\/)*[a-z0-9_-]+\.[a-z0-9]+$/iu.test(value)
  ) {
    throw new CapabilityError(
      "invalid-part",
      `${label} must be a safe package-relative path`,
    );
  }
  return value;
}

function packagePartDescriptor(
  part: { readonly pathTemplate?: unknown; readonly contentType?: unknown },
  relationshipType: unknown,
  relationships: unknown,
  content: unknown,
  label: string,
  allowsSlideTarget = false,
): {
  readonly pathTemplate: string;
  readonly contentType: string;
  readonly relationshipType?: string;
  readonly relationships: readonly CapabilityPartRelationship[];
} {
  if (
    typeof part.pathTemplate !== "string" ||
    typeof part.contentType !== "string" ||
    part.contentType.trim() === "" ||
    (relationshipType !== undefined &&
      (typeof relationshipType !== "string" ||
        relationshipType.trim() === "")) ||
    typeof content !== "function" ||
    (relationships !== undefined && !Array.isArray(relationships)) ||
    (Array.isArray(relationships) &&
      relationships.some(
        (relationship) =>
          typeof relationship !== "object" ||
          relationship === null ||
          typeof relationship.relationshipType !== "string" ||
          relationship.relationshipType.trim() === "" ||
          typeof relationship.target !== "object" ||
          relationship.target === null ||
          !["slide", "part", "path"].includes(relationship.target.kind) ||
          (relationship.target.kind === "part" &&
            (typeof relationship.target.key !== "string" ||
              relationship.target.key === "")) ||
          (relationship.target.kind === "path" &&
            typeof relationship.target.path !== "string") ||
          (relationship.target.kind === "slide" && !allowsSlideTarget),
      ))
  ) {
    throw new CapabilityError(
      "invalid-part",
      `${label} declaration is invalid`,
    );
  }
  try {
    assertXml10Characters(part.contentType);
    if (typeof relationshipType === "string") {
      assertXml10Characters(relationshipType);
    }
    for (const relationship of relationships ?? []) {
      assertXml10Characters(relationship.relationshipType);
      if (relationship.target.kind === "part") {
        assertXml10Characters(relationship.target.key);
      } else if (relationship.target.kind === "path") {
        assertXml10Characters(relationship.target.path);
      }
    }
  } catch {
    throw new CapabilityError(
      "invalid-part",
      `${label} contains characters not allowed in XML 1.0`,
    );
  }
  return {
    pathTemplate: part.pathTemplate,
    contentType: part.contentType,
    ...(typeof relationshipType === "string" ? { relationshipType } : {}),
    relationships: relationships ?? [],
  };
}

function targetPartPath(
  target: CapabilityPartRelationship["target"],
  context: PackageContext,
): string {
  if (target.kind === "slide") {
    return `ppt/slides/slide${context.slide.index}.xml`;
  }
  if (target.kind === "part") {
    const shared = context.document.sharedParts.get(target.key);
    if (shared === undefined) {
      throw new CapabilityError(
        "missing-part-target",
        `Package part relationship references unregistered shared part ${JSON.stringify(target.key)}`,
      );
    }
    return shared.path;
  }
  return packagePath(target.path, "Package part relationship target");
}

function partContent(
  content: string | Uint8Array,
  label: string,
): string | Uint8Array {
  if (
    !(typeof content === "string"
      ? content.length > 0
      : content instanceof Uint8Array && content.byteLength > 0)
  ) {
    throw new CapabilityError("invalid-part", `${label} must contain content`);
  }
  if (typeof content === "string") assertXml10Characters(content);
  return content;
}

function addSlidePackagePart(
  part: CapabilitySlidePackagePart,
  context: PackageContext,
): string {
  const descriptor = packagePartDescriptor(
    part.part,
    part.slideRelationshipType,
    part.relationships,
    part.content,
    "slide part",
    true,
  );
  const slideRelationshipType = descriptor.relationshipType;
  if (slideRelationshipType === undefined) {
    throw new CapabilityError(
      "invalid-part",
      "Slide part relationship type is required",
    );
  }
  const sequence = context.document.nextPackagePartId++;
  const partPath = packagePartPath(descriptor.pathTemplate, sequence);
  if (
    context.document.packageParts.has(partPath) ||
    context.document.presentationPartPaths.has(partPath)
  ) {
    throw new CapabilityError("duplicate-part", "Duplicate PPTX part path");
  }
  const relationships = descriptor.relationships.map((relationship, index) => ({
    id: `rId${index + 1}`,
    type: relationship.relationshipType,
    target: relationshipTarget(
      partPath,
      targetPartPath(relationship.target, context),
    ),
  }));
  const content = partContent(
    part.content(relationships.map(({ id }) => id)),
    "Slide part",
  );
  context.document.presentationPartPaths.add(partPath);
  context.document.packageParts.set(partPath, {
    path: partPath,
    content,
    contentType: descriptor.contentType,
  });
  if (relationships.length > 0) {
    context.document.packageRelationships.set(partPath, relationships);
  }
  const relationshipId = `rId${context.slide.relationships.length + context.slide.baseRelationshipCount + 1}`;
  context.slide.relationships.push({
    id: relationshipId,
    type: slideRelationshipType,
    target: relationshipTarget(context.slide.partPath, partPath),
  });
  return relationshipId;
}

function addSharedPackagePart(
  key: string,
  part: CapabilitySharedPackagePart,
  context: PackageContext,
): { readonly path: string; readonly relationshipId?: string } {
  if (typeof key !== "string" || key.trim() === "") {
    throw new CapabilityError(
      "invalid-part",
      "Shared part key must be a non-empty string",
    );
  }
  const descriptor = packagePartDescriptor(
    part.part,
    part.presentationRelationshipType,
    part.relationships,
    part.content,
    "shared part",
  );
  const presentationRelationshipType = descriptor.relationshipType;
  if (
    part.presentationElement !== undefined &&
    presentationRelationshipType === undefined
  ) {
    throw new CapabilityError(
      "invalid-part",
      "Shared part presentation element requires a presentation relationship type",
    );
  }
  const existing = context.document.sharedParts.get(key);
  if (existing !== undefined) {
    const same =
      existing.descriptor.part.pathTemplate === descriptor.pathTemplate &&
      existing.descriptor.part.contentType === descriptor.contentType &&
      existing.descriptor.presentationRelationshipType ===
        descriptor.relationshipType &&
      JSON.stringify(existing.descriptor.relationships ?? []) ===
        JSON.stringify(descriptor.relationships);
    if (!same) {
      throw new CapabilityError(
        "conflicting-shared-part",
        `Shared part ${JSON.stringify(key)} has conflicting declarations`,
      );
    }
    return {
      path: existing.path,
      ...(existing.relationshipId === undefined
        ? {}
        : { relationshipId: existing.relationshipId }),
    };
  }
  const partPath = allocateSharedPackagePartPath(
    descriptor.pathTemplate,
    context.document,
  );
  const relationships = descriptor.relationships.map((relationship, index) => ({
    id: `rId${index + 1}`,
    type: relationship.relationshipType,
    target: relationshipTarget(
      partPath,
      targetPartPath(relationship.target, context),
    ),
  }));
  const content = partContent(
    part.content(relationships.map(({ id }) => id)),
    "Shared part",
  );
  context.document.presentationPartPaths.add(partPath);
  context.document.packageParts.set(partPath, {
    path: partPath,
    content,
    contentType: descriptor.contentType,
  });
  if (relationships.length > 0) {
    context.document.packageRelationships.set(partPath, relationships);
  }
  let relationshipId: string | undefined;
  if (presentationRelationshipType !== undefined) {
    relationshipId = `rId${context.document.nextPresentationRelationshipId++}`;
    const presentationRelationships =
      context.document.packageRelationships.get("ppt/presentation.xml") ?? [];
    presentationRelationships.push({
      id: relationshipId,
      type: presentationRelationshipType,
      target: relationshipTarget("ppt/presentation.xml", partPath),
    });
    context.document.packageRelationships.set(
      "ppt/presentation.xml",
      presentationRelationships,
    );
    const element = part.presentationElement;
    if (element !== undefined) {
      if (
        (element.position !== "before-slides" &&
          element.position !== "after-notes-size") ||
        typeof element.content !== "function"
      ) {
        throw new CapabilityError(
          "invalid-part",
          "Shared part presentation element declaration is invalid",
        );
      }
      const xml = element.content(relationshipId);
      if (typeof xml !== "string" || xml.trim() === "") {
        throw new CapabilityError(
          "invalid-part",
          "Shared part presentation element must generate content",
        );
      }
      assertXml10Characters(xml);
      context.document.presentationElements.push({
        position: element.position,
        content: xml,
      });
    }
  }
  context.document.sharedParts.set(key, {
    path: partPath,
    ...(relationshipId === undefined ? {} : { relationshipId }),
    descriptor: {
      part: {
        pathTemplate: descriptor.pathTemplate,
        contentType: descriptor.contentType,
      },
      ...(presentationRelationshipType === undefined
        ? {}
        : { presentationRelationshipType }),
      ...(part.relationships === undefined
        ? {}
        : { relationships: descriptor.relationships }),
      ...(part.presentationElement === undefined
        ? {}
        : { presentationElement: part.presentationElement }),
      content: part.content,
    },
  });
  return {
    path: partPath,
    ...(relationshipId === undefined ? {} : { relationshipId }),
  };
}

function xmlName(value: string): boolean {
  return /^(?:[A-Za-z_][A-Za-z0-9_.-]*:)?[A-Za-z_][A-Za-z0-9_.-]*$/u.test(
    value,
  );
}

function addPresentationContribution(
  contribution: CapabilityPresentationContribution,
  context: {
    readonly document: DocumentEmitState;
    readonly slide: SlideEmitState;
  },
  path: string,
  objectId?: number,
): void {
  const { collection, record, item } = contribution;
  if (
    !xmlName(collection) ||
    typeof record.key !== "string" ||
    record.key === "" ||
    typeof record.identity !== "string" ||
    typeof record.tagName !== "string" ||
    !xmlName(record.tagName) ||
    typeof record.content !== "string" ||
    typeof item.key !== "string" ||
    item.key === "" ||
    typeof item.identity !== "string" ||
    !Number.isInteger(item.order) ||
    item.order < 0 ||
    typeof item.relationshipType !== "string" ||
    item.relationshipType.trim() === "" ||
    typeof item.pathTemplate !== "string" ||
    typeof item.contentType !== "string" ||
    item.contentType.trim() === "" ||
    (item.bytes !== undefined &&
      (!(item.bytes instanceof Uint8Array) || item.bytes.byteLength === 0)) ||
    (item.bytes === undefined && item.resolveBytes === undefined) ||
    (item.resourceClaim === undefined) !== (item.resolveBytes === undefined) ||
    (item.resourceClaim !== undefined &&
      (typeof item.resourceClaim.kind !== "string" ||
        item.resourceClaim.kind === "" ||
        typeof item.resourceClaim.key !== "string" ||
        item.resourceClaim.key === "")) ||
    (item.resolveBytes !== undefined &&
      typeof item.resolveBytes !== "function") ||
    typeof item.content !== "function"
  ) {
    throw new CapabilityError(
      "invalid-presentation-contribution",
      "Presentation-level contribution record, relationship, or resource is invalid",
    );
  }
  try {
    assertXml10Characters(collection);
    assertXml10Characters(record.content);
    assertXml10Characters(item.relationshipType);
    assertXml10Characters(item.contentType);
    for (const [name, value] of Object.entries(
      contribution.presentationAttributes ?? {},
    )) {
      if (!xmlName(name) || typeof value !== "string") {
        throw new Error("invalid presentation attribute");
      }
      assertXml10Characters(value);
    }
  } catch {
    throw new CapabilityError(
      "invalid-presentation-contribution",
      "Presentation-level contribution contains invalid XML name or characters",
    );
  }

  for (const [name, value] of Object.entries(
    contribution.presentationAttributes ?? {},
  )) {
    const existing = context.document.presentationAttributes.get(name);
    if (existing !== undefined && existing !== value) {
      context.slide.diagnostics.push(
        diagnostic(
          path,
          "conflicting-presentation-attribute",
          `Conflicting contribution value for presentation attribute ${name}`,
        ),
      );
    } else {
      context.document.presentationAttributes.set(name, value);
    }
  }

  let targetCollection =
    context.document.presentationCollections.get(collection);
  if (targetCollection === undefined) {
    targetCollection = { tagName: collection, records: new Map() };
    context.document.presentationCollections.set(collection, targetCollection);
  }
  let targetRecord = targetCollection.records.get(record.key);
  if (targetRecord === undefined) {
    targetRecord = {
      identity: record.identity,
      tagName: record.tagName,
      content: record.content,
      items: new Map(),
    };
    targetCollection.records.set(record.key, targetRecord);
  } else if (
    targetRecord.identity !== record.identity ||
    targetRecord.tagName !== record.tagName ||
    targetRecord.content !== record.content
  ) {
    context.slide.diagnostics.push(
      diagnostic(
        path,
        "conflicting-presentation-record",
        `Conflicting declaration for presentation-level record ${record.key}`,
      ),
    );
    return;
  }

  let preview: string;
  try {
    preview = item.content("rId0");
    assertXml10Characters(preview);
  } catch {
    throw new CapabilityError(
      "invalid-presentation-contribution",
      "Presentation-level contribution failed to generate XML content",
    );
  }
  if (preview.trim() === "") {
    throw new CapabilityError(
      "invalid-presentation-contribution",
      "Presentation-level contribution must generate content",
    );
  }
  const existingItem = targetRecord.items.get(item.key);
  if (existingItem !== undefined) {
    const conflicts =
      existingItem.identity !== item.identity ||
      existingItem.order !== item.order ||
      existingItem.preview !== preview ||
      existingItem.resourceClaim?.kind !== item.resourceClaim?.kind ||
      existingItem.resourceClaim?.key !== item.resourceClaim?.key;
    if (conflicts) {
      context.slide.diagnostics.push(
        diagnostic(
          path,
          "conflicting-presentation-item",
          `Presentation-level record ${record.key} item ${item.key} uses conflicting resources`,
        ),
      );
    } else {
      const objectIds =
        existingItem.resourceOwners.get(context.slide) ?? new Set<number>();
      if (objectId !== undefined) objectIds.add(objectId);
      existingItem.resourceOwners.set(context.slide, objectIds);
    }
    return;
  }

  let sequence = 1;
  let partPath = packagePartPath(item.pathTemplate, sequence);
  while (
    context.document.packageParts.has(partPath) ||
    context.document.presentationPartPaths.has(partPath)
  ) {
    sequence += 1;
    partPath = packagePartPath(item.pathTemplate, sequence);
  }
  context.document.presentationPartPaths.add(partPath);
  const relationshipId = `rId${context.document.nextPresentationRelationshipId++}`;
  const content = item.content(relationshipId);
  targetRecord.items.set(item.key, {
    identity: item.identity,
    order: item.order,
    preview,
    content,
    partPath,
    relationshipId,
    relationshipType: item.relationshipType,
    contentType: item.contentType,
    path,
    resourceOwners: new Map([
      [
        context.slide,
        objectId === undefined ? new Set<number>() : new Set([objectId]),
      ],
    ]),
    ...(item.bytes === undefined ? {} : { bytes: item.bytes }),
    ...(item.resourceClaim === undefined
      ? {}
      : { resourceClaim: item.resourceClaim }),
    ...(item.resolveBytes === undefined
      ? {}
      : { resolveBytes: item.resolveBytes }),
  });
}

function finalizePresentationContributions(document: DocumentEmitState): void {
  for (const collection of document.presentationCollections.values()) {
    for (const record of collection.records.values()) {
      for (const item of record.items.values()) {
        let bytes = item.bytes;
        if (item.resourceClaim !== undefined) {
          const claimKey = `${item.resourceClaim.kind}\u0000${item.resourceClaim.key}`;
          const claim = document.resourceClaims.get(claimKey);
          if (claim === undefined || claim.identity !== item.identity) {
            throw new BuildError([
              diagnostic(
                item.path,
                "missing-resource-claim",
                `Presentation resource ${item.resourceClaim.key} is missing matching character claims`,
              ),
            ]);
          }
          try {
            bytes = item.resolveBytes!(claim.characters);
          } catch (error) {
            if (error instanceof CapabilityError) {
              throw new BuildError([
                capabilityErrorToDiagnostic(error, item.path),
              ]);
            }
            throw new BuildError([
              diagnostic(
                item.path,
                "resource-contribution-failed",
                error instanceof Error ? error.message : String(error),
              ),
            ]);
          }
        }
        if (!(bytes instanceof Uint8Array) || bytes.byteLength === 0) {
          throw new BuildError([
            diagnostic(
              item.path,
              "invalid-presentation-contribution",
              "Presentation-level contribution did not produce valid resource bytes",
            ),
          ]);
        }
        document.packageParts.set(item.partPath, {
          path: item.partPath,
          content: new Uint8Array(bytes),
          contentType: item.contentType,
        });
        const relationships =
          document.packageRelationships.get("ppt/presentation.xml") ?? [];
        relationships.push({
          id: item.relationshipId,
          type: item.relationshipType,
          target: relationshipTarget("ppt/presentation.xml", item.partPath),
        });
        document.packageRelationships.set(
          "ppt/presentation.xml",
          relationships,
        );
        for (const [slide, objectIds] of item.resourceOwners) {
          for (const objectId of objectIds) {
            recordResourceObservation(slide, objectId, {
              part: item.partPath,
              contentType: item.contentType,
              byteLength: bytes.byteLength,
            });
          }
        }
      }
    }
  }
}

function presentationContent(document: DocumentEmitState): {
  readonly beforeSlides: string;
  readonly afterNotesSize: string;
} {
  const rendered = [...document.presentationCollections.values()]
    .map(
      ({ tagName, records }) =>
        `<${tagName}>${[...records.values()]
          .map(
            (record) =>
              `<${record.tagName}>${record.content}${[...record.items.values()]
                .sort((left, right) => left.order - right.order)
                .map(({ content }) => content)
                .join("")}</${record.tagName}>`,
          )
          .join("")}</${tagName}>`,
    )
    .join("");
  const elements = (position: "before-slides" | "after-notes-size") =>
    document.presentationElements
      .filter((element) => element.position === position)
      .map(({ content }) => content)
      .join("");
  return {
    beforeSlides: elements("before-slides"),
    afterNotesSize: rendered + elements("after-notes-size"),
  };
}

function registerLinkRelationship(
  link: ResolvedLink,
  slideIndexById: ReadonlyMap<string, number>,
  slide: PartPackageState,
): NativeLink {
  const { target } = link;
  if ("id" in target) {
    const targetIndex = slideIndexById.get(target.id);
    if (targetIndex === undefined) {
      throw new Error(
        `Missing validated target slide ${JSON.stringify(target.id)}`,
      );
    }
    const relationshipKey = `link:slide:${target.id}`;
    let relationshipId = slide.relationshipIds.get(relationshipKey);
    if (relationshipId === undefined) {
      relationshipId = `rId${slide.relationships.length + slide.baseRelationshipCount + 1}`;
      slide.relationshipIds.set(relationshipKey, relationshipId);
      slide.relationships.push({
        id: relationshipId,
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
        target: relationshipTarget(
          slide.partPath,
          `ppt/slides/slide${targetIndex}.xml`,
        ),
      });
    }
    return { relationshipId, action: "ppaction://hlinksldjump" };
  }
  if ("url" in target) {
    const relationshipKey = `link:url:${target.url}`;
    let relationshipId = slide.relationshipIds.get(relationshipKey);
    if (relationshipId === undefined) {
      relationshipId = `rId${slide.relationships.length + slide.baseRelationshipCount + 1}`;
      slide.relationshipIds.set(relationshipKey, relationshipId);
      slide.relationships.push({
        id: relationshipId,
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        target: target.url,
        targetMode: "External",
      });
    }
    return { relationshipId };
  }
  return {
    relationshipId: "",
    action: "ppaction://hlinkshowjump?jump=firstslide",
  };
}

function slideSlotOrder(slot: SlideCapabilitySlot): number {
  return slot === "transition" ? 0 : 1;
}

function createDocumentEmitState(
  internal: InternalResolution,
): DocumentEmitState {
  const themePartCount = countAllocatedMasterThemeParts(
    internal.masters,
    internal.theme.theme,
  );
  return {
    usedExtensions: new Set(),
    resourceClaims: internal.resourceClaims,
    resources: new Map(),
    resourceTypes: new Map(),
    packageParts: new Map(),
    packageRelationships: new Map(),
    presentationAttributes: new Map(),
    // Master theme parts are generated outside shared part registration; reserve all of them so
    // capability-provided theme parts (e.g. notes master) get distinct OPC identities.
    presentationPartPaths: new Set(
      Array.from(
        { length: themePartCount },
        (_, index) => `ppt/theme/theme${index + 1}.xml`,
      ),
    ),
    presentationCollections: new Map(),
    sharedParts: new Map(),
    presentationElements: [],
    nextPackagePartId: 1,
    nextPresentationRelationshipId:
      (internal.masters?.length ?? 1) + internal.slides.length + 1,
    nextTextFieldId: 1,
  };
}

function checkExtensions(
  extensions: readonly Extension[],
  target: Target,
  usedExtensions: Set<Extension>,
  diagnostics: BuildDiagnostic[],
  path: string,
): void {
  for (const extension of extensions) {
    if (!target.extensions.includes(extension)) {
      diagnostics.push(
        diagnostic(path, "forbidden-mechanism", `Target forbids ${extension}`, {
          target: target.name,
          extension,
        }),
      );
    } else {
      usedExtensions.add(extension);
    }
  }
}

function emitLayer(
  layer: ResolvedObjectLayer,
  slideIndexById: ReadonlyMap<string, number>,
  docEmit: DocumentEmitState,
  slideEmit: SlideEmitState,
  target: Target,
  options: EmitOptions,
): string {
  const packageCtx: PackageContext = { document: docEmit, slide: slideEmit };

  checkExtensions(
    layer.extensions,
    target,
    docEmit.usedExtensions,
    slideEmit.diagnostics,
    layer.path,
  );

  const childrenXml =
    layer.childComposition === "members"
      ? layer.memberLayers
          .map((child) =>
            emitLayer(
              child,
              slideIndexById,
              docEmit,
              slideEmit,
              target,
              options,
            ),
          )
          .join("")
      : "";
  const siblingXml =
    layer.childComposition === "siblings"
      ? layer.siblingLayers
          .map((child) =>
            emitLayer(
              child,
              slideIndexById,
              docEmit,
              slideEmit,
              target,
              options,
            ),
          )
          .join("")
      : "";

  let ownXml = "";
  try {
    const nativeLink =
      layer.resolvedLink === undefined
        ? undefined
        : registerLinkRelationship(
            layer.resolvedLink,
            slideIndexById,
            slideEmit,
          );
    const emitContext: EmitContext = {
      id: layer.objectId,
      path: layer.path,
      theme: layer.theme,
      options,
      slideNumber: slideEmit.index,
      ...(nativeLink === undefined ? {} : { nativeLink }),
      childrenXml,
      createTextFieldId: () => allocateTextFieldId(docEmit),
      registerResource: (resource) =>
        addResource(resource, packageCtx, layer.objectId),
      registerPackagePart: (part) =>
        addPackagePart(part, packageCtx, layer.objectId),
    };
    const output: EmitOutput = layer.emit(emitContext);
    ownXml = output.xml;
    if (layer.morphKey !== undefined && layer.emitsObject && ownXml === "") {
      slideEmit.diagnostics.push(
        diagnostic(
          layer.path,
          "morph-needs-object",
          "morphKey can only be used with capabilities that produce native objects",
        ),
      );
    }
    for (const contribution of output.presentationContributions ?? []) {
      addPresentationContribution(
        contribution,
        { document: docEmit, slide: slideEmit },
        layer.path,
        layer.objectId,
      );
    }
  } catch (error) {
    slideEmit.diagnostics.push(
      error instanceof CapabilityError
        ? capabilityErrorToDiagnostic(error, layer.path, true)
        : diagnostic(
            layer.path,
            "capability-failure",
            error instanceof Error ? error.message : String(error),
          ),
    );
  }

  return layer.childComposition === "members" ? ownXml : ownXml + siblingXml;
}

const emittedQueries = new WeakMap<BuildReport, SlideLocatorIndex>();

function maxResolvedObjectId(layers: readonly ResolvedObjectLayer[]): number {
  let maxId = 1;
  for (const layer of layers) {
    if (layer.objectId > maxId) maxId = layer.objectId;
    const memberMax = maxResolvedObjectId(layer.memberLayers);
    if (memberMax > maxId) maxId = memberMax;
    const siblingMax = maxResolvedObjectId(layer.siblingLayers);
    if (siblingMax > maxId) maxId = siblingMax;
  }
  return maxId;
}

/**
 * Returns the slide locator index enriched with emitted package resources for a BuildReport produced by `emit()`.
 */
export function emittedBuildQuery(
  report: BuildReport,
  fallbackQuery: SlideLocatorIndex,
): SlideLocatorIndex {
  return emittedQueries.get(report) ?? fallbackQuery;
}

function emitMasterOrLayoutCapabilities(
  capabilities: readonly ResolvedSlideCapabilityRecord[] | undefined,
  internal: InternalResolution,
  docEmit: DocumentEmitState,
  partEmit: SlideEmitState,
  target: Target,
  options: EmitOptions,
): { readonly transitionXml: string; readonly timingXml: string } {
  if (capabilities === undefined || capabilities.length === 0) {
    return { transitionXml: "", timingXml: "" };
  }
  const packageCtx: PackageContext = { document: docEmit, slide: partEmit };
  const transitionParts: string[] = [];
  const timingParts: string[] = [];

  for (const cap of capabilities) {
    checkExtensions(
      cap.extensions,
      target,
      docEmit.usedExtensions,
      partEmit.diagnostics,
      cap.path,
    );
    try {
      const output = cap.emit({
        path: cap.path,
        dimensions: internal.dimensions,
        theme: cap.theme,
        options,
        slideNumber: 0,
        createTextFieldId: () => allocateTextFieldId(docEmit),
        registerResource: (resource) => addResource(resource, packageCtx),
        registerPackagePart: (part) => addSlidePackagePart(part, packageCtx),
        registerSharedPart: (key, part) =>
          addSharedPackagePart(key, part, packageCtx),
      });
      if (cap.slot === "transition" && output.xml !== "") {
        transitionParts.push(output.xml);
      } else if (cap.slot === "timing" && output.xml !== "") {
        timingParts.push(output.xml);
      }
    } catch (error) {
      partEmit.diagnostics.push(
        error instanceof CapabilityError
          ? capabilityErrorToDiagnostic(error, cap.path)
          : diagnostic(
              cap.path,
              "capability-failure",
              error instanceof Error ? error.message : String(error),
            ),
      );
    }
  }
  return {
    transitionXml: transitionParts.join(""),
    timingXml: timingParts.join(""),
  };
}

function emitOwnerPlaceholderVisuals(
  placeholderVisuals: ResolvedSlideCapabilityRecord | undefined,
  internal: InternalResolution,
  docEmit: DocumentEmitState,
  ownerEmit: SlideEmitState,
  ownerCtx: PackageContext,
  target: Target,
  options: EmitOptions,
): Readonly<Record<string, string>> | undefined {
  if (placeholderVisuals === undefined) return undefined;
  checkExtensions(
    placeholderVisuals.extensions,
    target,
    docEmit.usedExtensions,
    ownerEmit.diagnostics,
    placeholderVisuals.path,
  );
  const phOut = placeholderVisuals.emit({
    path: placeholderVisuals.path,
    dimensions: internal.dimensions,
    theme: placeholderVisuals.theme,
    options,
    slideNumber: 1,
    createTextFieldId: () => allocateTextFieldId(docEmit),
    registerResource: (resource) => addResource(resource, ownerCtx),
    registerPackagePart: (part) => addSlidePackagePart(part, ownerCtx),
    registerSharedPart: (key, part) =>
      addSharedPackagePart(key, part, ownerCtx),
  });
  return phOut.placeholderSpPrByKey;
}

/**
 * Emits a PPTX archive and BuildReport from an existing Resolution without re-resolving the document.
 */
export function emit(
  resolution: Resolution,
  options: EmitOptions = {},
): EmitResult {
  const target = validateOptions(options);
  if (options.signal?.aborted) throw options.signal.reason;
  const internal = getInternalResolution(resolution);

  const docEmit = createDocumentEmitState(internal);
  let sectionXml = "";
  if (internal.sectionsEmitter !== undefined) {
    const forbiddenExtensions = internal.sectionsEmitter.extensions.filter(
      (ext) => !target.extensions.includes(ext),
    );
    if (forbiddenExtensions.length > 0) {
      throw new BuildError([
        diagnostic(
          "deck/sections",
          "forbidden-mechanism",
          "Current target forbids p14 section mechanism",
          {
            target: target.name,
            extension: forbiddenExtensions[0],
            mechanism: "sections",
          },
        ),
      ]);
    }
    for (const ext of internal.sectionsEmitter.extensions) {
      docEmit.usedExtensions.add(ext);
    }
    sectionXml = internal.sectionsEmitter.emit().xml;
  }

  const mastersEmitDriver = internal.mastersEmitDriver;
  const slideIndexById = new Map(
    internal.slides.map((slide) => [slide.id, slide.index] as const),
  );

  let globalLayoutNumber = 0;
  const emittedMasters: EmittedSlideMasterSpec[] | undefined =
    internal.masters?.map((master, masterIndex) => {
      const masterPartPath = `ppt/slideMasters/slideMaster${masterIndex + 1}.xml`;
      const masterEmit: SlideEmitState = {
        partPath: masterPartPath,
        baseRelationshipCount: master.layouts.length + 1,
        index: 1,
        diagnostics: [],
        relationships: [],
        relationshipIds: new Map(),
        resourceObservations: [],
      };
      const masterCtx: PackageContext = {
        document: docEmit,
        slide: masterEmit,
      };
      const masterCapXml = emitMasterOrLayoutCapabilities(
        master.slideCapabilities,
        internal,
        docEmit,
        masterEmit,
        target,
        options,
      );
      let masterBackgroundXml = "";
      if (master.background !== undefined) {
        checkExtensions(
          master.background.extensions,
          target,
          docEmit.usedExtensions,
          masterEmit.diagnostics,
          master.background.path,
        );
        const bgOut = master.background.emit({
          path: master.background.path,
          dimensions: internal.dimensions,
          theme: master.background.theme,
          options,
          slideNumber: 1,
          createTextFieldId: () => allocateTextFieldId(docEmit),
          registerResource: (resource) => addResource(resource, masterCtx),
          registerPackagePart: (part) => addSlidePackagePart(part, masterCtx),
          registerSharedPart: (key, part) =>
            addSharedPackagePart(key, part, masterCtx),
        });
        masterBackgroundXml = bgOut.xml;
      }
      let nextMasterObjectId =
        Math.max(
          maxResolvedObjectId(master.layers ?? []),
          1 + Object.keys(master.placeholders ?? {}).length,
        ) + 1;
      const masterShapeFragments = (master.layers ?? []).map((layer) =>
        emitLayer(layer, slideIndexById, docEmit, masterEmit, target, options),
      );
      if (master.footerPlaceholders !== undefined) {
        checkExtensions(
          master.footerPlaceholders.extensions,
          target,
          docEmit.usedExtensions,
          masterEmit.diagnostics,
          master.footerPlaceholders.path,
        );
        const fpOut = master.footerPlaceholders.emit({
          path: master.footerPlaceholders.path,
          dimensions: internal.dimensions,
          theme: master.footerPlaceholders.theme,
          options,
          slideNumber: 1,
          createTextFieldId: () => allocateTextFieldId(docEmit),
          allocateObjectId: () => nextMasterObjectId++,
          registerResource: (resource) => addResource(resource, masterCtx),
          registerPackagePart: (part) => addSlidePackagePart(part, masterCtx),
          registerSharedPart: (key, part) =>
            addSharedPackagePart(key, part, masterCtx),
        });
        masterShapeFragments.push(fpOut.xml);
        for (const contribution of fpOut.presentationContributions ?? []) {
          addPresentationContribution(
            contribution,
            { document: docEmit, slide: masterEmit },
            master.footerPlaceholders.path,
          );
        }
      }
      const masterShapeXml =
        mastersEmitDriver !== undefined
          ? mastersEmitDriver.emitSharedCanvas(masterShapeFragments)
          : masterShapeFragments.join("");
      const masterPlaceholderSpPrByKey = emitOwnerPlaceholderVisuals(
        master.placeholderVisuals,
        internal,
        docEmit,
        masterEmit,
        masterCtx,
        target,
        options,
      );
      if (masterEmit.diagnostics.length > 0) {
        throw new BuildError(masterEmit.diagnostics);
      }

      const emittedLayouts = master.layouts.map((layout) => {
        const layoutNumber = ++globalLayoutNumber;
        const layoutPartPath = `ppt/slideLayouts/slideLayout${layoutNumber}.xml`;
        const layoutEmit: SlideEmitState = {
          partPath: layoutPartPath,
          baseRelationshipCount: hasThemeOverrideContent(layout.themeOverride)
            ? 2
            : 1,
          index: 1,
          diagnostics: [],
          relationships: [],
          relationshipIds: new Map(),
          resourceObservations: [],
        };
        const layoutCtx: PackageContext = {
          document: docEmit,
          slide: layoutEmit,
        };
        const layoutCapXml = emitMasterOrLayoutCapabilities(
          layout.slideCapabilities,
          internal,
          docEmit,
          layoutEmit,
          target,
          options,
        );
        let layoutBackgroundXml = "";
        if (layout.background !== undefined) {
          checkExtensions(
            layout.background.extensions,
            target,
            docEmit.usedExtensions,
            layoutEmit.diagnostics,
            layout.background.path,
          );
          const bgOut = layout.background.emit({
            path: layout.background.path,
            dimensions: internal.dimensions,
            theme: layout.background.theme,
            options,
            slideNumber: 1,
            createTextFieldId: () => allocateTextFieldId(docEmit),
            registerResource: (resource) => addResource(resource, layoutCtx),
            registerPackagePart: (part) => addSlidePackagePart(part, layoutCtx),
            registerSharedPart: (key, part) =>
              addSharedPackagePart(key, part, layoutCtx),
          });
          layoutBackgroundXml = bgOut.xml;
        }
        let nextLayoutObjectId =
          Math.max(
            maxResolvedObjectId(layout.layers ?? []),
            1 + Object.keys(layout.placeholders ?? {}).length,
          ) + 1;
        const layoutShapeFragments = (layout.layers ?? []).map((layer) =>
          emitLayer(
            layer,
            slideIndexById,
            docEmit,
            layoutEmit,
            target,
            options,
          ),
        );
        if (layout.footerPlaceholders !== undefined) {
          checkExtensions(
            layout.footerPlaceholders.extensions,
            target,
            docEmit.usedExtensions,
            layoutEmit.diagnostics,
            layout.footerPlaceholders.path,
          );
          const fpOut = layout.footerPlaceholders.emit({
            path: layout.footerPlaceholders.path,
            dimensions: internal.dimensions,
            theme: layout.footerPlaceholders.theme,
            options,
            slideNumber: 1,
            createTextFieldId: () => allocateTextFieldId(docEmit),
            allocateObjectId: () => nextLayoutObjectId++,
            registerResource: (resource) => addResource(resource, layoutCtx),
            registerPackagePart: (part) => addSlidePackagePart(part, layoutCtx),
            registerSharedPart: (key, part) =>
              addSharedPackagePart(key, part, layoutCtx),
          });
          layoutShapeFragments.push(fpOut.xml);
          for (const contribution of fpOut.presentationContributions ?? []) {
            addPresentationContribution(
              contribution,
              { document: docEmit, slide: layoutEmit },
              layout.footerPlaceholders.path,
            );
          }
        }
        const layoutShapeXml =
          mastersEmitDriver !== undefined
            ? mastersEmitDriver.emitSharedCanvas(layoutShapeFragments)
            : layoutShapeFragments.join("");
        const layoutPlaceholderSpPrByKey = emitOwnerPlaceholderVisuals(
          layout.placeholderVisuals,
          internal,
          docEmit,
          layoutEmit,
          layoutCtx,
          target,
          options,
        );
        if (layoutEmit.diagnostics.length > 0) {
          throw new BuildError(layoutEmit.diagnostics);
        }
        return {
          key: layout.key,
          ...(layout.name === undefined ? {} : { name: layout.name }),
          ...(layout.type === undefined ? {} : { type: layout.type }),
          ...(layout.preserve === undefined
            ? {}
            : { preserve: layout.preserve }),
          ...(layout.matchingName === undefined
            ? {}
            : { matchingName: layout.matchingName }),
          ...(layout.userDrawn === undefined
            ? {}
            : { userDrawn: layout.userDrawn }),
          ...(layout.showMasterShapes === undefined
            ? {}
            : { showMasterShapes: layout.showMasterShapes }),
          ...(layout.showMasterPlaceholderAnimations === undefined
            ? {}
            : {
                showMasterPlaceholderAnimations:
                  layout.showMasterPlaceholderAnimations,
              }),
          ...(layout.placeholders === undefined
            ? {}
            : { placeholders: layout.placeholders }),
          ...(layoutPlaceholderSpPrByKey === undefined
            ? {}
            : { placeholderSpPrByKey: layoutPlaceholderSpPrByKey }),
          ...(layout.colorMapOverride === undefined
            ? {}
            : { colorMapOverride: layout.colorMapOverride }),
          ...(layout.themeOverride === undefined
            ? {}
            : { themeOverride: layout.themeOverride }),
          ...(layout.headerFooter === undefined
            ? {}
            : { headerFooter: layout.headerFooter }),
          ...(layoutBackgroundXml === ""
            ? {}
            : { backgroundXml: layoutBackgroundXml }),
          ...(layoutShapeXml === "" ? {} : { shapeXml: layoutShapeXml }),
          ...(layoutEmit.relationships.length === 0
            ? {}
            : { extraRelationships: layoutEmit.relationships }),
          ...(layoutCapXml.transitionXml === ""
            ? {}
            : { transitionXml: layoutCapXml.transitionXml }),
          ...(layoutCapXml.timingXml === ""
            ? {}
            : { timingXml: layoutCapXml.timingXml }),
        };
      });

      return {
        key: master.key,
        ...(master.name === undefined ? {} : { name: master.name }),
        ...(master.preserve === undefined ? {} : { preserve: master.preserve }),
        theme: master.theme,
        themeContext: master.themeContext,
        colorMap: master.colorMap,
        ...(master.placeholders === undefined
          ? {}
          : { placeholders: master.placeholders }),
        ...(masterPlaceholderSpPrByKey === undefined
          ? {}
          : { placeholderSpPrByKey: masterPlaceholderSpPrByKey }),
        ...(master.textStyles === undefined
          ? {}
          : { textStyles: master.textStyles }),
        ...(master.headerFooter === undefined
          ? {}
          : { headerFooter: master.headerFooter }),
        ...(masterBackgroundXml === ""
          ? {}
          : { backgroundXml: masterBackgroundXml }),
        ...(masterShapeXml === "" ? {} : { shapeXml: masterShapeXml }),
        ...(masterEmit.relationships.length === 0
          ? {}
          : { extraRelationships: masterEmit.relationships }),
        ...(masterCapXml.transitionXml === ""
          ? {}
          : { transitionXml: masterCapXml.transitionXml }),
        ...(masterCapXml.timingXml === ""
          ? {}
          : { timingXml: masterCapXml.timingXml }),
        layouts: emittedLayouts,
      };
    });

  const emittedSlides = internal.slides.map((resolvedSlide) => {
    const slideEmit: SlideEmitState = {
      partPath: `ppt/slides/slide${resolvedSlide.index}.xml`,
      baseRelationshipCount: 1,
      index: resolvedSlide.index,
      diagnostics: [],
      relationships: [],
      relationshipIds: new Map(),
      resourceObservations: [],
    };
    const packageCtx: PackageContext = { document: docEmit, slide: slideEmit };

    let nextSlideObjectId = maxResolvedObjectId(resolvedSlide.layers) + 1;
    let shapeXml = resolvedSlide.layers
      .map((layer) =>
        emitLayer(layer, slideIndexById, docEmit, slideEmit, target, options),
      )
      .join("");

    let backgroundXml = "";
    const slideXmlParts: {
      readonly slot: SlideCapabilitySlot;
      readonly xml: string;
    }[] = [];

    for (const slideCap of resolvedSlide.slideCapabilities) {
      checkExtensions(
        slideCap.extensions,
        target,
        docEmit.usedExtensions,
        slideEmit.diagnostics,
        slideCap.path,
      );
      try {
        const slideEmitContext: SlideEmitContext = {
          path: slideCap.path,
          dimensions: internal.dimensions,
          theme: slideCap.theme,
          options,
          slideNumber: resolvedSlide.index,
          createTextFieldId: () => allocateTextFieldId(docEmit),
          allocateObjectId: () => nextSlideObjectId++,
          registerResource: (resource) => addResource(resource, packageCtx),
          registerPackagePart: (part) => addSlidePackagePart(part, packageCtx),
          registerSharedPart: (key, part) =>
            addSharedPackagePart(key, part, packageCtx),
        };
        const output = slideCap.emit(slideEmitContext);
        if (slideCap.slot === "background") backgroundXml += output.xml;
        else if (slideCap.slot === "shape-tree") shapeXml += output.xml;
        else slideXmlParts.push({ slot: slideCap.slot, xml: output.xml });
        for (const contribution of output.presentationContributions ?? []) {
          addPresentationContribution(
            contribution,
            { document: docEmit, slide: slideEmit },
            slideCap.path,
          );
        }
      } catch (error) {
        slideEmit.diagnostics.push(
          error instanceof CapabilityError
            ? capabilityErrorToDiagnostic(error, slideCap.path)
            : diagnostic(
                slideCap.path,
                "capability-failure",
                error instanceof Error ? error.message : String(error),
              ),
        );
      }
    }

    if (slideEmit.diagnostics.length > 0) {
      throw new BuildError(slideEmit.diagnostics);
    }

    return {
      id: resolvedSlide.id,
      ...(resolvedSlide.name === undefined ? {} : { name: resolvedSlide.name }),
      ...(resolvedSlide.layoutKey === undefined
        ? {}
        : { layoutKey: resolvedSlide.layoutKey }),
      ...(resolvedSlide.colorMapOverride === undefined
        ? {}
        : { colorMapOverride: resolvedSlide.colorMapOverride }),
      ...(resolvedSlide.themeOverride === undefined
        ? {}
        : { themeOverride: resolvedSlide.themeOverride }),
      hidden: resolvedSlide.hidden,
      ...(resolvedSlide.showMasterPlaceholderAnimations === undefined
        ? {}
        : {
            showMasterPlaceholderAnimations:
              resolvedSlide.showMasterPlaceholderAnimations,
          }),
      ...(resolvedSlide.showMasterShapes === undefined
        ? {}
        : { showMasterShapes: resolvedSlide.showMasterShapes }),
      shapeXml,
      backgroundXml,
      slideXml: slideXmlParts
        .map((item, index) => ({ item, index }))
        .sort(
          (left, right) =>
            slideSlotOrder(left.item.slot) - slideSlotOrder(right.item.slot) ||
            left.index - right.index,
        )
        .map(({ item }) => item.xml)
        .join(""),
      relationships: slideEmit.relationships,
      resourceObservations: slideEmit.resourceObservations,
    };
  });

  finalizePresentationContributions(docEmit);

  const parts = pptxParts(
    internal.dimensions,
    emittedSlides.map(
      ({
        id,
        name,
        layoutKey,
        colorMapOverride,
        themeOverride,
        hidden,
        showMasterPlaceholderAnimations,
        showMasterShapes,
        shapeXml,
        backgroundXml,
        slideXml,
        relationships,
      }) => ({
        id,
        ...(name === undefined ? {} : { name }),
        ...(layoutKey === undefined ? {} : { layoutKey }),
        ...(colorMapOverride === undefined ? {} : { colorMapOverride }),
        ...(themeOverride === undefined ? {} : { themeOverride }),
        hidden,
        ...(showMasterPlaceholderAnimations === undefined
          ? {}
          : { showMasterPlaceholderAnimations }),
        ...(showMasterShapes === undefined ? {} : { showMasterShapes }),
        shapeXml,
        backgroundXml,
        slideXml,
        relationships,
      }),
    ),
    [...docEmit.resources.values()],
    [...docEmit.packageParts.values()],
    [...docEmit.packageRelationships].map(([sourcePath, relationships]) => ({
      sourcePath,
      relationships,
    })) satisfies PptxPackageRelationshipSet[],
    presentationContent(docEmit),
    [...docEmit.presentationAttributes].map(([name, value]) => ({
      name,
      value,
    })),
    internal.theme.theme,
    internal.title,
    sectionXml,
    emittedMasters,
    mastersEmitDriver,
  );
  const pptx = zipSync(
    Object.fromEntries(
      Object.entries(parts).map(([partPath, content]) => [
        partPath,
        typeof content === "string" ? strToU8(content) : content,
      ]),
    ),
    { level: 6, mtime: new Date("2020-01-01T00:00:00Z") },
  );

  // Preserve slide identity from resolution.report.slides so WeakMap-backed
  // slide queries (such as motionSchedule) remain bound across resolve and emit.
  const report: BuildReport = Object.freeze({
    evidence: "generation",
    ...(typeof internal.title === "string" ? { title: internal.title } : {}),
    dimensions: resolution.report.dimensions,
    target: Object.freeze({
      name: target.name,
      format: target.format,
      extensions: Object.freeze([...target.extensions]),
    }),
    extensions: Object.freeze([...docEmit.usedExtensions]),
    fonts: resolution.report.fonts,
    ...(resolution.report.masters === undefined
      ? {}
      : { masters: resolution.report.masters }),
    slides: resolution.report.slides,
  });
  const query = createBuildQuery(
    internal.slides.map((slide, index) => ({
      observation: report.slides[index]!,
      dimensions: report.dimensions,
      content: slide.content,
      resources: emittedSlides[index]!.resourceObservations,
    })),
  );
  emittedQueries.set(report, query);
  return Object.freeze({ pptx, report });
}
