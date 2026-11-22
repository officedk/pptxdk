import { posix } from "node:path";

import { strFromU8, unzipSync } from "fflate";
import { z } from "zod";

import { ImportError, type ImportDiagnostic } from "#/diagnostics.ts";
import {
  OPC_REL_NS,
  P_NS,
  REL_OFFICE_DOCUMENT,
  STRICT_NS_PREFIX,
} from "#/namespaces.ts";
import {
  MalformedXmlPartError,
  XmlElementCursor,
  parseXmlPart,
} from "#/xml-reader.ts";

/** Zod schema for an OPC `<Relationship>` element attributes. */
export const opcRelationshipSchema = z.object({
  Id: z.string().min(1),
  Type: z.string().min(1),
  Target: z.string().min(1),
  TargetMode: z.enum(["Internal", "External"]).default("Internal"),
});
export type OpcRelationship = z.infer<typeof opcRelationshipSchema>;

export type ParsedRelationshipEntry = OpcRelationship & {
  readonly xpath: string;
};

/** Computes the OPC relationships part path for a given source part (or `""` for package root). */
export function relsPathFor(partPath: string): string {
  if (partPath === "") return "_rels/.rels";
  const dir = posix.dirname(partPath);
  const base = posix.basename(partPath);
  return dir === "." ? `_rels/${base}.rels` : `${dir}/_rels/${base}.rels`;
}

/** Resolves a relative or root-anchored OPC target against its owning part path. */
export function resolvePartPath(
  sourcePartPath: string,
  target: string,
): string | undefined {
  const normalizedTarget = target.replace(/\\/gu, "/");
  if (normalizedTarget === "") return undefined;
  const baseDir =
    sourcePartPath === "" || sourcePartPath === "_rels/.rels"
      ? ""
      : posix.dirname(sourcePartPath);
  const joined = normalizedTarget.startsWith("/")
    ? normalizedTarget.slice(1)
    : baseDir === "" || baseDir === "."
      ? normalizedTarget
      : posix.join(baseDir, normalizedTarget);
  const normalized = posix.normalize(joined);
  if (
    normalized === "." ||
    normalized === ".." ||
    normalized.startsWith("../") ||
    normalized.startsWith("/")
  ) {
    return undefined;
  }
  return normalized;
}

/** Parses and validates all `<Relationship>` entries in an OPC `.rels` XML part. */
export function parseRelationshipsXml(
  relsPartPath: string,
  xmlText: string,
): {
  readonly entries: readonly ParsedRelationshipEntry[];
  readonly diagnostics: readonly ImportDiagnostic[];
} {
  const root = parseXmlPart(relsPartPath, xmlText);
  if (
    (root.namespaceURI ?? "") !== OPC_REL_NS ||
    (root.localName ?? root.tagName) !== "Relationships"
  ) {
    throw new MalformedXmlPartError(
      relsPartPath,
      `Expected OPC Relationships root, got ${root.namespaceURI ?? ""}:${root.localName ?? root.tagName}`,
    );
  }
  const cursor = XmlElementCursor.createRoot(
    root,
    relsPartPath,
    "/Relationships",
  );
  const entries: ParsedRelationshipEntry[] = [];
  const diagnostics: ImportDiagnostic[] = [];
  for (const relCursor of cursor.children(OPC_REL_NS, "Relationship")) {
    const raw = relCursor.attrsObject(["Id", "Type", "Target", "TargetMode"]);
    const parsed = opcRelationshipSchema.safeParse(raw);
    if (!parsed.success) {
      diagnostics.push({
        code: "broken-relationship",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part: relsPartPath, xpath: relCursor.xpath },
        relationshipId: raw["Id"] ?? "",
      });
      continue;
    }
    entries.push({
      ...parsed.data,
      xpath: relCursor.xpath,
    });
  }
  return { entries, diagnostics };
}

export type PptxPackage = {
  readonly presentationPath: string;
  readonly presentationRoot: XmlElementCursor;
  readonly presentationRels: readonly ParsedRelationshipEntry[];
  readonly relDiagnostics: readonly ImportDiagnostic[];
  hasPart(partPath: string): boolean;
  readTextPart(partPath: string): string | undefined;
  readDetachedBytes(partPath: string): Uint8Array | undefined;
  readRelationships(partPath: string): {
    readonly entries: readonly ParsedRelationshipEntry[];
    readonly diagnostics: readonly ImportDiagnostic[];
  };
};

function detectStrictBaselineIssue(
  part: string,
  entries: readonly ParsedRelationshipEntry[],
): ImportError | undefined {
  for (const entry of entries) {
    if (entry.Type.startsWith(STRICT_NS_PREFIX)) {
      return new ImportError(
        {
          code: "unsupported-format-baseline",
          category: "unsupported",
          disposition: "aborted",
          baseline: "ISO/IEC 29500 Strict",
          detectedNamespace: entry.Type,
          location: { part, xpath: entry.xpath },
        },
        `Unsupported ISO/IEC 29500 Strict OOXML relationship (${entry.Type}) in ${part}`,
      );
    }
  }
  return undefined;
}

/**
 * Unzips a PPTX archive into detached entries, validates OPC package entry points,
 * rejects ISO/IEC 29500 Strict OOXML with structured diagnostics, and loads `presentation.xml`.
 */
export function openPptxPackage(input: Uint8Array): PptxPackage {
  let rawEntries: Record<string, Uint8Array>;
  try {
    rawEntries = unzipSync(input);
  } catch (error) {
    throw new ImportError(
      {
        code: "corrupt-archive",
        category: "corrupt-input",
        disposition: "aborted",
      },
      `Failed to unzip PPTX archive: ${error instanceof Error ? error.message : String(error)}`,
    );
  }

  // Detach every entry from any shared buffer slice immediately so callers may mutate or zero input.
  const entries = new Map<string, Uint8Array>();
  for (const [name, bytes] of Object.entries(rawEntries)) {
    entries.set(name.replace(/^\/+/u, ""), new Uint8Array(bytes));
  }

  const rootRelsPath = "_rels/.rels";
  const rootRelsBytes = entries.get(rootRelsPath);
  if (rootRelsBytes === undefined) {
    throw new ImportError(
      {
        code: "missing-entry",
        category: "corrupt-input",
        disposition: "aborted",
        location: { part: rootRelsPath },
        expectedRelationshipType: REL_OFFICE_DOCUMENT,
      },
      "Missing package entry relationships (_rels/.rels)",
    );
  }

  let rootRels: readonly ParsedRelationshipEntry[];
  try {
    rootRels = parseRelationshipsXml(
      rootRelsPath,
      strFromU8(rootRelsBytes),
    ).entries;
  } catch (error) {
    throw new ImportError(
      {
        code: "malformed-xml",
        category: "corrupt-input",
        disposition: "aborted",
        location: { part: rootRelsPath },
        ...(error instanceof Error ? { detail: error.message } : {}),
      },
      `Malformed XML in ${rootRelsPath}`,
    );
  }

  const strictFromRels = detectStrictBaselineIssue(rootRelsPath, rootRels);
  if (strictFromRels !== undefined) {
    throw strictFromRels;
  }

  const officeDocRel = rootRels.find(
    (entry) => entry.Type === REL_OFFICE_DOCUMENT,
  );
  if (officeDocRel === undefined) {
    throw new ImportError(
      {
        code: "missing-entry",
        category: "corrupt-input",
        disposition: "aborted",
        location: { part: rootRelsPath },
        expectedRelationshipType: REL_OFFICE_DOCUMENT,
      },
      "Missing Transitional officeDocument relationship in _rels/.rels",
    );
  }

  const presentationPath =
    officeDocRel.TargetMode === "External"
      ? undefined
      : resolvePartPath("", officeDocRel.Target);
  if (presentationPath === undefined) {
    throw new ImportError(
      {
        code: "broken-entry-relationship",
        category: "corrupt-input",
        disposition: "aborted",
        location: { part: rootRelsPath, xpath: officeDocRel.xpath },
        relationshipId: officeDocRel.Id,
        target: officeDocRel.Target,
      },
      `Invalid officeDocument relationship target ${officeDocRel.Target}`,
    );
  }

  const presentationBytes = entries.get(presentationPath);
  if (presentationBytes === undefined) {
    throw new ImportError(
      {
        code: "broken-entry-relationship",
        category: "corrupt-input",
        disposition: "aborted",
        location: { part: rootRelsPath, xpath: officeDocRel.xpath },
        relationshipId: officeDocRel.Id,
        target: presentationPath,
      },
      `Missing presentation part ${presentationPath} referenced by ${officeDocRel.Id}`,
    );
  }

  let presElement: ReturnType<typeof parseXmlPart>;
  try {
    presElement = parseXmlPart(presentationPath, strFromU8(presentationBytes));
  } catch (error) {
    throw new ImportError(
      {
        code: "malformed-xml",
        category: "corrupt-input",
        disposition: "aborted",
        location: { part: presentationPath },
        ...(error instanceof Error ? { detail: error.message } : {}),
      },
      `Malformed XML in ${presentationPath}`,
    );
  }

  const presNs = presElement.namespaceURI ?? "";
  const presLocal = presElement.localName ?? presElement.tagName;
  if (presNs.startsWith(STRICT_NS_PREFIX)) {
    throw new ImportError(
      {
        code: "unsupported-format-baseline",
        category: "unsupported",
        disposition: "aborted",
        baseline: "ISO/IEC 29500 Strict",
        detectedNamespace: presNs,
        location: { part: presentationPath, xpath: `/${presLocal}` },
      },
      `Unsupported ISO/IEC 29500 Strict presentation namespace ${presNs}`,
    );
  }
  if (presNs !== P_NS || presLocal !== "presentation") {
    throw new ImportError(
      {
        code: "invalid-presentation-root",
        category: "corrupt-input",
        disposition: "aborted",
        location: { part: presentationPath, xpath: `/${presLocal}` },
        namespaceUri: presNs,
        localName: presLocal,
      },
      `Expected Transitional PresentationML root element, got ${presNs}:${presLocal}`,
    );
  }

  const presentationRoot = XmlElementCursor.createRoot(
    presElement,
    presentationPath,
    "/p:presentation",
  );

  const readRelationships = (partPath: string) => {
    const relsPart = relsPathFor(partPath);
    const bytes = entries.get(relsPart);
    if (bytes === undefined) {
      return { entries: [], diagnostics: [] };
    }
    try {
      return parseRelationshipsXml(relsPart, strFromU8(bytes));
    } catch {
      return {
        entries: [],
        diagnostics: [
          {
            code: "corrupt-part" as const,
            category: "corrupt-input" as const,
            disposition: "omitted" as const,
            location: { part: relsPart },
          },
        ],
      };
    }
  };

  const presRels = readRelationships(presentationPath);

  return {
    presentationPath,
    presentationRoot,
    presentationRels: presRels.entries,
    relDiagnostics: presRels.diagnostics,
    hasPart(partPath: string): boolean {
      return entries.has(partPath);
    },
    readTextPart(partPath: string): string | undefined {
      const bytes = entries.get(partPath);
      return bytes === undefined ? undefined : strFromU8(bytes);
    },
    readDetachedBytes(partPath: string): Uint8Array | undefined {
      const bytes = entries.get(partPath);
      return bytes === undefined ? undefined : new Uint8Array(bytes);
    },
    readRelationships,
  };
}
