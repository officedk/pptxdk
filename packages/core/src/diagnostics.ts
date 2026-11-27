import type {
  BuildDiagnostic,
  BuildDiagnosticFontUnavailableReason,
} from "#/stage-options.ts";

function freezeDiagnostic(diagnostic: BuildDiagnostic): BuildDiagnostic {
  const copy: Record<string, unknown> = { ...diagnostic };
  delete copy["severity"];
  if ("font" in diagnostic) {
    copy["font"] = Object.freeze({ ...diagnostic.font });
  }
  if ("characters" in diagnostic) {
    copy["characters"] = Object.freeze([...diagnostic.characters]);
  }
  return Object.freeze(copy) as BuildDiagnostic;
}

export class BuildError extends Error {
  readonly diagnostics: readonly BuildDiagnostic[];

  constructor(diagnostics: readonly BuildDiagnostic[]) {
    super(
      diagnostics
        .map((diagnostic) => `${diagnostic.path}: ${diagnostic.message}`)
        .join("\n"),
    );
    this.name = "BuildError";
    this.diagnostics = Object.freeze(diagnostics.map(freezeDiagnostic));
  }
}

export class CapabilityError extends Error {
  readonly code: string;
  readonly path: string | undefined;
  readonly details: Readonly<Record<string, unknown>> | undefined;

  /** Subpath relative to the native object triggering the capability, facilitating localization of cell declarations inside the object. */
  constructor(
    code: string,
    message: string,
    path?: string,
    details?: Readonly<Record<string, unknown>>,
  ) {
    super(message);
    this.name = "CapabilityError";
    this.code = code;
    this.path = path;
    this.details =
      details === undefined ? undefined : Object.freeze({ ...details });
  }
}

function readOptionalString(
  details: Readonly<Record<string, unknown>> | undefined,
  key: string,
): string | undefined {
  const val = details?.[key];
  return typeof val === "string" && val !== "" ? val : undefined;
}

function readFontDescriptor(
  details: Readonly<Record<string, unknown>> | undefined,
): {
  readonly family: string;
  readonly weight?: number;
  readonly style?: "normal" | "italic";
} {
  const rawFont =
    typeof details?.["font"] === "object" && details["font"] !== null
      ? (details["font"] as Record<string, unknown>)
      : details;
  const family =
    typeof rawFont?.["family"] === "string" && rawFont["family"] !== ""
      ? rawFont["family"]
      : "unknown";
  const weight =
    typeof rawFont?.["weight"] === "number" &&
    Number.isFinite(rawFont["weight"])
      ? rawFont["weight"]
      : undefined;
  const style =
    rawFont?.["style"] === "normal" || rawFont?.["style"] === "italic"
      ? rawFont["style"]
      : undefined;
  return {
    family,
    ...(weight === undefined ? {} : { weight }),
    ...(style === undefined ? {} : { style }),
  };
}

function extractUnknownPropName(
  details: Readonly<Record<string, unknown>> | undefined,
  message: string,
): string {
  const explicit = readOptionalString(details, "prop");
  if (explicit !== undefined) return explicit;
  const match =
    /(?:unknown|unsupported|unimplemented) (?:parameter|property|style(?: property)?)\s+["']?([A-Za-z0-9_.-]+)["']?/iu.exec(
      message,
    ) ??
    /does not support (?:style |property )?["']?([A-Za-z0-9_.-]+)["']?/iu.exec(
      message,
    );
  return match?.[1] ?? "unknown";
}

function inferStyleField(
  code: string,
  details: Readonly<Record<string, unknown>> | undefined,
  message: string,
): string | undefined {
  const explicit = readOptionalString(details, "field");
  if (explicit !== undefined) return explicit;
  switch (code) {
    case "invalid-font-size":
      return "style.fontSize";
    case "invalid-letter-spacing":
      return "style.letterSpacing";
    case "invalid-paragraph-spacing":
      return "style.spacingAfter";
    case "invalid-list":
      return "style.list";
    case "invalid-text-alignment":
      return "style.textAlign";
    case "invalid-text-decoration":
      return "style.textDecoration";
    case "invalid-text-lang":
      return "style.lang";
    case "invalid-text-transform":
      return "style.textTransform";
    case "invalid-text-wrap":
      return "style.whiteSpace";
    case "invalid-vertical-align":
      return "style.verticalAlign";
    case "invalid-shadow":
      return message.includes("textShadow")
        ? "style.textShadow"
        : "style.boxShadow";
    case "invalid-line-end-type":
      return "style.border";
    default: {
      const stylePrefix = /\bstyle\.([A-Za-z0-9_.[\]]+)/u.exec(message);
      if (stylePrefix?.[1] !== undefined) {
        return `style.${stylePrefix[1]}`;
      }
      const keyColon = /^([A-Za-z][A-Za-z0-9_]*):/u.exec(message);
      if (keyColon?.[1] !== undefined && keyColon[1] !== "style") {
        return `style.${keyColon[1]}`;
      }
      if (/^style\b/u.test(message)) {
        return "style";
      }
      return undefined;
    }
  }
}

const styleErrorCodes = new Set([
  "invalid-style",
  "invalid-visual-style",
  "invalid-layout-style",
  "unsupported-layout-style",
  "unsupported-layout-value",
  "unsupported-style-combination",
  "unsupported-grid-alignment",
  "unsupported-grid-positioning",
  "invalid-grid-placement",
  "invalid-grid-tracks",
  "invalid-color",
  "invalid-shadow",
  "invalid-font-size",
  "invalid-letter-spacing",
  "invalid-paragraph-spacing",
  "invalid-list",
  "invalid-text-alignment",
  "invalid-text-decoration",
  "invalid-text-lang",
  "invalid-text-length",
  "invalid-text-transform",
  "invalid-text-wrap",
  "invalid-vertical-align",
  "invalid-line-end-type",
  "invalid-table-border",
  "invalid-table-padding",
]);

const structureErrorCodes = new Set([
  "invalid-structure",
  "document-root",
  "missing-slide",
  "unexpected-text",
  "unexpected-section-text",
  "mixed-sections",
  "nested-section",
  "unexpected-declaration",
  "unsupported-children",
  "invalid-text-child",
  "unsupported-text-child",
  "conflicting-text-content",
  "empty-group",
  "empty-group-projection",
  "member-without-geometry",
  "empty-slide-master",
  "placeholder-inside-container",
  "placeholder-requires-layout",
  "incompatible-placeholder-content",
  "duplicate-notes",
  "invalid-layout-context",
  "layout-unavailable",
  "unsupported-baseline-context",
  "baseline-measurement-unavailable",
  "invalid-table-structure",
  "conflicting-table-merge",
  "link-needs-object",
  "motion-target-geometry",
  "cross-step-motion",
  "motion-cycle",
  "motion-plan-conflict",
  "no-natural-completion",
  "conflicting-presentation-attribute",
  "conflicting-presentation-item",
  "conflicting-presentation-record",
  "invalid-presentation-contribution",
  "conflicting-resource-type",
  "conflicting-shared-part",
]);

const fontUnavailableReasonByLegacyCode: Readonly<
  Record<string, BuildDiagnosticFontUnavailableReason>
> = {
  "font-measurement-unavailable": "missing-resource",
  "font-embedding-resource-missing": "missing-resource",
  "font-face-mismatch": "style-mismatch",
  "font-family-mismatch": "family-mismatch",
  "font-style-mismatch": "style-mismatch",
  "font-weight-mismatch": "style-mismatch",
  "font-substitution-required": "substitute-required",
  "font-substitute-not-embeddable": "not-embeddable",
  "font-not-embeddable": "not-embeddable",
  "font-format-unsupported": "unsupported-format",
  "unsupported-font-format": "unsupported-format",
  "font-embedding-bitmap-only": "unsupported-format",
  "font-metrics-invalid": "invalid-metrics",
  "invalid-font-permissions": "license-restricted",
  "font-embedding-restricted": "license-restricted",
  "font-embedding-preview-print": "license-restricted",
  "font-subsetting-forbidden": "subsetting-forbidden",
  "font-subsetting-failed": "subsetting-failed",
};

const validFontReasons = new Set<string>(
  Object.values(fontUnavailableReasonByLegacyCode).concat(["missing-face"]),
);

/**
 * Normalizes a direct or capability-originated error code and optional metadata
 * into a schema-valid `BuildDiagnostic` union member.
 */
export function toBuildDiagnostic(
  path: string,
  code: string,
  message: string,
  details?: Readonly<Record<string, unknown>>,
): BuildDiagnostic {
  const field = readOptionalString(details, "field");
  const expected = readOptionalString(details, "expected");
  const conflictPath = readOptionalString(details, "conflictPath");

  // 1. invalid-options
  if (
    code === "invalid-options" ||
    code === "invalid-target" ||
    code === "unsupported-format" ||
    code === "invalid-capability-selection"
  ) {
    return {
      code: "invalid-options",
      path,
      message,
      ...(field === undefined ? {} : { field }),
      ...(expected === undefined ? {} : { expected }),
    };
  }

  // 2. unknown-prop
  if (
    code === "unknown-prop" ||
    code === "unknown-parameter" ||
    code === "unknown-section-property" ||
    code === "unknown-style" ||
    code === "unknown-style-property"
  ) {
    const prop = extractUnknownPropName(details, message);
    const rawScope = details?.["scope"];
    const scope: "props" | "style" | "options" =
      rawScope === "props" || rawScope === "style" || rawScope === "options"
        ? rawScope
        : code === "unknown-style" || code === "unknown-style-property"
          ? "style"
          : path === "resolve" ||
              path === "build" ||
              path.startsWith("build.") ||
              path.startsWith("resolve.")
            ? "options"
            : "props";
    const resolvedField =
      field ??
      (prop === "unknown"
        ? undefined
        : scope === "style" && !prop.startsWith("style.")
          ? `style.${prop}`
          : prop);
    return {
      code: "unknown-prop",
      path,
      message,
      prop,
      scope,
      ...(resolvedField === undefined ? {} : { field: resolvedField }),
    };
  }

  // 3. font-unavailable
  const mappedFontReason = fontUnavailableReasonByLegacyCode[code];
  if (code === "font-unavailable" || mappedFontReason !== undefined) {
    const rawReason = readOptionalString(details, "reason");
    const reason: BuildDiagnosticFontUnavailableReason =
      rawReason !== undefined && validFontReasons.has(rawReason)
        ? (rawReason as BuildDiagnosticFontUnavailableReason)
        : (mappedFontReason ?? "missing-resource");
    return {
      code: "font-unavailable",
      path,
      message,
      font: readFontDescriptor(details),
      reason,
      ...(field === undefined ? {} : { field }),
    };
  }

  // 4. font-missing-glyph
  if (code === "font-missing-glyph") {
    const descriptor = readFontDescriptor(details);
    const rawChars = details?.["characters"];
    const characters = Array.isArray(rawChars)
      ? rawChars.filter((item): item is string => typeof item === "string")
      : [];
    return {
      code: "font-missing-glyph",
      path,
      message,
      font: {
        family: descriptor.family,
        weight: descriptor.weight ?? 400,
        style: descriptor.style ?? "normal",
      },
      characters,
      ...(field === undefined ? {} : { field }),
    };
  }

  // 5. duplicate-id
  if (
    code === "duplicate-id" ||
    code === "duplicate-author-id" ||
    code === "duplicate-morph-key" ||
    code === "duplicate-slide-placeholder" ||
    code === "duplicate-scope" ||
    code === "duplicate-cue" ||
    code === "duplicate-part" ||
    code === "conflicting-resource"
  ) {
    const identifier =
      readOptionalString(details, "identifier") ??
      /"([^"]+)"/u.exec(message)?.[1];
    const resolvedField =
      field ??
      (code === "duplicate-author-id"
        ? "id"
        : code === "duplicate-morph-key"
          ? "morphKey"
          : code === "duplicate-slide-placeholder"
            ? "placeholder"
            : undefined);
    return {
      code: "duplicate-id",
      path,
      message,
      ...(resolvedField === undefined ? {} : { field: resolvedField }),
      ...(identifier === undefined ? {} : { identifier }),
      ...(conflictPath === undefined ? {} : { conflictPath }),
    };
  }

  // 6. broken-reference
  if (
    code === "broken-reference" ||
    code === "missing-link-target" ||
    code === "ambiguous-slide-layout" ||
    code === "foreign-slide-layout" ||
    code === "unknown-slide-placeholder" ||
    code === "missing-placeholder-binding" ||
    code === "unknown-chart-axis" ||
    code === "missing-motion-target" ||
    code === "missing-cue" ||
    code === "missing-scope" ||
    code === "unknown-motion-start" ||
    code === "missing-part-target" ||
    code === "missing-resource-claim"
  ) {
    const reference =
      readOptionalString(details, "reference") ??
      (code === "unknown-slide-placeholder"
        ? /does not define placeholder "([^"]+)"/u.exec(message)?.[1]
        : undefined) ??
      /"([^"]+)"/u.exec(message)?.[1];
    const rawTargetKind = details?.["targetKind"];
    const targetKind =
      rawTargetKind === "slide" ||
      rawTargetKind === "layout" ||
      rawTargetKind === "placeholder" ||
      rawTargetKind === "object" ||
      rawTargetKind === "cue" ||
      rawTargetKind === "scope" ||
      rawTargetKind === "part" ||
      rawTargetKind === "resource" ||
      rawTargetKind === "axis"
        ? rawTargetKind
        : code === "missing-link-target"
          ? "slide"
          : code === "ambiguous-slide-layout" || code === "foreign-slide-layout"
            ? "layout"
            : code === "unknown-slide-placeholder" ||
                code === "missing-placeholder-binding"
              ? "placeholder"
              : code === "unknown-chart-axis"
                ? "axis"
                : code === "missing-motion-target"
                  ? "object"
                  : code === "missing-cue" || code === "unknown-motion-start"
                    ? "cue"
                    : code === "missing-scope"
                      ? "scope"
                      : code === "missing-part-target"
                        ? "part"
                        : code === "missing-resource-claim"
                          ? "resource"
                          : undefined;
    const resolvedField =
      field ??
      (code === "missing-link-target"
        ? "link"
        : code === "ambiguous-slide-layout" || code === "foreign-slide-layout"
          ? "layout"
          : code === "unknown-slide-placeholder" ||
              code === "missing-placeholder-binding"
            ? "placeholder"
            : undefined);
    return {
      code: "broken-reference",
      path,
      message,
      ...(resolvedField === undefined ? {} : { field: resolvedField }),
      ...(reference === undefined ? {} : { reference }),
      ...(targetKind === undefined ? {} : { targetKind }),
    };
  }

  // 7. invalid-continuity
  if (
    code === "invalid-continuity" ||
    code === "morph-inside-group" ||
    code === "morph-on-master" ||
    code === "morph-needs-object" ||
    code === "unsupported-morph-object" ||
    code === "morph-object-kind-mismatch" ||
    code === "morph-object-name-conflict"
  ) {
    const rawReason = details?.["reason"];
    const reason =
      rawReason === "inside-group" ||
      rawReason === "on-master" ||
      rawReason === "needs-object" ||
      rawReason === "unsupported-object" ||
      rawReason === "kind-mismatch" ||
      rawReason === "name-conflict"
        ? rawReason
        : code === "morph-inside-group"
          ? "inside-group"
          : code === "morph-on-master"
            ? "on-master"
            : code === "morph-needs-object"
              ? "needs-object"
              : code === "unsupported-morph-object"
                ? "unsupported-object"
                : code === "morph-object-kind-mismatch"
                  ? "kind-mismatch"
                  : code === "morph-object-name-conflict"
                    ? "name-conflict"
                    : undefined;
    const resolvedField = field ?? "morphKey";
    return {
      code: "invalid-continuity",
      path,
      message,
      field: resolvedField,
      ...(reason === undefined ? {} : { reason }),
      ...(conflictPath === undefined ? {} : { conflictPath }),
    };
  }

  // 8. missing-capability
  if (
    code === "missing-capability" ||
    code === "unsupported-declaration" ||
    code === "unknown-capability" ||
    code === "duplicate-capability" ||
    code === "invalid-capability-module" ||
    code === "invalid-capability-binding"
  ) {
    const capability =
      readOptionalString(details, "capability") ??
      /registered for ([A-Za-z0-9_.-]+)/u.exec(message)?.[1] ??
      /^([A-Za-z0-9_.-]+) declarations are only supported/u.exec(
        message,
      )?.[1] ??
      "unknown";
    return {
      code: "missing-capability",
      path,
      message,
      capability,
      ...(field === undefined ? {} : { field }),
    };
  }

  // 9. invalid-resource
  if (
    code === "invalid-resource" ||
    code === "invalid-image-source" ||
    code === "unsupported-image-format" ||
    code === "image-read-failed" ||
    code === "image-resource-failure" ||
    code === "invalid-part" ||
    code === "resource-contribution-failed" ||
    code === "capability-failure"
  ) {
    const rawKind = details?.["resourceKind"];
    const resourceKind: "image" | "part" | "capability" =
      rawKind === "image" || rawKind === "part" || rawKind === "capability"
        ? rawKind
        : code === "invalid-image-source" ||
            code === "unsupported-image-format" ||
            code === "image-read-failed" ||
            code === "image-resource-failure"
          ? "image"
          : code === "capability-failure"
            ? "capability"
            : "part";
    return {
      code: "invalid-resource",
      path,
      message,
      resourceKind,
      ...(field === undefined ? {} : { field }),
    };
  }

  // 10. target-violation
  if (
    code === "target-violation" ||
    code === "forbidden-mechanism" ||
    code === "unsupported-extension" ||
    code === "duplicate-extension" ||
    code === "unsupported-font-embedding"
  ) {
    const target = readOptionalString(details, "target") ?? "transitional";
    const extension = readOptionalString(details, "extension");
    const mechanism = readOptionalString(details, "mechanism");
    return {
      code: "target-violation",
      path,
      message,
      target,
      ...(field === undefined ? {} : { field }),
      ...(extension === undefined ? {} : { extension }),
      ...(mechanism === undefined ? {} : { mechanism }),
    };
  }

  // 11. render-failed
  if (
    code === "render-failed" ||
    code === "react-render" ||
    code === "append"
  ) {
    const rawPhase = details?.["phase"];
    const phase: "expand" | "append" =
      rawPhase === "expand" || rawPhase === "append"
        ? rawPhase
        : code === "append"
          ? "append"
          : "expand";
    return {
      code: "render-failed",
      path,
      message,
      phase,
      ...(field === undefined ? {} : { field }),
    };
  }

  // 12. invalid-style
  if (styleErrorCodes.has(code)) {
    const styleField = inferStyleField(code, details, message);
    return {
      code: "invalid-style",
      path,
      message,
      ...(styleField === undefined ? {} : { field: styleField }),
      ...(expected === undefined ? {} : { expected }),
    };
  }

  // 13. invalid-structure
  if (structureErrorCodes.has(code)) {
    return {
      code: "invalid-structure",
      path,
      message,
      ...(field === undefined ? {} : { field }),
      ...(expected === undefined ? {} : { expected }),
      ...(conflictPath === undefined ? {} : { conflictPath }),
    };
  }

  // 14. invalid-prop (covers all remaining prop/author parameter errors)
  const resolvedField =
    field ??
    (code === "invalid-author-id" || code === "invalid-id"
      ? "id"
      : code === "invalid-link"
        ? "link"
        : code === "invalid-morph-key"
          ? "morphKey"
          : code === "invalid-notes"
            ? "notes"
            : undefined);
  return {
    code: "invalid-prop",
    path,
    message,
    ...(resolvedField === undefined ? {} : { field: resolvedField }),
    ...(expected === undefined ? {} : { expected }),
  };
}

export function capabilityErrorToDiagnostic(
  error: CapabilityError,
  basePath: string,
  relativeToBase = false,
): BuildDiagnostic {
  const path =
    error.path === undefined || error.path === ""
      ? basePath
      : !relativeToBase || error.path.startsWith("deck/")
        ? error.path
        : `${basePath}/${error.path}`;
  return toBuildDiagnostic(path, error.code, error.message, error.details);
}

/** Converts a Zod validation issue into a structured `BuildDiagnostic`, mapping unrecognized keys to `unknown-prop`. */
export function zodIssueToBuildDiagnostic(
  path: string,
  fallbackCode: string,
  fallbackMessage: string,
  issue:
    | {
        readonly code: string;
        readonly message: string;
        readonly path: readonly PropertyKey[];
        readonly keys?: readonly string[];
      }
    | undefined,
  pathPrefix?: string,
): BuildDiagnostic {
  const pathSegments = [
    ...(pathPrefix === undefined ? [] : [pathPrefix]),
    ...(issue === undefined ? [] : issue.path.map(String)),
  ];
  if (
    issue?.code === "unrecognized_keys" &&
    issue.keys !== undefined &&
    issue.keys[0] !== undefined
  ) {
    const prop = issue.keys[0];
    const inStyle = pathSegments.includes("style");
    const field = [...pathSegments, prop].join(".");
    return toBuildDiagnostic(
      path,
      inStyle ? "unknown-style" : "unknown-parameter",
      issue.message,
      {
        prop,
        field,
        scope: inStyle ? "style" : "props",
      },
    );
  }
  const field = pathSegments.length > 0 ? pathSegments.join(".") : undefined;
  return toBuildDiagnostic(
    path,
    field?.startsWith("style") ? "invalid-style" : fallbackCode,
    issue?.message ?? fallbackMessage,
    field === undefined ? undefined : { field },
  );
}
