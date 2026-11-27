import { z } from "zod";

type AssertExact<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2
    ? true
    : never;

/** Source package location attached to import diagnostics and fatal errors. */
export type ImportLocation = {
  part: string;
  xpath?: string | undefined;
  slideIndex?: number | undefined;
  objectId?: number | undefined;
  objectName?: string | undefined;
};

const rawImportLocationSchema = z.object({
  part: z.string().min(1),
  xpath: z.string().min(1).optional(),
  slideIndex: z.number().int().nonnegative().optional(),
  objectId: z.number().int().optional(),
  objectName: z.string().optional(),
});
/** Source package location attached to import diagnostics and fatal errors. */
export const importLocationSchema: z.ZodType<ImportLocation> =
  rawImportLocationSchema satisfies z.ZodType<
    AssertExact<
      ImportLocation,
      z.infer<typeof rawImportLocationSchema>
    > extends true
      ? ImportLocation
      : never
  >;

/** Native editing capability lost when a fallback representation is selected. */
export type EditabilityLoss = {
  lost:
    | "native-mechanism"
    | "master-layout-inheritance"
    | "preferred-compatibility-branch"
    | "custom-print-layout";
  retainedAs: string;
};

const rawEditabilityLossSchema = z.object({
  lost: z.enum([
    "native-mechanism",
    "master-layout-inheritance",
    "preferred-compatibility-branch",
    "custom-print-layout",
  ]),
  retainedAs: z.string().min(1),
});
/** Native editing capability lost when a fallback representation is selected. */
export const editabilityLossSchema: z.ZodType<EditabilityLoss> =
  rawEditabilityLossSchema satisfies z.ZodType<
    AssertExact<
      EditabilityLoss,
      z.infer<typeof rawEditabilityLossSchema>
    > extends true
      ? EditabilityLoss
      : never
  >;

/** Dependent item omitted because an ancestor or target part could not be imported. */
export type CascadedOmission = {
  location: ImportLocation;
  role: string;
};

const rawCascadedOmissionSchema = z.object({
  location: rawImportLocationSchema,
  role: z.string().min(1),
});
/** Dependent item omitted because an ancestor or target part could not be imported. */
export const cascadedOmissionSchema: z.ZodType<CascadedOmission> =
  rawCascadedOmissionSchema satisfies z.ZodType<
    AssertExact<
      CascadedOmission,
      z.infer<typeof rawCascadedOmissionSchema>
    > extends true
      ? CascadedOmission
      : never
  >;

/** Fatal package/entry failure reason carried by `ImportError`. */
export type FatalImportIssue =
  | {
      code: "corrupt-archive";
      category: "corrupt-input";
      disposition: "aborted";
    }
  | {
      code: "missing-entry";
      category: "corrupt-input";
      disposition: "aborted";
      location: ImportLocation;
      expectedRelationshipType?: string | undefined;
    }
  | {
      code: "malformed-xml";
      category: "corrupt-input";
      disposition: "aborted";
      location: ImportLocation;
      detail?: string | undefined;
    }
  | {
      code: "broken-entry-relationship";
      category: "corrupt-input";
      disposition: "aborted";
      location: ImportLocation;
      relationshipId?: string | undefined;
      target?: string | undefined;
    }
  | {
      code: "invalid-presentation-root";
      category: "corrupt-input";
      disposition: "aborted";
      location: ImportLocation;
      namespaceUri?: string | undefined;
      localName?: string | undefined;
    }
  | {
      code: "unsupported-format-baseline";
      category: "unsupported";
      disposition: "aborted";
      baseline: "ISO/IEC 29500 Strict" | "unknown";
      detectedNamespace: string;
      location: ImportLocation;
    };

const rawFatalImportIssueSchema = z.discriminatedUnion("code", [
  z.object({
    code: z.literal("corrupt-archive"),
    category: z.literal("corrupt-input"),
    disposition: z.literal("aborted"),
  }),
  z.object({
    code: z.literal("missing-entry"),
    category: z.literal("corrupt-input"),
    disposition: z.literal("aborted"),
    location: rawImportLocationSchema,
    expectedRelationshipType: z.string().optional(),
  }),
  z.object({
    code: z.literal("malformed-xml"),
    category: z.literal("corrupt-input"),
    disposition: z.literal("aborted"),
    location: rawImportLocationSchema,
    detail: z.string().optional(),
  }),
  z.object({
    code: z.literal("broken-entry-relationship"),
    category: z.literal("corrupt-input"),
    disposition: z.literal("aborted"),
    location: rawImportLocationSchema,
    relationshipId: z.string().optional(),
    target: z.string().optional(),
  }),
  z.object({
    code: z.literal("invalid-presentation-root"),
    category: z.literal("corrupt-input"),
    disposition: z.literal("aborted"),
    location: rawImportLocationSchema,
    namespaceUri: z.string().optional(),
    localName: z.string().optional(),
  }),
  z.object({
    code: z.literal("unsupported-format-baseline"),
    category: z.literal("unsupported"),
    disposition: z.literal("aborted"),
    baseline: z.enum(["ISO/IEC 29500 Strict", "unknown"]),
    detectedNamespace: z.string().min(1),
    location: rawImportLocationSchema,
  }),
]);
/** Fatal package/entry failure reason carried by `ImportError`. */
export const fatalImportIssueSchema: z.ZodType<FatalImportIssue> =
  rawFatalImportIssueSchema satisfies z.ZodType<
    AssertExact<
      FatalImportIssue,
      z.infer<typeof rawFatalImportIssueSchema>
    > extends true
      ? FatalImportIssue
      : never
  >;

/** Recoverable or isolatable diagnostic event emitted in `ImportReport.diagnostics`. */
export type ImportDiagnostic =
  | {
      code: "corrupt-part";
      category: "corrupt-input";
      disposition: "omitted";
      location: ImportLocation;
      cascaded?: readonly CascadedOmission[] | undefined;
    }
  | {
      code: "missing-target-part";
      category: "corrupt-input";
      disposition: "omitted";
      location: ImportLocation;
      relationshipId: string;
      targetPart: string;
      cascaded?: readonly CascadedOmission[] | undefined;
    }
  | {
      code: "broken-relationship";
      category: "corrupt-input";
      disposition: "omitted";
      location: ImportLocation;
      relationshipId: string;
      cascaded?: readonly CascadedOmission[] | undefined;
    }
  | {
      code: "invalid-xml-value";
      category: "corrupt-input";
      disposition: "omitted" | "recovered";
      location: ImportLocation;
      attributeOrElement: string;
      rawValue?: string | undefined;
    }
  | {
      code: "unsupported-mechanism";
      category: "unsupported";
      disposition: "omitted";
      location: ImportLocation;
      namespaceUri: string;
      localName: string;
      mechanismHint?: string | undefined;
      loss?: EditabilityLoss | undefined;
      cascaded?: readonly CascadedOmission[] | undefined;
    }
  | {
      code: "compatibility-fallback";
      category: "unsupported";
      disposition: "fallback";
      location: ImportLocation;
      preferredMechanism: string;
      selectedFallback: string;
      loss: EditabilityLoss;
    }
  | {
      code: "unconsumed-xml";
      category: "unsupported";
      disposition: "omitted";
      kind: "attribute" | "element";
      location: ImportLocation;
      namespaceUri: string;
      localName: string;
    };

const rawImportDiagnosticSchema = z.discriminatedUnion("code", [
  z.object({
    code: z.literal("corrupt-part"),
    category: z.literal("corrupt-input"),
    disposition: z.literal("omitted"),
    location: rawImportLocationSchema,
    cascaded: z.array(rawCascadedOmissionSchema).readonly().optional(),
  }),
  z.object({
    code: z.literal("missing-target-part"),
    category: z.literal("corrupt-input"),
    disposition: z.literal("omitted"),
    location: rawImportLocationSchema,
    relationshipId: z.string(),
    targetPart: z.string(),
    cascaded: z.array(rawCascadedOmissionSchema).readonly().optional(),
  }),
  z.object({
    code: z.literal("broken-relationship"),
    category: z.literal("corrupt-input"),
    disposition: z.literal("omitted"),
    location: rawImportLocationSchema,
    relationshipId: z.string(),
    cascaded: z.array(rawCascadedOmissionSchema).readonly().optional(),
  }),
  z.object({
    code: z.literal("invalid-xml-value"),
    category: z.literal("corrupt-input"),
    disposition: z.enum(["omitted", "recovered"]),
    location: rawImportLocationSchema,
    attributeOrElement: z.string(),
    rawValue: z.string().optional(),
  }),
  z.object({
    code: z.literal("unsupported-mechanism"),
    category: z.literal("unsupported"),
    disposition: z.literal("omitted"),
    location: rawImportLocationSchema,
    namespaceUri: z.string(),
    localName: z.string(),
    mechanismHint: z.string().optional(),
    loss: rawEditabilityLossSchema.optional(),
    cascaded: z.array(rawCascadedOmissionSchema).readonly().optional(),
  }),
  z.object({
    code: z.literal("compatibility-fallback"),
    category: z.literal("unsupported"),
    disposition: z.literal("fallback"),
    location: rawImportLocationSchema,
    preferredMechanism: z.string(),
    selectedFallback: z.string(),
    loss: rawEditabilityLossSchema,
  }),
  z.object({
    code: z.literal("unconsumed-xml"),
    category: z.literal("unsupported"),
    disposition: z.literal("omitted"),
    kind: z.enum(["attribute", "element"]),
    location: rawImportLocationSchema,
    namespaceUri: z.string(),
    localName: z.string(),
  }),
]);
/** Recoverable or isolatable diagnostic event emitted in `ImportReport.diagnostics`. */
export const importDiagnosticSchema: z.ZodType<ImportDiagnostic> =
  rawImportDiagnosticSchema satisfies z.ZodType<
    AssertExact<
      ImportDiagnostic,
      z.infer<typeof rawImportDiagnosticSchema>
    > extends true
      ? ImportDiagnostic
      : never
  >;

/** Structured summary and per-item diagnostics produced by `createImporter`. */
export type ImportReport = {
  evidence: "import";
  sourceFormat: "ECMA-376 5th edition Transitional";
  slideCount: number;
  diagnostics: readonly ImportDiagnostic[];
};

const rawImportReportSchema = z.object({
  evidence: z.literal("import"),
  sourceFormat: z.literal("ECMA-376 5th edition Transitional"),
  slideCount: z.number().int().nonnegative(),
  diagnostics: z.array(rawImportDiagnosticSchema).readonly(),
});
/** Structured summary and per-item diagnostics produced by `createImporter`. */
export const importReportSchema: z.ZodType<ImportReport> =
  rawImportReportSchema satisfies z.ZodType<
    AssertExact<
      ImportReport,
      z.infer<typeof rawImportReportSchema>
    > extends true
      ? ImportReport
      : never
  >;

export type ImportOptions = {
  readonly signal?: AbortSignal | undefined;
};

export const importOptionsSchema: z.ZodType<ImportOptions> = z
  .object({
    signal: z
      .custom<AbortSignal>(
        (value) =>
          typeof value === "object" &&
          value !== null &&
          "aborted" in value &&
          typeof (value as { readonly aborted?: unknown }).aborted ===
            "boolean",
      )
      .optional(),
  })
  .strict();

/** Error thrown when PPTX entry points cannot be opened or use an unsupported baseline. */
export class ImportError extends Error {
  readonly issue: FatalImportIssue;

  constructor(issue: FatalImportIssue, message?: string) {
    super(message ?? issue.code);
    this.name = "ImportError";
    this.issue = Object.freeze(fatalImportIssueSchema.parse(issue));
  }
}

export type UnconsumedXmlItem = {
  readonly kind: "attribute" | "element";
  readonly location: ImportLocation;
  readonly namespaceUri: string;
  readonly localName: string;
};
