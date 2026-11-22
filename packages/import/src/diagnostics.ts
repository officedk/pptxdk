import { z } from "zod";

/** Source package location attached to import diagnostics and fatal errors. */
export const importLocationSchema = z.object({
  part: z.string().min(1),
  xpath: z.string().min(1).optional(),
  slideIndex: z.number().int().nonnegative().optional(),
  objectId: z.number().int().optional(),
  objectName: z.string().optional(),
});
export type ImportLocation = z.infer<typeof importLocationSchema>;

/** Native editing capability lost when a fallback representation is selected. */
export const editabilityLossSchema = z.object({
  lost: z.enum([
    "native-mechanism",
    "master-layout-inheritance",
    "preferred-compatibility-branch",
    "custom-print-layout",
  ]),
  retainedAs: z.string().min(1),
});
export type EditabilityLoss = z.infer<typeof editabilityLossSchema>;

/** Dependent item omitted because an ancestor or target part could not be imported. */
export const cascadedOmissionSchema = z.object({
  location: importLocationSchema,
  role: z.string().min(1),
});
export type CascadedOmission = z.infer<typeof cascadedOmissionSchema>;

/** Fatal package/entry failure reason carried by `ImportError`. */
export const fatalImportIssueSchema = z.discriminatedUnion("code", [
  z.object({
    code: z.literal("corrupt-archive"),
    category: z.literal("corrupt-input"),
    disposition: z.literal("aborted"),
  }),
  z.object({
    code: z.literal("missing-entry"),
    category: z.literal("corrupt-input"),
    disposition: z.literal("aborted"),
    location: importLocationSchema,
    expectedRelationshipType: z.string().optional(),
  }),
  z.object({
    code: z.literal("malformed-xml"),
    category: z.literal("corrupt-input"),
    disposition: z.literal("aborted"),
    location: importLocationSchema,
    detail: z.string().optional(),
  }),
  z.object({
    code: z.literal("broken-entry-relationship"),
    category: z.literal("corrupt-input"),
    disposition: z.literal("aborted"),
    location: importLocationSchema,
    relationshipId: z.string().optional(),
    target: z.string().optional(),
  }),
  z.object({
    code: z.literal("invalid-presentation-root"),
    category: z.literal("corrupt-input"),
    disposition: z.literal("aborted"),
    location: importLocationSchema,
    namespaceUri: z.string().optional(),
    localName: z.string().optional(),
  }),
  z.object({
    code: z.literal("unsupported-format-baseline"),
    category: z.literal("unsupported"),
    disposition: z.literal("aborted"),
    baseline: z.enum(["ISO/IEC 29500 Strict", "unknown"]),
    detectedNamespace: z.string().min(1),
    location: importLocationSchema,
  }),
]);
export type FatalImportIssue = z.infer<typeof fatalImportIssueSchema>;

/** Recoverable or isolatable diagnostic event emitted in `ImportReport.diagnostics`. */
export const importDiagnosticSchema = z.discriminatedUnion("code", [
  z.object({
    code: z.literal("corrupt-part"),
    category: z.literal("corrupt-input"),
    disposition: z.literal("omitted"),
    location: importLocationSchema,
    cascaded: z.array(cascadedOmissionSchema).readonly().optional(),
  }),
  z.object({
    code: z.literal("missing-target-part"),
    category: z.literal("corrupt-input"),
    disposition: z.literal("omitted"),
    location: importLocationSchema,
    relationshipId: z.string(),
    targetPart: z.string(),
    cascaded: z.array(cascadedOmissionSchema).readonly().optional(),
  }),
  z.object({
    code: z.literal("broken-relationship"),
    category: z.literal("corrupt-input"),
    disposition: z.literal("omitted"),
    location: importLocationSchema,
    relationshipId: z.string(),
    cascaded: z.array(cascadedOmissionSchema).readonly().optional(),
  }),
  z.object({
    code: z.literal("invalid-xml-value"),
    category: z.literal("corrupt-input"),
    disposition: z.enum(["omitted", "recovered"]),
    location: importLocationSchema,
    attributeOrElement: z.string(),
    rawValue: z.string().optional(),
  }),
  z.object({
    code: z.literal("unsupported-mechanism"),
    category: z.literal("unsupported"),
    disposition: z.literal("omitted"),
    location: importLocationSchema,
    namespaceUri: z.string(),
    localName: z.string(),
    mechanismHint: z.string().optional(),
    loss: editabilityLossSchema.optional(),
    cascaded: z.array(cascadedOmissionSchema).readonly().optional(),
  }),
  z.object({
    code: z.literal("compatibility-fallback"),
    category: z.literal("unsupported"),
    disposition: z.literal("fallback"),
    location: importLocationSchema,
    preferredMechanism: z.string(),
    selectedFallback: z.string(),
    loss: editabilityLossSchema,
  }),
  z.object({
    code: z.literal("unconsumed-xml"),
    category: z.literal("unsupported"),
    disposition: z.literal("omitted"),
    kind: z.enum(["attribute", "element"]),
    location: importLocationSchema,
    namespaceUri: z.string(),
    localName: z.string(),
  }),
]);
export type ImportDiagnostic = z.infer<typeof importDiagnosticSchema>;

/** Structured summary and per-item diagnostics produced by `createImporter`. */
export const importReportSchema = z.object({
  evidence: z.literal("import"),
  sourceFormat: z.literal("ECMA-376 5th edition Transitional"),
  slideCount: z.number().int().nonnegative(),
  diagnostics: z.array(importDiagnosticSchema).readonly(),
});
export type ImportReport = z.infer<typeof importReportSchema>;

export const importOptionsSchema = z
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

export type ImportOptions = {
  readonly signal?: AbortSignal;
};

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
