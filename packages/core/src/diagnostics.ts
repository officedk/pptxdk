import type { Diagnostic } from "#/types.ts";

export class BuildError extends Error {
  readonly diagnostics: readonly Diagnostic[];

  constructor(diagnostics: readonly Diagnostic[]) {
    super(
      diagnostics
        .map((diagnostic) => `${diagnostic.path}: ${diagnostic.message}`)
        .join("\n"),
    );
    this.name = "BuildError";
    this.diagnostics = Object.freeze(
      diagnostics.map((diagnostic) => Object.freeze({ ...diagnostic })),
    );
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
