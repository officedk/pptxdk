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

  /** Subpath relative to the native object triggering the capability, facilitating localization of cell declarations inside the object. */
  constructor(code: string, message: string, path?: string) {
    super(message);
    this.name = "CapabilityError";
    this.code = code;
    this.path = path;
  }
}
