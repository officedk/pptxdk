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

  /** 子路径相对触发能力的原生对象，方便定位对象内部的单元格声明。 */
  constructor(code: string, message: string, path?: string) {
    super(message);
    this.name = "CapabilityError";
    this.code = code;
    this.path = path;
  }
}
