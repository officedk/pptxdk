const morphKeyBrand: unique symbol = Symbol("MorphKey");
const morphKeyLabel: unique symbol = Symbol("MorphKeyLabel");

/** Handle created by MorphScope to explicitly declare cross-slide object continuity. */
export type MorphKey = {
  readonly [morphKeyBrand]: true;
  readonly [morphKeyLabel]: string;
};

/**
 * Continuity handle factory for a single logical component instance.
 *
 * Pass the same MorphScope when reusing across slides; duplicated components on the same slide each create their own scope. When created in a React component,
 * use useMemo or useState to keep the scope stable; ordinary function component boundaries do not automatically establish a scope.
 */
export type MorphScope = {
  /** Obtain a stable handle by name within this scope. */
  handle(name: string): MorphKey;
};

/**
 * Create an independent continuity scope; handles correspond across slides only when explicitly reused.
 *
 * When created in a React component, use useMemo or useState to keep the return value stable; cross-slide correspondence requires reusing the same scope.
 */
export function createMorphScope(): MorphScope {
  const handles = new Map<string, MorphKey>();
  return Object.freeze({
    handle(name: string): MorphKey {
      if (typeof name !== "string" || name.trim() === "") {
        throw new TypeError(
          "MorphScope.handle(name) requires a non-empty name",
        );
      }
      let handle = handles.get(name);
      if (handle === undefined) {
        handle = Object.freeze({
          [morphKeyBrand]: true as const,
          [morphKeyLabel]: name,
        });
        handles.set(name, handle);
      }
      return handle;
    },
  });
}

export function isMorphKey(value: unknown): value is MorphKey {
  return (
    typeof value === "object" &&
    value !== null &&
    morphKeyBrand in value &&
    morphKeyLabel in value
  );
}

export function getMorphKeyLabel(value: MorphKey): string {
  return value[morphKeyLabel];
}
