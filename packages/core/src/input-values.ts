import { replaceDeclarationProps } from "#/document-internals.ts";
import {
  isImmutableBytes,
  isSealedFont,
  markImmutableBytes,
} from "#/font-seal.ts";
import { isMorphKey } from "#/morph.ts";

const slideLayoutDefinitionSymbol = Symbol.for(
  "@pptxdk/core/slide-layout-definition",
);
const themeBrandSymbol = Symbol.for("@pptxdk/core/theme");

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function preservesIdentity(value: object): boolean {
  if (
    isSealedFont(value) ||
    isMorphKey(value) ||
    slideLayoutDefinitionSymbol in value ||
    themeBrandSymbol in value
  ) {
    return true;
  }
  // Common nodes expose this adapter hook; projecting their fields would discard node identity.
  if (replaceDeclarationProps in value) return true;
  // Motion scope handles carry a callable identity used by choreography; copying them would sever it.
  return (
    isRecord(value) &&
    Object.getOwnPropertyDescriptor(value, "type")?.value ===
      "motion-scope-handle"
  );
}

function isRecordValue(value: object): value is Record<PropertyKey, unknown> {
  // Structural content records may be class instances; built-in opaque values have their own tags.
  return Object.prototype.toString.call(value) === "[object Object]";
}

function copyValue<T>(value: T, cloneImmutableBytes: boolean): T {
  if (value === null || typeof value !== "object") return value;
  if (value instanceof Uint8Array) {
    if (!cloneImmutableBytes && isImmutableBytes(value)) {
      return value;
    }
    const copy = new Uint8Array(value);
    return (cloneImmutableBytes ? copy : markImmutableBytes(copy)) as T;
  }
  if (value instanceof ArrayBuffer) return value.slice(0) as T;
  if (value instanceof URL) return new URL(value.href) as T;
  if (Array.isArray(value)) {
    return Object.freeze(
      value.map((item) => copyValue(item, cloneImmutableBytes)),
    ) as T;
  }
  if (preservesIdentity(value)) return value;
  if (!isRecordValue(value)) return value;
  const copy: Record<PropertyKey, unknown> = Object.create(null);
  for (const key of Reflect.ownKeys(value)) {
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (descriptor?.enumerable !== true) continue;
    const observedValue = Reflect.get(value, key);
    Object.defineProperty(copy, key, {
      configurable: false,
      enumerable: true,
      value: copyValue(observedValue, cloneImmutableBytes),
      writable: false,
    });
  }
  return Object.freeze(copy) as T;
}

/**
 * Copies author-owned values received by a declaration entry.
 *
 * Record values, arrays, URLs, and byte views are isolated recursively. Enumerable getters are read
 * at receipt, while behavior handles, sealed fonts, already-immutable bytes, and opaque objects
 * keep their identity. This is a field-level authoring rule, not a general document clone.
 */
export function snapshotValue<T>(value: T): T {
  return copyValue(value, false);
}

/**
 * Clones internal declaration values for public inspection (`CommonNode.prototype.props`),
 * always returning fresh unmarked `Uint8Array` copies so external reads cannot mutate internal
 * `immutableBytes` buffers.
 */
export function cloneExternalValue<T>(value: T): T {
  return copyValue(value, true);
}
