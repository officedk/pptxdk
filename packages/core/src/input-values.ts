import { isMorphKey } from "#/morph.ts";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function preservesIdentity(value: object): boolean {
  if (isMorphKey(value)) return true;
  // Motion scope handles carry a callable identity used by choreography; copying them would sever it.
  return isRecord(value) && value["type"] === "motion-scope-handle";
}

/**
 * Copies author-owned values received by a declaration entry.
 *
 * Plain records, arrays, URLs, and byte views are isolated recursively. Behavior handles and opaque
 * objects keep their identity so structural and choreography associations are not cloned. This is a
 * field-level authoring rule, not a general document clone.
 */
export function snapshotValue<T>(value: T): T {
  if (value === null || typeof value !== "object") return value;
  if (value instanceof Uint8Array) return new Uint8Array(value) as T;
  if (value instanceof ArrayBuffer) return value.slice(0) as T;
  if (value instanceof URL) return new URL(value.href) as T;
  if (Array.isArray(value)) {
    return Object.freeze(value.map((item) => snapshotValue(item))) as T;
  }
  if (preservesIdentity(value)) return value;
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) return value;
  const copy: Record<PropertyKey, unknown> = Object.create(null);
  for (const key of Reflect.ownKeys(value)) {
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (descriptor?.enumerable !== true) continue;
    Object.defineProperty(copy, key, {
      configurable: false,
      enumerable: true,
      value: snapshotValue(descriptor.value),
      writable: false,
    });
  }
  return Object.freeze(copy) as T;
}
