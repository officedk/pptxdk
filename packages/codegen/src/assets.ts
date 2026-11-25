import { createHash } from "node:crypto";

import {
  IdentifierAllocator,
  slugifyAscii,
  toCamelCaseIdentifier,
} from "./identifiers.ts";

export type RegisteredAsset =
  | {
      readonly kind: "image";
      readonly varName: string;
      readonly relPath: string;
      readonly bytes: Uint8Array;
    }
  | {
      readonly kind: "font";
      readonly varName: string;
      readonly relPath: string;
      readonly bytes: Uint8Array;
    };

function normalizeImageExtension(ext: string): string {
  const cleaned = ext.trim().toLowerCase().replace(/^\./u, "");
  if (!/^[a-z0-9]+$/u.test(cleaned)) {
    throw new Error(`Invalid image asset extension: ${JSON.stringify(ext)}`);
  }
  return cleaned === "jpeg" ? "jpg" : cleaned;
}

function detectFontExtension(
  bytes: Uint8Array,
  explicit?: "ttf" | "otf",
): "ttf" | "otf" {
  if (explicit !== undefined) return explicit;
  if (
    bytes.byteLength >= 4 &&
    bytes[0] === 0x4f &&
    bytes[1] === 0x54 &&
    bytes[2] === 0x54 &&
    bytes[3] === 0x4f
  ) {
    return "otf";
  }
  return "ttf";
}

/** Content-addressed asset registry emitting deterministic `./assets/*` relative paths and bindings. */
export class AssetRegistry {
  readonly #allocator: IdentifierAllocator;
  readonly #byKey = new Map<string, RegisteredAsset>();
  readonly #usedPaths = new Set<string>();
  readonly #ordered: RegisteredAsset[] = [];

  constructor(allocator: IdentifierAllocator) {
    this.#allocator = allocator;
  }

  registerImage(bytes: Uint8Array, ext: string): string {
    const normalizedExt = normalizeImageExtension(ext);
    const fullHash = createHash("sha256").update(bytes).digest("hex");
    const dedupKey = `image:${normalizedExt}:${fullHash}`;
    const existing = this.#byKey.get(dedupKey);
    if (existing !== undefined) {
      return existing.varName;
    }
    const relPath = this.#allocateUniquePath(
      `assets/image-${fullHash.slice(0, 12)}`,
      normalizedExt,
    );
    const varName = this.#allocator.allocate("imageAsset");
    const entry: RegisteredAsset = {
      kind: "image",
      varName,
      relPath,
      bytes: new Uint8Array(bytes),
    };
    this.#byKey.set(dedupKey, entry);
    this.#ordered.push(entry);
    return varName;
  }

  registerFont(bytes: Uint8Array, family: string, ext?: "ttf" | "otf"): string {
    const resolvedExt = detectFontExtension(bytes, ext);
    const fullHash = createHash("sha256").update(bytes).digest("hex");
    const dedupKey = `font:${resolvedExt}:${fullHash}`;
    const existing = this.#byKey.get(dedupKey);
    if (existing !== undefined) {
      return existing.varName;
    }
    const slug = slugifyAscii(family);
    const baseStem =
      slug === ""
        ? `assets/font-${fullHash.slice(0, 12)}`
        : `assets/font-${slug}-${fullHash.slice(0, 12)}`;
    const relPath = this.#allocateUniquePath(baseStem, resolvedExt);
    const baseVar = toCamelCaseIdentifier(family, "font", "FontBytes");
    const varName = this.#allocator.allocate(baseVar);
    const entry: RegisteredAsset = {
      kind: "font",
      varName,
      relPath,
      bytes: new Uint8Array(bytes),
    };
    this.#byKey.set(dedupKey, entry);
    this.#ordered.push(entry);
    return varName;
  }

  get entries(): readonly RegisteredAsset[] {
    return this.#ordered;
  }

  #allocateUniquePath(stem: string, ext: string): string {
    const initial = `${stem}.${ext}`;
    if (!this.#usedPaths.has(initial)) {
      this.#usedPaths.add(initial);
      return initial;
    }
    let counter = 2;
    while (this.#usedPaths.has(`${stem}-${counter}.${ext}`)) {
      counter += 1;
    }
    const unique = `${stem}-${counter}.${ext}`;
    this.#usedPaths.add(unique);
    return unique;
  }
}
