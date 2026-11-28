const reservedIdentifiers = new Set([
  // ECMAScript / TypeScript keywords & strict-mode reserved words
  "break",
  "case",
  "catch",
  "class",
  "const",
  "continue",
  "debugger",
  "default",
  "delete",
  "do",
  "else",
  "enum",
  "export",
  "extends",
  "false",
  "finally",
  "for",
  "function",
  "if",
  "implements",
  "import",
  "in",
  "instanceof",
  "interface",
  "let",
  "new",
  "null",
  "package",
  "private",
  "protected",
  "public",
  "return",
  "static",
  "super",
  "switch",
  "this",
  "throw",
  "true",
  "try",
  "typeof",
  "undefined",
  "var",
  "void",
  "while",
  "with",
  "yield",
  "await",
  "async",
  "as",
  "from",
  "of",
  "type",
  "declare",
  "satisfies",
  // Common globals
  "Object",
  "Array",
  "Map",
  "Set",
  "URL",
  "Uint8Array",
  "Math",
  "Number",
  "String",
  "Boolean",
  "Symbol",
  "Error",
  "Promise",
  "globalThis",
  "process",
  "console",
  "readFileSync",
  // Public SDK value imports (vanilla & react) to prevent self-shadowing
  "deck",
  "slide",
  "slideMaster",
  "slideLayout",
  "defineSlideLayout",
  "section",
  "defineTheme",
  "defaultTheme",
  "darker",
  "lighter",
  "createMorphScope",
  "font",
  "linearGradient",
  "radialGradient",
  "transitions",
  "preset",
  "path",
  "blockArc",
  "createMotionScopeHandle",
  "step",
  "motionScope",
  "after",
  "appear",
  "fade",
  "wipe",
  "zoom",
  "fly",
  "floatUp",
  "floatDown",
  "spinIn",
  "stretch",
  "spin",
  "fillColor",
  "breathe",
  "disappear",
  "fadeOut",
  "wipeOut",
  "flyOut",
  "zoomOut",
  "floatOut",
  "spinOut",
  "collapse",
  "onStep",
  "repeat",
  "text",
  "paragraph",
  "span",
  "slideNumber",
  "notes",
  "image",
  "placeholder",
  "shape",
  "box",
  "group",
  "table",
  "row",
  "cell",
  "chart",
  "Deck",
  "Slide",
  "SlideMaster",
  "SlideLayout",
  "Section",
  "Text",
  "P",
  "Span",
  "SlideNumber",
  "Notes",
  "Image",
  "Placeholder",
  "Shape",
  "Box",
  "Group",
  "Table",
  "Row",
  "Cell",
  "Chart",
  "Step",
  "MotionScope",
  "useMotionScope",
  "GeneratedDeck",
]);

function splitAsciiWords(raw: string): string[] {
  const normalized = raw
    .replace(/([a-z0-9])([A-Z])/gu, "$1 $2")
    .replace(/[^A-Za-z0-9]+/gu, " ")
    .trim();
  if (normalized === "") return [];
  return normalized.split(/\s+/u).filter((word) => word.length > 0);
}

export function slugifyAscii(raw: string): string {
  return splitAsciiWords(raw)
    .map((word) => word.toLowerCase())
    .join("-");
}

export function toCamelCaseIdentifier(
  raw: string | undefined,
  fallbackBase: string,
  suffix = "",
): string {
  const words = raw === undefined ? [] : splitAsciiWords(raw);
  if (words.length === 0) {
    return fallbackBase;
  }
  const camel = words
    .map((word, index) => {
      const lower = word.toLowerCase();
      if (index === 0) return lower;
      return `${lower.charAt(0).toUpperCase()}${lower.slice(1)}`;
    })
    .join("");
  const prefixed = /^[A-Za-z_]/u.test(camel)
    ? camel
    : `${fallbackBase}${camel}`;
  if (suffix !== "" && !prefixed.toLowerCase().endsWith(suffix.toLowerCase())) {
    return `${prefixed}${suffix}`;
  }
  return prefixed;
}

export function toPascalCaseIdentifier(
  raw: string | undefined,
  fallbackBase: string,
  suffix = "",
): string {
  const camel = toCamelCaseIdentifier(raw, fallbackBase, suffix);
  const pascal = `${camel.charAt(0).toUpperCase()}${camel.slice(1)}`;
  return /^[A-Z]/u.test(pascal) ? pascal : `${fallbackBase}${pascal}`;
}

/** Deterministic identifier allocator preventing collisions with keywords, globals, and SDK imports. */
export class IdentifierAllocator {
  readonly #used: Set<string>;

  constructor(initialReserved: Iterable<string> = []) {
    this.#used = new Set([...reservedIdentifiers, ...initialReserved]);
  }

  allocate(baseCandidate: string): string {
    const clean = /^[A-Za-z_$][A-Za-z0-9_$]*$/u.test(baseCandidate)
      ? baseCandidate
      : "item";
    if (!this.#used.has(clean)) {
      this.#used.add(clean);
      return clean;
    }
    let counter = 2;
    while (this.#used.has(`${clean}${counter}`)) {
      counter += 1;
    }
    const allocated = `${clean}${counter}`;
    this.#used.add(allocated);
    return allocated;
  }

  fork(): IdentifierAllocator {
    return new IdentifierAllocator(this.#used);
  }
}
