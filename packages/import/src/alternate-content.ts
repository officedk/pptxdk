import { MC_NS } from "#/namespaces.ts";
import type { XmlElementCursor } from "#/xml-reader.ts";

export type AlternateContentSelection =
  | {
      readonly branch: "choice";
      readonly cursor: XmlElementCursor;
      readonly children: readonly XmlElementCursor[];
      readonly requires: readonly string[];
      readonly resolvedNamespaces: readonly string[];
    }
  | {
      readonly branch: "fallback";
      readonly cursor: XmlElementCursor;
      readonly children: readonly XmlElementCursor[];
      readonly preferredMechanism: string;
    }
  | {
      readonly branch: "unsupported";
      readonly preferredMechanism: string;
    };

/**
 * Evaluates `<mc:AlternateContent>` according to ECMA-376 Part 3 Markup Compatibility:
 * selects the first `<mc:Choice>` whose `Requires` prefixes all resolve to supported namespaces,
 * or falls back to `<mc:Fallback>` while suppressing spurious unconsumed-xml on rejected branches.
 */
export function readAlternateContent(
  mcCursor: XmlElementCursor,
  isNamespaceSupported: (namespaceUri: string) => boolean,
): AlternateContentSelection {
  const choices = mcCursor.children(MC_NS, "Choice");
  const fallback = mcCursor.child(MC_NS, "Fallback");
  let preferredMechanism = "unknown";

  for (const choice of choices) {
    const rawRequires = choice.attr("Requires") ?? "";
    const prefixes = rawRequires.trim().split(/\s+/u).filter(Boolean);
    const resolvedNamespaces: string[] = [];
    let allSupported = prefixes.length > 0;

    for (const prefix of prefixes) {
      const nsUri = choice.element.lookupNamespaceURI(prefix);
      if (nsUri === null || nsUri === "") {
        allSupported = false;
        break;
      }
      resolvedNamespaces.push(nsUri);
      if (!isNamespaceSupported(nsUri)) {
        allSupported = false;
      }
    }

    if (preferredMechanism === "unknown") {
      preferredMechanism =
        resolvedNamespaces[0] ?? (rawRequires.trim() || "mc:Choice");
    }

    if (allSupported) {
      for (const other of choices) {
        if (other !== choice) {
          other.consumeAttr("Requires");
          other.suppressUnconsumed();
        }
      }
      fallback?.suppressUnconsumed();
      return {
        branch: "choice",
        cursor: choice,
        children: choice.remainingElementChildren(),
        requires: prefixes,
        resolvedNamespaces,
      };
    }

    choice.suppressUnconsumed();
  }

  if (fallback !== undefined) {
    return {
      branch: "fallback",
      cursor: fallback,
      children: fallback.remainingElementChildren(),
      preferredMechanism,
    };
  }

  return {
    branch: "unsupported",
    preferredMechanism,
  };
}
