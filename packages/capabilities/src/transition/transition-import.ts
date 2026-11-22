import {
  MC_NS,
  ooxmlNonNegativeIntSchema,
  P14_NS,
  P159_NS,
  P_NS,
  canonicalQName,
  type ImportDiagnostic,
  type ImportLocation,
  type SlideImportContext,
  type SlideImportOutcome,
  type XmlElementCursor,
} from "@pptxdk/import";
import { z } from "zod";

import {
  transitionSpeedSchema,
  transitions,
  type StandardTransition,
  type Transition,
} from "./transition-api.ts";

const p14DurationPattern = /^(\d+(?:\.\d+)?)(h|min|s|ms|µs|ns)?$/u;

const P14DurXmlSchema = z.string().transform((raw, ctx) => {
  const match = p14DurationPattern.exec(raw.trim());
  if (match === null) {
    ctx.addIssue({ code: "custom", message: "Invalid p14:dur value" });
    return z.NEVER;
  }
  const amount = Number(match[1]);
  const unit = match[2];
  const multiplier =
    unit === "h"
      ? 3_600_000
      : unit === "min"
        ? 60_000
        : unit === "s"
          ? 1_000
          : unit === "µs"
            ? 0.001
            : unit === "ns"
              ? 0.000001
              : 1;
  const ms = amount * multiplier;
  if (!Number.isFinite(ms) || ms < 0 || ms > Number.MAX_SAFE_INTEGER) {
    ctx.addIssue({ code: "custom", message: "p14:dur out of range" });
    return z.NEVER;
  }
  return ms;
});

const AdvTmXmlSchema = ooxmlNonNegativeIntSchema.refine(
  (value) => value <= 4_294_967_295,
);

const StripsDirXmlSchema = z
  .enum(["lu", "ru", "ld", "rd"])
  .transform((value) => {
    switch (value) {
      case "lu":
        return "leftUp" as const;
      case "ru":
        return "rightUp" as const;
      case "ld":
        return "leftDown" as const;
      case "rd":
        return "rightDown" as const;
    }
  });

const PushDirXmlSchema = z.enum(["l", "r", "u", "d"]).transform((value) => {
  switch (value) {
    case "l":
      return "left" as const;
    case "r":
      return "right" as const;
    case "u":
      return "top" as const;
    case "d":
      return "bottom" as const;
  }
});

const FlipDirXmlSchema = z.enum(["l", "r"]).transform((value) => {
  return value === "l" ? ("left" as const) : ("right" as const);
});

const MorphOptionXmlSchema = z
  .enum(["byObject", "byWord", "byChar"])
  .transform((value) => {
    switch (value) {
      case "byObject":
        return "object" as const;
      case "byWord":
        return "word" as const;
      case "byChar":
        return "character" as const;
    }
  });

const supportedTransitionNamespaces = new Set<string>([P_NS, P14_NS, P159_NS]);

type ParsedTransitionEffect =
  | {
      readonly status: "standard";
      readonly transition: StandardTransition;
      readonly qname: string;
      readonly advanceAfter?: number;
    }
  | {
      readonly status: "flip";
      readonly direction: "left" | "right";
      readonly duration?: number;
      readonly qname: string;
      readonly advanceAfter?: number;
    }
  | {
      readonly status: "morph";
      readonly by: "object" | "word" | "character";
      readonly duration?: number;
      readonly qname: string;
      readonly advanceAfter?: number;
    }
  | {
      readonly status: "advance-only";
      readonly advanceAfter: number;
    }
  | {
      readonly status: "empty";
    }
  | {
      readonly status: "unsupported";
      readonly qname: string;
      readonly namespaceUri: string;
      readonly localName: string;
      readonly xpath: string;
      readonly advanceAfter?: number;
    };

function locationFor(
  context: SlideImportContext,
  xpath: string,
): ImportLocation {
  return {
    part: context.part,
    xpath,
    ...(context.slideIndex === undefined
      ? {}
      : { slideIndex: context.slideIndex }),
  };
}

function readTransitionElement(
  cursor: XmlElementCursor,
  context: SlideImportContext,
  diagnostics: ImportDiagnostic[],
): ParsedTransitionEffect {
  const rawSpd = cursor.attr("spd");
  if (
    rawSpd !== undefined &&
    !transitionSpeedSchema.safeParse(rawSpd).success
  ) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "recovered",
      location: locationFor(context, `${cursor.xpath}/@spd`),
      attributeOrElement: "@spd",
      rawValue: rawSpd,
    });
  }

  let duration: number | undefined;
  const rawDur = cursor.attr("dur", P14_NS);
  if (rawDur !== undefined) {
    const parsedDur = P14DurXmlSchema.safeParse(rawDur);
    if (parsedDur.success) {
      duration = parsedDur.data;
    } else {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: locationFor(context, `${cursor.xpath}/@p14:dur`),
        attributeOrElement: "@p14:dur",
        rawValue: rawDur,
      });
    }
  }

  let advanceAfter: number | undefined;
  const rawAdvTm = cursor.attr("advTm");
  if (rawAdvTm !== undefined) {
    const parsedAdv = AdvTmXmlSchema.safeParse(rawAdvTm);
    if (parsedAdv.success) {
      advanceAfter = parsedAdv.data;
    } else {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: locationFor(context, `${cursor.xpath}/@advTm`),
        attributeOrElement: "@advTm",
        rawValue: rawAdvTm,
      });
    }
  }

  const children = cursor.remainingElementChildren();
  if (children.length === 0) {
    return advanceAfter === undefined
      ? { status: "empty" }
      : { status: "advance-only", advanceAfter };
  }

  if (children.length > 1) {
    for (let i = 1; i < children.length; i += 1) {
      children[i]!.suppressUnconsumed();
    }
  }
  const effect = children[0]!;
  const durationOptions = duration === undefined ? {} : { duration };
  const advProps = advanceAfter === undefined ? {} : { advanceAfter };

  if (effect.namespaceUri === P_NS) {
    if (effect.localName === "fade") {
      return {
        status: "standard",
        transition: transitions.fade(durationOptions),
        qname: "p:fade",
        ...advProps,
      };
    }
    if (effect.localName === "strips") {
      const rawDir = effect.attr("dir");
      let direction: "leftUp" | "rightUp" | "leftDown" | "rightDown" =
        "rightDown";
      if (rawDir !== undefined) {
        const parsed = StripsDirXmlSchema.safeParse(rawDir);
        if (parsed.success) {
          direction = parsed.data;
        } else {
          diagnostics.push({
            code: "invalid-xml-value",
            category: "corrupt-input",
            disposition: "recovered",
            location: locationFor(context, `${effect.xpath}/@dir`),
            attributeOrElement: "@dir",
            rawValue: rawDir,
          });
        }
      }
      return {
        status: "standard",
        transition: transitions.strips({ ...durationOptions, direction }),
        qname: "p:strips",
        ...advProps,
      };
    }
    if (effect.localName === "push") {
      const rawDir = effect.attr("dir");
      let from: "left" | "right" | "top" | "bottom" = "left";
      if (rawDir !== undefined) {
        const parsed = PushDirXmlSchema.safeParse(rawDir);
        if (parsed.success) {
          from = parsed.data;
        } else {
          diagnostics.push({
            code: "invalid-xml-value",
            category: "corrupt-input",
            disposition: "recovered",
            location: locationFor(context, `${effect.xpath}/@dir`),
            attributeOrElement: "@dir",
            rawValue: rawDir,
          });
        }
      }
      return {
        status: "standard",
        transition: transitions.push({ ...durationOptions, from }),
        qname: "p:push",
        ...advProps,
      };
    }
  } else if (effect.namespaceUri === P14_NS && effect.localName === "flip") {
    const rawDir = effect.attr("dir");
    let direction: "left" | "right" = "left";
    if (rawDir !== undefined) {
      const parsed = FlipDirXmlSchema.safeParse(rawDir);
      if (parsed.success) {
        direction = parsed.data;
      } else {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "recovered",
          location: locationFor(context, `${effect.xpath}/@dir`),
          attributeOrElement: "@dir",
          rawValue: rawDir,
        });
      }
    }
    return {
      status: "flip",
      direction,
      ...(duration === undefined ? {} : { duration }),
      qname: "p14:flip",
      ...advProps,
    };
  } else if (effect.namespaceUri === P159_NS && effect.localName === "morph") {
    const rawOption = effect.attr("option");
    let by: "object" | "word" | "character" = "object";
    if (rawOption !== undefined) {
      const parsed = MorphOptionXmlSchema.safeParse(rawOption);
      if (parsed.success) {
        by = parsed.data;
      } else {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "recovered",
          location: locationFor(context, `${effect.xpath}/@option`),
          attributeOrElement: "@option",
          rawValue: rawOption,
        });
      }
    }
    return {
      status: "morph",
      by,
      ...(duration === undefined ? {} : { duration }),
      qname: "p159:morph",
      ...advProps,
    };
  }

  effect.suppressUnconsumed();
  return {
    status: "unsupported",
    qname: canonicalQName(effect.namespaceUri, effect.localName),
    namespaceUri: effect.namespaceUri,
    localName: effect.localName,
    xpath: effect.xpath,
    ...advProps,
  };
}

function finalizeSupportedTransition(
  parsed: Exclude<ParsedTransitionEffect, { readonly status: "unsupported" }>,
  fallback: StandardTransition = transitions.fade(),
): {
  readonly transition?: Transition;
  readonly advanceAfter?: number;
} {
  const advProps =
    "advanceAfter" in parsed && parsed.advanceAfter !== undefined
      ? { advanceAfter: parsed.advanceAfter }
      : {};
  switch (parsed.status) {
    case "standard":
      return { transition: parsed.transition, ...advProps };
    case "flip":
      return {
        transition: transitions.flip({
          direction: parsed.direction,
          ...(parsed.duration === undefined
            ? {}
            : { duration: parsed.duration }),
          fallback,
        }),
        ...advProps,
      };
    case "morph":
      return {
        transition: transitions.morph({
          by: parsed.by,
          ...(parsed.duration === undefined
            ? {}
            : { duration: parsed.duration }),
          fallback,
        }),
        ...advProps,
      };
    case "advance-only":
      return advProps;
    case "empty":
      return {};
  }
}

function choiceNamespacesSupported(choice: XmlElementCursor): {
  readonly supported: boolean;
  readonly preferredFallbackLabel: string;
} {
  const rawRequires = choice.attr("Requires") ?? "";
  const prefixes = rawRequires.trim().split(/\s+/u).filter(Boolean);
  if (prefixes.length === 0) {
    return { supported: false, preferredFallbackLabel: "mc:Choice" };
  }
  const resolved: string[] = [];
  let supported = true;
  for (const prefix of prefixes) {
    const nsUri = choice.element.lookupNamespaceURI(prefix);
    if (nsUri === null || nsUri === "") {
      supported = false;
      continue;
    }
    resolved.push(nsUri);
    if (!supportedTransitionNamespaces.has(nsUri)) {
      supported = false;
    }
  }
  return {
    supported,
    preferredFallbackLabel: resolved[0] ?? (rawRequires.trim() || "mc:Choice"),
  };
}

function readAlternateContentTransition(
  mcCursor: XmlElementCursor,
  context: SlideImportContext,
  diagnostics: ImportDiagnostic[],
): {
  readonly transition?: Transition;
  readonly advanceAfter?: number;
} {
  const choices = mcCursor.children(MC_NS, "Choice");
  const fallback = mcCursor.child(MC_NS, "Fallback");
  let preferredMechanism = "mc:Choice";
  let preferredUnsupported:
    | Extract<ParsedTransitionEffect, { readonly status: "unsupported" }>
    | undefined;
  let fallbackAdvTm: number | undefined;

  for (let i = 0; i < choices.length; i += 1) {
    const choice = choices[i]!;
    const nsCheck = choiceNamespacesSupported(choice);
    const transitionCursor = choice.child(P_NS, "transition");
    if (nsCheck.supported && transitionCursor !== undefined) {
      const choiceDiags: ImportDiagnostic[] = [];
      const parsedChoice = readTransitionElement(
        transitionCursor,
        context,
        choiceDiags,
      );
      if (parsedChoice.status !== "unsupported") {
        diagnostics.push(...choiceDiags);
        for (let j = i + 1; j < choices.length; j += 1) {
          choices[j]!.consumeAttr("Requires");
          choices[j]!.suppressUnconsumed();
        }
        let standardFallback: StandardTransition = transitions.fade();
        let fallbackAdv: number | undefined;
        if (fallback !== undefined) {
          const fallbackTransition = fallback.child(P_NS, "transition");
          if (fallbackTransition !== undefined) {
            const fbDiags: ImportDiagnostic[] = [];
            const parsedFb = readTransitionElement(
              fallbackTransition,
              context,
              fbDiags,
            );
            if (parsedFb.status === "standard") {
              diagnostics.push(...fbDiags);
              standardFallback = parsedFb.transition;
              fallbackAdv = parsedFb.advanceAfter;
            } else if (parsedFb.status === "advance-only") {
              diagnostics.push(...fbDiags);
              fallbackAdv = parsedFb.advanceAfter;
            } else if (
              parsedFb.status === "unsupported" &&
              (parsedChoice.status === "flip" ||
                parsedChoice.status === "morph")
            ) {
              diagnostics.push(...fbDiags);
              diagnostics.push({
                code: "unsupported-mechanism",
                category: "unsupported",
                disposition: "omitted",
                location: locationFor(context, parsedFb.xpath),
                namespaceUri: parsedFb.namespaceUri,
                localName: parsedFb.localName,
                mechanismHint: parsedFb.qname,
              });
            }
          }
          fallback.suppressUnconsumed();
        }
        const finalized = finalizeSupportedTransition(
          parsedChoice,
          standardFallback,
        );
        return {
          ...finalized,
          ...(finalized.advanceAfter !== undefined
            ? {}
            : fallbackAdv !== undefined
              ? { advanceAfter: fallbackAdv }
              : {}),
        };
      }
      if (preferredUnsupported === undefined) {
        preferredUnsupported = parsedChoice;
        preferredMechanism = parsedChoice.qname;
      }
      if (
        fallbackAdvTm === undefined &&
        parsedChoice.advanceAfter !== undefined
      ) {
        fallbackAdvTm = parsedChoice.advanceAfter;
      }
    } else if (preferredUnsupported === undefined) {
      if (transitionCursor !== undefined) {
        const parsedChoice = readTransitionElement(
          transitionCursor,
          context,
          [],
        );
        if (parsedChoice.status === "unsupported") {
          preferredUnsupported = parsedChoice;
          preferredMechanism = parsedChoice.qname;
        } else {
          preferredMechanism = nsCheck.preferredFallbackLabel;
        }
      } else {
        preferredMechanism = nsCheck.preferredFallbackLabel;
      }
    }
    choice.suppressUnconsumed();
  }

  if (fallback !== undefined) {
    const fallbackTransition = fallback.child(P_NS, "transition");
    if (fallbackTransition !== undefined) {
      const fbDiags: ImportDiagnostic[] = [];
      const parsedFb = readTransitionElement(
        fallbackTransition,
        context,
        fbDiags,
      );
      if (parsedFb.status !== "unsupported" && parsedFb.status !== "empty") {
        diagnostics.push(...fbDiags);
        const adopted = finalizeSupportedTransition(parsedFb);
        if (adopted.transition !== undefined) {
          diagnostics.push({
            code: "compatibility-fallback",
            category: "unsupported",
            disposition: "fallback",
            location: locationFor(context, mcCursor.xpath),
            preferredMechanism,
            selectedFallback:
              "qname" in parsedFb ? parsedFb.qname : "p:transition",
            loss: {
              lost: "preferred-compatibility-branch",
              retainedAs: `transition:${adopted.transition.kind}`,
            },
          });
        }
        fallback.suppressUnconsumed();
        return adopted;
      }
      if (parsedFb.status === "unsupported") {
        diagnostics.push(...fbDiags);
        diagnostics.push({
          code: "unsupported-mechanism",
          category: "unsupported",
          disposition: "omitted",
          location: locationFor(context, parsedFb.xpath),
          namespaceUri: parsedFb.namespaceUri,
          localName: parsedFb.localName,
          mechanismHint: preferredMechanism,
        });
        fallback.suppressUnconsumed();
        const adv = parsedFb.advanceAfter ?? fallbackAdvTm;
        return adv === undefined ? {} : { advanceAfter: adv };
      }
    }
    fallback.suppressUnconsumed();
  }

  diagnostics.push({
    code: "unsupported-mechanism",
    category: "unsupported",
    disposition: "omitted",
    location: locationFor(
      context,
      preferredUnsupported?.xpath ?? mcCursor.xpath,
    ),
    namespaceUri: preferredUnsupported?.namespaceUri ?? MC_NS,
    localName: preferredUnsupported?.localName ?? "AlternateContent",
    mechanismHint: preferredMechanism,
  });
  mcCursor.suppressUnconsumed();
  return fallbackAdvTm === undefined ? {} : { advanceAfter: fallbackAdvTm };
}

function isTransitionAlternateContentElement(
  element: XmlElementCursor["element"],
): boolean {
  if (
    (element.namespaceURI ?? "") !== MC_NS ||
    (element.localName ?? element.tagName) !== "AlternateContent"
  ) {
    return false;
  }
  for (
    let branch = element.firstChild;
    branch !== null;
    branch = branch.nextSibling
  ) {
    if (branch.nodeType !== 1) continue;
    const branchEl = branch as XmlElementCursor["element"];
    if ((branchEl.namespaceURI ?? "") !== MC_NS) continue;
    const branchLocal = branchEl.localName ?? branchEl.tagName;
    if (branchLocal !== "Choice" && branchLocal !== "Fallback") continue;
    for (
      let child = branchEl.firstChild;
      child !== null;
      child = child.nextSibling
    ) {
      if (child.nodeType !== 1) continue;
      const childEl = child as XmlElementCursor["element"];
      if (
        (childEl.namespaceURI ?? "") === P_NS &&
        (childEl.localName ?? childEl.tagName) === "transition"
      ) {
        return true;
      }
    }
  }
  return false;
}

/**
 * Imports `<p:transition>` or `<mc:AlternateContent>` transition declarations onto
 * a `Slide`, `SlideLayout`, or `SlideMaster` node while preserving exact ownership and fallbacks.
 */
export function readSlideTransitionImport(
  context: SlideImportContext,
): SlideImportOutcome {
  const candidates = context.slideCursor.matchingChildren((element) => {
    const ns = element.namespaceURI ?? "";
    const local = element.localName ?? element.tagName;
    return (
      (ns === P_NS && local === "transition") ||
      isTransitionAlternateContentElement(element)
    );
  });
  if (candidates.length === 0) {
    return {};
  }

  const diagnostics: ImportDiagnostic[] = [];
  if (candidates.length > 1) {
    for (let i = 1; i < candidates.length; i += 1) {
      const duplicate = candidates[i]!;
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: locationFor(context, duplicate.xpath),
        attributeOrElement: duplicate.localName,
      });
      duplicate.suppressUnconsumed();
    }
  }

  const primary = candidates[0]!;
  let resolved: {
    readonly transition?: Transition;
    readonly advanceAfter?: number;
  };
  if (primary.namespaceUri === MC_NS) {
    resolved = readAlternateContentTransition(primary, context, diagnostics);
  } else {
    const parsed = readTransitionElement(primary, context, diagnostics);
    if (parsed.status === "unsupported") {
      diagnostics.push({
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: locationFor(context, parsed.xpath),
        namespaceUri: parsed.namespaceUri,
        localName: parsed.localName,
        mechanismHint: parsed.qname,
      });
      resolved =
        parsed.advanceAfter === undefined
          ? {}
          : { advanceAfter: parsed.advanceAfter };
    } else {
      const finalized = finalizeSupportedTransition(parsed);
      resolved = finalized;
    }
  }

  if (
    resolved.transition !== undefined ||
    resolved.advanceAfter !== undefined
  ) {
    context.slideNode.set({
      ...(resolved.transition === undefined
        ? {}
        : { transition: resolved.transition }),
      ...(resolved.advanceAfter === undefined
        ? {}
        : { advanceAfter: resolved.advanceAfter }),
    });
  }

  return diagnostics.length === 0 ? {} : { diagnostics };
}
