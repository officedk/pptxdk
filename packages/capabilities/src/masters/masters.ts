import {
  defaultColorMap,
  defaultTheme,
  visualStyleSchema,
  type ColorMap,
  type ColorMapOverride,
  type LayoutPlaceholderSpec,
  type NativeColorSlot,
  type PlaceholderMap,
  type SlideCapabilityModule,
  type Theme,
  type ThemeContext,
  type ThemeOverride,
} from "@pptxdk/core/capability";
import {
  colorMapAttributesXml,
  colorMapOverrideXml,
  layoutXml,
  masterXml,
  themeOverrideXml,
  themeXml,
  type EmittedSlideLayoutSpec,
  type EmittedSlideMasterSpec,
} from "@pptxdk/emit";
import {
  A_NS,
  TransformAttributesFragmentSchema,
  defineCanvasMechanism,
  importMastersAndLayouts,
  type DeckImportContext,
  type DeckImportOutcome,
  type ImportDiagnostic,
  type ImportedPlaceholderSpPr,
  type PlaceholderVisualReader,
} from "@pptxdk/import";

import type { Geometry } from "../shared/geometry.ts";
import { hasLocalDeclaration } from "../shared/style-inheritance.ts";
import {
  emitPlaceholderVisualSpPr,
  paintVisualModel,
  readSpPrAppearance,
  readTransformXml,
  resolveVisualModel,
  type PlaceholderSpPrFlags,
  type ResolvedVisualModel,
} from "../shared/visual.ts";

const supportedPlaceholderVisualStyleKeys = new Set(
  Object.keys(visualStyleSchema.unwrap().shape),
);

type ResolvedPlaceholderVisualEntry = {
  readonly key: string;
  readonly visual: ResolvedVisualModel;
  readonly flags: PlaceholderSpPrFlags;
  readonly hasPaintableBox: boolean;
};

const readMasterPlaceholderVisuals: PlaceholderVisualReader = ({
  spPr,
  pkg,
  location,
  diagnostics,
}): ImportedPlaceholderSpPr => {
  let x: number | undefined;
  let y: number | undefined;
  let width: number | undefined;
  let height: number | undefined;
  let rotate: number | undefined;
  let flipH: boolean | undefined;
  let flipV: boolean | undefined;
  let fallbackUnits = { w: 0, h: 0 };

  const xfrm = spPr.child(A_NS, "xfrm");
  if (xfrm !== undefined) {
    if (
      xfrm.child(A_NS, "off") !== undefined &&
      xfrm.child(A_NS, "ext") !== undefined
    ) {
      const tf = readTransformXml(xfrm, location, diagnostics);
      if (tf !== undefined) {
        x = tf.x;
        y = tf.y;
        width = tf.width;
        height = tf.height;
        rotate = tf.rotate;
        flipH = tf.flipH;
        flipV = tf.flipV;
        fallbackUnits = { w: tf.cxEmu, h: tf.cyEmu };
      }
    } else {
      const parsedXfrm = TransformAttributesFragmentSchema.safeParse(
        xfrm.attrsObject(["rot", "flipH", "flipV"]),
      );
      if (parsedXfrm.success) {
        if (parsedXfrm.data.rot !== undefined && parsedXfrm.data.rot !== 0) {
          rotate = parsedXfrm.data.rot / 60_000;
        }
        if (parsedXfrm.data.flipH === true) flipH = true;
        if (parsedXfrm.data.flipV === true) flipV = true;
      }
    }
  }

  const appearance = readSpPrAppearance(
    spPr,
    pkg,
    fallbackUnits,
    location,
    diagnostics,
  );

  let borderRadius: number | undefined;
  let exportedGeometry: unknown;
  if (appearance.geometry === "roundRect") {
    if (width !== undefined && height !== undefined) {
      borderRadius = (16667 / 100_000) * Math.min(width, height);
    } else {
      exportedGeometry = appearance.geometry;
    }
  } else if (
    typeof appearance.geometry === "object" &&
    appearance.geometry.type === "preset"
  ) {
    if (appearance.geometry.name === "roundRect") {
      if (width !== undefined && height !== undefined) {
        const adj = appearance.geometry.adjust?.["adj"] ?? 16667;
        if (adj > 0) {
          borderRadius = (adj / 100_000) * Math.min(width, height);
        }
      } else {
        exportedGeometry = appearance.geometry;
      }
    } else if (appearance.geometry.name !== "rect") {
      exportedGeometry = appearance.geometry;
    }
  } else if (
    appearance.geometry !== undefined &&
    appearance.geometry !== "rect"
  ) {
    exportedGeometry = appearance.geometry;
  }

  const style: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(appearance.style)) {
    if (value !== undefined && supportedPlaceholderVisualStyleKeys.has(key)) {
      style[key] = value;
    }
  }
  if (borderRadius !== undefined) {
    style["borderRadius"] = borderRadius;
  }
  if (rotate !== undefined) {
    style["rotate"] = rotate;
  }

  return {
    ...(x === undefined ? {} : { x }),
    ...(y === undefined ? {} : { y }),
    ...(width === undefined ? {} : { width }),
    ...(height === undefined ? {} : { height }),
    ...(exportedGeometry === undefined ? {} : { geometry: exportedGeometry }),
    ...(flipH ? { flipH: true } : {}),
    ...(flipV ? { flipV: true } : {}),
    ...(Object.keys(style).length === 0 ? {} : { style }),
  };
};

function importDeckMasters(context: DeckImportContext): DeckImportOutcome {
  if (context.deckNode.masters.length > 0) {
    return {};
  }
  const diagnostics: ImportDiagnostic[] = [];
  const presRelById =
    context.relById ??
    new Map(context.package.presentationRels.map((entry) => [entry.Id, entry]));
  const usedDeckIds = context.usedDeckIds ?? new Set<string>();
  const clrMapByLayoutPartPath =
    context.clrMapByLayoutPartPath ??
    new Map<string, Readonly<Record<string, NativeColorSlot>>>();
  const layoutNodes = importMastersAndLayouts(
    context.presentationCursor,
    context.package,
    presRelById,
    context.deckNode,
    usedDeckIds,
    diagnostics,
    context.slideHandlers ?? [],
    context.shapeTreeHandlers ?? [],
    context.registerDeferredSlideLink ?? (() => {}),
    clrMapByLayoutPartPath,
    readMasterPlaceholderVisuals,
  );
  if (context.layoutNodeByPartPath !== undefined) {
    for (const [partPath, layoutNode] of layoutNodes) {
      context.layoutNodeByPartPath.set(partPath, layoutNode);
    }
  }
  return diagnostics.length === 0 ? {} : { diagnostics };
}

function resolvePlaceholderVisualEntries(
  placeholders: PlaceholderMap | undefined,
  theme: ThemeContext,
): readonly ResolvedPlaceholderVisualEntry[] {
  if (placeholders === undefined) return [];
  const entries: ResolvedPlaceholderVisualEntry[] = [];
  for (const [key, spec] of Object.entries(placeholders)) {
    const hasBox =
      spec.x !== undefined &&
      spec.y !== undefined &&
      spec.width !== undefined &&
      spec.height !== undefined;
    const styleRecord = spec.style as
      | Readonly<Record<string, unknown>>
      | undefined;
    const hasRotate = hasLocalDeclaration(styleRecord, "rotate");
    const hasFlipH = spec.flipH === true;
    const hasFlipV = spec.flipV === true;
    const hasGeometryOrRadius =
      spec.geometry !== undefined ||
      hasLocalDeclaration(styleRecord, "borderRadius");
    const hasFill =
      hasLocalDeclaration(styleRecord, "backgroundColor") ||
      hasLocalDeclaration(styleRecord, "backgroundImage");
    const hasBorder =
      hasLocalDeclaration(styleRecord, "borderColor") &&
      spec.style?.borderWidth !== 0;
    const hasEffects =
      hasLocalDeclaration(styleRecord, "boxShadow") ||
      hasLocalDeclaration(styleRecord, "softEdge");

    if (
      !hasBox &&
      !hasRotate &&
      !hasFlipH &&
      !hasFlipV &&
      !hasGeometryOrRadius &&
      !hasFill &&
      !hasBorder &&
      !hasEffects
    ) {
      continue;
    }

    const box = hasBox
      ? {
          x: spec.x,
          y: spec.y,
          width: spec.width,
          height: spec.height,
        }
      : { x: 0, y: 0, width: 100, height: 100 };
    const visual = resolveVisualModel({
      theme,
      style: styleRecord ?? {},
      box,
      ...(spec.geometry === undefined
        ? {}
        : { geometry: spec.geometry as Geometry }),
      ...(hasFlipH ? { flipH: true } : {}),
      ...(hasFlipV ? { flipV: true } : {}),
      layoutOnlyWhenInvisible: !hasBox,
    });
    entries.push({
      key,
      visual,
      flags: {
        hasBox,
        hasRotate,
        hasFlipH,
        hasFlipV,
        hasGeometryOrRadius,
        hasFill,
        hasBorder,
        hasEffects,
        includeDefaultRectWhenHasBox: true,
      },
      hasPaintableBox: hasBox && visual.hasVisibleStyle,
    });
  }
  return entries;
}

function paintPlaceholderVisualEntries(
  entries: readonly ResolvedPlaceholderVisualEntry[],
  slideNumber: number,
): string {
  return entries
    .filter((entry) => entry.hasPaintableBox)
    .map(
      (entry, index) =>
        paintVisualModel({
          model: entry.visual,
          slideNumber,
          objectId: 2 + index,
        }).svg,
    )
    .join("");
}

type SlideMasterMechanismModel = {
  readonly spec: EmittedSlideMasterSpec;
  readonly layoutIds: readonly { readonly id: number; readonly rId: string }[];
  readonly themeContext: ThemeContext;
  readonly placeholders?: readonly ResolvedPlaceholderVisualEntry[];
  readonly slideNumber?: number;
};

/** Native canvas mechanism for `<p:sldMaster>` ownership, `preserve`, and master-layout relationships. */
export const slideMasterMechanism = defineCanvasMechanism({
  key: "slide-master",
  role: "canvas",
  emit: (model: SlideMasterMechanismModel) => ({
    xml: masterXml(model.spec, model.layoutIds, model.themeContext),
  }),
  paint: (model: SlideMasterMechanismModel) => ({
    svg: paintPlaceholderVisualEntries(
      model.placeholders ?? [],
      model.slideNumber ?? 1,
    ),
  }),
  import: {
    target: "deck",
    stage: "pre-slides",
    read: importDeckMasters,
  },
});

type SlideLayoutMechanismModel = {
  readonly spec: EmittedSlideLayoutSpec;
  readonly themeContext: ThemeContext;
  readonly placeholders?: readonly ResolvedPlaceholderVisualEntry[];
  readonly slideNumber?: number;
};

/** Native canvas mechanism for `<p:sldLayout>` ownership, `type`, `preserve`, and idle layout retention. */
export const slideLayoutMechanism = defineCanvasMechanism({
  key: "slide-layout",
  role: "canvas",
  emit: (model: SlideLayoutMechanismModel) => ({
    xml: layoutXml(model.spec, model.themeContext),
  }),
  paint: (model: SlideLayoutMechanismModel) => ({
    svg: paintPlaceholderVisualEntries(
      model.placeholders ?? [],
      model.slideNumber ?? 1,
    ),
  }),
  import: {
    target: "deck",
    stage: "pre-slides",
    read: importDeckMasters,
  },
});

function serializeColorMapModel(model: {
  readonly colorMap?: ColorMap;
  readonly colorMapOverride?: ColorMapOverride;
}): string {
  if (model.colorMapOverride !== undefined) {
    return colorMapOverrideXml(model.colorMapOverride);
  }
  return `<p:clrMap ${colorMapAttributesXml(model.colorMap ?? defaultColorMap)}/>`;
}

type NativeThemeMechanismModel = {
  readonly theme: Theme;
  readonly themeContext: ThemeContext;
};

/** Native canvas mechanism for `<a:theme>` (`clrScheme`, `fontScheme`, `fmtScheme`) part identity and mutation. */
export const nativeThemeMechanism = defineCanvasMechanism({
  key: "native-theme",
  role: "canvas",
  emit: (model: NativeThemeMechanismModel) => ({ xml: themeXml(model.theme) }),
  paint: (model: NativeThemeMechanismModel) => {
    const bg = model.themeContext.resolveColor("bg1");
    return {
      svg: `<rect width="100%" height="100%" fill="#${bg.hex}" fill-opacity="${bg.alpha ?? 1}"/>`,
    };
  },
  import: {
    target: "deck",
    stage: "pre-slides",
    read: importDeckMasters,
  },
});

type ColorMapMechanismModel = {
  readonly colorMap?: ColorMap;
  readonly colorMapOverride?: ColorMapOverride;
  readonly themeContext: ThemeContext;
};

/** Native canvas mechanism for master `<p:clrMap>` and layout/slide `<p:clrMapOvr>` color-slot mapping. */
export const colorMapMechanism = defineCanvasMechanism({
  key: "color-map",
  role: "canvas",
  emit: (model: ColorMapMechanismModel) => ({
    xml: serializeColorMapModel(model),
  }),
  paint: (model: ColorMapMechanismModel) => {
    const bg = model.themeContext.resolveColor("bg1");
    const tx = model.themeContext.resolveColor("tx1");
    return {
      svg: `<g data-color-map-bg="#${bg.hex}" data-color-map-tx="#${tx.hex}"/>`,
    };
  },
  import: {
    target: "deck",
    stage: "pre-slides",
    read: importDeckMasters,
  },
});

type ThemeOverrideMechanismModel = {
  readonly override: ThemeOverride;
  readonly baseTheme?: Theme;
  readonly themeContext: ThemeContext;
};

/** Native canvas mechanism for `<a:themeOverride>` layout and slide palette/font/format overrides. */
export const themeOverrideMechanism = defineCanvasMechanism({
  key: "theme-override",
  role: "canvas",
  emit: (model: ThemeOverrideMechanismModel) => ({
    xml: themeOverrideXml(model.override, model.baseTheme ?? defaultTheme),
  }),
  paint: (model: ThemeOverrideMechanismModel) => {
    const accent1 = model.themeContext.resolveColor("accent1");
    return {
      svg: `<g data-theme-override-accent1="#${accent1.hex}"/>`,
    };
  },
  import: {
    target: "deck",
    stage: "pre-slides",
    read: importDeckMasters,
  },
});

/** Native canvas mechanism for non-placeholder shared canvas objects (`sp`/`pic`/`grpSp`) on SlideMaster and SlideLayout. */
export const masterSharedCanvasMechanism = defineCanvasMechanism({
  key: "master-shared-canvas",
  role: "canvas",
  emit: (model: { readonly shapeXml: string; readonly svg: string }) => ({
    xml: model.shapeXml,
  }),
  paint: (model: { readonly shapeXml: string; readonly svg: string }) => ({
    svg: model.svg,
  }),
  import: {
    target: "deck",
    stage: "pre-slides",
    read: importDeckMasters,
  },
});

export {
  footerPlaceholderCapability,
  footerPlaceholderMechanism,
  footerPlaceholderShapeImportHandler,
} from "./footer-placeholder.ts";

/** Capability declaration wiring native slide master, layout, theme, color-map, theme-override, and shared-canvas mechanisms into builds. */
export const mastersCapability: SlideCapabilityModule & {
  readonly mechanisms: readonly [
    typeof slideMasterMechanism,
    typeof slideLayoutMechanism,
    typeof nativeThemeMechanism,
    typeof colorMapMechanism,
    typeof themeOverrideMechanism,
    typeof masterSharedCanvasMechanism,
  ];
} = {
  declaration: "slide-masters",
  slideProps: [],
  slot: "background",
  always: false,
  mechanisms: [
    slideMasterMechanism,
    slideLayoutMechanism,
    nativeThemeMechanism,
    colorMapMechanism,
    themeOverrideMechanism,
    masterSharedCanvasMechanism,
  ] as const,
  bindSlide(input, context) {
    const placeholders = input.props["placeholders"] as
      | Readonly<Record<string, LayoutPlaceholderSpec>>
      | undefined;
    const entries = resolvePlaceholderVisualEntries(
      placeholders,
      context.theme,
    );
    return {
      resolveSlide() {
        return {
          extensions: [],
          emit(emitContext) {
            if (entries.length === 0) {
              return { xml: "" };
            }
            const placeholderSpPrByKey: Record<string, string> = {};
            for (const entry of entries) {
              placeholderSpPrByKey[entry.key] = emitPlaceholderVisualSpPr(
                entry.visual,
                entry.flags,
                (resource) => emitContext.registerResource(resource),
              );
            }
            return {
              xml: "",
              placeholderSpPrByKey,
            };
          },
          paint(paintContext) {
            return {
              svg: paintPlaceholderVisualEntries(
                entries,
                paintContext.slideNumber,
              ),
            };
          },
        };
      },
    };
  },
};
