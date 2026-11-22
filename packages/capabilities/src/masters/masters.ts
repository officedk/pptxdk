import {
  defaultTheme,
  visualStyleSchema,
  type ColorMap,
  type ColorMapOverride,
  type LayoutPlaceholderSpec,
  type NativeColorSlot,
  type PaintOutput,
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
  type MastersEmitDriver,
} from "@pptxdk/emit";
import {
  defineCanvasMechanism,
  importMastersAndLayouts,
  readPlaceholderSpPrVisuals,
  type DeckImportContext,
  type DeckImportOutcome,
  type ImportDiagnostic,
  type PlaceholderVisualReader,
} from "@pptxdk/import";

import { hasLocalDeclaration } from "../shared/style-inheritance.ts";
import {
  emitPlaceholderVisualSpPr,
  readSpPrAppearance,
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
};

const readMasterPlaceholderVisuals: PlaceholderVisualReader = (args) =>
  readPlaceholderSpPrVisuals(args, {
    readAppearance: readSpPrAppearance,
    allowedStyleKeys: supportedPlaceholderVisualStyleKeys,
  });

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
      ...(spec.geometry === undefined ? {} : { geometry: spec.geometry }),
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
    });
  }
  return entries;
}

function paintThemeCanvasBackdrop(model: {
  readonly themeContext: ThemeContext;
}): PaintOutput {
  const bg = model.themeContext.resolveColor("bg1");
  if (
    bg.hex.toUpperCase() === "FFFFFF" &&
    (bg.alpha === undefined || bg.alpha === 1)
  ) {
    return { svg: '<rect width="100%" height="100%" fill="white"/>' };
  }
  const opacityAttr =
    bg.alpha === undefined || bg.alpha === 1
      ? ""
      : ` fill-opacity="${bg.alpha}"`;
  return {
    svg: `<rect width="100%" height="100%" fill="#${bg.hex}"${opacityAttr}/>`,
  };
}

function paintSharedCanvasFragments(model: {
  readonly fragments: readonly string[];
}): PaintOutput {
  return { svg: model.fragments.join("") };
}

type SlideMasterMechanismModel =
  | {
      readonly spec: EmittedSlideMasterSpec;
      readonly layoutIds: readonly {
        readonly id: number;
        readonly rId: string;
      }[];
      readonly themeContext: ThemeContext;
      readonly fragments: readonly string[];
    }
  | { readonly fragments: readonly string[] };

/** Native canvas mechanism for `<p:sldMaster>` ownership, `preserve`, and master-layout relationships. */
export const slideMasterMechanism = defineCanvasMechanism<
  "slide-master",
  SlideMasterMechanismModel,
  void,
  { readonly xml: string },
  void
>({
  key: "slide-master",
  role: "canvas",
  emit: (model) => {
    if (!("spec" in model)) {
      throw new Error(
        "SlideMaster emit requires spec, layoutIds, and themeContext",
      );
    }
    return { xml: masterXml(model.spec, model.layoutIds, model.themeContext) };
  },
  paint: paintSharedCanvasFragments,
  import: {
    target: "deck",
    stage: "pre-slides",
    read: importDeckMasters,
  },
});

type SlideLayoutMechanismModel =
  | {
      readonly spec: EmittedSlideLayoutSpec;
      readonly themeContext: ThemeContext;
      readonly fragments: readonly string[];
    }
  | { readonly fragments: readonly string[] };

/** Native canvas mechanism for `<p:sldLayout>` ownership, `type`, `preserve`, and idle layout retention. */
export const slideLayoutMechanism = defineCanvasMechanism<
  "slide-layout",
  SlideLayoutMechanismModel,
  void,
  { readonly xml: string },
  void
>({
  key: "slide-layout",
  role: "canvas",
  emit: (model) => {
    if (!("spec" in model)) {
      throw new Error("SlideLayout emit requires spec and themeContext");
    }
    return { xml: layoutXml(model.spec, model.themeContext) };
  },
  paint: paintSharedCanvasFragments,
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
  if (model.colorMap !== undefined) {
    return `<p:clrMap ${colorMapAttributesXml(model.colorMap)}/>`;
  }
  return colorMapOverrideXml(model.colorMapOverride);
}

type NativeThemeMechanismModel = {
  readonly theme: Theme;
  readonly themeContext: ThemeContext;
};

/** Native canvas mechanism for `<a:theme>` (`clrScheme`, `fontScheme`, `fmtScheme`) part identity and mutation. */
export const nativeThemeMechanism = defineCanvasMechanism<
  "native-theme",
  NativeThemeMechanismModel,
  void,
  { readonly xml: string },
  void
>({
  key: "native-theme",
  role: "canvas",
  emit: (model) => ({ xml: themeXml(model.theme) }),
  paint: paintThemeCanvasBackdrop,
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
export const colorMapMechanism = defineCanvasMechanism<
  "color-map",
  ColorMapMechanismModel,
  void,
  { readonly xml: string },
  void
>({
  key: "color-map",
  role: "canvas",
  emit: (model) => ({
    xml: serializeColorMapModel(model),
  }),
  paint: paintThemeCanvasBackdrop,
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
export const themeOverrideMechanism = defineCanvasMechanism<
  "theme-override",
  ThemeOverrideMechanismModel,
  void,
  { readonly xml: string },
  void
>({
  key: "theme-override",
  role: "canvas",
  emit: (model) => ({
    xml: themeOverrideXml(model.override, model.baseTheme ?? defaultTheme),
  }),
  paint: paintThemeCanvasBackdrop,
  import: {
    target: "deck",
    stage: "pre-slides",
    read: importDeckMasters,
  },
});

type MasterSharedCanvasMechanismModel = {
  readonly fragments: readonly string[];
};

/** Native canvas mechanism for non-placeholder shared canvas objects (`sp`/`pic`/`grpSp`) on SlideMaster and SlideLayout. */
export const masterSharedCanvasMechanism = defineCanvasMechanism<
  "master-shared-canvas",
  MasterSharedCanvasMechanismModel,
  void,
  { readonly xml: string },
  void
>({
  key: "master-shared-canvas",
  role: "canvas",
  emit: (model) => ({
    xml: model.fragments.join(""),
  }),
  paint: paintSharedCanvasFragments,
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

const mastersEmitDriver: MastersEmitDriver = {
  emitMaster(spec, layoutIds, themeContext) {
    return slideMasterMechanism.emit({
      spec,
      layoutIds,
      themeContext,
      fragments: [],
    }).xml;
  },
  emitLayout(spec, themeContext) {
    return slideLayoutMechanism.emit({ spec, themeContext, fragments: [] }).xml;
  },
  emitTheme(theme) {
    return nativeThemeMechanism.emit({
      theme,
      themeContext: undefined as never,
    }).xml;
  },
  emitColorMap(model) {
    return colorMapMechanism.emit({
      ...model,
      themeContext: undefined as never,
    }).xml;
  },
  emitThemeOverride(override, baseTheme) {
    return themeOverrideMechanism.emit({
      override,
      ...(baseTheme === undefined ? {} : { baseTheme }),
      themeContext: undefined as never,
    }).xml;
  },
  emitSharedCanvas(fragments) {
    return masterSharedCanvasMechanism.emit({ fragments }).xml;
  },
};

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
  mastersCanvas: {
    paintBackdrop(themeContext) {
      return colorMapMechanism.paint({ themeContext });
    },
    paintMasterLayers(fragments) {
      return slideMasterMechanism.paint({ fragments });
    },
    paintLayoutLayers(fragments) {
      return slideLayoutMechanism.paint({ fragments });
    },
    paintSharedObject(svg) {
      return masterSharedCanvasMechanism.paint({ fragments: [svg] });
    },
  },
  mastersEmitDriver,
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
        };
      },
    };
  },
};
