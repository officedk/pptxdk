import {
  defaultColorMap,
  defaultTheme,
  resolvePlaceholderBoxFromLayers,
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
import { mastersCodegen } from "./masters-codegen.ts";

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
  dimensions: { readonly width: number; readonly height: number },
): readonly ResolvedPlaceholderVisualEntry[] {
  if (placeholders === undefined) return [];
  const entries: ResolvedPlaceholderVisualEntry[] = [];
  for (const [key, spec] of Object.entries(placeholders)) {
    const resolvedBox = resolvePlaceholderBoxFromLayers([spec], dimensions);
    const hasBox =
      typeof resolvedBox.x === "number" &&
      typeof resolvedBox.y === "number" &&
      typeof resolvedBox.width === "number" &&
      typeof resolvedBox.height === "number";
    const styleRecord = spec.style as
      | Readonly<Record<string, unknown>>
      | undefined;
    const hasRotate = hasLocalDeclaration(styleRecord, "rotate");
    const hasFlipH =
      hasLocalDeclaration(styleRecord, "flipH") || spec.flipH === true;
    const hasFlipV =
      hasLocalDeclaration(styleRecord, "flipV") || spec.flipV === true;
    const hasGeometryOrRadius =
      spec.geometry !== undefined ||
      hasLocalDeclaration(styleRecord, "borderRadius");
    const hasFill =
      hasLocalDeclaration(styleRecord, "background") ||
      hasLocalDeclaration(styleRecord, "backgroundColor") ||
      hasLocalDeclaration(styleRecord, "backgroundImage");
    const hasBorder =
      hasLocalDeclaration(styleRecord, "border") ||
      (hasLocalDeclaration(styleRecord, "borderColor") &&
        spec.style?.borderWidth !== 0);
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
          x: resolvedBox.x as number,
          y: resolvedBox.y as number,
          width: resolvedBox.width as number,
          height: resolvedBox.height as number,
        }
      : { x: 0, y: 0, width: 100, height: 100 };
    const visual = resolveVisualModel({
      theme,
      style: styleRecord ?? {},
      box,
      ...(spec.geometry === undefined ? {} : { geometry: spec.geometry }),
      ...(spec.flipH === undefined ? {} : { flipH: spec.flipH }),
      ...(spec.flipV === undefined ? {} : { flipV: spec.flipV }),
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

function paintThemeCanvasBackdrop(themeContext: ThemeContext): PaintOutput {
  const bg = themeContext.resolveColor("bg1");
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

function paintSharedCanvasFragments(fragments: readonly string[]): PaintOutput {
  return { svg: fragments.join("") };
}

type SlideMasterEmitContext = {
  readonly spec: EmittedSlideMasterSpec;
  readonly layoutIds: readonly {
    readonly id: number;
    readonly rId: string;
  }[];
  readonly themeContext: ThemeContext;
  readonly colorMapXml: string;
};

/** Native canvas mechanism for `<p:sldMaster>` ownership, `preserve`, and master-layout relationships. */
export const slideMasterMechanism = defineCanvasMechanism<
  "slide-master",
  void,
  SlideMasterEmitContext,
  { readonly xml: string },
  readonly string[]
>({
  key: "slide-master",
  role: "canvas",
  emit: (_model, context) => ({
    xml: masterXml(
      context.spec,
      context.layoutIds,
      context.themeContext,
      context.colorMapXml,
    ),
  }),
  paint: (_model, fragments) => paintSharedCanvasFragments(fragments),
  import: {
    target: "deck",
    stage: "pre-slides",
    read: importDeckMasters,
  },
});

type SlideLayoutEmitContext = {
  readonly spec: EmittedSlideLayoutSpec;
  readonly themeContext: ThemeContext;
  readonly colorMapOvrXml: string;
};

/** Native canvas mechanism for `<p:sldLayout>` ownership, `type`, `preserve`, and idle layout retention. */
export const slideLayoutMechanism = defineCanvasMechanism<
  "slide-layout",
  void,
  SlideLayoutEmitContext,
  { readonly xml: string },
  readonly string[]
>({
  key: "slide-layout",
  role: "canvas",
  emit: (_model, context) => ({
    xml: layoutXml(context.spec, context.themeContext, context.colorMapOvrXml),
  }),
  paint: (_model, fragments) => paintSharedCanvasFragments(fragments),
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
};

/** Native canvas mechanism for `<a:theme>` (`clrScheme`, `fontScheme`, `fmtScheme`) part identity and mutation. */
export const nativeThemeMechanism = defineCanvasMechanism<
  "native-theme",
  NativeThemeMechanismModel,
  void,
  { readonly xml: string },
  ThemeContext
>({
  key: "native-theme",
  role: "canvas",
  emit: (model) => ({ xml: themeXml(model.theme) }),
  paint: (_model, themeContext) => paintThemeCanvasBackdrop(themeContext),
  import: {
    target: "deck",
    stage: "pre-slides",
    read: importDeckMasters,
  },
});

type ColorMapMechanismModel = {
  readonly colorMap?: ColorMap;
  readonly colorMapOverride?: ColorMapOverride;
};

/** Native canvas mechanism for master `<p:clrMap>` and layout/slide `<p:clrMapOvr>` color-slot mapping. */
export const colorMapMechanism = defineCanvasMechanism<
  "color-map",
  ColorMapMechanismModel,
  void,
  { readonly xml: string },
  ThemeContext
>({
  key: "color-map",
  role: "canvas",
  emit: (model) => ({
    xml: serializeColorMapModel(model),
  }),
  paint: (_model, themeContext) => paintThemeCanvasBackdrop(themeContext),
  import: {
    target: "deck",
    stage: "pre-slides",
    read: importDeckMasters,
  },
});

type ThemeOverrideMechanismModel = {
  readonly override: ThemeOverride;
  readonly baseTheme?: Theme;
};

/** Native canvas mechanism for `<a:themeOverride>` layout and slide palette/font/format overrides. */
export const themeOverrideMechanism = defineCanvasMechanism<
  "theme-override",
  ThemeOverrideMechanismModel,
  void,
  { readonly xml: string },
  ThemeContext
>({
  key: "theme-override",
  role: "canvas",
  emit: (model) => ({
    xml: themeOverrideXml(model.override, model.baseTheme ?? defaultTheme),
  }),
  paint: (_model, themeContext) => paintThemeCanvasBackdrop(themeContext),
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
  paint: (model) => paintSharedCanvasFragments(model.fragments),
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
    const colorMapXml = colorMapMechanism.emit({
      colorMap: spec.colorMap ?? defaultColorMap,
    }).xml;
    return slideMasterMechanism.emit(undefined, {
      spec,
      layoutIds,
      themeContext,
      colorMapXml,
    }).xml;
  },
  emitLayout(spec, themeContext) {
    const colorMapOvrXml = colorMapMechanism.emit(
      spec.colorMapOverride === undefined
        ? {}
        : { colorMapOverride: spec.colorMapOverride },
    ).xml;
    return slideLayoutMechanism.emit(undefined, {
      spec,
      themeContext,
      colorMapOvrXml,
    }).xml;
  },
  emitTheme(theme) {
    return nativeThemeMechanism.emit({ theme }).xml;
  },
  emitColorMap(model) {
    return colorMapMechanism.emit(model).xml;
  },
  emitThemeOverride(override, baseTheme) {
    return themeOverrideMechanism.emit({
      override,
      ...(baseTheme === undefined ? {} : { baseTheme }),
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
  codegen: mastersCodegen,
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
    paintBackdrop(input) {
      if (input.themeOverride !== undefined) {
        return themeOverrideMechanism.paint(
          { override: input.themeOverride },
          input.theme,
        );
      }
      if (
        input.colorMapOverride !== undefined ||
        input.hasCustomColorMap === true
      ) {
        return colorMapMechanism.paint(
          input.colorMapOverride === undefined
            ? {}
            : { colorMapOverride: input.colorMapOverride },
          input.theme,
        );
      }
      return nativeThemeMechanism.paint(
        { theme: input.theme.theme },
        input.theme,
      );
    },
    paintMasterLayers(fragments) {
      return slideMasterMechanism.paint(undefined, fragments);
    },
    paintLayoutLayers(fragments) {
      return slideLayoutMechanism.paint(undefined, fragments);
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
      context.dimensions,
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
