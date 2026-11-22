import {
  defaultColorMap,
  defaultTheme,
  type ColorMap,
  type ColorMapOverride,
  type SlideCapabilityModule,
  type SlideLayoutProps,
  type SlideMasterProps,
  type Theme,
  type ThemeOverride,
} from "@pptxdk/core/capability";
import {
  colorMapAttributesXml,
  colorMapOverrideXml,
  themeOverrideXml,
  themeXml,
  xmlEscape,
} from "@pptxdk/emit";
import {
  defineCanvasMechanism,
  importMastersAndLayouts,
  type DeckImportContext,
} from "@pptxdk/import";

function serializeSlideMasterXml(model: SlideMasterProps): string {
  const preserveAttr =
    model.preserve === undefined
      ? ""
      : ` preserve="${model.preserve ? "1" : "0"}"`;
  const nameAttr =
    model.name === undefined ? "" : ` name="${xmlEscape(model.name)}"`;
  return `<p:sldMaster${preserveAttr}><p:cSld${nameAttr}/></p:sldMaster>`;
}

function serializeSlideLayoutXml(model: SlideLayoutProps): string {
  const typeAttr =
    model.type === undefined ? "" : ` type="${xmlEscape(model.type)}"`;
  const preserveAttr =
    model.preserve === undefined
      ? ""
      : ` preserve="${model.preserve ? "1" : "0"}"`;
  const matchingNameAttr =
    model.matchingName === undefined
      ? ""
      : ` matchingName="${xmlEscape(model.matchingName)}"`;
  const userDrawnAttr =
    model.userDrawn === undefined
      ? ""
      : ` userDrawn="${model.userDrawn ? "1" : "0"}"`;
  const showMasterSpAttr =
    model.showMasterShapes === false ? ' showMasterSp="0"' : "";
  const showMasterPhAnimAttr =
    model.showMasterPlaceholderAnimations === undefined
      ? ""
      : ` showMasterPhAnim="${model.showMasterPlaceholderAnimations ? "1" : "0"}"`;
  const nameAttr =
    model.name === undefined ? "" : ` name="${xmlEscape(model.name)}"`;
  return `<p:sldLayout${typeAttr}${preserveAttr}${matchingNameAttr}${userDrawnAttr}${showMasterSpAttr}${showMasterPhAnimAttr}><p:cSld${nameAttr}/></p:sldLayout>`;
}

function importDeckMastersIfUnpopulated(context: DeckImportContext) {
  if ("masters" in context.deckNode) {
    return {};
  }
  void importMastersAndLayouts;
  return {};
}

/** Native canvas mechanism for `<p:sldMaster>` ownership, `preserve`, and master-layout relationships. */
export const slideMasterMechanism = defineCanvasMechanism({
  key: "slide-master",
  role: "canvas",
  emit: (model: SlideMasterProps) => ({ xml: serializeSlideMasterXml(model) }),
  paint: () => ({ svg: "" }),
  import: {
    target: "deck",
    read: importDeckMastersIfUnpopulated,
  },
});

/** Native canvas mechanism for `<p:sldLayout>` ownership, `type`, `preserve`, and idle layout retention. */
export const slideLayoutMechanism = defineCanvasMechanism({
  key: "slide-layout",
  role: "canvas",
  emit: (model: SlideLayoutProps) => ({ xml: serializeSlideLayoutXml(model) }),
  paint: () => ({ svg: "" }),
  import: {
    target: "deck",
    read: importDeckMastersIfUnpopulated,
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

/** Native canvas mechanism for `<a:theme>` (`clrScheme`, `fontScheme`, `fmtScheme`) part identity and mutation. */
export const nativeThemeMechanism = defineCanvasMechanism({
  key: "native-theme",
  role: "canvas",
  emit: (model: Theme) => ({ xml: themeXml(model) }),
  paint: () => ({ svg: "" }),
  import: {
    target: "deck",
    read: importDeckMastersIfUnpopulated,
  },
});

/** Native canvas mechanism for master `<p:clrMap>` and layout/slide `<p:clrMapOvr>` color-slot mapping. */
export const colorMapMechanism = defineCanvasMechanism({
  key: "color-map",
  role: "canvas",
  emit: (model: {
    readonly colorMap?: ColorMap;
    readonly colorMapOverride?: ColorMapOverride;
  }) => ({ xml: serializeColorMapModel(model) }),
  paint: () => ({ svg: "" }),
  import: {
    target: "deck",
    read: importDeckMastersIfUnpopulated,
  },
});

/** Native canvas mechanism for `<a:themeOverride>` layout and slide palette/font/format overrides. */
export const themeOverrideMechanism = defineCanvasMechanism({
  key: "theme-override",
  role: "canvas",
  emit: (model: {
    readonly override: ThemeOverride;
    readonly baseTheme?: Theme;
  }) => ({
    xml: themeOverrideXml(model.override, model.baseTheme ?? defaultTheme),
  }),
  paint: () => ({ svg: "" }),
  import: {
    target: "deck",
    read: importDeckMastersIfUnpopulated,
  },
});

/** Capability declaration wiring native slide master, layout, theme, color-map, and theme-override mechanisms into builds. */
export const mastersCapability = {
  declaration: "slide-masters",
  slideProp: "slideMasters",
  slideProps: [],
  slot: "background",
  always: false,
  mechanisms: [
    slideMasterMechanism,
    slideLayoutMechanism,
    nativeThemeMechanism,
    colorMapMechanism,
    themeOverrideMechanism,
  ] as const,
  bindSlide() {
    return {
      resolveSlide() {
        return {
          extensions: [],
          emit: () => ({ xml: "" }),
        };
      },
    };
  },
} satisfies SlideCapabilityModule;
