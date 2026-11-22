import type {
  SlideCapabilityModule,
  SlideLayoutProps,
  SlideMasterProps,
} from "@pptxdk/core/capability";
import { xmlEscape } from "@pptxdk/emit";
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
    model.showMasterPlaceholderAnimations === false
      ? ' showMasterPhAnim="0"'
      : "";
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

/** Capability declaration wiring native slide master and slide layout mechanisms into builds. */
export const mastersCapability = {
  declaration: "slide-masters",
  slideProp: "slideMasters",
  slideProps: [],
  slot: "background",
  always: false,
  mechanisms: [slideMasterMechanism, slideLayoutMechanism] as const,
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
