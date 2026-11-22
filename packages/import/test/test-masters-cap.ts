import type { SlideCapabilityModule } from "@pptxdk/core/capability";
import {
  defineCanvasMechanism,
  importMastersAndLayouts,
  type DeckImportContext,
} from "@pptxdk/import";

export const testMastersCap: SlideCapabilityModule = {
  declaration: "slide-masters",
  slideProps: [],
  slot: "background",
  mechanisms: [
    defineCanvasMechanism({
      key: "test-masters",
      role: "canvas",
      emit: () => ({ xml: "" }),
      paint: () => ({ svg: "" }),
      import: {
        target: "deck" as const,
        stage: "pre-slides" as const,
        read(context: DeckImportContext) {
          const diagnostics: Parameters<typeof importMastersAndLayouts>[5] = [];
          const layouts = importMastersAndLayouts(
            context.presentationCursor,
            context.package,
            context.relById ?? new Map(),
            context.deckNode,
            context.usedDeckIds ?? new Set(),
            diagnostics,
            context.slideHandlers ?? [],
            context.shapeTreeHandlers ?? [],
            context.registerDeferredSlideLink ?? (() => {}),
            context.clrMapByLayoutPartPath ?? new Map(),
          );
          if (context.layoutNodeByPartPath !== undefined) {
            for (const [k, v] of layouts) {
              context.layoutNodeByPartPath.set(k, v);
            }
          }
          return diagnostics.length === 0 ? {} : { diagnostics };
        },
      },
    }),
  ],
  bindSlide: () => ({
    resolveSlide: () => ({ extensions: [], emit: () => ({ xml: "" }) }),
  }),
};
