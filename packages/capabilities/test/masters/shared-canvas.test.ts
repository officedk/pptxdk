import { documentCapabilities } from "@pptxdk/capabilities/all";
import { createResolver } from "@pptxdk/core/build";
import { describe, it } from "vitest";

import { verifyMasterLogoSingleEditPropagates } from "./shared-canvas-mechanisms.ts";

describe("SlideMaster and SlideLayout shared canvas", () => {
  it("updates associated slides when moving the master Logo once without modifying slide local overrides", async () => {
    const resolve = createResolver(documentCapabilities);
    await verifyMasterLogoSingleEditPropagates(documentCapabilities, resolve);
  });
});
