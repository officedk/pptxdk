import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  CommonNode,
  DeckContainer,
  SlideMasterContainer,
  type SlideLayoutObservation,
} from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { createSlideNode } from "@pptxdk/core/document";
import {
  A_NS,
  OPC_REL_NS,
  P_NS,
  REL_OFFICE_DOCUMENT,
  REL_SLIDE_LAYOUT,
  REL_SLIDE_MASTER,
  R_NS,
  createImporter,
} from "@pptxdk/import";
import { strToU8, zipSync } from "fflate";
import { describe, it } from "vitest";

import { testMastersCap } from "./test-masters-cap.ts";

describe("SlideMaster and SlideLayout import", () => {
  it("imports prs-slide-masters.pptx with 0 slides, 2 masters, preserve distinction, and 2 idle layouts", () => {
    const bytes = new Uint8Array(
      readFileSync(
        new URL("./fixtures/prs-slide-masters.pptx", import.meta.url),
      ),
    );
    const importer = createImporter([testMastersCap]);
    const result = importer(bytes);
    bytes.fill(0);

    assert.equal(result.report.slideCount, 0);
    assert.ok(result.document instanceof DeckContainer);
    assert.equal(result.document.masters.length, 2);

    const [master1, master2] = result.document.masters;
    assert.ok(master1 instanceof SlideMasterContainer);
    assert.ok(master2 instanceof SlideMasterContainer);
    assert.equal(master1.id, "master-1");
    assert.equal(master1.props["name"], undefined);
    assert.equal(master1.props["preserve"], undefined);
    assert.equal(master1.layouts.length, 1);
    assert.equal(master1.layouts[0]?.id, "Blank");
    assert.equal(master1.layouts[0]?.props["name"], "Blank");
    assert.equal(master1.layouts[0]?.props["type"], "blank");
    assert.equal(master1.layouts[0]?.props["preserve"], true);

    assert.equal(master2.id, "master-2");
    assert.equal(master2.props["name"], undefined);
    assert.equal(master2.props["preserve"], true);
    assert.equal(master2.layouts.length, 1);
    assert.equal(master2.layouts[0]?.id, "Blank");
    assert.equal(master2.layouts[0]?.props["name"], "Blank");
    assert.equal(master2.layouts[0]?.props["type"], "blank");
    assert.equal(master2.layouts[0]?.props["preserve"], true);

    // Resolve the 0-slide deck as-is.
    const resolve = createResolver([testMastersCap]);
    const zeroSlideResolution = resolve(result.document);
    assert.equal(zeroSlideResolution.report.slides.length, 0);
    assert.equal(zeroSlideResolution.report.masters?.length, 2);
    assert.equal(zeroSlideResolution.report.masters?.[0]?.preserve, undefined);
    assert.equal(zeroSlideResolution.report.masters?.[1]?.preserve, true);

    // Append a new slide referencing master 2's previously idle layout.
    const idleLayout = master2.layouts[0];
    result.document.append(
      createSlideNode({ id: "cover", layout: idleLayout }),
    );
    const withSlideResolution = resolve(result.document);
    assert.equal(withSlideResolution.report.slides.length, 1);
    assert.equal(withSlideResolution.report.slides[0]?.id, "cover");
    assert.equal(withSlideResolution.report.slides[0]?.layout, "Blank");
  });

  it("imports layouts.pptx retaining all 11 layouts (9 used + 2 idle) and preserve flags", () => {
    const bytes = new Uint8Array(
      readFileSync(new URL("./fixtures/layouts.pptx", import.meta.url)),
    );
    const importer = createImporter([testMastersCap]);
    const result = importer(bytes);
    bytes.fill(0);

    assert.equal(result.report.slideCount, 10);
    assert.ok(result.document instanceof DeckContainer);
    assert.equal(result.document.masters.length, 1);

    const master = result.document.masters[0]!;
    assert.equal(master.props["preserve"], undefined);
    assert.equal(master.layouts.length, 11);

    const layoutNames = master.layouts.map((layout) => layout.props["name"]);
    assert.deepEqual(layoutNames, [
      "Title Slide",
      "Title and Content",
      "Section Header",
      "Two Content",
      "Comparison",
      "Title Only",
      "Blank",
      "Content with Caption",
      "Picture with Caption",
      "Title and Vertical Text",
      "Vertical Title and Text",
    ]);

    for (const layout of master.layouts) {
      assert.equal(layout.props["preserve"], true);
    }

    // Idle layouts 10 and 11 can be queried directly from the master scope.
    const queriedIdle = result.document.query
      .get(master.id!)
      .get("Title and Vertical Text").node;
    assert.equal(queriedIdle, master.layouts[9]);

    // Slide 1..9 bind to layouts 1..9; slide 10 binds to layout 2 ("Title and Content").
    const slides = result.document.children.filter(
      (child): child is CommonNode =>
        child instanceof CommonNode && child.kind === "slide",
    );
    assert.equal(slides.length, 10);
    assert.equal(slides[0]?.props["layout"], master.layouts[0]);
    assert.equal(slides[9]?.props["layout"], master.layouts[1]);

    const resolve = createResolver([testMastersCap]);
    const resolved = resolve(result.document);
    const resolvedLayouts: readonly SlideLayoutObservation[] =
      resolved.report.masters?.[0]?.layouts ?? [];
    assert.equal(resolvedLayouts.length, 11);
    assert.equal(resolvedLayouts[9]?.type, "vertTx");
    assert.equal(resolvedLayouts[10]?.type, "vertTitleAndTx");

    // Master txStyles and placeholders + deterministic Comparison placeholder keys.
    assert.ok(master.props["textStyles"] !== undefined);
    assert.ok(master.props["placeholders"] !== undefined);
    const comparisonPlaceholders = master.layouts[4]?.props["placeholders"] as
      | Readonly<
          Record<string, { readonly type: string; readonly idx?: number }>
        >
      | undefined;
    assert.ok(comparisonPlaceholders !== undefined);
    assert.deepEqual(Object.keys(comparisonPlaceholders), [
      "title",
      "body-1",
      "content-2",
      "body-3",
      "content-4",
      "dt",
      "ftr",
      "sldNum",
    ]);
    assert.equal(comparisonPlaceholders["body-1"]?.type, "body");
    assert.equal(comparisonPlaceholders["body-1"]?.idx, 1);
    assert.equal(comparisonPlaceholders["content-2"]?.type, "obj");
    assert.equal(comparisonPlaceholders["content-2"]?.idx, 2);
    assert.equal(comparisonPlaceholders["body-3"]?.type, "body");
    assert.equal(comparisonPlaceholders["body-3"]?.idx, 3);
    assert.equal(comparisonPlaceholders["content-4"]?.type, "obj");
    assert.equal(comparisonPlaceholders["content-4"]?.idx, 4);
  });

  it("imports placeholder-layout-color.pptx with layout listStyles (35pt #00FFFF), master clrMap, and disambiguated duplicate title keys", () => {
    const bytes = new Uint8Array(
      readFileSync(
        new URL("./fixtures/placeholder-layout-color.pptx", import.meta.url),
      ),
    );
    const importer = createImporter([testMastersCap]);
    const result = importer(bytes);
    bytes.fill(0);

    assert.equal(result.report.slideCount, 1);
    assert.ok(result.document instanceof DeckContainer);
    const master = result.document.masters[0]!;
    const layout = master.layouts[0]!;
    const layoutPlaceholders = layout.props["placeholders"] as
      | Readonly<
          Record<
            string,
            {
              readonly type: string;
              readonly idx?: number;
              readonly hasCustomPrompt?: boolean;
              readonly listStyles?: Readonly<
                Record<
                  number,
                  {
                    readonly fontSize?: string;
                    readonly color?: string;
                    readonly fontFamily?: string;
                  }
                >
              >;
            }
          >
        >
      | undefined;
    assert.ok(layoutPlaceholders !== undefined);
    assert.deepEqual(Object.keys(layoutPlaceholders), ["title", "title-2"]);
    assert.equal(
      layoutPlaceholders["title"]?.listStyles?.[1]?.fontSize,
      "35pt",
    );
    assert.equal(
      layoutPlaceholders["title"]?.listStyles?.[1]?.color,
      "#00FFFF",
    );
    assert.equal(
      layoutPlaceholders["title"]?.listStyles?.[1]?.fontFamily,
      "major",
    );
    assert.equal(layoutPlaceholders["title-2"]?.idx, 2);
    assert.equal(layoutPlaceholders["title-2"]?.hasCustomPrompt, true);
    assert.equal(
      layoutPlaceholders["title-2"]?.listStyles?.[1]?.fontSize,
      "80pt",
    );
    assert.equal(
      layoutPlaceholders["title-2"]?.listStyles?.[1]?.color,
      "#FF0000",
    );
  });

  it("disambiguates duplicate display names deterministically within a master while preserving display names", () => {
    const pptx = zipSync({
      "_rels/.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>`,
      ),
      "ppt/presentation.xml": strToU8(
        `<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="slideMasters/slideMaster1.xml"/></Relationships>`,
      ),
      "ppt/slideMasters/slideMaster1.xml": strToU8(
        `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld name="Brand"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/><p:sldLayoutId id="2147483650" r:id="rId2"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`,
      ),
      "ppt/slideMasters/_rels/slideMaster1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout2.xml"/></Relationships>`,
      ),
      "ppt/slideLayouts/slideLayout1.xml": strToU8(
        `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="cust" userDrawn="1"><p:cSld name="Shared Name"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
      ),
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
      ),
      "ppt/slideLayouts/slideLayout2.xml": strToU8(
        `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="blank" preserve="1"><p:cSld name="Shared Name"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
      ),
      "ppt/slideLayouts/_rels/slideLayout2.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
      ),
    });

    const importer = createImporter([testMastersCap]);
    const result = importer(pptx);
    assert.deepEqual(result.report.diagnostics, []);
    assert.ok(result.document instanceof DeckContainer);
    const master = result.document.masters[0]!;
    assert.equal(master.id, "Brand");
    assert.equal(master.props["name"], "Brand");
    assert.equal(master.layouts.length, 2);

    const [first, second] = master.layouts;
    assert.equal(first?.id, "Shared Name");
    assert.equal(first?.props["name"], "Shared Name");
    assert.equal(first?.props["userDrawn"], true);
    assert.equal(second?.id, "Shared Name-2");
    assert.equal(second?.props["name"], "Shared Name");
    assert.equal(second?.props["preserve"], true);
  });

  it("imports placeholder-layout-color.pptx headerFooter flags on SlideMaster", () => {
    const bytes = new Uint8Array(
      readFileSync(
        new URL("./fixtures/placeholder-layout-color.pptx", import.meta.url),
      ),
    );
    const importer = createImporter([testMastersCap]);
    const result = importer(bytes);
    bytes.fill(0);

    assert.ok(result.document instanceof DeckContainer);
    const master = result.document.masters[0]!;
    assert.deepEqual(
      { ...(master.props["headerFooter"] as Record<string, unknown>) },
      {
        dateTime: false,
        footer: false,
        header: false,
        slideNumber: false,
      },
    );
  });
});
