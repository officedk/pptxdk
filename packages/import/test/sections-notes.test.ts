import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  CommonNode,
  DeckContainer,
  sectionPropsSchema,
  slideCorePropsSchema,
} from "@pptxdk/core";
import {
  A_NS,
  OPC_REL_NS,
  P14_NS,
  P14_SECTION_LIST_EXT_URI,
  P_NS,
  REL_HANDOUT_MASTER,
  REL_NOTES_MASTER,
  REL_OFFICE_DOCUMENT,
  REL_SLIDE,
  REL_SLIDE_LAYOUT,
  REL_SLIDE_MASTER,
  REL_THEME,
  R_NS,
  createImporter,
  importDiagnosticSchema,
  defineShowMechanism,
  importPresentationSections,
  type DeckImportContext,
} from "@pptxdk/import";
import { strToU8, unzipSync, zipSync } from "fflate";
import { describe, it } from "vitest";

const testSectionsCap = {
  declaration: "test-sections",
  slideProps: [] as const,
  slot: "timing" as const,
  mechanisms: [
    defineShowMechanism({
      key: "test-sections",
      role: "show",
      emit: () => ({ xml: "" }),
      import: {
        target: "deck" as const,
        stage: "post-slides" as const,
        read(context: DeckImportContext) {
          const diagnostics: Parameters<typeof importPresentationSections>[4] =
            [];
          importPresentationSections(
            context.presentationCursor,
            context.deckNode,
            context.slideNodeByPresentationSldId ?? new Map(),
            context.package.presentationPath,
            diagnostics,
          );
          return diagnostics.length === 0 ? {} : { diagnostics };
        },
      },
    }),
  ],
  bindSlide: () => ({
    resolveSlide: () => ({ extensions: [], emit: () => ({ xml: "" }) }),
  }),
};

function blankSlideXml(name: string, extraAttrs = ""): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"${extraAttrs}>
  <p:cSld name="${name}">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
}

function baselineNotesMasterXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notesMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:cSld>
    <p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp><p:nvSpPr><p:cNvPr id="2" name="Header Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="hdr" sz="quarter"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="3" name="Date Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Image Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="5" name="Notes Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="6" name="Footer Placeholder 5"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="4"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="7" name="Slide Number Placeholder 6"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="5"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:notesStyle><a:lvl1pPr><a:defRPr sz="1200"/></a:lvl1pPr></p:notesStyle>
</p:notesMaster>`;
}

function buildMasterLayoutEntries(): Record<string, Uint8Array> {
  return {
    "ppt/slideMasters/slideMaster1.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>`,
    ),
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${OPC_REL_NS}">
  <Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`,
    ),
    "ppt/slideLayouts/slideLayout1.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="blank">
  <p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`,
    ),
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${OPC_REL_NS}">
  <Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`,
    ),
  };
}

describe("Presentation slide order, hidden state, sections, and notes/handout master inspection", () => {
  it("preserves non-sequential sldIdLst order, show='0'/'false' hidden flags, and p14:sectionLst grouping", () => {
    assert.equal(sectionPropsSchema.safeParse({ name: "Intro" }).success, true);
    assert.equal(
      slideCorePropsSchema.safeParse({ id: "s1", hidden: true }).success,
      true,
    );

    const pptx = zipSync({
      "_rels/.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>`,
      ),
      "ppt/presentation.xml": strToU8(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="263" r:id="rId4"/>
    <p:sldId id="256" r:id="rId2"/>
    <p:sldId id="260" r:id="rId3"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="5143500"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:extLst>
    <p:ext uri="${P14_SECTION_LIST_EXT_URI}">
      <p14:sectionLst xmlns:p14="${P14_NS}">
        <p14:section name="Opening" id="{11111111-1111-1111-1111-111111111111}">
          <p14:sldIdLst>
            <p14:sldId id="263"/>
          </p14:sldIdLst>
        </p14:section>
        <p14:section name="Details" id="{22222222-2222-2222-2222-222222222222}">
          <p14:sldIdLst>
            <p14:sldId id="256"/>
            <p14:sldId id="260"/>
          </p14:sldIdLst>
        </p14:section>
      </p14:sectionLst>
    </p:ext>
  </p:extLst>
</p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": strToU8(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${OPC_REL_NS}">
  <Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="${REL_SLIDE}" Target="slides/slide1.xml"/>
  <Relationship Id="rId3" Type="${REL_SLIDE}" Target="slides/slide2.xml"/>
  <Relationship Id="rId4" Type="${REL_SLIDE}" Target="slides/slide3.xml"/>
</Relationships>`,
      ),
      ...buildMasterLayoutEntries(),
      "ppt/slides/slide1.xml": strToU8(
        blankSlideXml("mid-hidden", ' show="0"'),
      ),
      "ppt/slides/_rels/slide1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
      ),
      "ppt/slides/slide2.xml": strToU8(
        blankSlideXml("end-hidden", ' show="false"'),
      ),
      "ppt/slides/_rels/slide2.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
      ),
      "ppt/slides/slide3.xml": strToU8(
        blankSlideXml("cover-first", ' show="1"'),
      ),
      "ppt/slides/_rels/slide3.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
      ),
    });

    const importer = createImporter([testSectionsCap]);
    const result = importer(pptx);
    assert.deepEqual(result.report.diagnostics, []);
    assert.equal(result.report.slideCount, 3);
    assert.ok(result.document instanceof DeckContainer);

    const sections = result.document.children.filter(
      (child): child is CommonNode =>
        child instanceof CommonNode && child.type === "section",
    );
    assert.equal(sections.length, 2);
    assert.equal(sections[0]?.props["name"], "Opening");
    assert.equal(sections[1]?.props["name"], "Details");

    const openingSlides = sections[0].children.filter(
      (child): child is CommonNode => child instanceof CommonNode,
    );
    const detailsSlides = sections[1].children.filter(
      (child): child is CommonNode => child instanceof CommonNode,
    );
    assert.deepEqual(
      openingSlides.map((s) => ({
        id: s.id,
        hidden: s.props["hidden"] ?? false,
      })),
      [{ id: "cover-first", hidden: false }],
    );
    assert.deepEqual(
      detailsSlides.map((s) => ({
        id: s.id,
        hidden: s.props["hidden"] ?? false,
      })),
      [
        { id: "mid-hidden", hidden: true },
        { id: "end-hidden", hidden: true },
      ],
    );
  });

  it("imports standard baseline notesMaster1.xml with zero false loss diagnostics", () => {
    const pptx = zipSync({
      "_rels/.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>`,
      ),
      "ppt/presentation.xml": strToU8(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:notesMasterIdLst><p:notesMasterId r:id="rId3"/></p:notesMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="5143500"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": strToU8(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${OPC_REL_NS}">
  <Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="${REL_SLIDE}" Target="slides/slide1.xml"/>
  <Relationship Id="rId3" Type="${REL_NOTES_MASTER}" Target="notesMasters/notesMaster1.xml"/>
</Relationships>`,
      ),
      ...buildMasterLayoutEntries(),
      "ppt/slides/slide1.xml": strToU8(blankSlideXml("s1")),
      "ppt/slides/_rels/slide1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
      ),
      "ppt/notesMasters/notesMaster1.xml": strToU8(baselineNotesMasterXml()),
      "ppt/notesMasters/_rels/notesMaster1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_THEME}" Target="../theme/theme2.xml"/></Relationships>`,
      ),
    });

    const importer = createImporter([testSectionsCap]);
    const result = importer(pptx);
    assert.deepEqual(result.report.diagnostics, []);
  });

  it("reports structured custom-print-layout loss for custom notesMaster graphics/bgPr and handoutMaster1.xml", () => {
    const pptx = zipSync({
      "_rels/.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>`,
      ),
      "ppt/presentation.xml": strToU8(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:notesMasterIdLst><p:notesMasterId r:id="rId3"/></p:notesMasterIdLst>
  <p:handoutMasterIdLst><p:handoutMasterId r:id="rId4"/></p:handoutMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="5143500"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": strToU8(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${OPC_REL_NS}">
  <Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="${REL_SLIDE}" Target="slides/slide1.xml"/>
  <Relationship Id="rId3" Type="${REL_NOTES_MASTER}" Target="notesMasters/notesMaster1.xml"/>
  <Relationship Id="rId4" Type="${REL_HANDOUT_MASTER}" Target="handoutMasters/handoutMaster1.xml"/>
</Relationships>`,
      ),
      ...buildMasterLayoutEntries(),
      "ppt/slides/slide1.xml": strToU8(blankSlideXml("s1")),
      "ppt/slides/_rels/slide1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
      ),
      "ppt/notesMasters/notesMaster1.xml": strToU8(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notesMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="F1F5F9"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder 1"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="9" name="Custom Watermark"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:notesStyle/>
</p:notesMaster>`,
      ),
      "ppt/handoutMasters/handoutMaster1.xml": strToU8(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:handoutMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp><p:nvSpPr><p:cNvPr id="5" name="Handout Banner"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:handoutStyle/>
</p:handoutMaster>`,
      ),
    });

    const importer = createImporter([]);
    const result = importer(pptx);
    const diags = result.report.diagnostics.map((d) =>
      importDiagnosticSchema.parse(d),
    );
    const notesMasterDiags = diags.filter(
      (d): d is Extract<typeof d, { readonly code: "unsupported-mechanism" }> =>
        d.code === "unsupported-mechanism" &&
        d.location.part === "ppt/notesMasters/notesMaster1.xml",
    );
    assert.equal(notesMasterDiags.length, 2);
    for (const d of notesMasterDiags) {
      assert.equal(d.loss?.lost, "custom-print-layout");
      assert.equal(d.loss?.retainedAs, "default-notes-master");
    }
    assert.ok(
      notesMasterDiags.some(
        (d) =>
          d.location.objectId === 9 &&
          d.location.objectName === "Custom Watermark",
      ),
    );

    const handoutDiags = diags.filter(
      (d): d is Extract<typeof d, { readonly code: "unsupported-mechanism" }> =>
        d.code === "unsupported-mechanism" &&
        d.location.part === "ppt/handoutMasters/handoutMaster1.xml",
    );
    assert.equal(handoutDiags.length, 1);
    assert.equal(handoutDiags[0]?.loss?.lost, "custom-print-layout");
    assert.equal(handoutDiags[0]?.loss?.retainedAs, "omitted");
    assert.equal(handoutDiags[0]?.location.objectName, "Handout Banner");
  });

  it("reads prs-notes.pptx fixture, zero false loss on standard placeholders, and flags Picture 7 on notesMaster1.xml", () => {
    const raw = new Uint8Array(
      readFileSync(new URL("./fixtures/prs-notes.pptx", import.meta.url)),
    );
    const importer = createImporter([]);
    const result = importer(raw);
    raw.fill(0);

    assert.equal(result.report.slideCount, 1);
    const notesMasterLosses = result.report.diagnostics.filter(
      (d) => d.location.part === "ppt/notesMasters/notesMaster1.xml",
    );
    assert.equal(notesMasterLosses.length, 1);
    assert.equal(notesMasterLosses[0]?.code, "unsupported-mechanism");
    if (notesMasterLosses[0]?.code === "unsupported-mechanism") {
      assert.equal(notesMasterLosses[0].loss?.lost, "custom-print-layout");
      assert.equal(notesMasterLosses[0].location.objectName, "Picture 7");
      assert.equal(notesMasterLosses[0].location.objectId, 8);
    }

    // Verify layouts.pptx baseline notesMaster1.xml produces zero notesMaster diagnostics.
    const layoutsRaw = new Uint8Array(
      readFileSync(new URL("./fixtures/layouts.pptx", import.meta.url)),
    );
    const layoutsResult = importer(layoutsRaw);
    layoutsRaw.fill(0);
    const layoutsNotesDiags = layoutsResult.report.diagnostics.filter((d) =>
      d.location.part.startsWith("ppt/notesMasters/"),
    );
    assert.deepEqual(layoutsNotesDiags, []);
    void unzipSync;
  });
});
