import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import {
  deck,
  font,
  notes,
  paragraph,
  section,
  shape,
  slide,
  slideNumber,
  span,
  text,
} from "@pptxdk/capabilities/document";
import { notesDeclarationPropsSchema } from "@pptxdk/capabilities/notes";
import { CommonContainer, CommonNode, DeckContainer } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import {
  A_NS,
  OPC_REL_NS,
  P14_NS,
  P14_SECTION_LIST_EXT_URI,
  P_NS,
  REL_NOTES_MASTER,
  REL_NOTES_SLIDE,
  REL_OFFICE_DOCUMENT,
  REL_SLIDE,
  REL_SLIDE_LAYOUT,
  REL_SLIDE_MASTER,
  R_NS,
  createImporter,
} from "@pptxdk/import";
import { paint } from "@pptxdk/preview";
import { validatePptx } from "@pptxdk/validation";
import { strToU8, zipSync } from "fflate";
import { describe, it } from "vitest";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
  {
    weight: 700,
    data: readFileSync(
      new URL("../../../test/assets/fonts/Jost-700-Bold.ttf", import.meta.url),
    ),
  },
]);

const resolveDocument = createResolver(documentCapabilities);
const importDocument = createImporter(documentCapabilities);

describe("Speaker notes, slide order, hidden slides, and sections roundtrip", () => {
  it("validates pure-data NotesDeclarationProps via Zod schema", () => {
    assert.equal(
      notesDeclarationPropsSchema.safeParse({
        id: "n1",
        style: { fontSize: "12pt" },
      }).success,
      true,
    );
    assert.equal(
      notesDeclarationPropsSchema.safeParse({ id: "n1", layout: { x: 1 } })
        .success,
      false,
    );
  });

  it("roundtrips sections, hidden slides, speaker notes formatting, and keeps notes attached after post-import reorder & edit", async () => {
    const original = deck({
      width: 640,
      height: 360,
      theme: { fonts: { major: jostFont, minor: jostFont } },
    }).append(
      section({ name: "Part One" }).append(
        slide({ id: "intro" }).append(
          text("Slide One Title", {
            id: "s1-title",
            style: {
              position: "absolute",
              left: 32,
              top: 32,
              width: 300,
              height: 48,
              fontFamily: jostFont,
              fontSize: "20pt",
            },
          }),
          notes([
            paragraph("First slide cue"),
            paragraph([
              span("Important: ", {
                style: { fontWeight: "bold", color: "#DC2626" },
              }),
              "remember pause",
            ]),
            paragraph(["Page ", slideNumber()], {
              style: { list: { type: "bullet" }, align: "center" },
            }),
          ]),
        ),
        slide({ id: "secret", hidden: true }).append(
          shape({
            id: "secret-box",
            style: {
              position: "absolute",
              left: 40,
              top: 40,
              width: 120,
              height: 80,
              background: "#334155",
            },
          }),
          notes("Hidden slide notes"),
        ),
      ),
      section({ name: "Part Two" }).append(
        slide({ id: "closing" }).append(
          shape({
            id: "closing-box",
            style: {
              position: "absolute",
              left: 40,
              top: 40,
              width: 120,
              height: 80,
              background: "#059669",
            },
          }),
        ),
      ),
    );

    const initialPptx = emit(resolveDocument(original), {
      embedFonts: "subset",
    }).pptx;
    const copy = new Uint8Array(initialPptx);
    const imported = importDocument(copy);
    copy.fill(0);

    assert.deepEqual(imported.report.diagnostics, []);
    assert.ok(imported.document instanceof DeckContainer);

    const sections = imported.document.children.filter(
      (child): child is CommonContainer =>
        child instanceof CommonContainer && child.type === "section",
    );
    assert.deepEqual(
      sections.map((s) => s.props["name"]),
      ["Part One", "Part Two"],
    );

    const [partOne, partTwo] = sections;
    const s1 = partOne!.children[0] as CommonNode;
    const s2 = partOne!.children[1] as CommonNode;
    const s3 = partTwo!.children[0] as CommonNode;
    assert.equal(s1.id, "intro");
    assert.equal(s2.id, "secret");
    assert.equal(s2.props["hidden"], true);
    assert.equal(s3.id, "closing");

    // Edit Slide 1 notes text and reorder Slide 1 into Part Two after Closing.
    const s1Notes = s1.children.find(
      (child): child is CommonNode =>
        child instanceof CommonNode && child.type === "notes",
    );
    assert.ok(s1Notes !== undefined);
    (s1Notes as unknown as { content(...items: unknown[]): void }).content(
      paragraph("Edited first slide cue"),
      paragraph(["Now at index ", slideNumber()]),
    );
    s1.moveTo(partTwo!);

    const rebuiltResolution = resolveDocument(imported.document);
    const rebuiltPptx = emit(rebuiltResolution).pptx;
    const painted = paint(rebuiltResolution);

    // Paint never renders speaker notes onto the visible slide canvas.
    assert.equal(painted.pages.length, 3);
    for (const page of painted.pages) {
      assert.deepEqual(page.errors, []);
      assert.ok(page.png instanceof Uint8Array);
      assert.ok(typeof page.svg === "string");
      assert.equal(page.svg.includes("Edited first slide cue"), false);
      assert.equal(page.svg.includes("Hidden slide notes"), false);
    }

    const validated = await validatePptx(rebuiltPptx);
    assert.deepEqual(validated.errors, []);
    assert.deepEqual(
      validated.presentation.sections?.map((sec) => ({
        name: sec.name,
        slideIds: sec.slideIds,
      })),
      [
        { name: "Part One", slideIds: ["secret"] },
        { name: "Part Two", slideIds: ["closing", "intro"] },
      ],
    );
    assert.equal(validated.slides[0]?.id, "secret");
    assert.equal(validated.slides[0]?.hidden, true);
    assert.equal(validated.slides[0]?.notes?.text.text, "Hidden slide notes");
    assert.equal(validated.slides[1]?.id, "closing");
    assert.equal(validated.slides[1]?.notes, undefined);
    assert.equal(validated.slides[2]?.id, "intro");
    assert.equal(
      validated.slides[2]?.notes?.text.text,
      "Edited first slide cue\nNow at index 3",
    );
  });

  it("imports non-sequential sldIds, show='0', custom notesMaster/handoutMaster loss, and preserves all slide notes bodies", async () => {
    const pptx = zipSync({
      "_rels/.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>`,
      ),
      "ppt/presentation.xml": strToU8(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:notesMasterIdLst><p:notesMasterId r:id="rId4"/></p:notesMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="264" r:id="rId3"/>
    <p:sldId id="256" r:id="rId2"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="5143500"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:extLst>
    <p:ext uri="${P14_SECTION_LIST_EXT_URI}">
      <p14:sectionLst xmlns:p14="${P14_NS}">
        <p14:section name="Section A" id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}">
          <p14:sldIdLst><p14:sldId id="264"/><p14:sldId id="256"/></p14:sldIdLst>
        </p14:section>
      </p14:sectionLst>
    </p:ext>
  </p:extLst>
</p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}">
  <Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="${REL_SLIDE}" Target="slides/slide1.xml"/>
  <Relationship Id="rId3" Type="${REL_SLIDE}" Target="slides/slide2.xml"/>
  <Relationship Id="rId4" Type="${REL_NOTES_MASTER}" Target="notesMasters/notesMaster1.xml"/>
</Relationships>`,
      ),
      "ppt/slideMasters/slideMaster1.xml": strToU8(
        `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`,
      ),
      "ppt/slideMasters/_rels/slideMaster1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
      ),
      "ppt/slideLayouts/slideLayout1.xml": strToU8(
        `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="blank"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
      ),
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
      ),
      "ppt/slides/slide1.xml": strToU8(
        `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" show="0"><p:cSld name="second-in-order"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
      ),
      "ppt/slides/_rels/slide1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL_NOTES_SLIDE}" Target="../notesSlides/notesSlide1.xml"/></Relationships>`,
      ),
      "ppt/slides/slide2.xml": strToU8(
        `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld name="first-in-order"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
      ),
      "ppt/slides/_rels/slide2.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL_NOTES_SLIDE}" Target="../notesSlides/notesSlide2.xml"/></Relationships>`,
      ),
      "ppt/notesMasters/notesMaster1.xml": strToU8(
        `<p:notesMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:notesStyle/></p:notesMaster>`,
      ),
      "ppt/notesSlides/notesSlide1.xml": strToU8(
        `<p:notes xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr b="1"/><a:t>Notes for hidden second slide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`,
      ),
      "ppt/notesSlides/notesSlide2.xml": strToU8(
        `<p:notes xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Notes for first slide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`,
      ),
    });

    const imported = importDocument(pptx);
    assert.equal(imported.report.diagnostics.length, 1);
    const diag = imported.report.diagnostics[0]!;
    assert.equal(diag.code, "unsupported-mechanism");
    assert.equal(diag.location.part, "ppt/notesMasters/notesMaster1.xml");
    if (diag.code === "unsupported-mechanism") {
      assert.equal(diag.loss?.lost, "custom-print-layout");
    }

    const rebuilt = emit(resolveDocument(imported.document)).pptx;
    const validated = await validatePptx(rebuilt);
    assert.deepEqual(validated.errors, []);
    assert.equal(validated.slides[0]?.id, "first-in-order");
    assert.equal(validated.slides[0]?.hidden, undefined);
    assert.equal(
      validated.slides[0]?.notes?.text.text,
      "Notes for first slide",
    );
    assert.equal(validated.slides[1]?.id, "second-in-order");
    assert.equal(validated.slides[1]?.hidden, true);
    assert.equal(
      validated.slides[1]?.notes?.text.text,
      "Notes for hidden second slide",
    );
    assert.equal(
      validated.slides[1]?.notes?.text.paragraphs[0]?.runs[0]?.bold,
      true,
    );
  });
});
