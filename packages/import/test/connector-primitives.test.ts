import assert from "node:assert/strict";

import {
  A_NS,
  P_NS,
  XmlElementCursor,
  parseXmlPart,
  peekShapeTreeObjectIdentity,
  readTransformXml,
  type ImportDiagnostic,
} from "@pptxdk/import";
import { describe, it } from "vitest";

describe("Connector and single-axis zero-extent import primitives", () => {
  it("reads p:nvCxnSpPr/p:cNvPr identity via peekShapeTreeObjectIdentity", () => {
    const el = parseXmlPart(
      "ppt/slides/slide1.xml",
      `<p:cxnSp xmlns:p="${P_NS}" xmlns:a="${A_NS}"><p:nvCxnSpPr><p:cNvPr id="7" name="Connector 7"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr></p:cxnSp>`,
    );
    assert.deepEqual(peekShapeTreeObjectIdentity(el), {
      objectId: 7,
      objectName: "Connector 7",
    });
  });

  it("allows single-axis zero extent in readTransformXml only when allowSingleZeroExtent is enabled", () => {
    const horizXml = `<a:xfrm xmlns:a="${A_NS}"><a:off x="95250" y="190500"/><a:ext cx="952500" cy="0"/></a:xfrm>`;
    const zeroBothXml = `<a:xfrm xmlns:a="${A_NS}"><a:off x="95250" y="190500"/><a:ext cx="0" cy="0"/></a:xfrm>`;

    const diagsDefault: ImportDiagnostic[] = [];
    const cursorDefault = XmlElementCursor.createRoot(
      parseXmlPart("ppt/slides/slide1.xml", horizXml),
      "ppt/slides/slide1.xml",
    );
    assert.equal(
      readTransformXml(
        cursorDefault,
        { part: "ppt/slides/slide1.xml" },
        diagsDefault,
      ),
      undefined,
    );
    assert.equal(diagsDefault[0]?.code, "invalid-xml-value");

    const diagsLine: ImportDiagnostic[] = [];
    const cursorLine = XmlElementCursor.createRoot(
      parseXmlPart("ppt/slides/slide1.xml", horizXml),
      "ppt/slides/slide1.xml",
    );
    const parsed = readTransformXml(
      cursorLine,
      { part: "ppt/slides/slide1.xml" },
      diagsLine,
      { allowSingleZeroExtent: true },
    );
    assert.deepEqual(diagsLine, []);
    assert.deepEqual(parsed, {
      x: 10,
      y: 20,
      width: 100,
      height: 0,
      xEmu: 95250,
      yEmu: 190500,
      cxEmu: 952500,
      cyEmu: 0,
    });

    const diagsBoth: ImportDiagnostic[] = [];
    const cursorBoth = XmlElementCursor.createRoot(
      parseXmlPart("ppt/slides/slide1.xml", zeroBothXml),
      "ppt/slides/slide1.xml",
    );
    assert.equal(
      readTransformXml(
        cursorBoth,
        { part: "ppt/slides/slide1.xml" },
        diagsBoth,
        { allowSingleZeroExtent: true },
      ),
      undefined,
    );
    assert.equal(diagsBoth[0]?.code, "invalid-xml-value");
  });
});
