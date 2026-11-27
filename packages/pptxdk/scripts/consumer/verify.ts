import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import { registerHooks } from "node:module";

function patchTtfWeightClass(
  source: Uint8Array,
  weightClass: number,
): Uint8Array {
  const view = new DataView(
    source.buffer,
    source.byteOffset,
    source.byteLength,
  );
  const sfntVersion = view.getUint32(0);
  const numTables = view.getUint16(4);
  const tables = new Map<string, Uint8Array>();

  for (let index = 0; index < numTables; index += 1) {
    const recOffset = 12 + index * 16;
    const tag = String.fromCharCode(
      ...source.subarray(recOffset, recOffset + 4),
    );
    const offset = view.getUint32(recOffset + 8);
    const length = view.getUint32(recOffset + 12);
    const tableBytes = source.slice(offset, offset + length);
    if (tag === "OS/2") {
      new DataView(tableBytes.buffer).setUint16(4, weightClass);
    }
    tables.set(tag, tableBytes);
  }

  const sorted = [...tables.entries()].toSorted(([a], [b]) =>
    a < b ? -1 : a > b ? 1 : 0,
  );
  const count = sorted.length;
  const headerSize = 12 + count * 16;
  let totalSize = headerSize;
  for (const [, data] of sorted) {
    totalSize += Math.ceil(data.byteLength / 4) * 4;
  }

  const checksum = (bytes: Uint8Array): number => {
    const padded = Math.ceil(bytes.byteLength / 4) * 4;
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let sum = 0;
    for (let o = 0; o < padded; o += 4) {
      if (o + 4 <= bytes.byteLength) {
        sum = (sum + dv.getUint32(o)) >>> 0;
      } else {
        let word = 0;
        for (let b = 0; b < 4; b += 1) {
          word = (word << 8) | (o + b < bytes.byteLength ? bytes[o + b]! : 0);
        }
        sum = (sum + (word >>> 0)) >>> 0;
      }
    }
    return sum >>> 0;
  };

  const out = new Uint8Array(totalSize);
  const ov = new DataView(out.buffer);
  ov.setUint32(0, sfntVersion >>> 0);
  ov.setUint16(4, count);
  const entrySelector = Math.floor(Math.log2(Math.max(1, count)));
  const searchRange = (1 << entrySelector) * 16;
  ov.setUint16(6, searchRange);
  ov.setUint16(8, entrySelector);
  ov.setUint16(10, count * 16 - searchRange);

  let dataOff = headerSize;
  let headOff: number | undefined;
  for (let i = 0; i < sorted.length; i += 1) {
    const [tag, raw] = sorted[i]!;
    const bytes =
      tag === "head" && raw.byteLength >= 12 ? new Uint8Array(raw) : raw;
    if (tag === "head" && bytes.byteLength >= 12) {
      new DataView(bytes.buffer).setUint32(8, 0);
      headOff = dataOff;
    }
    out.set(bytes, dataOff);
    const recOff = 12 + i * 16;
    for (let c = 0; c < 4; c += 1) out[recOff + c] = tag.charCodeAt(c);
    ov.setUint32(recOff + 4, checksum(bytes));
    ov.setUint32(recOff + 8, dataOff);
    ov.setUint32(recOff + 12, bytes.byteLength);
    dataOff += Math.ceil(bytes.byteLength / 4) * 4;
  }
  if (headOff !== undefined) {
    ov.setUint32(headOff + 8, (0xb1b0afba - checksum(out)) >>> 0);
  }
  return out;
}

// Loading the vanilla path must not activate CLI process handling, React or rasterization.
const hook = registerHooks({
  resolve(specifier, context, nextResolve) {
    assert.doesNotMatch(
      specifier,
      /^(?:@boringparts\/cli-contract|@resvg\/resvg-js|react|react-reconciler)(?:\/|$)/,
    );
    return nextResolve(specifier, context);
  },
});
const { default: document } = await import("./deck.ts");
const { PreviewPageRangeError, build, codegen, importPptx, isTheme, preview } =
  await import("@officedk/pptxdk");
const { selectCapabilities } = await import("@officedk/pptxdk/build");
assert.equal(typeof selectCapabilities, "function");
assert.equal(typeof PreviewPageRangeError, "function");
const built = await build(document, { embedFonts: "subset" });
await writeFile("sdk.pptx", built.pptx);
hook.deregister();

const {
  buildFromSource,
  importPptxFile,
  previewFromSource,
  resolveFromSource,
} = await import("@officedk/pptxdk/node");
const resolved = await resolveFromSource("./deck.ts");
assert.equal(resolved.slides.length, 1);
const { document: loaded } = await importPptxFile("sdk.pptx");
assert(loaded);
const result = await preview(loaded);
assert.equal(result.pages.length, 1);
assert.deepEqual(result.pages[0]?.errors, []);
assert(result.pages[0]?.png?.byteLength);

// Verify React consumer deck with Mount, fromVanilla (closed + slotted), placeholders CommonNode,
// omitted-metadata FontFace, and non-hundred fontWeight through buildFromSource & previewFromSource.
const reactBuilt = await buildFromSource("./react-components.tsx", {
  exportName: "ReactConsumerDeck",
  embedFonts: "subset",
});
assert.equal(reactBuilt.report.slides.length, 2);
assert(
  reactBuilt.report.fonts.some(
    (entry) =>
      entry.family === "Jost*" && entry.weight === 400 && entry.hasResource,
  ),
);
const reactPreview = await previewFromSource("./react-components.tsx", {
  exportName: "ReactConsumerDeck",
});
assert.equal(reactPreview.pages.length, 2);
assert.deepEqual(reactPreview.pages[0]?.errors, []);
assert.deepEqual(reactPreview.pages[1]?.errors, []);
assert(reactPreview.pages[0]?.png?.byteLength);
assert(reactPreview.pages[1]?.png?.byteLength);

function collectNodeText(node: unknown, out: string[] = []): string[] {
  if (typeof node === "string") {
    out.push(node);
    return out;
  }
  if (typeof node === "object" && node !== null) {
    if ("text" in node && typeof node.text === "string") {
      out.push(node.text);
    }
    if ("children" in node && Array.isArray(node.children)) {
      for (const child of node.children) {
        collectNodeText(child, out);
      }
    }
  }
  return out;
}

const { document: importedReact } = importPptx(reactBuilt.pptx);
const importedSlides = importedReact.children.filter(
  (child) => child.type === "slide",
);
assert.equal(importedSlides.length, 2);
const reactText = importedSlides
  .flatMap((slideNode) => collectNodeText(slideNode))
  .join("\n");
for (const expectedText of [
  "Bridged closed badge",
  "Bridged slotted shell",
  "Adopted closed node",
  "Mounted slot child",
  "Imperative placeholder body",
]) {
  assert(
    reactText.includes(expectedText),
    `Expected imported React deck to preserve ${expectedText}`,
  );
}

// Verify expandFragment returns frozen ownerless CommonNode[] from the packed tarball.
const { Fragment, createElement } = await import("react");
const { Text, expandFragment } = await import("@officedk/pptxdk/react");
const fragmentNodes = await expandFragment(
  createElement(
    Fragment,
    null,
    createElement(Text, { id: "frag-1" }, "Fragment A"),
    createElement(Text, { id: "frag-2" }, "Fragment B"),
  ),
  { signal: new AbortController().signal },
);
assert.equal(fragmentNodes.length, 2);
assert.equal(Object.isFrozen(fragmentNodes), true);
assert.equal(
  fragmentNodes.every((node) => node.parent === undefined),
  true,
);

// Verify omitted-metadata FontFace infers a non-hundred usWeightClass (350) and roundtrips through importPptx.
const { deck, defineTheme, font, shape, slide } =
  await import("@officedk/pptxdk/vanilla");
const jost350Bytes = patchTtfWeightClass(
  new Uint8Array(await readFile(new URL("./Jost.ttf", import.meta.url))),
  350,
);
const jost350Family = font("Jost*", [{ data: jost350Bytes }]);
const customWeightDeck = deck({
  width: 640,
  height: 360,
  theme: defineTheme({
    name: "Weight350",
    fonts: { major: jost350Family, minor: jost350Family },
  }),
}).append(
  slide({ id: "w350" }).append(
    shape("DemiLight 350", {
      id: "w350-shape",
      style: {
        position: "absolute",
        left: 24,
        top: 24,
        width: 300,
        height: 48,
        fontFamily: jost350Family,
        fontWeight: 350,
      },
    }),
  ),
);
const built350 = await build(customWeightDeck, { embedFonts: "subset" });
assert(
  built350.report.fonts.some(
    (entry) =>
      entry.family === "Jost*" && entry.weight === 350 && entry.hasResource,
  ),
);
const { document: imported350 } = importPptx(built350.pptx);
const importedTheme = imported350.props["theme"];
assert(isTheme(importedTheme));
const majorFont = importedTheme.fonts?.major;
assert(typeof majorFont === "object" && majorFont !== null);
assert.equal(majorFont.faces?.[0]?.weight, 350);
const generated350 = codegen(imported350, { target: "vanilla" });
const generatedSource350 = generated350.files[generated350.entry];
assert.equal(typeof generatedSource350, "string");
assert.match(generatedSource350 as string, /weight:\s*350\b/u);
assert.match(generatedSource350 as string, /fontWeight:\s*350\b/u);
