import assert from "node:assert/strict";
import { writeFile } from "node:fs/promises";
import { registerHooks } from "node:module";

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
const { PreviewPageRangeError, build, preview } =
  await import("@officedk/pptxdk");
const { selectCapabilities } = await import("@officedk/pptxdk/build");
assert.equal(typeof selectCapabilities, "function");
assert.equal(typeof PreviewPageRangeError, "function");
const built = await build(document, { embedFonts: "subset" });
await writeFile("sdk.pptx", built.pptx);
hook.deregister();

const { importPptxFile, resolveFromSource } =
  await import("@officedk/pptxdk/node");
const resolved = await resolveFromSource("./deck.ts");
assert.equal(resolved.slides.length, 1);
const { document: loaded } = await importPptxFile("sdk.pptx");
assert(loaded);
const result = await preview(loaded);
assert.equal(result.pages.length, 1);
assert.deepEqual(result.pages[0]?.errors, []);
assert(result.pages[0]?.png?.byteLength);
