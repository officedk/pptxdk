import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, it } from "vitest";

const packageRoot = fileURLToPath(new URL("../", import.meta.url));

function runIsolatedNode(
  script: string,
  conditions: "source" | "dist",
): string {
  const args =
    conditions === "source"
      ? ["--conditions=source", "--input-type=module", "--eval", script]
      : ["--input-type=module", "--eval", script];
  const child = spawnSync(process.execPath, args, {
    cwd: packageRoot,
    encoding: "utf8",
  });
  assert.equal(child.status, 0, `${conditions} stderr:\n${child.stderr}`);
  return child.stdout;
}

describe("@pptxdk/capabilities production runtime closure", () => {
  it("imports all 29 public export subpaths without resolving preview, react adapter, validation, resvg, react-reconciler, or eager image decoders", () => {
    const pkg = JSON.parse(
      readFileSync(new URL("../package.json", import.meta.url), "utf8"),
    ) as { readonly exports: Record<string, unknown> };
    const subpaths = Object.keys(pkg.exports);
    assert.equal(subpaths.length, 29);

    const specifiers = subpaths.map((key) =>
      key === "."
        ? "@pptxdk/capabilities"
        : `@pptxdk/capabilities/${key.slice(2)}`,
    );

    const script = `
import { registerHooks } from "node:module";

function checkForbidden(target, label) {
  if (
    target === "@pptxdk/preview" ||
    target.startsWith("@pptxdk/preview/") ||
    target.includes("/packages/preview/") ||
    target === "@pptxdk/react" ||
    target.startsWith("@pptxdk/react/") ||
    target.includes("/packages/react/") ||
    target === "@pptxdk/validation" ||
    target.startsWith("@pptxdk/validation/") ||
    target.includes("/packages/validation/") ||
    target.includes("@resvg/resvg-js") ||
    target === "react-reconciler" ||
    target.startsWith("react-reconciler/") ||
    target.includes("/react-reconciler/") ||
    target === "pngjs" ||
    target.startsWith("pngjs/") ||
    target.includes("/pngjs/") ||
    target === "jpeg-js" ||
    target.startsWith("jpeg-js/") ||
    target.includes("/jpeg-js/")
  ) {
    throw new Error("Loaded forbidden dependency (" + label + "): " + target);
  }
}

registerHooks({
  resolve(specifier, context, nextResolve) {
    checkForbidden(specifier, "specifier");
    const resolved = nextResolve(specifier, context);
    checkForbidden(resolved.url, "resolved.url");
    return resolved;
  },
});

const specifiers = ${JSON.stringify(specifiers)};
for (const specifier of specifiers) {
  const mod = await import(specifier);
  if (Object.keys(mod).length === 0) {
    throw new Error("Empty module export for " + specifier);
  }
}
process.stdout.write("ok");
`;

    for (const condition of ["source", "dist"] as const) {
      assert.equal(runIsolatedNode(script, condition), "ok");
    }
  });

  it("runs resolve + emit across multi-capability documents without loading preview, rasterizer, react, or reconciler", () => {
    const pngBase64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/n1sAAAAASUVORK5CYII=";
    const jpegBase64 =
      "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////2wBDAf//////////////////////////////////////////////////////////////////////////////////////wAARCAABAAEDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAX/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIQAxAAAAFR/8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABBQJ//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAwEBPwF//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAgEBPwF//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQAGPwJ//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABPyF//9oADAMBAAIRAxEAPwD/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/9oACAEBAAE/IV//2Q==";

    const script = `
import { registerHooks } from "node:module";

function checkForbidden(target, label) {
  if (
    target === "react" ||
    target.startsWith("react/") ||
    target === "react-reconciler" ||
    target.startsWith("react-reconciler/") ||
    target.includes("/react-reconciler/") ||
    target === "@pptxdk/preview" ||
    target.startsWith("@pptxdk/preview/") ||
    target.includes("/packages/preview/") ||
    target === "@pptxdk/react" ||
    target.startsWith("@pptxdk/react/") ||
    target.includes("/packages/react/") ||
    target === "@pptxdk/validation" ||
    target.startsWith("@pptxdk/validation/") ||
    target.includes("/packages/validation/") ||
    target.includes("@resvg/resvg-js") ||
    target === "pngjs" ||
    target.startsWith("pngjs/") ||
    target.includes("/pngjs/") ||
    target === "jpeg-js" ||
    target.startsWith("jpeg-js/") ||
    target.includes("/jpeg-js/")
  ) {
    throw new Error("resolve + emit loaded forbidden dependency (" + label + "): " + target);
  }
}

registerHooks({
  resolve(specifier, context, nextResolve) {
    checkForbidden(specifier, "specifier");
    const resolved = nextResolve(specifier, context);
    checkForbidden(resolved.url, "resolved.url");
    return resolved;
  },
});

const { createResolver } = await import("@pptxdk/core/build");
const { emit } = await import("@pptxdk/emit");
const { documentCapabilities } = await import("@pptxdk/capabilities/all");
const {
  box,
  cell,
  chart,
  deck,
  group,
  image,
  notes,
  row,
  shape,
  slide,
  step,
  table,
  text,
} = await import("@pptxdk/capabilities/document");
const { fade } = await import("@pptxdk/capabilities/motion");
const { transitions } = await import("@pptxdk/capabilities/transition");

const pngBytes = new Uint8Array(Buffer.from("${pngBase64}", "base64"));
const jpegBytes = new Uint8Array(Buffer.from("${jpegBase64}", "base64"));

const doc = deck({ width: 960, height: 540 }).append(
  slide({ id: "s1", background: "#F8FAFC", transition: transitions.fade() }).append(
    box({ id: "b1", style: { width: 200, height: 100, backgroundColor: "#FFFFFF" } }).append(
      shape({ id: "sh1", x: 10, y: 10, width: 80, height: 40, fill: "#2563EB" }),
    ),
    image({ id: "img1", src: pngBytes, x: 220, y: 20, width: 40, height: 40 }),
    image({ id: "img2", src: jpegBytes, x: 270, y: 20, width: 40, height: 40 }),
    group({ id: "g1" }).append(
      text("Group text", { id: "t1", x: 320, y: 20, width: 120, height: 32 }),
    ),
    table({ id: "tbl1", x: 20, y: 140, width: 240, columns: [120, "1fr"] }).append(
      row({ style: { height: 28 } }).append(cell("A"), cell("B")),
    ),
    chart({
      id: "ch1",
      x: 280,
      y: 140,
      width: 240,
      height: 160,
      groups: [
        {
          kind: "column",
          data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
          series: [{ name: "S", values: [1, 2] }],
        },
      ],
    }),
    step({ n: 1 }).append(
      shape({ id: "sh2", x: 540, y: 140, width: 60, height: 40, enter: fade() }),
    ),
    notes("Speaker note"),
  ),
);

const resolution = createResolver(documentCapabilities)(doc);
const emitted = emit(resolution);
if (!(emitted.pptx instanceof Uint8Array) || emitted.pptx.byteLength === 0) {
  throw new Error("Expected non-empty PPTX output");
}
process.stdout.write("ok");
`;

    for (const condition of ["source", "dist"] as const) {
      assert.equal(runIsolatedNode(script, condition), "ok");
    }
  });

  it("excludes all capability test directories, test files, and test fixtures from dist/", () => {
    const distRoot = join(packageRoot, "dist");
    const forbiddenEntries: string[] = [];
    const visit = (dir: string, relativePrefix: string): void => {
      for (const entry of readdirSync(dir, { withFileTypes: true })) {
        const rel =
          relativePrefix === ""
            ? entry.name
            : `${relativePrefix}/${entry.name}`;
        if (entry.isDirectory()) {
          if (entry.name === "test") {
            forbiddenEntries.push(rel);
            continue;
          }
          visit(join(dir, entry.name), rel);
        } else if (
          /\.test(?:\.d)?\.[cm]?[jt]sx?(?:\.map)?$/u.test(entry.name) ||
          /\.(?:pptx|ttf|woff2|png)$/u.test(entry.name)
        ) {
          forbiddenEntries.push(rel);
        }
      }
    };
    visit(distRoot, "");
    assert.deepEqual(forbiddenEntries, []);
  });
});
