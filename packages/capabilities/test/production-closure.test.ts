import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import type {
  ObjectCodegenContribution,
  ProductionDeclarationKind,
  ProductionSlideSlotProp,
  SlideCodegenContribution,
} from "@pptxdk/codegen";
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
  it("imports all 31 public export subpaths without resolving preview, react adapter, validation, resvg, react-reconciler, or eager image decoders", () => {
    const pkg = JSON.parse(
      readFileSync(new URL("../package.json", import.meta.url), "utf8"),
    ) as { readonly exports: Record<string, unknown> };
    const subpaths = Object.keys(pkg.exports);
    assert.equal(subpaths.length, 31);

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
  line,
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
    box({ id: "b1", style: { width: 200, height: 100, background: "#FFFFFF" } }).append(
      shape({ id: "sh1", style: { position: "absolute", left: 10, top: 10, width: 80, height: 40, background: "#2563EB" } }),
    ),
    line({ id: "ln1", from: { id: "b1", port: "right" }, to: { x: 240, y: 50 }, stroke: { color: "#0F172A", width: 2, endArrow: "triangle" } }),
    image({ id: "img1", src: pngBytes, style: { position: "absolute", left: 220, top: 20, width: 40, height: 40 } }),
    image({ id: "img2", src: jpegBytes, style: { position: "absolute", left: 270, top: 20, width: 40, height: 40 } }),
    group({ id: "g1" }).append(
      text("Group text", { id: "t1", style: { position: "absolute", left: 320, top: 20, width: 120, height: 32 } }),
    ),
    table({ id: "tbl1", columns: [120, "1fr"], style: { position: "absolute", left: 20, top: 140, width: 240 } }).append(
      row({ style: { height: 28 } }).append(cell("A"), cell("B")),
    ),
    chart({
      id: "ch1",
      style: { position: "absolute", left: 280, top: 140, width: 240, height: 160 },
      groups: [
        {
          kind: "column",
          data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
          series: [{ name: "S", values: [1, 2] }],
        },
      ],
    }),
    step({ n: 1 }).append(
      shape({ id: "sh2", style: { position: "absolute", left: 540, top: 140, width: 60, height: 40 }, enter: fade() }),
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

  it("mounts non-empty valid codegen contributions on every object and slide capability covering all production declarations, owner slots, and object motion props", () => {
    const expectedDeclarations: readonly ProductionDeclarationKind[] = [
      "box",
      "cell",
      "chart",
      "group",
      "image",
      "line",
      "motion-scope",
      "motion-step",
      "notes",
      "paragraph",
      "placeholder",
      "row",
      "shape",
      "slide-number",
      "span",
      "sub",
      "sup",
      "table",
      "text",
    ];
    const expectedSlotProps: readonly ProductionSlideSlotProp[] = [
      "advanceAfter",
      "background",
      "dateTime",
      "footer",
      "footerPlaceholders",
      "headerFooter",
      "motion",
      "slideNumber",
      "transition",
    ];
    const expectedObjectProps = [
      "emphasis",
      "enter",
      "exit",
      "motion",
      "publishEnd",
      "publishStart",
      "startAt",
    ] as const;

    const coveredDeclarations = new Set<string>();
    const coveredSlotProps = new Set<string>();
    const coveredObjectProps = new Set<string>();

    for (const mod of documentCapabilities) {
      const isObjectOrSlide = "bind" in mod || "bindSlide" in mod;
      assert.equal(
        isObjectOrSlide,
        true,
        `Capability ${mod.declaration} must be an object or slide capability module`,
      );
      const rawCodegen = (mod as { readonly codegen?: unknown }).codegen;
      assert.ok(
        typeof rawCodegen === "object" && rawCodegen !== null,
        `Capability ${mod.declaration} is missing a codegen contribution`,
      );

      if ("declarations" in rawCodegen) {
        const contribution = rawCodegen as ObjectCodegenContribution;
        const declEntries = Object.entries(contribution.declarations);
        assert.ok(
          declEntries.length > 0,
          `Capability ${mod.declaration} object codegen must register at least one declaration`,
        );
        for (const [kind, handler] of declEntries) {
          assert.equal(
            typeof handler.emitVanilla,
            "function",
            `Capability ${mod.declaration} declaration ${kind} missing emitVanilla`,
          );
          assert.equal(
            typeof handler.emitReact,
            "function",
            `Capability ${mod.declaration} declaration ${kind} missing emitReact`,
          );
          assert.equal(
            coveredDeclarations.has(kind),
            false,
            `Duplicate codegen declaration ${kind} on ${mod.declaration}`,
          );
          coveredDeclarations.add(kind);
        }
        if (contribution.objectProps !== undefined) {
          for (const [prop, serializer] of Object.entries(
            contribution.objectProps,
          )) {
            assert.equal(typeof serializer, "function");
            coveredObjectProps.add(prop);
          }
        }
        if (contribution.ownerProps !== undefined) {
          for (const [prop, serializer] of Object.entries(
            contribution.ownerProps,
          )) {
            assert.equal(typeof serializer, "function");
            coveredSlotProps.add(prop);
          }
        }
        if (contribution.valueSerializers !== undefined) {
          for (const serializer of contribution.valueSerializers) {
            assert.equal(typeof serializer, "function");
          }
        }
      } else if ("slideProps" in rawCodegen) {
        const contribution = rawCodegen as SlideCodegenContribution;
        const propEntries = Object.entries(contribution.slideProps);
        assert.ok(
          propEntries.length > 0,
          `Capability ${mod.declaration} slide codegen must register at least one slideProp`,
        );
        for (const [prop, serializer] of propEntries) {
          assert.equal(
            typeof serializer,
            "function",
            `Capability ${mod.declaration} slideProp ${prop} must be a function`,
          );
          assert.equal(
            coveredSlotProps.has(prop),
            false,
            `Duplicate codegen slideProp ${prop} on ${mod.declaration}`,
          );
          coveredSlotProps.add(prop);
        }
        if (contribution.valueSerializers !== undefined) {
          for (const serializer of contribution.valueSerializers) {
            assert.equal(typeof serializer, "function");
          }
        }
      } else {
        assert.fail(
          `Capability ${mod.declaration} codegen contribution must define declarations or slideProps`,
        );
      }
    }

    assert.deepEqual(
      [...coveredDeclarations].toSorted(),
      [...expectedDeclarations].toSorted(),
    );
    assert.deepEqual(
      [...coveredSlotProps].toSorted(),
      [...expectedSlotProps].toSorted(),
    );
    assert.deepEqual(
      [...coveredObjectProps].toSorted(),
      [...expectedObjectProps].toSorted(),
    );
  });

  it("runs createCodegen(documentCapabilities) across multi-capability documents without loading preview, rasterizer, react, or reconciler", () => {
    const pngBase64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/n1sAAAAASUVORK5CYII=";
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
    throw new Error("codegen loaded forbidden dependency (" + label + "): " + target);
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

const { createCodegen } = await import("@pptxdk/codegen");
const { documentCapabilities } = await import("@pptxdk/capabilities/all");
const { deck, image, notes, shape, slide, step, text } = await import("@pptxdk/capabilities/document");
const { fade } = await import("@pptxdk/capabilities/motion");
const { transitions } = await import("@pptxdk/capabilities/transition");

const pngBytes = new Uint8Array(Buffer.from("${pngBase64}", "base64"));
const doc = deck({ width: 960, height: 540 }).append(
  slide({ id: "s1", background: "#F8FAFC", transition: transitions.fade() }).append(
    shape({ id: "sh1", style: { position: "absolute", left: 10, top: 10, width: 80, height: 40, background: "#2563EB" } }),
    image({ id: "img1", src: pngBytes, style: { position: "absolute", left: 100, top: 20, width: 40, height: 40 } }),
    step({ n: 1 }).append(text("Animated", { id: "t1", style: { position: "absolute", left: 160, top: 20, width: 120, height: 32 }, enter: fade() })),
    notes("Note"),
  ),
);

const generate = createCodegen(documentCapabilities);
const vanilla = generate(doc, { target: "vanilla" });
const react = generate(doc, { target: "react" });
if (typeof vanilla.files[vanilla.entry] !== "string" || typeof react.files[react.entry] !== "string") {
  throw new Error("Expected string entry output");
}
process.stdout.write("ok");
`;

    for (const condition of ["source", "dist"] as const) {
      assert.equal(runIsolatedNode(script, condition), "ok");
    }
  });
});
