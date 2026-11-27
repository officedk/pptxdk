import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { existsSync } from "node:fs";
import {
  cp,
  copyFile,
  mkdtemp,
  readdir,
  readFile,
  rm,
  writeFile,
} from "node:fs/promises";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { z } from "zod";

async function collectDtsFiles(directory: string): Promise<readonly string[]> {
  const entries = await readdir(directory, { withFileTypes: true });
  const files: string[] = [];
  for (const entry of entries) {
    const fullPath = join(directory, entry.name);
    if (entry.isDirectory()) {
      files.push(...(await collectDtsFiles(fullPath)));
    } else if (entry.name.endsWith(".d.ts")) {
      files.push(fullPath);
    }
  }
  return files.toSorted();
}

const maxTotalDtsBytes = 350 * 1024;

const spec = process.argv[2];
assert(spec, "Expected an npm tarball or package spec");
const require = createRequire(import.meta.url);
const version = (name: string): string =>
  z.object({ version: z.string() }).parse(require(name + "/package.json"))
    .version;
const consumer = await mkdtemp(join(tmpdir(), "pptxdk-npm-consumer-"));
try {
  await cp(new URL("./consumer/", import.meta.url), consumer, {
    recursive: true,
  });
  await copyFile(
    new URL("../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
    join(consumer, "Jost.ttf"),
  );
  // Resolve the consumer toolchain versions from the locked SDK development installation.
  await writeFile(
    join(consumer, "package.json"),
    JSON.stringify({
      private: true,
      type: "module",
      devDependencies: {
        "@types/node": version("@types/node"),
        "@types/react": version("@types/react"),
      },
    }),
  );
  execFileSync(
    "npm",
    [
      "install",
      "--no-audit",
      "--no-fund",
      "--registry=https://registry.npmjs.org",
      spec,
    ],
    { cwd: consumer, stdio: "inherit" },
  );
  const run = (bin: string, args: string[]) =>
    execFileSync(bin, args, { cwd: consumer, stdio: "inherit" });
  const registryIntegrity = process.argv[3];
  if (registryIntegrity) {
    // Confirm that resolving latest during installation selected the inspected tarball.
    const lockfile: unknown = JSON.parse(
      await readFile(join(consumer, "package-lock.json"), "utf8"),
    );
    assert.partialDeepStrictEqual(lockfile, {
      packages: {
        "node_modules/@officedk/pptxdk": { integrity: registryIntegrity },
      },
    });
    run("npm", [
      "audit",
      "signatures",
      "--omit=dev",
      "--registry=https://registry.npmjs.org",
    ]);
  }
  const installedPackage = join(consumer, "node_modules/@officedk/pptxdk");
  const manifest: unknown = JSON.parse(
    await readFile(join(installedPackage, "package.json"), "utf8"),
  );
  assert.partialDeepStrictEqual(manifest, {
    name: "@officedk/pptxdk",
    license: "MIT",
  });
  assert.match(
    await readFile(join(installedPackage, "LICENSE"), "utf8"),
    /MIT License[\s\S]+Copyright \(c\) 2026 officedk/u,
  );
  assert.equal(
    existsSync(join(installedPackage, "README.md")),
    false,
    "Packed @officedk/pptxdk must not ship README.md",
  );
  const dtsFiles = await collectDtsFiles(join(installedPackage, "dist"));
  assert(dtsFiles.length > 0, "Expected packed .d.ts files under dist/");
  let totalDtsBytes = 0;
  const dtsChunks: string[] = [];
  for (const file of dtsFiles) {
    const content = await readFile(file, "utf8");
    totalDtsBytes += Buffer.byteLength(content, "utf8");
    assert.doesNotMatch(
      content,
      /\b(?:fontkit|harfbuzzjs)\b/u,
      `Unexpected internal font library reference leaked in ${file}`,
    );
    assert.doesNotMatch(
      content,
      /\bZodObject<|\bZodReadonly<|\bZodDiscriminatedUnion<|z\.infer</u,
      `Unexpected expanded Zod AST or z.infer leaked in ${file}`,
    );
    dtsChunks.push(content);
  }
  assert(
    totalDtsBytes < maxTotalDtsBytes,
    `Packed dist/*.d.ts size ${totalDtsBytes} bytes exceeds ${maxTotalDtsBytes} byte budget`,
  );
  const combinedDts = dtsChunks.join("\n");
  for (const [label, pattern] of [
    [
      "GradientStop.position JSDoc",
      /Normalized stop position in the closed interval `\[0, 1\]`/u,
    ],
    [
      "ChartProps.groups example JSDoc",
      /groups:\s*\[\s*\*\s*\{\s*\*\s*kind: "column"/u,
    ],
    ["ChartPlotArea / plotArea px JSDoc", /top-left corner/u],
    [
      "FontSize unit JSDoc",
      /Positive numbers default to pixels \(`px`\) converted at `96 dpi`/u,
    ],
    ["Image src / SVG JSDoc", /\.svg` files are not\s+\*\s*supported/u],
    ["LayoutStyle.flex JSDoc", /minWidth: "auto"/u],
    [
      "FontFace metadata inference & Name ID 1 JSDoc",
      /When `data` is a static single-font binary and `weight` or `style` is omitted[\s\S]*?native Unicode legacy family name \(`Name ID 1`\)/u,
    ],
    [
      "DeckProps pixel dimensions JSDoc",
      /canvas `width` and `height` are in CSS pixels \(`1280 × 720` default\)/u,
    ],
    [
      "DeckProps default width JSDoc",
      /defaults to `1280` \(16:9 widescreen\) when omitted/u,
    ],
    ["SectionProps section JSDoc", /native `p14:sectionLst` section/u],
    [
      "SectionProps extension JSDoc",
      /Office 2010 `p14:sectionLst` presentation extension/u,
    ],
  ] as const) {
    assert.match(combinedDts, pattern, `Missing ${label} in packed .d.ts`);
  }
  const reactDts = await readFile(
    join(installedPackage, "dist/react.d.ts"),
    "utf8",
  );
  for (const [label, pattern] of [
    [
      "Mount flat declaration",
      /export declare function Mount\(props: MountProps\): ReactNode;/u,
    ],
    ["MountProps flat type", /type MountProps =/u],
    ["MountProps export", /\btype MountProps\b/u],
    [
      "fromVanilla flat declaration",
      /export declare function fromVanilla<[\s\S]*?>\(/u,
    ],
    [
      "VanillaSlotResult flat type",
      /type VanillaSlotResult<Node extends CommonNode = CommonNode, Slot extends CommonContainer = CommonContainer> =/u,
    ],
    ["VanillaSlotResult export", /\btype VanillaSlotResult\b/u],
    [
      "expandFragment flat declaration",
      /export declare function expandFragment\(element: ReactNode, options\?: ExpandFragmentOptions\): Promise<readonly CommonNode\[\]>;/u,
    ],
    [
      "ExpandFragmentOptions flat type",
      /type ExpandFragmentOptions = ExpandDocumentOptions;/u,
    ],
    ["ExpandFragmentOptions export", /\btype ExpandFragmentOptions\b/u],
  ] as const) {
    assert.match(reactDts, pattern, `Missing ${label} in dist/react.d.ts`);
  }
  for (const [label, pattern] of [
    ["HexColor flat type", /type HexColor = `#\$\{string\}`;/u],
    [
      "ThemeColor flat type",
      /type ThemeColor = SchemeColorName \| HexColor \| ThemeShade;/u,
    ],
    [
      "ThemeFontSlot flat type",
      /declare const themeFontSlots: readonly \["major", "minor"\];\s*type ThemeFontSlot = \(typeof themeFontSlots\)\[number\];/u,
    ],
    [
      "FontReference flat type",
      /type FontReference = ThemeFontSlot \| \(string & \{\}\) \| Font;/u,
    ],
  ] as const) {
    assert.match(combinedDts, pattern, `Missing ${label} in packed .d.ts`);
  }
  run(process.execPath, ["verify.ts"]);
  const cli = join(consumer, "node_modules", ".bin", "pptxdk");
  for (const [target, entry] of [
    ["vanilla", "deck.ts"],
    ["react", "deck.tsx"],
  ] as const) {
    run(cli, ["codegen", "sdk.pptx", "--target", target, "--out-dir", target]);
    run(cli, ["build", join(target, entry), "--out", target + ".pptx"]);
    const pptx = await readFile(join(consumer, target + ".pptx"));
    assert.equal(
      pptx.readUInt32LE(),
      0x04034b50,
      "Expected a PPTX ZIP archive",
    );
  }
  run(join(import.meta.dirname, "../node_modules/.bin/tsc"), [
    "--project",
    "tsconfig.json",
    "--pretty",
    "false",
  ]);
} finally {
  await rm(consumer, { recursive: true, force: true });
}
