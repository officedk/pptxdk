import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import {
  cp,
  copyFile,
  mkdtemp,
  readFile,
  rm,
  writeFile,
} from "node:fs/promises";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { z } from "zod";

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
