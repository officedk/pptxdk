import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";

import { z } from "zod";

const packageName = "@officedk/pptxdk";
const latest = packageName + "@latest";
const planSchema = z.object({
  plan: z.array(
    z.array(
      z.object({
        name: z.string(),
        version: z.string(),
        kind: z.string(),
        tarball: z.object({ path: z.string() }).optional(),
      }),
    ),
  ),
});
const metadataSchema = z.object({
  version: z.string(),
  dist: z.object({
    integrity: z.string(),
    attestations: z.object({
      url: z.url(),
      provenance: z.object({
        predicateType: z.literal("https://slsa.dev/provenance/v1"),
      }),
    }),
  }),
});

async function verifyRegistry(
  version: string,
  tarball?: string,
): Promise<string> {
  const metadata = execFileSync(
    "npm",
    [
      "view",
      tarball ? latest : packageName + "@" + version,
      "version",
      "dist",
      "--json",
      "--prefer-online",
      "--registry=https://registry.npmjs.org",
    ],
    { encoding: "utf8" },
  );
  const info = metadataSchema.parse(JSON.parse(metadata));
  assert.equal(info.version, version);
  const integrity = info.dist.integrity;
  assert.match(integrity, /^sha512-/);
  const hash = Buffer.from(integrity.slice("sha512-".length), "base64");
  if (tarball) {
    assert.deepEqual(
      createHash("sha512")
        .update(await readFile(tarball))
        .digest(),
      hash,
    );
  }

  const response = await fetch(info.dist.attestations.url, {
    signal: AbortSignal.timeout(30_000),
  });
  assert(
    response.ok,
    "Unable to retrieve npm release attestations: " + response.status,
  );
  const { attestations } = z
    .object({
      attestations: z.array(
        z.object({
          predicateType: z.string(),
          bundle: z.object({ dsseEnvelope: z.object({ payload: z.string() }) }),
        }),
      ),
    })
    .parse(await response.json());
  const type = info.dist.attestations.provenance.predicateType;
  const provenance = attestations.find(
    ({ predicateType }) => predicateType === type,
  );
  assert(provenance, "Missing SLSA provenance bundle");
  const payload: unknown = JSON.parse(
    Buffer.from(provenance.bundle.dsseEnvelope.payload, "base64").toString(
      "utf8",
    ),
  );
  const tag = packageName + "@" + version;
  const api = "https://api.github.com/repos/officedk/pptxdk";
  const encodedTag = encodeURIComponent(tag);
  const [release, tagged] = await Promise.all([
    fetch(api + "/releases/tags/" + encodedTag, {
      signal: AbortSignal.timeout(30_000),
    }),
    fetch(api + "/commits/" + encodedTag, {
      signal: AbortSignal.timeout(30_000),
    }),
  ]);
  assert(
    release.ok && tagged.ok,
    "Verify or restore the GitHub Release and tag " +
      tag +
      " at the commit recorded in npm provenance (GitHub returned " +
      release.status +
      "/" +
      tagged.status +
      ")",
  );
  assert.partialDeepStrictEqual(await release.json(), {
    tag_name: tag,
    draft: false,
  });
  const { sha: commit } = z
    .object({ sha: z.string().regex(/^[a-f0-9]{40}$/) })
    .parse(await tagged.json());
  if (tarball) {
    assert.equal(
      commit,
      execFileSync("git", ["rev-parse", "HEAD"], { encoding: "utf8" }).trim(),
    );
  }
  assert.partialDeepStrictEqual(payload, {
    predicateType: type,
    subject: [{ digest: { sha512: hash.toString("hex") } }],
    predicate: {
      buildDefinition: {
        externalParameters: {
          workflow: {
            repository: "https://github.com/officedk/pptxdk",
            ref: "refs/heads/main",
            path: ".github/workflows/release.yml",
          },
        },
        resolvedDependencies: [{ digest: { gitCommit: commit } }],
      },
    },
  });
  return integrity;
}

function checkConsumer(spec: string, registryIntegrity?: string): void {
  execFileSync(
    "pnpm",
    [
      "--filter",
      packageName,
      "exec",
      "node",
      "scripts/check-consumer.ts",
      spec,
      ...(registryIntegrity ? [registryIntegrity] : []),
    ],
    { stdio: "inherit" },
  );
}

async function main(): Promise<void> {
  const packed = process.env.PPTXDK_RELEASE_PACK_DIR;
  const published = process.argv[2] === "--published";
  if (published && !packed) throw new Error("Missing release pack directory");
  if (packed) {
    const dir = resolve(packed);
    if (!published) {
      execFileSync("pnpm", ["exec", "changeset", "pack", "--out-dir", dir], {
        stdio: "inherit",
      });
    }
    const manifest = await readFile(join(dir, "publish-plan.json"), "utf8");
    const releases = planSchema
      .parse(JSON.parse(manifest))
      .plan.flat()
      .filter(({ kind }) => kind === "publish");
    if (releases.length) {
      assert.equal(releases.length, 1, "Expected a single public npm package");
      const [release] = releases;
      assert(release?.tarball);
      assert.equal(release.name, packageName);
      const tarball = join(dir, release.tarball.path);
      if (published) {
        const integrity = await verifyRegistry(release.version, tarball);
        checkConsumer(latest, integrity);
      } else {
        checkConsumer(tarball);
      }
      return;
    }
    if (published) {
      // A retry after npm accepted the version may still be missing the GitHub release.
      const workspace = execFileSync(
        "pnpm",
        ["--filter", packageName, "list", "--depth=-1", "--json"],
        { encoding: "utf8" },
      );
      const [sdk] = z
        .tuple([
          z.object({ name: z.literal(packageName), version: z.string() }),
        ])
        .parse(JSON.parse(workspace));
      const integrity = await verifyRegistry(sdk.version);
      checkConsumer(packageName + "@" + sdk.version, integrity);
      return;
    }
  }
  // PRs still need a consumer check when their version already exists on npm.
  const temporary = await mkdtemp(join(tmpdir(), "pptxdk-pack-"));
  try {
    const tarball = join(temporary, "sdk.tgz");
    execFileSync(
      "pnpm",
      ["--filter", packageName, "pack", "--out", tarball, "--silent"],
      { stdio: "inherit" },
    );
    checkConsumer(tarball);
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
}

await main();
