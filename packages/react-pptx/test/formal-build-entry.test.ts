import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

import { createFormalBuild } from "@react-pptx/react-pptx/build";
import { describe, it } from "vitest";

describe("pure formal build entry", () => {
  it("rejects a pre-cancelled formal build before compiling", () => {
    const controller = new AbortController();
    const reason = new Error("pre-cancelled formal build");
    controller.abort(reason);

    assert.throws(
      () =>
        createFormalBuild([])({ children: [] }, { signal: controller.signal }),
      (error: unknown) => error === reason,
    );
  });

  it("loads without React and exposes the shared formal build", () => {
    const packageRoot = fileURLToPath(new URL("../", import.meta.url));
    const script = `
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier === "react" ||
      specifier === "react-reconciler" ||
      specifier.startsWith("react/")
    ) {
      throw new Error("pure build entry loaded " + specifier);
    }
    return nextResolve(specifier, context);
  },
});

const entry = await import("@react-pptx/react-pptx/build");
if (typeof entry.createFormalBuild !== "function") {
  throw new Error("pure build entry does not expose createFormalBuild");
}
process.stdout.write("ok");
`;

    const child = spawnSync(
      process.execPath,
      ["--conditions=source", "--input-type=module", "--eval", script],
      { cwd: packageRoot, encoding: "utf8" },
    );

    assert.equal(child.status, 0, child.stderr);
    assert.equal(child.stdout, "ok");
  });
});
