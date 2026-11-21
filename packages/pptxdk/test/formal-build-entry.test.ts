import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

import { createResolver, emit } from "pptxdk/build";
import { describe, it } from "vitest";

describe("pure formal build entry", () => {
  it("rejects a pre-cancelled resolve or emit before running", () => {
    const controller = new AbortController();
    const reason = new Error("pre-cancelled formal build");
    controller.abort(reason);

    assert.throws(
      () => createResolver([])({ children: [] }, { signal: controller.signal }),
      (error: unknown) => error === reason,
    );
  });

  it("loads without React and exposes createResolver and emit", () => {
    const packageRoot = fileURLToPath(new URL("../", import.meta.url));
    const script = `
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier === "react" ||
      specifier === "react-reconciler" ||
      specifier.startsWith("react/") ||
      specifier === "@pptxdk/react" ||
      specifier.startsWith("@pptxdk/react/")
    ) {
      throw new Error("pure build entry loaded " + specifier);
    }
    return nextResolve(specifier, context);
  },
});

const entry = await import("pptxdk/build");
if (typeof entry.createResolver !== "function" || typeof entry.emit !== "function") {
  throw new Error("pure build entry does not expose createResolver and emit");
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
    void emit;
  });
});
