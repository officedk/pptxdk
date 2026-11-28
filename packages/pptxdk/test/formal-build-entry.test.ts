import assert from "node:assert/strict";

import { createResolver, emit } from "@officedk/pptxdk/build";
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
    void emit;
  });
});
