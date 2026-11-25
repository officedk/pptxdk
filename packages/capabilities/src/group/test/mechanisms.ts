import assert from "node:assert/strict";

import type { ExhaustiveMechanismVerificationMap } from "../../../test/mechanism-verification.ts";
import { shapeAndGroupRoundtripScenario } from "../../shape/test/mechanisms.ts";

export const groupMechanismVerifications: Pick<
  ExhaustiveMechanismVerificationMap,
  "group"
> = {
  group: {
    scenario: shapeAndGroupRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} errors`);
      const s2 = validation.slides[1]!;
      const outer = s2.objects[0];
      assert.equal(outer?.kind, "group");
      assert.equal(outer?.name, "Outer Cluster");
      assert.equal(outer?.rotationDegrees, 12);
      assert.deepEqual(
        outer?.members?.map((m) => m.name),
        ["Member Shape", "Inner Cluster"],
      );
      assert.equal(outer?.members?.[0]?.geometry, "hexagon");
      const inner = outer?.members?.[1];
      assert.equal(inner?.kind, "group");
      assert.equal(inner?.name, "Inner Cluster");
      assert.equal(inner?.rotationDegrees, 342);
      assert.equal(inner?.members?.[0]?.name, "Nested Leaf");
      assert.equal(inner?.members?.[0]?.geometry, "diamond");

      const s3 = validation.slides[2]!;
      const textGroup = s3.objects[1];
      assert.equal(textGroup?.kind, "group");
      assert.equal(textGroup?.name, "Text Cluster");
      assert.equal(textGroup?.members?.[0]?.name, "Labeled Hex");
      assert.equal(textGroup?.members?.[0]?.text?.text, "Hex");
      assert.equal(textGroup?.members?.[1]?.name, "Member Text");
      assert.equal(textGroup?.members?.[1]?.text?.text, "Group Label");
    },
    assertPaint({ initial, roundtrip }) {
      assert.deepEqual(initial.pages[1]?.diagnostics, []);
      assert.deepEqual(roundtrip.pages[1]?.diagnostics, []);
      assert.ok(initial.pages[1]?.svg?.includes("rotate(12 "));
      assert.ok(initial.pages[1]?.svg?.includes("rotate(-18 "));
      assert.ok(roundtrip.pages[1]?.svg?.includes("rotate(12 "));
      assert.ok(roundtrip.pages[1]?.svg?.includes("rotate(342 "));
    },
  },
};
