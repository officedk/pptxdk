import assert from "node:assert/strict";

import { BuildError, createResolver } from "@pptxdk/core/build";
import {
  CapabilityError,
  parseCapabilityLayoutContract,
  type ObjectCapabilityModule,
  type ScopeTargetObject,
} from "@pptxdk/core/capability";
import {
  CommonContainer,
  CommonNode,
  createDeckNode,
  createSlideNode,
} from "@pptxdk/core/document";
import { describe, it } from "vitest";

const boxOnlyModule: ObjectCapabilityModule = {
  declaration: "test-box",
  bind(input, context) {
    const contract = parseCapabilityLayoutContract({
      nativeKind: "sp",
      props: input.props,
      allowedStyleKeys: new Set([
        "display",
        "position",
        "left",
        "top",
        "width",
        "height",
        "flexDirection",
        "gap",
      ]),
      allowPadding: false,
      isContainer: true,
      path: context.path,
    });
    const emitsObject = input.props["emits"] === true;
    return {
      contract,
      container: {
        children: input.children,
        emitsObject,
        childComposition: "siblings",
      },
      resolve(ctx) {
        const box = ctx.layoutBox ?? { x: 0, y: 0, width: 100, height: 100 };
        return {
          observation: {
            id: ctx.id,
            kind: "shape",
            name:
              typeof input.props["id"] === "string"
                ? input.props["id"]
                : String(ctx.id),
            box,
          },
          extensions: [],
          emit: () => ({ xml: "" }),
        };
      },
    };
  },
};

const groupOnlyModule: ObjectCapabilityModule = {
  declaration: "test-group",
  bind(input, context) {
    const contract = parseCapabilityLayoutContract({
      nativeKind: "grpSp",
      props: input.props,
      allowedStyleKeys: new Set([
        "display",
        "position",
        "left",
        "top",
        "width",
        "height",
      ]),
      allowPadding: false,
      isContainer: true,
      path: context.path,
    });
    return {
      contract,
      container: {
        children: input.children,
        emitsObject: true,
        childComposition: "members",
        restrictsMemberContinuity: true,
        observesLayout: false,
        propagatesBaseline: false,
      },
      resolve(ctx) {
        const members = ctx.resolveChildren(input.children);
        return {
          observation: {
            id: ctx.id,
            kind: "group",
            name:
              typeof input.props["id"] === "string"
                ? input.props["id"]
                : String(ctx.id),
            box: { x: 0, y: 0, width: 200, height: 200 },
            members: members.objects,
          },
          extensions: [],
          emit: () => ({ xml: "" }),
        };
      },
    };
  },
};

const targetShapeModule: ObjectCapabilityModule = {
  declaration: "test-shape",
  bind(input, context) {
    const contract = parseCapabilityLayoutContract({
      nativeKind: "sp",
      props: input.props,
      allowedStyleKeys: new Set([
        "display",
        "position",
        "left",
        "top",
        "width",
        "height",
        "rotate",
        "flipH",
        "flipV",
      ]),
      allowPadding: false,
      path: context.path,
    });
    return {
      contract,
      resolve(ctx) {
        const box = ctx.layoutBox ?? { x: 0, y: 0, width: 80, height: 40 };
        return {
          observation: {
            id: ctx.id,
            kind: "shape",
            name:
              typeof input.props["id"] === "string"
                ? input.props["id"]
                : String(ctx.id),
            box,
          },
          transform: {
            box,
            angle: Number(
              (input.props["style"] as Record<string, unknown> | undefined)?.[
                "rotate"
              ] ?? 0,
            ),
            flipH: Boolean(
              (input.props["style"] as Record<string, unknown> | undefined)?.[
                "flipH"
              ],
            ),
            flipV: Boolean(
              (input.props["style"] as Record<string, unknown> | undefined)?.[
                "flipV"
              ],
            ),
          },
          extensions: [],
          emit: () => ({ xml: "" }),
        };
      },
    };
  },
};

const connectorProbeModule: ObjectCapabilityModule = {
  declaration: "test-connector",
  bind(input, context) {
    const hasEndpoints =
      input.props["fromId"] !== undefined || input.props["toId"] !== undefined;
    const contract = parseCapabilityLayoutContract({
      nativeKind: "cxnSp",
      props: input.props,
      allowedStyleKeys: new Set([
        "display",
        "position",
        "left",
        "top",
        "width",
        "height",
      ]),
      allowPadding: false,
      excludedFromLayout: hasEndpoints,
      path: context.path,
    });
    return {
      contract,
      resolve(ctx) {
        if (!hasEndpoints) {
          return {
            observation: {
              id: ctx.id,
              kind: "line",
              name:
                typeof input.props["id"] === "string"
                  ? input.props["id"]
                  : String(ctx.id),
              box: ctx.layoutBox ?? { x: 0, y: 0, width: 10, height: 0 },
            },
            extensions: [],
            emit: () => ({ xml: "" }),
          };
        }
        const fromId = String(input.props["fromId"]);
        const toId = String(input.props["toId"]);
        const fromTarget: ScopeTargetObject | undefined =
          ctx.resolveTargetObject?.(fromId);
        const toTarget: ScopeTargetObject | undefined =
          ctx.resolveTargetObject?.(toId);
        if (fromTarget === undefined || toTarget === undefined) {
          throw new CapabilityError(
            "broken-reference",
            `Missing target ${fromTarget === undefined ? fromId : toId}`,
          );
        }
        const x1 = fromTarget.box.x + fromTarget.box.width;
        const y1 = fromTarget.box.y + fromTarget.box.height / 2;
        const x2 = toTarget.box.x;
        const y2 = toTarget.box.y + toTarget.box.height / 2;
        return {
          observation: {
            id: ctx.id,
            kind: "line",
            name: `${fromTarget.spid}->${toTarget.spid}`,
            box: {
              x: Math.min(x1, x2),
              y: Math.min(y1, y2),
              width: Math.abs(x2 - x1),
              height: Math.abs(y2 - y1),
            },
          },
          extensions: [],
          emit: () => ({ xml: "" }),
        };
      },
    };
  },
};

const resolveTest = createResolver([
  boxOnlyModule,
  groupOnlyModule,
  targetShapeModule,
  connectorProbeModule,
]);

describe("Core scope target lookup and excludedFromLayout contract", () => {
  it("excludes endpoint-mode nodes from flex layout while resolving both backward and forward targets", () => {
    const doc = createDeckNode({ width: 960, height: 540 }).append(
      createSlideNode({ id: "s1" }).append(
        new CommonContainer("test-box", "test-box", {
          style: {
            display: "flex",
            flexDirection: "row",
            gap: 20,
            width: 400,
            height: 100,
          },
        }).append(
          new CommonNode("test-shape", "test-shape", {
            id: "a",
            style: { width: 80, height: 40 },
          }),
          // Forward reference inside flex row should NOT add a 20px gap!
          new CommonNode("test-connector", "test-connector", {
            id: "conn",
            fromId: "a",
            toId: "b",
          }),
          new CommonNode("test-shape", "test-shape", {
            id: "b",
            style: { width: 80, height: 40 },
          }),
        ),
      ),
    );

    const resolution = resolveTest(doc);
    const s1 = resolution.slide("s1");
    assert.deepEqual(s1.get("a").box, { x: 0, y: 0, width: 80, height: 40 });
    // Without excludedFromLayout, b.x would be 80 + 20 + 0 + 20 = 120. With excludedFromLayout, b.x is 100!
    assert.deepEqual(s1.get("b").box, { x: 100, y: 0, width: 80, height: 40 });
    assert.deepEqual(s1.get("conn").box, {
      x: 80,
      y: 20,
      width: 20,
      height: 0,
    });
    assert.equal(
      s1.get("conn").objects[0]?.name,
      `${s1.get("a").objects[0]!.id}->${s1.get("b").objects[0]!.id}`,
    );
  });

  it("resolves targets across anonymous containers and inside groups, while rejecting hidden, non-native, and cross-scope targets", () => {
    const validGroupDoc = createDeckNode({ width: 960, height: 540 }).append(
      createSlideNode({ id: "s1" }).append(
        new CommonContainer("test-box", "test-box", {
          style: { width: 300, height: 200 },
        }).append(
          new CommonNode("test-shape", "test-shape", {
            id: "inside-anon",
            style: {
              position: "absolute",
              left: 10,
              top: 10,
              width: 50,
              height: 30,
            },
          }),
        ),
        new CommonNode("test-connector", "test-connector", {
          id: "c-anon",
          fromId: "inside-anon",
          toId: "outer",
        }),
        new CommonNode("test-shape", "test-shape", {
          id: "outer",
          style: {
            position: "absolute",
            left: 200,
            top: 10,
            width: 50,
            height: 30,
          },
        }),
        new CommonContainer("test-group", "test-group", { id: "g1" }).append(
          new CommonNode("test-connector", "test-connector", {
            id: "g-conn",
            fromId: "g-a",
            toId: "g-b",
          }),
          new CommonNode("test-shape", "test-shape", {
            id: "g-a",
            style: {
              position: "absolute",
              left: 20,
              top: 100,
              width: 40,
              height: 20,
            },
          }),
          new CommonNode("test-shape", "test-shape", {
            id: "g-b",
            style: {
              position: "absolute",
              left: 120,
              top: 100,
              width: 40,
              height: 20,
            },
          }),
        ),
      ),
    );
    const res = resolveTest(validGroupDoc);
    assert.deepEqual(res.slide("s1").get("c-anon").box, {
      x: 60,
      y: 25,
      width: 140,
      height: 0,
    });
    assert.deepEqual(res.slide("s1").get("g1").get("g-conn").box, {
      x: 60,
      y: 110,
      width: 60,
      height: 0,
    });

    const assertBrokenReference = (fn: () => unknown): void => {
      assert.throws(
        fn,
        (error) =>
          error instanceof BuildError &&
          error.diagnostics.some((entry) => entry.code === "broken-reference"),
      );
    };

    // Hidden target -> broken-reference
    assertBrokenReference(() =>
      resolveTest(
        createDeckNode({ width: 960, height: 540 }).append(
          createSlideNode({ id: "s1" }).append(
            new CommonNode("test-shape", "test-shape", {
              id: "a",
              style: { display: "none", width: 40, height: 20 },
            }),
            new CommonNode("test-shape", "test-shape", {
              id: "b",
              style: {
                position: "absolute",
                left: 100,
                top: 0,
                width: 40,
                height: 20,
              },
            }),
            new CommonNode("test-connector", "test-connector", {
              fromId: "a",
              toId: "b",
            }),
          ),
        ),
      ),
    );

    // Pure logical Box (emitsObject: false) -> broken-reference
    assertBrokenReference(() =>
      resolveTest(
        createDeckNode({ width: 960, height: 540 }).append(
          createSlideNode({ id: "s1" }).append(
            new CommonContainer("test-box", "test-box", {
              id: "logical-box",
              emits: false,
            }),
            new CommonNode("test-shape", "test-shape", {
              id: "b",
              style: {
                position: "absolute",
                left: 100,
                top: 0,
                width: 40,
                height: 20,
              },
            }),
            new CommonNode("test-connector", "test-connector", {
              fromId: "logical-box",
              toId: "b",
            }),
          ),
        ),
      ),
    );

    // Cross-scope target inside named Box -> broken-reference
    assertBrokenReference(() =>
      resolveTest(
        createDeckNode({ width: 960, height: 540 }).append(
          createSlideNode({ id: "s1" }).append(
            new CommonContainer("test-box", "test-box", {
              id: "named-box",
            }).append(
              new CommonNode("test-shape", "test-shape", {
                id: "scoped",
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 40,
                  height: 20,
                },
              }),
            ),
            new CommonNode("test-shape", "test-shape", {
              id: "b",
              style: {
                position: "absolute",
                left: 100,
                top: 0,
                width: 40,
                height: 20,
              },
            }),
            new CommonNode("test-connector", "test-connector", {
              fromId: "scoped",
              toId: "b",
            }),
          ),
        ),
      ),
    );
  });
});
