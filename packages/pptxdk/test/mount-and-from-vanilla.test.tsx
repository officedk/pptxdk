import assert from "node:assert/strict";

import {
  build,
  BuildError,
  CommonContainer,
  CommonNode,
  createMotionScopeHandle,
  defineTheme,
  fade,
  motionSchedule,
  resolve,
} from "@officedk/pptxdk";
import {
  Box,
  Deck,
  expandDocument,
  expandFragment,
  fromVanilla,
  MotionScope,
  Mount,
  Slide,
  Step,
  Text,
  type MountProps,
  type VanillaSlotResult,
} from "@officedk/pptxdk/react";
import { box, deck, image, shape, slide, text } from "@officedk/pptxdk/vanilla";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import {
  createElement,
  Fragment,
  Suspense,
  use,
  useEffect,
  useState,
} from "react";
import { describe, it } from "vitest";

describe("public React Mount adoption", () => {
  it("adopts single node, node array, and empty values without adding wrapper nodes", async () => {
    const heroBox = box({
      id: "hero-box",
      style: {
        position: "absolute",
        left: 24,
        top: 24,
        width: 300,
        height: 120,
        display: "flex",
        flexDirection: "column",
        gap: 8,
        padding: 12,
        background: "#EEF2FF",
      },
    }).append(
      shape("Imperative header", {
        id: "hero-header",
        style: { height: 32, background: "#3B82F6", color: "#FFFFFF" },
      }),
      text("Imperative body", {
        id: "hero-body",
        style: { height: 28 },
      }),
    );

    const badgeA = shape("A", {
      id: "badge-a",
      style: {
        position: "absolute",
        left: 360,
        top: 24,
        width: 80,
        height: 36,
      },
    });
    const badgeB = shape("B", {
      id: "badge-b",
      style: {
        position: "absolute",
        left: 460,
        top: 24,
        width: 80,
        height: 36,
      },
    });

    const nullProps: MountProps = { node: null };
    const undefProps: MountProps = { node: undefined };

    const doc = await expandDocument(
      createElement(
        Deck,
        { width: 640, height: 360 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(Mount, nullProps),
          createElement(Mount, { node: heroBox }),
          createElement(Mount, { node: [badgeA, badgeB] }),
          createElement(Mount, undefProps),
          createElement(
            Text,
            {
              id: "jsx-after",
              style: {
                position: "absolute",
                left: 24,
                top: 180,
                width: 200,
                height: 30,
              },
            },
            "JSX sibling",
          ),
        ),
      ),
    );

    const slideNode = doc.query.get("s1").node;
    assert.ok(slideNode instanceof CommonContainer);
    assert.equal(slideNode.children.length, 4);
    assert.equal(slideNode.children[0], heroBox);
    assert.equal(slideNode.children[1], badgeA);
    assert.equal(slideNode.children[2], badgeB);
    assert.ok(slideNode.children[3] instanceof CommonNode);
    assert.equal(slideNode.children[3].id, "jsx-after");

    // Imperative props and preset children remain intact.
    assert.equal(heroBox.children.length, 2);
    assert.equal(heroBox.props["id"], "hero-box");
    assert.equal(
      heroBox.query.get("hero-header").node.props["id"],
      "hero-header",
    );

    const built = await build(doc);
    assert.deepEqual((await validatePptx(built.pptx)).errors, []);
    assert.deepEqual(
      built.report.slides[0]?.objects.map((o) => o.authorId),
      [
        "hero-box",
        "hero-header",
        "hero-body",
        "badge-a",
        "badge-b",
        "jsx-after",
      ],
    );

    // Also works at top level of expandFragment returning the exact adopted instances.
    const topA = shape("Top A", {
      id: "top-a",
      style: { width: 100, height: 30 },
    });
    const topB = shape("Top B", {
      id: "top-b",
      style: { width: 100, height: 30 },
    });
    const frag = await expandFragment(
      createElement(
        Fragment,
        null,
        createElement(Mount, { node: [topA, topB] }),
        createElement(Mount, { node: null }),
      ),
    );
    assert.equal(frag.length, 2);
    assert.equal(frag[0], topA);
    assert.equal(frag[1], topB);
    assert.equal(topA.parent, undefined);
    assert.equal(topB.parent, undefined);
  });

  it("stays idempotent across re-renders, swaps references cleanly, and detaches on unmount", async () => {
    const persistent = shape("Initial", {
      id: "persistent",
      style: {
        position: "absolute",
        left: 20,
        top: 20,
        width: 160,
        height: 40,
      },
    });
    const firstSwap = shape("Swap A", {
      id: "swap-a",
      style: {
        position: "absolute",
        left: 20,
        top: 80,
        width: 120,
        height: 32,
      },
    });
    const secondSwap = shape("Swap B", {
      id: "swap-b",
      style: {
        position: "absolute",
        left: 20,
        top: 80,
        width: 120,
        height: 32,
      },
    });
    const ephemeral = box({ id: "ephemeral" }).append(
      shape("Child stays inside", { id: "ephemeral-child" }),
    );

    const gate = Promise.withResolvers<void>();
    function ReRenderingMount() {
      const [phase, setPhase] = useState(0);
      useEffect(() => {
        if (phase === 0) {
          // Mutate the adopted node imperatively before a re-render with the same node reference.
          persistent.content("Imperative edit").style({ color: "#DC2626" });
          setPhase(1);
        } else if (phase === 1) {
          setPhase(2);
          gate.resolve();
        }
      }, [phase]);

      return createElement(
        Slide,
        { id: "swap-slide" },
        createElement(
          Text,
          {
            id: "anchor-before",
            style: {
              position: "absolute",
              left: 0,
              top: 0,
              width: 80,
              height: 20,
            },
          },
          "Before",
        ),
        createElement(Mount, { node: persistent }),
        createElement(Mount, {
          node: phase === 0 ? null : phase === 1 ? [firstSwap] : secondSwap,
        }),
        phase < 2 ? createElement(Mount, { node: ephemeral }) : null,
        createElement(
          Text,
          {
            id: "anchor-after",
            style: {
              position: "absolute",
              left: 0,
              top: 140,
              width: 80,
              height: 20,
            },
          },
          "After",
        ),
      );
    }

    const docPromise = expandDocument(
      createElement(
        Deck,
        { width: 640, height: 360 },
        createElement(ReRenderingMount),
      ),
    );
    await gate.promise;
    const doc = await docPromise;
    const slideNode = doc.query.get("swap-slide").node;
    assert.deepEqual(
      slideNode.children.map((c) =>
        c instanceof CommonNode ? c.id : undefined,
      ),
      ["anchor-before", "persistent", "swap-b", "anchor-after"],
    );
    assert.equal(
      persistent.props["style"] &&
        (persistent.props["style"] as Record<string, unknown>)["color"],
      "#DC2626",
    );
    assert.equal(firstSwap.parent, undefined);
    assert.equal(ephemeral.parent, undefined);
    assert.equal(ephemeral.children.length, 1);

    const built = await build(doc);
    assert.deepEqual((await validatePptx(built.pptx)).errors, []);
    assert.equal(
      built.report.slides[0]?.objects.find((o) => o.authorId === "persistent")
        ?.text,
      "Imperative edit",
    );
  });

  it("enforces strict single-parent ownership and handles Suspense retries without false conflict", async () => {
    // 1. Node already held by an external container is rejected without stealing it.
    const externalParent = box({ id: "ext-parent" });
    const ownedChild = shape("Owned", { id: "owned" });
    externalParent.append(ownedChild);

    await assert.rejects(
      expandDocument(
        createElement(
          Deck,
          { width: 320, height: 180 },
          createElement(
            Slide,
            { id: "s1" },
            createElement(Mount, { node: ownedChild }),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-structure");
        return true;
      },
    );
    assert.equal(ownedChild.parent, externalParent);
    assert.deepEqual([...externalParent.children], [ownedChild]);

    // 2. Same node mounted in two active <Mount> elements simultaneously is rejected.
    const shared = shape("Shared", { id: "shared" });
    await assert.rejects(
      expandDocument(
        createElement(
          Deck,
          { width: 320, height: 180 },
          createElement(
            Slide,
            { id: "s1" },
            createElement(Mount, { node: shared }),
            createElement(Mount, { node: shared }),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-structure");
        return true;
      },
    );
    assert.equal(shared.parent, undefined);

    // 3. Same node duplicated inside node={[a, a]} is rejected.
    await assert.rejects(
      expandFragment(createElement(Mount, { node: [shared, shared] })),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-structure");
        return true;
      },
    );

    // 4. Node previously owned by another Deck (even after remove()) is rejected.
    const foreignNode = shape("Foreign", { id: "foreign" });
    deck({ width: 320, height: 180 }).append(
      slide({ id: "foreign-slide" }).append(foreignNode),
    );
    foreignNode.remove();
    assert.equal(foreignNode.parent, undefined);

    await assert.rejects(
      expandDocument(
        createElement(
          Deck,
          { width: 320, height: 180 },
          createElement(
            Slide,
            { id: "s1" },
            createElement(Mount, { node: foreignNode }),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-structure");
        return true;
      },
    );

    // 5. Suspense retry next to <Mount> does not trigger false double-mount.
    const suspenseNode = shape("Suspense safe", {
      id: "suspense-safe",
      style: {
        position: "absolute",
        left: 10,
        top: 10,
        width: 120,
        height: 30,
      },
    });
    const gate = Promise.withResolvers<string>();
    function AsyncSibling() {
      return createElement(
        Text,
        {
          id: "async-sibling",
          style: {
            position: "absolute",
            left: 10,
            top: 50,
            width: 120,
            height: 30,
          },
        },
        use(gate.promise),
      );
    }
    const expanding = expandDocument(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Suspense,
          { fallback: null },
          createElement(
            Slide,
            { id: "suspense-slide" },
            createElement(Mount, { node: suspenseNode }),
            createElement(AsyncSibling),
          ),
        ),
      ),
    );
    gate.resolve("Loaded");
    const doc = await expanding;
    assert.equal(
      doc.query.get("suspense-slide").get("suspense-safe").node,
      suspenseNode,
    );
  });

  it("reconciles JSX children into descendant slot after preset children and rejects invalid slots", async () => {
    const card = box({
      id: "panel",
      style: {
        position: "absolute",
        left: 24,
        top: 24,
        width: 320,
        height: 220,
        display: "flex",
        flexDirection: "column",
        gap: 8,
        padding: 12,
        background: "#F8FAFC",
      },
    });
    const presetHeader = shape("Header", {
      id: "preset-header",
      style: { height: 32, background: "#1E293B", color: "#FFFFFF" },
    });
    const slotBody = box({
      id: "slot-body",
      style: { flex: 1, display: "flex", flexDirection: "column", gap: 6 },
    });
    const presetLead = shape("Preset lead", {
      id: "preset-lead",
      style: { height: 28 },
    });
    slotBody.append(presetLead);
    card.append(presetHeader, slotBody);

    const gate = Promise.withResolvers<string>();
    function AsyncSlotItem() {
      return createElement(
        Text,
        { id: "jsx-async", style: { height: 28 } },
        use(gate.promise),
      );
    }

    const docPromise = expandDocument(
      createElement(
        Deck,
        { width: 640, height: 360 },
        createElement(
          Slide,
          { id: "slot-slide" },
          createElement(
            Mount,
            { node: card, slot: slotBody },
            createElement(
              Suspense,
              {
                fallback: createElement(
                  Text,
                  { id: "slot-fallback", style: { height: 28 } },
                  "Loading slot",
                ),
              },
              createElement(AsyncSlotItem),
            ),
            createElement(
              Text,
              { id: "jsx-tail", style: { height: 28 } },
              "Static tail",
            ),
          ),
        ),
      ),
    );
    gate.resolve("Resolved slot item");
    const doc = await docPromise;

    assert.equal(card.children.length, 2);
    assert.equal(card.children[0], presetHeader);
    assert.equal(card.children[1], slotBody);
    assert.deepEqual(
      slotBody.children.map((c) =>
        c instanceof CommonNode ? c.id : undefined,
      ),
      ["preset-lead", "jsx-async", "jsx-tail"],
    );

    const built = await build(doc);
    assert.deepEqual((await validatePptx(built.pptx)).errors, []);
    assert.deepEqual(
      built.report.slides[0]?.objects.map((o) => o.authorId),
      ["panel", "preset-header", "preset-lead", "jsx-async", "jsx-tail"],
    );

    // Slot === node itself also works.
    const selfSlot = box({ id: "self-slot" }).append(
      shape("Preset first", { id: "self-preset", style: { height: 24 } }),
    );
    const selfFrag = await expandFragment(
      createElement(
        Mount,
        { node: selfSlot, slot: selfSlot },
        createElement(
          Text,
          { id: "self-jsx", style: { height: 24 } },
          "Appended",
        ),
      ),
    );
    assert.equal(selfFrag[0], selfSlot);
    assert.deepEqual(
      selfSlot.children.map((c) =>
        c instanceof CommonNode ? c.id : undefined,
      ),
      ["self-preset", "self-jsx"],
    );

    // Rejects unrelated slot container.
    const targetNode = box({ id: "target-node" });
    const unrelatedSlot = box({ id: "unrelated-slot" });
    await assert.rejects(
      expandFragment(
        createElement(Mount, { node: targetNode, slot: unrelatedSlot }),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-structure");
        return true;
      },
    );

    // Rejects non-container slot (ImageNode is CommonNode, not CommonContainer).
    const imgChild = image({
      src: new Uint8Array([137, 80, 78, 71]),
      style: { width: 10, height: 10 },
    });
    const hostBox = box().append(imgChild);
    await assert.rejects(
      expandFragment(
        createElement(Mount, {
          node: hostBox,
          // @ts-expect-error ImageNode is not a CommonContainer
          slot: imgChild,
        }),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-structure");
        return true;
      },
    );
  });

  it("participates dynamically in outer Flex/Grid layout, Step, MotionScope stagger, and ThemeScope", async () => {
    const scopeHandle = createMotionScopeHandle();
    const outerTheme = defineTheme({
      name: "OuterTheme",
      colors: { accent: "#1D4ED8" },
    });
    const scopedTheme = {
      colors: { accent: "#BE123C" },
    };

    const createTile = (id: string) =>
      shape(id, {
        id,
        style: {
          flex: 1,
          height: 44,
          background: "accent",
          color: "#FFFFFF",
        },
        enter: fade({ duration: 150 }),
      });

    const narrowA = createTile("narrow-a");
    const narrowB = createTile("narrow-b");
    const wideA = createTile("wide-a");
    const wideB = createTile("wide-b");

    const doc = await expandDocument(
      createElement(
        Deck,
        { width: 800, height: 450, theme: outerTheme },
        createElement(
          Slide,
          { id: "dynamic-slide" },
          createElement(
            Step,
            { n: 1 },
            createElement(
              MotionScope,
              { scope: scopeHandle, mode: "stagger", each: 80 },
              createElement(
                Box,
                {
                  id: "narrow-grid",
                  style: {
                    position: "absolute",
                    left: 20,
                    top: 20,
                    width: 240,
                    height: 80,
                    display: "grid",
                    gridTemplateColumns: "1fr 1fr",
                    gap: 20,
                  },
                },
                createElement(Mount, { node: narrowA }),
                createElement(Mount, { node: narrowB }),
              ),
              createElement(
                Box,
                {
                  id: "wide-flex",
                  theme: scopedTheme,
                  style: {
                    position: "absolute",
                    left: 20,
                    top: 120,
                    width: 500,
                    height: 80,
                    display: "flex",
                    flexDirection: "row",
                    gap: 20,
                  },
                },
                createElement(Mount, { node: wideA }),
                createElement(Mount, { node: wideB }),
              ),
            ),
          ),
        ),
      ),
    );

    const built = await build(doc);
    assert.deepEqual((await validatePptx(built.pptx)).errors, []);
    const resolved = resolve(doc);
    const objs = resolved.slides[0]?.objects ?? [];
    const narrowABox = objs.find((o) => o.authorId === "narrow-a")?.box;
    const narrowBBox = objs.find((o) => o.authorId === "narrow-b")?.box;
    const wideABox = objs.find((o) => o.authorId === "wide-a")?.box;
    const wideBBox = objs.find((o) => o.authorId === "wide-b")?.box;
    // Narrow grid (240 - 20 gap) / 2 = 110 each; Wide flex (500 - 20 gap) / 2 = 240 each.
    assert.equal(narrowABox?.width, 110);
    assert.equal(narrowBBox?.width, 110);
    assert.equal(wideABox?.width, 240);
    assert.equal(wideBBox?.width, 240);

    // Scoped ThemeScope overrides accent color on wide tiles vs outer theme on narrow tiles.
    const xml = strFromU8(unzipSync(built.pptx)["ppt/slides/slide1.xml"]!);
    assert.match(xml, /1D4ED8/u);
    assert.match(xml, /BE123C/u);

    // MotionScope stagger + Step 1 applies to all 4 mounted tiles.
    const schedule = motionSchedule(built.slides[0]!);
    assert.ok(schedule);
    const step1 = schedule.steps.find((s) => s.number === 1);
    assert.ok(step1);
    assert.equal(step1.trigger, "click");
    assert.deepEqual(
      step1.effects.map((e) => (e.start.state === "known" ? e.start.ms : -1)),
      [0, 80, 160, 240],
    );
  });
});

describe("public React fromVanilla bridge", () => {
  it("distinguishes closed and slot container factories statically and runs closed components dynamically", async () => {
    interface MetricTileProps {
      readonly id: string;
      readonly label: string;
      readonly value: string;
    }

    const createBadge = () =>
      shape("Zero-arg", {
        id: "zero-arg",
        style: {
          width: 100,
          height: 32,
          background: "accent",
          color: "#FFFFFF",
        },
      });
    const createMetricTile = ({ id, label, value }: MetricTileProps) =>
      box({
        id,
        style: {
          flex: 1,
          display: "flex",
          flexDirection: "column",
          gap: 4,
          padding: 8,
          background: "#F1F5F9",
        },
        enter: fade({ duration: 120 }),
      }).append(
        shape(label, { id: `${id}-label`, style: { height: 18 } }),
        text(value, {
          id: `${id}-value`,
          style: { height: 24, color: "accent" },
        }),
      );

    const Badge = fromVanilla(createBadge);
    const MetricTile = fromVanilla(createMetricTile);

    // Static type checks: closed components reject children.
    const checkClosedRejectsChildren = () => {
      const badZero = (
        <Badge>
          {/* @ts-expect-error closed zero-arg vanilla component does not accept children */}
          <Text>nope</Text>
        </Badge>
      );
      const badParam = (
        <MetricTile id="x" label="L" value="1">
          {/* @ts-expect-error closed parameterized vanilla component does not accept children */}
          <Text>nope</Text>
        </MetricTile>
      );
      return [badZero, badParam];
    };
    void checkClosedRejectsChildren;

    const scopeHandle = createMotionScopeHandle();
    const doc = await expandDocument(
      <Deck
        width={800}
        height={450}
        theme={defineTheme({
          name: "BridgeTheme",
          colors: { accent: "#0F766E" },
        })}
      >
        <Slide id="bridge-slide">
          <Step n={1}>
            <MotionScope scope={scopeHandle} mode="stagger" each={60}>
              <Box
                id="narrow"
                style={{
                  position: "absolute",
                  left: 20,
                  top: 20,
                  width: 240,
                  height: 90,
                  display: "flex",
                  flexDirection: "row",
                  gap: 20,
                }}
              >
                <MetricTile id="tile-n1" label="N1" value="10" />
                <MetricTile id="tile-n2" label="N2" value="20" />
              </Box>
              <Box
                id="wide"
                theme={{ colors: { accent: "#9333EA" } }}
                style={{
                  position: "absolute",
                  left: 20,
                  top: 130,
                  width: 500,
                  height: 90,
                  display: "flex",
                  flexDirection: "row",
                  gap: 20,
                }}
              >
                <MetricTile id="tile-w1" label="W1" value="30" />
                <MetricTile id="tile-w2" label="W2" value="40" />
              </Box>
            </MotionScope>
          </Step>
          <Box
            style={{
              position: "absolute",
              left: 20,
              top: 240,
              width: 120,
              height: 40,
            }}
          >
            <Badge />
          </Box>
        </Slide>
      </Deck>,
    );

    const built = await build(doc);
    assert.deepEqual((await validatePptx(built.pptx)).errors, []);
    const resolved = resolve(doc);
    const objs = resolved.slides[0]?.objects ?? [];
    assert.equal(objs.find((o) => o.authorId === "tile-n1")?.box?.width, 110);
    assert.equal(objs.find((o) => o.authorId === "tile-w1")?.box?.width, 240);
    const xml = strFromU8(unzipSync(built.pptx)["ppt/slides/slide1.xml"]!);
    assert.match(xml, /0F766E/u);
    assert.match(xml, /9333EA/u);

    const schedule = motionSchedule(built.slides[0]!);
    assert.ok(schedule);
    assert.deepEqual(
      schedule.steps[0]?.effects.map((e) =>
        e.start.state === "known" ? e.start.ms : -1,
      ),
      [0, 0, 0, 60, 60, 60, 120, 120, 120, 180, 180, 180],
    );

    // Runtime guard if a caller bypasses TypeScript and passes children to a closed component.
    const UnsafeBadge = Badge as (props: { children?: unknown }) => unknown;
    await assert.rejects(
      expandFragment(
        createElement(
          UnsafeBadge as never,
          null,
          createElement(Text, null, "x"),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-structure");
        return true;
      },
    );
  });

  it("stabilizes slot container shells across re-renders and Suspense, and replaces on prop changes", async () => {
    let builds = 0;
    const createdNodes: CommonNode[] = [];

    function createPanel(props: { readonly title: string }): VanillaSlotResult {
      builds += 1;
      const root = box({
        id: `panel-${builds}`,
        style: {
          position: "absolute",
          left: 24,
          top: 24,
          width: 400,
          height: 220,
          display: "flex",
          flexDirection: "column",
          gap: 8,
          padding: 12,
          background: "#F8FAFC",
        },
      });
      const header = shape(props.title, {
        id: `header-${builds}`,
        style: { height: 32, background: "#0F172A", color: "#FFFFFF" },
      });
      const slot = box({
        id: `slot-${builds}`,
        style: { flex: 1, display: "flex", flexDirection: "column", gap: 6 },
      });
      root.append(header, slot);
      createdNodes.push(root);
      return { node: root, slot };
    }

    const Panel = fromVanilla(createPanel);
    const createZeroSlot = (): VanillaSlotResult => {
      const root = box({ id: "zero-slot-root" });
      return { node: root, slot: root };
    };
    const ZeroSlot = fromVanilla(createZeroSlot);

    const asyncGate = Promise.withResolvers<string>();
    const rerenderGate = Promise.withResolvers<void>();

    function AsyncSlotChild() {
      return (
        <Text id="async-child" style={{ height: 24 }}>
          {use(asyncGate.promise)}
        </Text>
      );
    }

    function Harness() {
      const [tick, setTick] = useState(0);
      const [title, setTitle] = useState("Initial Panel");
      useEffect(() => {
        if (tick === 0) {
          // Parent re-render with identical non-children props must reuse the same shell instance.
          setTick(1);
        } else if (tick === 1) {
          asyncGate.resolve("Loaded child");
          rerenderGate.resolve();
        }
      }, [tick]);

      // Expose setTitle trigger via another state step? Keep two separate tests for clarity:
      void setTitle;
      return (
        <Slide id="panel-slide">
          <Panel title={title}>
            <Suspense
              fallback={
                <Text id="panel-loading" style={{ height: 24 }}>
                  Loading
                </Text>
              }
            >
              <AsyncSlotChild />
            </Suspense>
            <Text id="tick-text" style={{ height: 24 }}>
              {`Tick ${tick}`}
            </Text>
          </Panel>
          <ZeroSlot>
            <Text id="zero-slot-child" style={{ height: 20 }}>
              Zero slot child
            </Text>
          </ZeroSlot>
        </Slide>
      );
    }

    const docPromise = expandDocument(
      <Deck width={640} height={360}>
        <Harness />
      </Deck>,
    );
    await rerenderGate.promise;
    const doc = await docPromise;

    // Only built once across parent re-render + Suspense resolution!
    assert.equal(builds, 1);
    assert.equal(createdNodes.length, 1);
    const panelNode = doc.query.get("panel-slide").get("panel-1").node;
    assert.equal(panelNode, createdNodes[0]);
    const slotNode = panelNode.query.get("slot-1").node;
    assert.deepEqual(
      slotNode.children.map((c) =>
        c instanceof CommonNode ? c.id : undefined,
      ),
      ["async-child", "tick-text"],
    );
    assert.equal(panelNode.query.get("header-1").node.props["id"], "header-1");

    // Changing a non-children prop swaps the shell cleanly.
    const swapGate = Promise.withResolvers<void>();
    function PropSwapHarness() {
      const [title, setTitle] = useState("Before Swap");
      useEffect(() => {
        if (title === "Before Swap") {
          setTitle("After Swap");
          swapGate.resolve();
        }
      }, [title]);
      return (
        <Slide id="swap-panel-slide">
          <Panel title={title}>
            <Text id="persisted-child" style={{ height: 24 }}>
              Carried over
            </Text>
          </Panel>
        </Slide>
      );
    }

    const swapPromise = expandDocument(
      <Deck width={640} height={360}>
        <PropSwapHarness />
      </Deck>,
    );
    await swapGate.promise;
    const swappedDoc = await swapPromise;
    assert.equal(builds, 3);
    assert.equal(createdNodes[1]?.parent, undefined);
    const activePanel = swappedDoc.query
      .get("swap-panel-slide")
      .get("panel-3").node;
    assert.equal(activePanel, createdNodes[2]);
    assert.equal(
      activePanel.query.get("slot-3").get("persisted-child").node.id,
      "persisted-child",
    );
  });

  it("infers exact JSX props for zero-arg, required-prop, optional-prop, and default-prop closed and slotted vanilla factories", async () => {
    const ZeroClosed = fromVanilla(() => shape("zero", { id: "zero-closed" }));
    const ReqClosed = fromVanilla(
      (props: { readonly id: string; readonly label: string }) =>
        shape(props.label, { id: props.id }),
    );
    const OptClosed = fromVanilla(
      (props?: { readonly id?: string; readonly tone?: "warm" | "cool" }) =>
        shape(props?.tone ?? "neutral", {
          id: props?.id ?? "opt-closed",
        }),
    );
    const DefClosed = fromVanilla(
      (props: { readonly id?: string; readonly tone?: "warm" | "cool" } = {}) =>
        shape(props.tone ?? "default", {
          id: props.id ?? "def-closed",
        }),
    );

    const ZeroSlot = fromVanilla((): VanillaSlotResult => {
      const root = box({ id: "zero-slot" });
      return { node: root, slot: root };
    });
    const ReqSlot = fromVanilla(
      (props: { readonly id: string }): VanillaSlotResult => {
        const root = box({ id: props.id });
        return { node: root, slot: root };
      },
    );
    const OptSlot = fromVanilla(
      (props?: {
        readonly id?: string;
        readonly tone?: "warm" | "cool";
      }): VanillaSlotResult => {
        const slot = box({ id: `${props?.id ?? "opt-slot"}-inner` });
        const root = box({ id: props?.id ?? "opt-slot" }).append(
          shape(props?.tone ?? "neutral", {
            id: `${props?.id ?? "opt-slot"}-header`,
          }),
          slot,
        );
        return { node: root, slot };
      },
    );
    const DefSlot = fromVanilla(
      (
        props: {
          readonly id?: string;
          readonly tone?: "warm" | "cool";
        } = {},
      ): VanillaSlotResult => {
        const slot = box({ id: `${props.id ?? "def-slot"}-inner` });
        const root = box({ id: props.id ?? "def-slot" }).append(
          shape(props.tone ?? "default", {
            id: `${props.id ?? "def-slot"}-header`,
          }),
          slot,
        );
        return { node: root, slot };
      },
    );

    const checkTypes = () => (
      <>
        <ZeroClosed />
        {/* @ts-expect-error zero-arg closed component must reject unknown props */}
        <ZeroClosed unknownProp={1} />
        <ZeroClosed>
          {/* @ts-expect-error zero-arg closed component must reject children */}
          <Text>nope</Text>
        </ZeroClosed>

        {/* @ts-expect-error required-prop closed component must require props */}
        <ReqClosed />
        <ReqClosed id="r1" label="ok">
          {/* @ts-expect-error required-prop closed component must reject children */}
          <Text>nope</Text>
        </ReqClosed>

        <OptClosed />
        <OptClosed id="o1" tone="warm" />
        {/* @ts-expect-error optional-prop closed component must reject wrong prop type */}
        <OptClosed tone={123} />
        {/* @ts-expect-error optional-prop closed component must reject unknown props */}
        <OptClosed unknownProp={1} />
        <OptClosed tone="warm">
          {/* @ts-expect-error optional-prop closed component must reject children */}
          <Text>nope</Text>
        </OptClosed>

        <DefClosed />
        <DefClosed id="d1" tone="cool" />
        {/* @ts-expect-error default-prop closed component must reject wrong prop type */}
        <DefClosed tone={123} />
        <DefClosed tone="cool">
          {/* @ts-expect-error default-prop closed component must reject children */}
          <Text>nope</Text>
        </DefClosed>

        <ZeroSlot />
        {/* @ts-expect-error zero-arg slot component must reject unknown props */}
        <ZeroSlot unknownProp={1} />

        {/* @ts-expect-error required-prop slot component must require props */}
        <ReqSlot />

        <OptSlot />
        <OptSlot id="os1" tone="warm">
          <Text id="os1-child">in opt</Text>
        </OptSlot>
        {/* @ts-expect-error optional-prop slot component must reject wrong prop type */}
        <OptSlot tone={123} />

        <DefSlot />
        <DefSlot id="ds1" tone="cool">
          <Text id="ds1-child">in def</Text>
        </DefSlot>
        {/* @ts-expect-error default-prop slot component must reject wrong prop type */}
        <DefSlot tone={123} />
      </>
    );
    void checkTypes;

    const doc = await expandDocument(
      <Deck width={640} height={360}>
        <Slide id="s-shapes">
          <OptClosed id="opt-default" />
          <OptClosed id="opt-warm" tone="warm" />
          <DefClosed id="def-default" />
          <DefClosed id="def-cool" tone="cool" />
          <OptSlot id="slot-opt-default">
            <Text id="slot-opt-default-child">child A</Text>
          </OptSlot>
          <OptSlot id="slot-opt-warm" tone="warm">
            <Text id="slot-opt-warm-child">child B</Text>
          </OptSlot>
          <DefSlot id="slot-def-default">
            <Text id="slot-def-default-child">child C</Text>
          </DefSlot>
          <DefSlot id="slot-def-cool" tone="cool">
            <Text id="slot-def-cool-child">child D</Text>
          </DefSlot>
        </Slide>
      </Deck>,
    );

    const slideNode = doc.query.get("s-shapes").node;
    const firstText = (nodeId: string, scopeId?: string) => {
      const target =
        scopeId === undefined
          ? slideNode.query.get(nodeId).node
          : slideNode.query.get(scopeId).get(nodeId).node;
      const child = target.children[0];
      return child !== undefined && "text" in child ? child.text : undefined;
    };
    assert.equal(firstText("opt-default"), "neutral");
    assert.equal(firstText("opt-warm"), "warm");
    assert.equal(firstText("def-default"), "default");
    assert.equal(firstText("def-cool"), "cool");
    assert.equal(
      firstText("slot-opt-default-header", "slot-opt-default"),
      "neutral",
    );
    assert.equal(firstText("slot-opt-warm-header", "slot-opt-warm"), "warm");
    assert.equal(
      firstText("slot-def-default-header", "slot-def-default"),
      "default",
    );
    assert.equal(firstText("slot-def-cool-header", "slot-def-cool"), "cool");
    assert.equal(
      slideNode.query
        .get("slot-opt-warm")
        .get("slot-opt-warm-inner")
        .get("slot-opt-warm-child").node.id,
      "slot-opt-warm-child",
    );
    assert.equal(
      slideNode.query
        .get("slot-def-cool")
        .get("slot-def-cool-inner")
        .get("slot-def-cool-child").node.id,
      "slot-def-cool-child",
    );
  });
});
