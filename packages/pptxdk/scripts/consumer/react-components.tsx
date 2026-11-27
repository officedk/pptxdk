import { readFileSync } from "node:fs";

import {
  Box,
  Cell,
  Chart,
  Deck,
  Group,
  Mount,
  Row,
  Shape,
  Slide,
  SlideLayout,
  SlideMaster,
  Table,
  Text,
  darker,
  defaultColorMap,
  defineSlideLayout,
  defineTheme,
  expandFragment,
  font,
  fromVanilla,
  lighter,
  type ColorMapOverride,
  type CommonContainer,
  type CommonNode,
  type ExpandFragmentOptions,
  type FontReference,
  type HexColor,
  type MountProps,
  type NativeColorSlot,
  type SchemeColorName,
  type Theme,
  type ThemeColor,
  type ThemeFontSlot,
  type ThemeInput,
  type ThemeOverride,
  type VanillaSlotResult,
} from "@officedk/pptxdk/react";
import { box, shape, text } from "@officedk/pptxdk/vanilla";
import type { ReactElement } from "react";

const sharedPadding = [8, 12] as const;

const consumerLayout = defineSlideLayout({
  id: "consumer-layout",
  name: "Consumer Layout",
  type: "obj",
  placeholders: {
    title: {
      type: "title",
      style: {
        position: "absolute",
        left: 24,
        top: 16,
        width: 592,
        height: 32,
        fontSize: 18,
        fontWeight: 400,
      },
    },
    body: {
      type: "body",
      idx: 1,
      style: {
        position: "absolute",
        left: 24,
        top: 56,
        width: 592,
        height: 260,
        fontSize: 13,
      },
    },
  },
});

function createClosedBadge(props: {
  readonly id: string;
  readonly label: string;
}): CommonNode {
  return shape(props.label, {
    id: props.id,
    style: { height: 28, fontSize: 12 },
  });
}

function createCardShell(props: {
  readonly id: string;
  readonly title: string;
}): VanillaSlotResult<CommonContainer, CommonContainer> {
  const slot = box({
    id: `${props.id}-slot`,
    style: { display: "flex", flexDirection: "column", gap: 6 },
  });
  const node = box({
    id: props.id,
    style: {
      display: "flex",
      flexDirection: "column",
      gap: 6,
      padding: sharedPadding,
    },
  }).append(
    shape(props.title, {
      id: `${props.id}-header`,
      style: { height: 24, fontSize: 12 },
    }),
    slot,
  );
  return { node, slot };
}

const ClosedBadge = fromVanilla(createClosedBadge);
const CardShell = fromVanilla(createCardShell);
const OptionalNote = fromVanilla(
  (props?: { readonly tone?: "warm" | "cool" }) =>
    shape(props?.tone ?? "neutral", {
      style: { height: 22, fontSize: 11 },
    }),
);
const DefaultNote = fromVanilla(
  (props: { readonly tone?: "warm" | "cool" } = {}) =>
    shape(props.tone ?? "default", {
      style: { height: 22, fontSize: 11 },
    }),
);

// Statically verify that closed fromVanilla components and closed OOXML theme contracts reject invalid usage.
const rejectClosedChildren = () => (
  <>
    <ClosedBadge id="invalid-closed" label="Disallowed">
      {/* @ts-expect-error closed fromVanilla component must reject JSX children */}
      <Text>disallowed</Text>
    </ClosedBadge>
    {/* @ts-expect-error optional-prop fromVanilla component must reject invalid prop type */}
    <OptionalNote tone={123} />
    {/* @ts-expect-error default-prop fromVanilla component must reject invalid prop type */}
    <DefaultNote tone={123} />
  </>
);
void rejectClosedChildren;

function verifyClosedThemeTypes(): ReactElement {
  // @ts-expect-error misspelled native theme color slot must fail type checking
  defineTheme({ colors: { followHlink: "#112233" } });
  // @ts-expect-error slot-to-slot alias in theme colors must fail type checking
  defineTheme({ colors: { dk1: "accent2" } });
  // @ts-expect-error ThemeShade in theme colors must fail type checking
  defineTheme({ colors: { accent1: lighter("accent2", 0.4) } });
  // @ts-expect-error misspelled theme font slot must fail type checking
  defineTheme({ fonts: { minro: "Aptos" } });
  return (
    <>
      <Slide
        id="bad-ovr"
        themeOverride={{
          colors: {
            // @ts-expect-error misspelled ThemeOverride color slot must fail type checking
            followHlink: "#112233",
          },
        }}
      />
      <Slide
        id="bad-font-ovr"
        themeOverride={{
          fonts: {
            // @ts-expect-error misspelled ThemeOverride font slot must fail type checking
            minro: "Aptos",
          },
        }}
      />
      {/* @ts-expect-error CSS color keyword must fail type checking */}
      <Text style={{ color: "red" }}>x</Text>
      {/* @ts-expect-error custom color token name must fail type checking */}
      <Text style={{ color: "tileAccent" }}>x</Text>
      {/* @ts-expect-error misspelled scheme color name must fail type checking */}
      <Text style={{ color: "acent1" }}>x</Text>
      {/* @ts-expect-error custom background color token name must fail type checking */}
      <Shape style={{ background: "tileAccent" }} />
      {/* @ts-expect-error removed container theme prop on Slide must fail type checking */}
      <Slide id="s" theme={{}} />
      {/* @ts-expect-error removed container theme prop on Box must fail type checking */}
      <Box theme={{}} />
      {/* @ts-expect-error removed container theme prop on Group must fail type checking */}
      <Group theme={{}} />
    </>
  );
}
void verifyClosedThemeTypes;

export async function expandConsumerFragment(
  options?: ExpandFragmentOptions,
): Promise<readonly CommonNode[]> {
  return expandFragment(
    <>
      <Text id="fragment-item-1" style={{ fontSize: 12, fontWeight: 350 }}>
        Fragment item 1
      </Text>
      <Text id="fragment-item-2" style={{ fontSize: 12, fontWeight: 450 }}>
        Fragment item 2
      </Text>
    </>,
    options,
  );
}

export function ReactConsumerDeck(): ReactElement {
  const jostFamily = font("Jost*", [
    { data: readFileSync(new URL("./Jost.ttf", import.meta.url)) },
  ]);
  const primarySlot: NativeColorSlot = "accent1";
  const primaryHex: HexColor = "#1d4ed8";
  const schemeText: SchemeColorName = "tx1";
  const schemeBg: ThemeColor = "bg1";
  const shadeLight: ThemeColor = lighter("accent2", 0.85);
  const shadeDark: ThemeColor = darker("accent1", 0.25);
  const majorSlot: ThemeFontSlot = "major";
  const minorFontRef: FontReference = "minor";

  const consumerThemeInput: ThemeInput = {
    name: "React Consumer",
    colors: {
      dk1: "#111827",
      lt1: "#fff",
      dk2: "#1e293b",
      lt2: "#f8fafc",
      accent1: primaryHex,
      accent2: "#abc",
      accent3: "#0f766e",
      accent4: "#b45309",
      accent5: "#7c3aed",
      accent6: "#be123c",
      hlink: "#2563eb",
      folHlink: "#6d28d9",
    },
    fonts: { major: jostFamily, minor: jostFamily },
  };
  const consumerTheme: Theme = defineTheme(consumerThemeInput);
  const slideThemeOverride: ThemeOverride = {
    colors: { accent1: "#0f766e", accent2: "#abc" },
    fonts: { major: jostFamily, minor: jostFamily },
  };
  const slideColorMapOverride: ColorMapOverride = {
    ...defaultColorMap,
    tx1: primarySlot,
  };

  const adoptedClosed = shape("Adopted closed node", {
    id: "adopted-closed",
    style: {
      height: 24,
      fontSize: 12,
      fontFamily: minorFontRef,
      color: "#123c",
    },
  });
  const adoptedSlot = box({
    id: "adopted-slot-inner",
    style: { display: "flex", flexDirection: "column", gap: 4 },
  });
  const adoptedContainer = box({
    id: "adopted-container",
    style: { display: "flex", flexDirection: "column", gap: 4 },
  }).append(adoptedSlot);
  const closedMountProps: MountProps = { node: adoptedClosed };
  const slottedMountProps: MountProps = {
    node: adoptedContainer,
    slot: adoptedSlot,
  };
  const placeholderBody = text("Imperative placeholder body", {
    id: "placeholder-body",
    style: { fontSize: 13, fontFamily: minorFontRef, color: schemeText },
  });

  return (
    <Deck width={640} height={360} theme={consumerTheme}>
      <SlideMaster id="consumer-master" name="Consumer Master">
        <SlideLayout layout={consumerLayout} />
      </SlideMaster>

      <Slide
        id="react-consumer"
        themeOverride={slideThemeOverride}
        colorMapOverride={slideColorMapOverride}
      >
        <Box
          id="react-container"
          style={{
            position: "absolute",
            left: 24,
            top: 24,
            width: 592,
            height: 312,
            display: "flex",
            flexDirection: "column",
            gap: 6,
            padding: sharedPadding,
            background: schemeBg,
          }}
        >
          <ClosedBadge id="bridged-badge" label="Bridged closed badge" />
          <Box style={{ display: "flex", flexDirection: "row", gap: 6 }}>
            <OptionalNote />
            <DefaultNote tone="warm" />
            <Shape
              id="closed-theme-chip"
              style={{
                height: 22,
                fontSize: 11,
                fontFamily: majorSlot,
                color: schemeText,
                background: shadeLight,
                border: { color: "#1d4ed8cc", width: 1 },
              }}
            >
              Closed theme chip
            </Shape>
          </Box>
          <CardShell id="bridged-shell" title="Bridged slotted shell">
            <Table id="react-summary" columns={["1fr", "1fr"]}>
              <Row>
                <Cell style={{ padding: sharedPadding, color: schemeText }}>
                  Metric
                </Cell>
                <Cell
                  style={{
                    padding: sharedPadding,
                    color: schemeBg,
                    background: shadeDark,
                  }}
                >
                  Value
                </Cell>
              </Row>
            </Table>
            <Chart
              id="react-trend"
              groups={[
                {
                  kind: "column",
                  data: {
                    kind: "shared-categories",
                    categories: ["Apr", "May"],
                  },
                  series: [
                    { name: "Revenue", values: [120, 150], color: "accent2" },
                  ],
                },
              ]}
              style={{ width: "100%", height: 80 }}
            />
          </CardShell>
          <Mount {...closedMountProps} />
          <Mount {...slottedMountProps}>
            <Text
              id="slotted-child"
              style={{ fontSize: 11, fontFamily: minorFontRef }}
            >
              Mounted slot child
            </Text>
          </Mount>
        </Box>
      </Slide>

      <Slide
        id="placeholder-adoption"
        layout={consumerLayout}
        placeholders={{
          title: "Placeholder Slide",
          body: placeholderBody,
        }}
      />
    </Deck>
  );
}
