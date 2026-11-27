import { readFileSync } from "node:fs";

import {
  Box,
  Cell,
  Chart,
  Deck,
  Mount,
  Row,
  Slide,
  SlideLayout,
  SlideMaster,
  Table,
  Text,
  defineSlideLayout,
  defineTheme,
  expandFragment,
  font,
  fromVanilla,
  type CommonContainer,
  type CommonNode,
  type ExpandFragmentOptions,
  type MountProps,
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

// Statically verify that closed fromVanilla components reject JSX children and invalid props.
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
  const consumerTheme = defineTheme({
    name: "React Consumer",
    fonts: { major: jostFamily, minor: jostFamily },
  });

  const adoptedClosed = shape("Adopted closed node", {
    id: "adopted-closed",
    style: { height: 24, fontSize: 12 },
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
    style: { fontSize: 13 },
  });

  return (
    <Deck width={640} height={360} theme={consumerTheme}>
      <SlideMaster id="consumer-master" name="Consumer Master">
        <SlideLayout layout={consumerLayout} />
      </SlideMaster>

      <Slide id="react-consumer">
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
          }}
        >
          <ClosedBadge id="bridged-badge" label="Bridged closed badge" />
          <Box style={{ display: "flex", flexDirection: "row", gap: 6 }}>
            <OptionalNote />
            <DefaultNote tone="warm" />
          </Box>
          <CardShell id="bridged-shell" title="Bridged slotted shell">
            <Table id="react-summary" columns={["1fr", "1fr"]}>
              <Row>
                <Cell style={{ padding: sharedPadding }}>Metric</Cell>
                <Cell style={{ padding: sharedPadding }}>Value</Cell>
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
                  series: [{ name: "Revenue", values: [120, 150] }],
                },
              ]}
              style={{ width: "100%", height: 80 }}
            />
          </CardShell>
          <Mount {...closedMountProps} />
          <Mount {...slottedMountProps}>
            <Text id="slotted-child" style={{ fontSize: 11 }}>
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
