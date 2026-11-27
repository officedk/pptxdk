import {
  Box,
  Cell,
  Chart,
  Deck,
  Row,
  Slide,
  Table,
} from "@officedk/pptxdk/react";
import type { ReactElement } from "react";

const sharedPadding = [8, 12] as const;

export function ReactConsumerDeck(): ReactElement {
  return (
    <Deck width={640} height={360}>
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
            gap: 8,
            padding: sharedPadding,
          }}
        >
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
                data: { kind: "shared-categories", categories: ["Apr", "May"] },
                series: [{ name: "Revenue", values: [120, 150] }],
              },
            ]}
            style={{ flex: 1, width: "100%", height: "100%" }}
          />
        </Box>
      </Slide>
    </Deck>
  );
}
