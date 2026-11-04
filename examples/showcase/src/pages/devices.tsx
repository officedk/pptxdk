import { Text } from "@react-pptx/react-pptx";

import { DeviceFrame } from "../components/device-frame.tsx";
import { PageFrame } from "../templates/page-frame.tsx";
import { colors } from "../theme.ts";

function ExampleContent({ width }: { width: number }) {
  return (
    <>
      <Text
        width={width}
        height="auto"
        style={{ fontSize: 28, color: colors.teal }}
      >
        内容先行
      </Text>
      <Text width={width} height="auto" style={{ fontSize: 22 }}>
        先说明要表达的内容，再选择合适的展示范围。
      </Text>
      <Text
        width={width}
        height="auto"
        style={{ fontSize: 18, color: colors.muted }}
      >
        标题、正文与留白，在不同尺寸中保持清晰。
      </Text>
    </>
  );
}

export function DevicesPage() {
  return (
    <PageFrame
      slideId="devices"
      title="给内容一个合适的框"
      subtitle="外框交代展示情境；真正需要读懂的内容仍然清晰可见。"
    >
      <DeviceFrame x={110} y={190} width={310} height={450}>
        <ExampleContent width={250} />
      </DeviceFrame>
      <DeviceFrame x={540} y={235} width={630} height={350}>
        <ExampleContent width={570} />
      </DeviceFrame>
    </PageFrame>
  );
}
