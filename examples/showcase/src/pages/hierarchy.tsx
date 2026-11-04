import { Box, Text } from "@react-pptx/react-pptx";

import { PageFrame } from "../templates/page-frame.tsx";
import { colors } from "../theme.ts";

export function HierarchyPage() {
  return (
    <PageFrame
      slideId="hierarchy"
      title="一页先形成阅读层级"
      subtitle="同一组内容，有了主次，才有明确的阅读起点。"
    >
      <Box
        style={{
          position: "absolute",
          left: 64,
          top: 210,
          width: 1152,
          display: "flex",
          gap: 64,
        }}
      >
        <Box
          style={{
            width: 650,
            display: "flex",
            flexDirection: "column",
            gap: 24,
          }}
        >
          <Text width={650} height="auto" style={{ fontSize: 52 }}>
            先让人知道{`\n`}这一页要说什么
          </Text>
          <Text
            width={610}
            height="auto"
            style={{ fontSize: 25, color: colors.muted }}
          >
            标题给出结论，正文提供依据。字号、颜色与留白一起建立清晰的阅读顺序。
          </Text>
        </Box>
        <Box
          style={{
            width: 438,
            padding: 32,
            backgroundColor: colors.pale,
            display: "flex",
            flexDirection: "column",
            gap: 24,
          }}
        >
          <Text
            width={374}
            height="auto"
            style={{ fontSize: 32, color: colors.teal }}
          >
            01 结论
          </Text>
          <Text width={374} height="auto" style={{ fontSize: 27 }}>
            02 关键依据
          </Text>
          <Text
            width={374}
            height="auto"
            style={{ fontSize: 22, color: colors.muted }}
          >
            03 补充说明
          </Text>
        </Box>
      </Box>
    </PageFrame>
  );
}
