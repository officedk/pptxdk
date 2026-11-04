import { Text } from "@react-pptx/react-pptx";

import { ProcessPath } from "../components/process-path.tsx";
import { PageFrame } from "../templates/page-frame.tsx";
import { colors } from "../theme.ts";

export function ProcessPage() {
  return (
    <PageFrame
      slideId="process"
      title="把过程组织成路径"
      subtitle="节点说明发生了什么，连线说明下一步去哪里。"
    >
      <ProcessPath
        x={64}
        y={240}
        width={1152}
        nodes={[
          { title: "提出问题", body: "明确这段表达要回答的核心问题" },
          { title: "组织材料", body: "选择能支撑观点的内容与顺序" },
          { title: "形成表达", body: "用结构和节奏让关系逐步清晰" },
        ]}
      />
      <Text
        x={64}
        y={550}
        width={1100}
        height={60}
        style={{ fontSize: 23, color: colors.muted }}
      >
        时间线强调先后；流程图还需要说明依赖与方向。
      </Text>
    </PageFrame>
  );
}
