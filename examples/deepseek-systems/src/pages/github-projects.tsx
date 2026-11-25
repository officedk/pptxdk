import { Chart, Text } from "@officedk/pptxdk/react";
import { use } from "react";

import type { fetchGitHubProjects } from "../github-projects.ts";
import { SlideFrame } from "../templates/slide-frame.tsx";
import { colors } from "../theme.ts";

export function GitHubProjectsPage({
  snapshot,
}: {
  readonly snapshot: ReturnType<typeof fetchGitHubProjects>;
}) {
  const { projects, fetchedAt } = use(snapshot);

  return (
    <SlideFrame
      id="github-projects"
      kicker="04 · DeepSeek Open-Source Week / GitHub"
      title="开源项目近况"
      source={`数据来源：GitHub · 获取时间：${fetchedAt.replace("T", " ").slice(0, 19)} UTC`}
    >
      <Text x={48} y={118} width={1184} height={36} style={{ fontSize: 18 }}>
        GitHub 星标数
      </Text>
      <Chart
        id="github-stars"
        x={48}
        y={166}
        width={1184}
        height={410}
        legend={false}
        axes={{
          category: { kind: "category", reversed: true, position: "left" },
          value: { kind: "value", min: 0, position: "bottom" },
        }}
        style={{ fontFamily: "body", fontSize: 14, color: colors.ink }}
        groups={[
          {
            kind: "bar",
            data: {
              kind: "shared-categories",
              categories: projects.map(({ name }) => name),
            },
            series: [
              {
                name: "GitHub 星标数",
                values: projects.map(({ stars }) => stars),
                color: colors.compute,
              },
            ],
            dataLabels: { value: true, position: "outsideEnd" },
          },
        ]}
      />
      {projects.map(({ name, url }, index) => (
        <Text
          key={name}
          x={48 + index * 400}
          y={610}
          width={384}
          height={36}
          style={{
            fontSize: 16,
            color: colors.compute,
            textDecoration: "underline",
          }}
          link={{ url }}
        >
          {name}
        </Text>
      ))}
    </SlideFrame>
  );
}
