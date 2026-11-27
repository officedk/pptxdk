import { Chart, Text } from "@officedk/pptxdk/react";
import { use } from "react";

import type { fetchGitHubProjects } from "../github-projects.ts";
import { SlideFrame } from "../templates/slide-frame.tsx";
import { colors, fonts } from "../theme.ts";

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
      <Text
        style={{
          position: "absolute",
          left: 48,
          top: 118,
          width: 1184,
          height: 36,
          fontSize: 18,
        }}
      >
        GitHub 星标数
      </Text>
      <Chart
        id="github-stars"
        legend={false}
        axes={{
          category: { kind: "category", reversed: true, position: "left" },
          value: { kind: "value", min: 0, position: "bottom" },
        }}
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
        style={{
          position: "absolute",
          left: 48,
          top: 166,
          width: 1184,
          height: 410,
          fontFamily: fonts.body,
          fontSize: 14,
          color: colors.ink,
        }}
      />
      {projects.map(({ name, url }, index) => (
        <Text
          key={name}
          link={{ url }}
          style={{
            position: "absolute",
            left: 48 + index * 400,
            top: 610,
            width: 384,
            height: 36,
            fontSize: 16,
            color: colors.compute,
            textDecoration: "underline",
          }}
        >
          {name}
        </Text>
      ))}
    </SlideFrame>
  );
}
