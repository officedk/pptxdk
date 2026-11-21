import { Text } from "pptxdk/react";
import { Fragment } from "react";

import { Rule } from "../components/primitives.tsx";
import { sources } from "../data/sources.ts";
import { SlideFrame } from "../templates/slide-frame.tsx";
import { colors } from "../theme.ts";

export function SourcesPage({ page }: { readonly page: 1 | 2 }) {
  const items = sources.slice(page === 1 ? 0 : 9, page === 1 ? 9 : 18);
  return (
    <SlideFrame
      id={`sources-${page}`}
      chapter="附录｜原始论文与官方发布"
      title={
        page === 1 ? "史料与延伸阅读｜基础与架构" : "史料与延伸阅读｜系统与前沿"
      }
      number={page === 1 ? 13 : 14}
      source="标题与网址均可点击；所有条目均为原始论文或官方页面"
    >
      {items.map((source, index) => {
        const y = 152 + index * 56;
        return (
          <Fragment key={index}>
            <Text
              x={64}
              y={y + 2}
              width={42}
              height={32}
              style={{
                fontFamily: "display",
                fontSize: 23,
                color: colors.cobalt,
              }}
            >
              {String(source.id).padStart(2, "0")}
            </Text>
            <Text
              x={116}
              y={y}
              width={748}
              height={34}
              link={{ url: source.url }}
              style={{ fontSize: 13, lineHeight: 16, color: colors.ink }}
            >
              {source.title}
            </Text>
            <Text
              x={116}
              y={y + 34}
              width={748}
              height={18}
              style={{ fontSize: 10, color: colors.muted, textWrap: "nowrap" }}
            >
              {source.detail}
            </Text>
            <Text
              x={888}
              y={y + 4}
              width={328}
              height={40}
              link={{ url: source.url }}
              style={{
                fontSize: 10,
                lineHeight: 15,
                color: colors.cobalt,
                textDecoration: "underline",
              }}
            >
              {source.url.replace("https://", "")}
            </Text>
            {index < items.length - 1 && (
              <Rule x={64} y={y + 53} width={1152} color={colors.mist} />
            )}
          </Fragment>
        );
      })}
    </SlideFrame>
  );
}
