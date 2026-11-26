import { Cell, Chart, P, Row, Table, Text, wipe } from "@officedk/pptxdk/react";

import { MetricCard, PanelCard } from "../components/primitives.tsx";
import { SlideFrame } from "../templates/slide-frame.tsx";
import { colors } from "../theme.ts";

const costRows = [
  {
    phase: "预训练 (Pre-Training)",
    gpuHours: "2,664K 小时",
    share: "95.55%",
    cost: "$5.328M",
    detail: "14.8T tokens FP8 混合精度训练（约 180K GPU 小时/万亿 token）",
    highlight: false,
  },
  {
    phase: "上下文扩展 (Context Ext.)",
    gpuHours: "119K 小时",
    share: "4.27%",
    cost: "$0.238M",
    detail: "YaRN 两阶段扩展上下文窗口（4K → 32K → 128K）",
    highlight: false,
  },
  {
    phase: "后训练 (Post-Training)",
    gpuHours: "5K 小时",
    share: "0.18%",
    cost: "$0.010M",
    detail: "监督微调（SFT）与强化学习（RL，含 R1 推理能力蒸馏）",
    highlight: false,
  },
  {
    phase: "完整训练总计 (Total)",
    gpuHours: "2,788K 小时",
    share: "100.00%",
    cost: "$5.576M",
    detail: "671B 总参数 / 37B 激活参数正式训练运行（0 次回滚）",
    highlight: true,
  },
] as const;

export function TrainingCostPage() {
  return (
    <SlideFrame
      id="training-cost"
      kicker="01 · 模型架构基座｜训练算力投入与成本拆解"
      title="278.8 万 H800 GPU 小时完成 671B 完整训练，预训练占总算力 95.55%"
      takeaway="核心记录：14.8T 语料正式训练耗资 $5.576M，全程 0 次不可恢复 Loss 尖峰"
      takeawayTone="copper"
      source="数据源：DeepSeek-V3 Technical Report (arXiv:2412.19437) Table 1，按 $2 / H800 GPU 小时租赁价格核算"
    >
      {/* Left Panel: Column Chart + Stacked Percent Bar Chart */}
      {/* These three charts use one full chart wipe because WPS replayed category cues as repeated full chart wipes. */}
      <PanelCard
        x={48}
        y={108}
        width={556}
        height={554}
        title="分阶段 H800 GPU 小时投入与成本占比"
        subtitle="预训练阶段消耗 2,664K 小时（$5.328M），后续阶段仅需 124K 小时（$0.248M）"
        rightBadge="单位：千 GPU 小时"
        tone="compute"
      >
        <Chart
          id="training-hours-chart"
          enter={wipe({ duration: 460 })}
          legend={false}
          plotArea={{ left: 58, top: 24, width: 442, height: 214 }}
          axes={{
            category: {
              kind: "category",
              line: { color: colors.steel, width: 1 },
              style: { fontFamily: "body", fontSize: 12, color: colors.ink },
            },
            value: {
              kind: "value",
              min: 0,
              max: 3000,
              majorUnit: 1000,
              numberFormat: '#,##0"K"',
              gridlines: { color: colors.line, width: 1 },
              line: false,
              style: { fontFamily: "body", fontSize: 11, color: colors.muted },
            },
          }}
          groups={[
            {
              kind: "column",
              data: {
                kind: "shared-categories",
                categories: [
                  "预训练 (14.8T)",
                  "上下文扩展 (128K)",
                  "后训练 (SFT+RL)",
                ],
              },
              series: [
                {
                  name: "H800 算力耗时（千小时）",
                  values: [2664, 119, 5],
                  color: colors.compute,
                  points: {
                    0: { color: colors.compute, labelColor: colors.compute },
                    1: { color: colors.comm, labelColor: colors.commText },
                    2: { color: colors.copper, labelColor: colors.copperText },
                  },
                },
              ],
              gapWidth: 95,
              dataLabels: {
                value: true,
                position: "outsideEnd",
                numberFormat: '#,##0"K"',
                style: { fontFamily: "body", fontSize: 12, color: colors.ink },
              },
            },
          ]}
          style={{
            position: "absolute",
            left: 66,
            top: 168,
            width: 520,
            height: 292,
            fontFamily: "body",
            fontSize: 12,
            color: colors.ink,
          }}
        />

        <Text
          style={{
            position: "absolute",
            left: 68,
            top: 472,
            width: 516,
            height: 20,
            fontSize: 12.5,
            color: colors.steel,
            textWrap: "nowrap",
          }}
        >
          完整训练总算力构成比例（100% 堆叠条形图 · 2,788K 小时 = $5.576M）
        </Text>

        <Chart
          id="training-share-chart"
          legend="bottom"
          plotArea={{ left: 68, top: 12, width: 432, height: 68 }}
          axes={{
            category: {
              kind: "category",
              line: { color: colors.steel, width: 1 },
              style: { fontFamily: "body", fontSize: 11.5, color: colors.ink },
            },
            value: {
              kind: "value",
              visible: true,
              min: 0,
              max: 100,
              majorUnit: 20,
              numberFormat: '0"%"',
              gridlines: { color: colors.line, width: 1 },
              line: false,
              style: {
                fontFamily: "body",
                fontSize: 10.5,
                color: colors.muted,
              },
            },
          }}
          groups={[
            {
              kind: "bar",
              data: { kind: "shared-categories", categories: ["成本结构"] },
              series: [
                {
                  name: "预训练 95.55%",
                  values: [95.55],
                  color: colors.compute,
                },
                {
                  name: "上下文扩展 4.27%",
                  values: [4.27],
                  color: colors.comm,
                },
                {
                  name: "后训练 0.18%",
                  values: [0.18],
                  color: colors.copper,
                },
              ],
              stacking: "stacked",
              gapWidth: 55,
            },
          ]}
          style={{
            position: "absolute",
            left: 66,
            top: 496,
            width: 520,
            height: 150,
            fontFamily: "body",
            fontSize: 11,
            color: colors.ink,
          }}
        />
      </PanelCard>

      {/* Right Panel: Detailed Table + 3 Quantitative Callouts */}
      <PanelCard
        x={620}
        y={108}
        width={612}
        height={426}
        title="DeepSeek-V3 训练阶段算力与成本明细表"
        subtitle="基于 2,048 张 H800 SXM5 GPU 集群实测统计，不含前期架构消融与小规模实验开销"
        rightBadge="$2 / GPU 小时"
        tone="copper"
      >
        <Table
          id="training-cost-table"
          columns={[132, 92, 74, 78, 200]}
          style={{
            position: "absolute",
            left: 638,
            top: 174,
            width: 576,
            fontFamily: "body",
            fontSize: 11.5,
            color: colors.ink,
          }}
        >
          <Row
            style={{
              height: 38,
              background: colors.steel,
              color: colors.surface,
              verticalAlign: "middle",
            }}
          >
            <Cell style={{ padding: [6, 8] }}>
              <P>训练阶段</P>
            </Cell>
            <Cell style={{ padding: [6, 8] }}>
              <P style={{ align: "right" }}>GPU 小时</P>
            </Cell>
            <Cell style={{ padding: [6, 8] }}>
              <P style={{ align: "right" }}>占比</P>
            </Cell>
            <Cell style={{ padding: [6, 8] }}>
              <P style={{ align: "right" }}>估算成本</P>
            </Cell>
            <Cell style={{ padding: [6, 8] }}>
              <P>核心工作内容</P>
            </Cell>
          </Row>
          {costRows.map((row, index) => {
            const bg = row.highlight
              ? colors.copperSoft
              : index % 2 === 1
                ? colors.paper
                : colors.surface;
            const textColor = row.highlight ? colors.copperText : colors.ink;
            return (
              <Row
                key={row.phase}
                style={{
                  height: 70,
                  background: bg,
                  color: textColor,
                  verticalAlign: "middle",
                }}
              >
                <Cell
                  style={{
                    padding: [6, 8],
                    border: { color: colors.line, width: 1 },
                  }}
                >
                  <P>{row.phase}</P>
                </Cell>
                <Cell
                  style={{
                    padding: [6, 8],
                    border: { color: colors.line, width: 1 },
                  }}
                >
                  <P style={{ align: "right" }}>{row.gpuHours}</P>
                </Cell>
                <Cell
                  style={{
                    padding: [6, 8],
                    border: { color: colors.line, width: 1 },
                  }}
                >
                  <P style={{ align: "right" }}>{row.share}</P>
                </Cell>
                <Cell
                  style={{
                    padding: [6, 8],
                    border: { color: colors.line, width: 1 },
                  }}
                >
                  <P style={{ align: "right" }}>{row.cost}</P>
                </Cell>
                <Cell
                  style={{
                    padding: [6, 8],
                    border: { color: colors.line, width: 1 },
                  }}
                >
                  <P>{row.detail}</P>
                </Cell>
              </Row>
            );
          })}
        </Table>
      </PanelCard>

      {/* Bottom Right 3 KPI Cards */}
      <MetricCard
        x={620}
        y={548}
        width={194}
        height={114}
        label="每万亿 Token 训练开销"
        value="180K GPUh"
        caption="2,048 卡约 3.7 天完成 1T 训练"
        tone="compute"
      />
      <MetricCard
        x={829}
        y={548}
        width={194}
        height={114}
        label="后训练与对齐算力占比"
        value="0.18%"
        caption="仅 5K 小时 ($10K) 完成对齐与蒸馏"
        tone="comm"
      />
      <MetricCard
        x={1038}
        y={548}
        width={194}
        height={114}
        label="全周期不可恢复尖峰"
        value="0 Rollbacks"
        caption="细粒度 FP8 保障 14.8T 稳定收敛"
        tone="copper"
      />
    </SlideFrame>
  );
}
