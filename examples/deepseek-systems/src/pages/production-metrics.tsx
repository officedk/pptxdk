import { Chart, Text } from "@officedk/pptxdk/react";

import { MetricCard, PanelCard, Rule } from "../components/primitives.tsx";
import { SlideFrame } from "../templates/slide-frame.tsx";
import { colors, fonts } from "../theme.ts";

const timeSlots = [
  "12:00",
  "15:00",
  "18:00",
  "21:00",
  "00:00",
  "03:00",
  "06:00",
  "09:00",
  "12:00+",
] as const;

const nodeOccupancyValues = [
  264, 278, 272, 268, 215, 92, 68, 196, 268,
] as const;

const costRevenueCategories = [
  "H800 租赁成本",
  "命中输入折算",
  "未命中输入折算",
  "输出 Token 折算",
  "全量理论总营收",
] as const;

const costRevenueValues = [87.1, 47.9, 146.3, 367.9, 562.0] as const;

/** Slide 11: 24-hour production inference telemetry dashboard (Feb 27-28, 2025). */
export function ProductionMetricsPage() {
  return (
    <SlideFrame
      id="production-metrics"
      kicker="03 · 线上生产环境 24 小时运行实测 / 2025-02-27 12:00 至 02-28 12:00 (UTC+8)"
      title="生产环境 24 小时全量运行实测：608B 输入与 168B 输出背后的能效比"
      takeaway="日均占用 226.75 节点 ($87.1K 成本)，按 R1 定价折算理论利润率达 545%"
      takeawayTone="copper"
      source="数据来源：DeepSeek-V3/R1 Inference System Overview 线上真实统计 (H800 按 $2/GPU/h 计价，R1 API 定价口径)"
    >
      {/* Top KPI Ribbon: 4 core 24h production metrics */}
      <MetricCard
        x={48}
        y={108}
        width={284}
        height={88}
        label="24 小时平均 / 峰值 H800 节点数"
        value="226.75 / 278"
        caption="8 卡/节点 · 夜间闲时释放至科研训练集群"
        tone="compute"
      />
      <MetricCard
        x={348}
        y={108}
        width={284}
        height={88}
        label="24 小时总输入 Token 吞吐"
        value="608B Tokens"
        caption="其中 342B 命中 3FS 远端 KVCache (56.3%)"
        tone="comm"
      />
      <MetricCard
        x={648}
        y={108}
        width={284}
        height={88}
        label="24 小时总输出 Token 吞吐"
        value="168B Tokens"
        caption="平均输出速率 20–22 Tokens/s · KV 长度 4,989"
        tone="compute"
      />
      <MetricCard
        x={948}
        y={108}
        width={284}
        height={88}
        label="24 小时算力成本 vs 理论利润率"
        value="$87.1K / 545%"
        caption="理论总营收 $562,027 (全部按 R1 费率折算)"
        tone="copper"
      />

      {/* Bottom Left: 24-hour Node Occupancy Area Chart */}
      <PanelCard
        x={48}
        y={210}
        width={420}
        height={450}
        title="24 小时推理节点占用曲线"
        subtitle="白天高峰扩容至 278 节点，凌晨低谷缩容至 68 节点动态削峰填谷"
        rightBadge="昼夜弹性调度"
        tone="compute"
      >
        <Chart
          id="node-occupancy-chart"
          legend="top"
          axes={{
            category: {
              kind: "category",
              line: { color: colors.steel, width: 1 },
              style: {
                fontFamily: fonts.body,
                fontSize: 10,
                color: colors.ink,
              },
            },
            value: {
              kind: "value",
              min: 0,
              max: 320,
              majorUnit: 80,
              numberFormat: "0",
              gridlines: { color: colors.line, width: 1 },
              line: false,
              style: {
                fontFamily: fonts.body,
                fontSize: 10,
                color: colors.muted,
              },
            },
          }}
          groups={[
            {
              kind: "area",
              data: { kind: "shared-categories", categories: timeSlots },
              series: [
                {
                  name: "在线推理 H800 节点数 (8xGPU/节点)",
                  values: nodeOccupancyValues,
                  color: "#9EBDF5",
                },
              ],
              dataLabels: {
                value: true,
                numberFormat: "0",
                style: {
                  fontFamily: fonts.body,
                  fontSize: 9.5,
                  color: colors.steel,
                },
              },
            },
          ]}
          style={{
            position: "absolute",
            left: 62,
            top: 272,
            width: 392,
            height: 292,
            fontFamily: fonts.body,
            fontSize: 10,
            color: colors.ink,
          }}
        />
        <Rule x={64} y={576} width={388} />
        <Text
          style={{
            position: "absolute",
            left: 64,
            top: 586,
            width: 388,
            padding: 0,
            fontSize: 10.5,
            lineHeight: 15.5,
            color: colors.muted,
          }}
        >
          动态潮汐调度：因夜间请求量显著回落，集群调度系统自动将闲置 H800
          节点切回训练与实验任务，避免固定容量预留造成的算力空转。
        </Text>
      </PanelCard>

      {/* Bottom Middle: KVCache Hit Rate Doughnut Chart */}
      <PanelCard
        x={484}
        y={210}
        width={340}
        height={450}
        title="3FS 缓存命中构成"
        subtitle="基于 3FS 全闪存储实现跨节点历史 KVCache 共享"
        rightBadge="56.3% 命中率"
        tone="comm"
      >
        <Chart
          id="kv-cache-hit-chart"
          legend="bottom"
          groups={[
            {
              kind: "doughnut",
              data: {
                kind: "shared-categories",
                categories: ["3FS 缓存命中 (342B)", "未命中重算 (266B)"],
              },
              series: [
                {
                  name: "输入 Token 占比 (%)",
                  values: [56.3, 43.7],
                  points: {
                    0: { color: colors.comm, labelColor: "#FFFFFF" },
                    1: { color: colors.baseline, labelColor: colors.ink },
                  },
                },
              ],
              holeSize: 58,
              dataLabels: {
                value: true,
                numberFormat: '0.0"%"',
                style: {
                  fontFamily: fonts.body,
                  fontSize: 10.5,
                  color: colors.ink,
                },
              },
            },
          ]}
          style={{
            position: "absolute",
            left: 500,
            top: 272,
            width: 308,
            height: 286,
            fontFamily: fonts.body,
            fontSize: 10.5,
            color: colors.ink,
          }}
        />
        <Rule x={500} y={576} width={308} />
        <Text
          style={{
            position: "absolute",
            left: 500,
            top: 586,
            width: 308,
            padding: 0,
            fontSize: 10.5,
            lineHeight: 15.5,
            color: colors.muted,
          }}
        >
          MLA 将每 Token 缓存压缩至紧凑潜向量后，342B 输入直接从 3FS 读取历史
          KV，节省超一半 Prefill 算力。
        </Text>
      </PanelCard>

      {/* Bottom Right: Cost vs Theoretical Revenue Breakdown Column Chart */}
      <PanelCard
        x={840}
        y={210}
        width={392}
        height={450}
        title="单日成本与理论营收 ($K)"
        subtitle="全量 Token 按 DeepSeek-R1 标准 API 价格计费对比"
        rightBadge="单位经济模型"
        tone="copper"
      >
        <Chart
          id="cost-revenue-chart"
          legend={false}
          axes={{
            category: {
              kind: "category",
              line: { color: colors.steel, width: 1 },
              style: {
                fontFamily: fonts.body,
                fontSize: 10,
                color: colors.ink,
              },
            },
            value: {
              kind: "value",
              min: 0,
              max: 640,
              majorUnit: 160,
              numberFormat: '"$"0',
              gridlines: { color: colors.line, width: 1 },
              line: false,
              style: {
                fontFamily: fonts.body,
                fontSize: 10,
                color: colors.muted,
              },
            },
          }}
          groups={[
            {
              kind: "column",
              data: {
                kind: "shared-categories",
                categories: costRevenueCategories,
              },
              series: [
                {
                  name: "单日金额 ($K USD)",
                  values: costRevenueValues,
                  color: colors.compute,
                  points: {
                    0: { color: colors.steel, labelColor: colors.steel },
                    1: { color: colors.comm, labelColor: colors.commText },
                    2: { color: colors.compute, labelColor: colors.compute },
                    3: { color: colors.copper, labelColor: colors.copperText },
                    4: {
                      color: colors.copperText,
                      labelColor: colors.copperText,
                    },
                  },
                },
              ],
              dataLabels: {
                value: true,
                position: "outsideEnd",
                numberFormat: '"$"0.0"K"',
                style: {
                  fontFamily: fonts.body,
                  fontSize: 9.5,
                  color: colors.ink,
                },
              },
            },
          ]}
          style={{
            position: "absolute",
            left: 854,
            top: 272,
            width: 364,
            height: 292,
            fontFamily: fonts.body,
            fontSize: 10,
            color: colors.ink,
          }}
        />
        <Rule x={856} y={576} width={360} />
        <Text
          style={{
            position: "absolute",
            left: 856,
            top: 586,
            width: 360,
            padding: 0,
            fontSize: 10.5,
            lineHeight: 15.5,
            color: colors.muted,
          }}
        >
          费率口径：命中输入 $0.14/M、未命中输入 $0.55/M、输出
          $2.19/M；实际综合收入受 V3
          低定价、网页端免费与夜间立减折扣影响低于理论值。
        </Text>
      </PanelCard>
    </SlideFrame>
  );
}
