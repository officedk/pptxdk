import { Chart, Shape, Text } from "@officedk/pptxdk/react";

import {
  MetricCard,
  PanelCard,
  Rule,
  SemanticTag,
} from "../components/primitives.tsx";
import { SlideFrame } from "../templates/slide-frame.tsx";
import { colors, fonts } from "../theme.ts";

const epScaleCategories = ["16 GPU", "32 GPU", "64 GPU", "128 GPU"] as const;
const dispatchBandwidth = [42.47, 58.02, 50.58, 45.34] as const;
const combineBandwidth = [43.05, 56.96, 48.54, 41.6] as const;

const dualModes = [
  {
    title: "高吞吐内核 (Normal Kernels · 训练与 Prefill)",
    badge: "NVLink + RDMA 流水转发",
    tone: "compute" as const,
    metrics: "机内 153 GB/s · 跨机 43–58 GB/s (24 SMs)",
    desc: "利用机内 NVLink (160 GB/s) 与机间 RDMA (50 GB/s) 带宽非对称性，先通过 IB 投递至目标节点同平面网卡，再经机内 NVLink 转发至目标 GPU，将单次跨机传输摊薄至最多 4 个节点。",
  },
  {
    title: "纯 RDMA 低延迟内核 (Low-Latency · Decode)",
    badge: "0 SM 占用 · IBGDA 直发",
    tone: "comm" as const,
    metrics: "8 专家 163 us Dispatch · 318 us Combine",
    desc: "绕过机内 NVLink 中转以消除微批次流水气泡，完全基于 InfiniBand GPUDirect Async (IBGDA) 直接收发，并采用挂钩 (Hook) 后台轮询机制实现零 SM 占用下的计算通信重叠。",
  },
] as const;

/** Slide 07: Multi-Plane Fat-Tree (MPFT) network and DeepEP communication library. */
export function MpftDeepEpPage() {
  return (
    <SlideFrame
      id="mpft-deepep"
      kicker="02 · 多平面集群网络与专家通信库 / MPFT · DeepEP · IBGDA"
      title="多平面胖树网络 (MPFT) 与 DeepEP：两级两层交换撑起万卡级专家通信"
      takeaway="相比三层胖树集群造价降低约 40%，跨机 All-to-All 实测达 58 GB/s"
      takeawayTone="comm"
      source="数据来源：ISCA '25 Insights into DeepSeek-V3 Fig. 7 · DeepEP 开源仓库实测 (8x400G RoCE/IB + H800 NVLink)"
    >
      {/* Left column: DeepEP All-to-All Bandwidth Chart across EP scales */}
      <PanelCard
        x={48}
        y={108}
        width={588}
        height={362}
        title="DeepEP 跨节点 All-to-All 有效带宽随 EP 规模扩展实测 (GB/s)"
        subtitle="基于 MPFT 多平面两层胖树网络，4,096 Tokens/GPU 下 Dispatch 与 Combine 单卡带宽"
        rightBadge="ISCA '25 Fig. 7"
        tone="comm"
      >
        <Chart
          id="deepep-bandwidth-chart"
          legend="top"
          axes={{
            category: {
              kind: "category",
              line: { color: colors.steel, width: 1 },
              style: {
                fontFamily: fonts.body,
                fontSize: 10.5,
                color: colors.ink,
              },
            },
            value: {
              kind: "value",
              min: 0,
              max: 70,
              majorUnit: 14,
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
              kind: "column",
              data: {
                kind: "shared-categories",
                categories: epScaleCategories,
              },
              series: [
                {
                  name: "Dispatch 分发带宽 (GB/s)",
                  values: dispatchBandwidth,
                  color: colors.compute,
                },
                {
                  name: "Combine 聚合带宽 (GB/s)",
                  values: combineBandwidth,
                  color: colors.comm,
                },
              ],
              dataLabels: {
                value: true,
                position: "outsideEnd",
                numberFormat: "0.0",
                style: {
                  fontFamily: fonts.body,
                  fontSize: 10,
                  color: colors.ink,
                },
              },
            },
          ]}
          style={{
            position: "absolute",
            left: 64,
            top: 168,
            width: 556,
            height: 252,
            fontFamily: fonts.body,
            fontSize: 10.5,
            color: colors.ink,
          }}
        />
        <Text
          style={{
            position: "absolute",
            left: 68,
            top: 428,
            width: 548,
            height: 32,
            fontSize: 10.5,
            lineHeight: 15,
            color: colors.muted,
          }}
        >
          注：32 GPU (4 节点) 时正好触发节点限制路由 (M=4) 的机内 NVLink
          满负荷转发甜点，单卡吞吐突破 50 GB/s 网卡线速；扩展至 128 GPU 仍保持
          41.6–45.3 GB/s 高水位。
        </Text>
      </PanelCard>

      {/* Left Bottom: 3 KPI metric cards */}
      <MetricCard
        x={48}
        y={484}
        width={188}
        height={88}
        label="机内 NVLink 分发峰值"
        value="158 GB/s"
        caption="逼近 H800 160 GB/s 硬件上限"
        tone="compute"
      />
      <MetricCard
        x={248}
        y={484}
        width={188}
        height={88}
        label="Decode 128 专家延迟"
        value="194 us"
        caption="纯 RDMA Dispatch (128 Tokens/卡)"
        tone="comm"
      />
      <MetricCard
        x={448}
        y={484}
        width={188}
        height={88}
        label="两层多平面网络降本"
        value="-40% Cost"
        caption="3,200 端口等效三层胖树对比"
        tone="copper"
      />

      {/* Left Bottom footer card: Traffic isolation */}
      <Shape
        style={{
          position: "absolute",
          left: 48,
          top: 584,
          width: 588,
          height: 76,
          background: colors.surface,
          border: { color: colors.line, width: 1 },
        }}
      />
      <Shape
        style={{
          position: "absolute",
          left: 48,
          top: 584,
          width: 4,
          height: 76,
          background: colors.comm,
        }}
      />
      <Text
        style={{
          position: "absolute",
          left: 64,
          top: 594,
          width: 556,
          height: 20,
          fontSize: 12.5,
          color: colors.steel,
          textWrap: "nowrap",
        }}
      >
        InfiniBand 虚拟通道 (VL) 硬件级流量隔离：
      </Text>
      <Text
        style={{
          position: "absolute",
          left: 64,
          top: 616,
          width: 556,
          height: 34,
          fontSize: 11,
          lineHeight: 15.5,
          color: colors.muted,
        }}
      >
        将高突发 EP All-to-All、流水线 PP 点对点传输与 DP 梯度归约映射至独立 VL
        并启用自适应路由 (AR)，消除多工作负载下的队头阻塞与链路拥塞扩散。
      </Text>

      {/* Right column: MPFT Topology + DeepEP Dual-Mode Architecture */}
      <PanelCard
        x={652}
        y={108}
        width={580}
        height={552}
        title="多平面两层胖树 (MPFT) 拓扑与 DeepEP 双模内核架构"
        subtitle="8 个独立物理平面消除跨平面哈希冲突，机内 NVSwitch 充当跨平面高速交换桥"
        rightBadge="8x400G 独立平面"
        tone="compute"
      >
        {/* Top half of right panel: MPFT Topology Diagram */}
        <Shape
          style={{
            position: "absolute",
            left: 672,
            top: 170,
            width: 540,
            height: 186,
            background: colors.paper,
            border: { color: colors.line, width: 1 },
          }}
        />

        {/* Layer 1: 8 Independent Two-Layer Fat-Tree Planes */}
        <Text
          style={{
            position: "absolute",
            left: 686,
            top: 178,
            width: 512,
            height: 18,
            fontSize: 11,
            color: colors.commText,
            textWrap: "nowrap",
          }}
        >
          8 个物理隔离的 2 层 InfiniBand Fat-Tree 平面 (Plane 0 – Plane 7 ·
          每平面仅 Spine + Leaf 两级交换)
        </Text>
        {[0, 1, 2, 3, 4, 5, 6, 7].map((plane) => (
          <Shape
            key={plane}
            style={{
              position: "absolute",
              left: 686 + plane * 64,
              top: 200,
              width: 58,
              height: 34,
              background: plane % 2 === 0 ? colors.comm : colors.commSoft,
              border: { color: colors.comm, width: 1 },
              fontFamily: fonts.metric,
              fontSize: 10,
              fontWeight: 700,
              color: plane % 2 === 0 ? colors.surface : colors.commText,
              align: "center",
              verticalAlign: "middle",
            }}
          >{`Plane ${plane}`}</Shape>
        ))}

        {/* Layer 2: 8x 400G IB NICs paired 1:1 with 8x H800 GPUs */}
        {[0, 1, 2, 3, 4, 5, 6, 7].map((gpu) => (
          <Shape
            key={`gpu-${gpu}`}
            style={{
              position: "absolute",
              left: 686 + gpu * 64,
              top: 246,
              width: 58,
              height: 44,
              background: colors.surface,
              border: { color: colors.compute, width: 1 },
              fontFamily: fonts.metric,
              fontSize: 9.5,
              fontWeight: 700,
              color: colors.compute,
              align: "center",
              verticalAlign: "middle",
            }}
          >{`GPU ${gpu}\nNIC ${gpu}`}</Shape>
        ))}

        {/* Layer 3: Intra-node NVLink / NVSwitch fabric */}
        <Shape
          style={{
            position: "absolute",
            left: 686,
            top: 302,
            width: 506,
            height: 42,
            background: colors.compute,
            fontSize: 11.5,
            color: colors.surface,
            align: "center",
            verticalAlign: "middle",
          }}
        >
          机内全互连 NVLink / NVSwitch 交换矩阵 (400 GB/s 双向带宽 ·
          承担跨网卡平面转发)
        </Shape>

        {/* Divider */}
        <Rule x={672} y={368} width={540} />

        {/* Bottom half of right panel: DeepEP Dual-Mode Kernels */}
        {dualModes.map((mode, idx) => {
          const boxY = 380 + idx * 136;
          return (
            <Shape
              key={mode.title}
              style={{
                position: "absolute",
                left: 672,
                top: boxY,
                width: 540,
                height: 124,
                background: colors.paper,
                border: { color: colors.line, width: 1 },
              }}
            />
          );
        })}
        {dualModes.map((mode, idx) => {
          const boxY = 380 + idx * 136;
          const accent = mode.tone === "compute" ? colors.compute : colors.comm;
          return (
            <Shape
              key={`bar-${idx}`}
              style={{
                position: "absolute",
                left: 672,
                top: boxY,
                width: 4,
                height: 124,
                background: accent,
              }}
            />
          );
        })}
        {dualModes.map((mode, idx) => {
          const boxY = 380 + idx * 136;
          return (
            <Text
              key={`title-${idx}`}
              style={{
                position: "absolute",
                left: 688,
                top: boxY + 10,
                width: 340,
                height: 22,
                fontSize: 13,
                color: colors.ink,
                textWrap: "nowrap",
              }}
            >
              {mode.title}
            </Text>
          );
        })}
        {dualModes.map((mode, idx) => {
          const boxY = 380 + idx * 136;
          return (
            <SemanticTag
              key={`badge-${idx}`}
              x={1034}
              y={boxY + 10}
              width={166}
              height={22}
              text={mode.badge}
              tone={mode.tone}
            />
          );
        })}
        {dualModes.map((mode, idx) => {
          const boxY = 380 + idx * 136;
          return (
            <Text
              key={`metrics-${idx}`}
              style={{
                position: "absolute",
                left: 688,
                top: boxY + 36,
                width: 508,
                height: 20,
                fontSize: 11.5,
                color:
                  mode.tone === "compute" ? colors.compute : colors.commText,
                textWrap: "nowrap",
              }}
            >
              实测性能：{mode.metrics}
            </Text>
          );
        })}
        {dualModes.map((mode, idx) => {
          const boxY = 380 + idx * 136;
          return (
            <Text
              key={`desc-${idx}`}
              style={{
                position: "absolute",
                left: 688,
                top: boxY + 60,
                width: 510,
                height: 54,
                fontSize: 11,
                lineHeight: 16,
                color: colors.muted,
              }}
            >
              {mode.desc}
            </Text>
          );
        })}
      </PanelCard>
    </SlideFrame>
  );
}
