import {
  floatUp,
  MotionScope,
  Shape,
  Slide,
  Text,
  transitions,
  useMotionScope,
} from "@officedk/pptxdk/react";

import { MetricCard, Rule } from "../components/primitives.tsx";
import { colors } from "../theme.ts";

const pillars = [
  {
    code: "PILLAR 01 · 架构层",
    title: "高稀疏度激活与低秩键值压缩",
    detail:
      "DeepSeekMoE 每层 256 路由专家仅激活 8 个（37B/671B），配合 MLA 低秩联合压缩将每 Token KV 缓存压至 70.3 KB。",
    tone: colors.compute,
  },
  {
    code: "PILLAR 02 · 训练系统层",
    title: "细粒度 FP8 量化与双向流水线重叠",
    detail:
      "1×128 Tile 激活量化与 128×128 Block 权重量化结合 DualPipe 与 DeepEP，跨节点 All-to-All 通信完全隐藏于计算背后。",
    tone: colors.comm,
  },
  {
    code: "PILLAR 03 · 推理与后训练层",
    title: "PD 分离异构并行与 GRPO 推理涌现",
    detail:
      "Prefill EP32 与 Decode EP144 异构部署配合 3FS 磁盘缓存卸载；基于 GRPO 强化学习激发 R1 复杂推理能力并向稠密小模型蒸馏。",
    tone: colors.copper,
  },
] as const;

export function OpeningPage() {
  const kpiScope = useMotionScope();

  return (
    <Slide
      id="opening"
      transition={transitions.fade({ duration: 500 })}
      style={{ background: colors.obsidian }}
    >
      {/* Top telemetry bar */}
      <Text
        style={{
          position: "absolute",
          left: 56,
          top: 28,
          width: 620,
          height: 22,
          fontFamily: "metric",
          fontSize: 12,
          fontWeight: 700,
          color: colors.computeBright,
          letterSpacing: 1,
          textWrap: "nowrap",
        }}
      >
        DEEPSEEK-V3 / R1 ARCHITECTURE &amp; SYSTEMS ENGINEERING
      </Text>
      <Text
        style={{
          position: "absolute",
          left: 680,
          top: 28,
          width: 544,
          height: 22,
          fontFamily: "metric",
          fontSize: 12,
          color: colors.darkMuted,
          align: "right",
          textWrap: "nowrap",
        }}
      >
        ARXIV:2412.19437 · 2501.12948 · ISCA &apos;25 · SC &apos;24
      </Text>
      <Rule x={56} y={56} width={1168} height={1} color={colors.darkLine} />

      {/* Left hero column */}
      <Shape
        style={{
          position: "absolute",
          left: 56,
          top: 84,
          width: 236,
          height: 28,
          background: colors.obsidianCard,
          border: { color: colors.compute, width: 1 },
          fontSize: 12,
          color: colors.computeBright,
          align: "center",
          verticalAlign: "middle",
        }}
      >
        AI-HPC 软硬协同系统工程剖析
      </Shape>
      <Shape
        style={{
          position: "absolute",
          left: 56,
          top: 84,
          width: 4,
          height: 28,
          background: colors.compute,
        }}
      />

      <Text
        style={{
          position: "absolute",
          left: 56,
          top: 126,
          width: 736,
          height: 98,
          fontSize: 35,
          lineHeight: 46,
          color: colors.surface,
        }}
      >
        稀疏激活与通信隐藏：DeepSeek-V3 / R1 训练与推理全栈架构剖析
      </Text>

      <Text
        style={{
          position: "absolute",
          left: 56,
          top: 234,
          width: 728,
          height: 64,
          fontSize: 14,
          lineHeight: 22,
          color: colors.darkMuted,
        }}
      >
        从 MLA 键值缓存压缩、FP8 细粒度混合精度与 DualPipe 双向流水线，到跨节点
        EP 预填充-解码分离与 3FS 并行文件系统——复盘 2,048 张 H800 如何以 557.6
        万美元完成 671B 基座训练并支撑大规模低延迟在线推理。
      </Text>

      {/* Three engineering pillars */}
      {pillars.map((pillar, index) => {
        const y = 318 + index * 110;
        return (
          <Shape
            key={pillar.code}
            id={`opening-pillar-${index + 1}`}
            style={{
              position: "absolute",
              left: 56,
              top: y,
              width: 728,
              height: 96,
              background: colors.obsidianCard,
              border: { color: colors.darkLine, width: 1 },
            }}
          />
        );
      })}
      {pillars.map((pillar, index) => {
        const y = 318 + index * 110;
        return (
          <Shape
            key={`${pillar.code}-bar`}
            style={{
              position: "absolute",
              left: 56,
              top: y,
              width: 4,
              height: 96,
              background: pillar.tone,
            }}
          />
        );
      })}
      {pillars.map((pillar, index) => {
        const y = 318 + index * 110;
        return (
          <Text
            key={`${pillar.code}-label`}
            style={{
              position: "absolute",
              left: 74,
              top: y + 10,
              width: 220,
              height: 18,
              fontSize: 11,
              color: colors.computeBright,
              textWrap: "nowrap",
            }}
          >
            {pillar.code}
          </Text>
        );
      })}
      {pillars.map((pillar, index) => {
        const y = 318 + index * 110;
        return (
          <Text
            key={`${pillar.code}-title`}
            style={{
              position: "absolute",
              left: 260,
              top: y + 8,
              width: 508,
              height: 22,
              fontSize: 15,
              color: colors.surface,
              textWrap: "nowrap",
            }}
          >
            {pillar.title}
          </Text>
        );
      })}
      {pillars.map((pillar, index) => {
        const y = 318 + index * 110;
        return (
          <Text
            key={`${pillar.code}-detail`}
            style={{
              position: "absolute",
              left: 74,
              top: y + 36,
              width: 694,
              height: 50,
              fontSize: 12.5,
              lineHeight: 20,
              color: colors.darkMuted,
            }}
          >
            {pillar.detail}
          </Text>
        );
      })}

      {/* Right column: Core quantitative telemetry stack */}
      <Text
        style={{
          position: "absolute",
          left: 816,
          top: 86,
          width: 408,
          height: 22,
          fontSize: 13,
          color: colors.darkMuted,
          textWrap: "nowrap",
        }}
      >
        核心架构与实测指标速览（2,048 × H800 SXM5 集群）
      </Text>

      <MotionScope scope={kpiScope} mode="stagger" each={120}>
        <MetricCard
          id="opening-kpi-1"
          enter={floatUp({ duration: 420 })}
          x={816}
          y={118}
          width={408}
          height={122}
          label="单 Token 激活参数 / 总参数量 (MoE)"
          value="37B / 671B"
          caption="每层 1 共享专家 + 256 路由专家（Top-8 激活），另附 14B MTP 模块"
          tone="compute"
          dark
        />
        <MetricCard
          id="opening-kpi-2"
          enter={floatUp({ duration: 420 })}
          x={816}
          y={256}
          width={408}
          height={122}
          label="MLA 每 Token KV 缓存显存占用 (BF16)"
          value="70.27 KB"
          caption="仅为 Qwen-2.5 72B (327.7 KB) 的 21.4%、LLaMA-3.1 405B (516.1 KB) 的 13.6%"
          tone="comm"
          dark
        />
        <MetricCard
          id="opening-kpi-3"
          enter={floatUp({ duration: 420 })}
          x={816}
          y={394}
          width={408}
          height={122}
          label="14.8T 完整训练 GPU 小时与折算总成本"
          value="2.788M h / $5.576M"
          caption="按 $2 / H800 小时核算，预训练 2,664K 小时，全程 0 次 Loss 尖峰回滚"
          tone="copper"
          dark
        />
        <MetricCard
          id="opening-kpi-4"
          enter={floatUp({ duration: 420 })}
          x={816}
          y={532}
          width={408}
          height={112}
          label="24 小时生产环境单节点吞吐与理论成本利润率"
          value="73.7k / 14.8k tok/s · 545%"
          caption="Prefill 4 节点 EP32 / Decode 18 节点 EP144，日均 226.75 节点服务 608B 输入"
          tone="comm"
          dark
        />
      </MotionScope>

      {/* Footer */}
      <Rule x={56} y={672} width={1168} height={1} color={colors.darkLine} />
      <Text
        style={{
          position: "absolute",
          left: 56,
          top: 680,
          width: 960,
          height: 20,
          fontSize: 11,
          color: colors.darkSubtle,
          textWrap: "nowrap",
        }}
      >
        一手数据源：DeepSeek-V3 Technical Report · DeepSeek-R1 · ISCA &apos;25
        Industry Track · SC &apos;24 Fire-Flyer AI-HPC · Open-Source Week
      </Text>
      <Text
        style={{
          position: "absolute",
          left: 1100,
          top: 680,
          width: 124,
          height: 20,
          fontFamily: "metric",
          fontSize: 11,
          fontWeight: 700,
          color: colors.darkMuted,
          align: "right",
          textWrap: "nowrap",
        }}
      >
        01 / 14
      </Text>
    </Slide>
  );
}
