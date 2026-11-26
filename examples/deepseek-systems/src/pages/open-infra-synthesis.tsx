import { Shape, Text } from "@officedk/pptxdk/react";

import { Rule } from "../components/primitives.tsx";
import { SlideFrame } from "../templates/slide-frame.tsx";
import { colors } from "../theme.ts";

const openSourceWeekItems = [
  {
    day: "Day 1",
    repo: "FlashMLA",
    layer: "注意力解码内核",
    kpi: "3,000 GB/s | 660 TFLOPS",
    desc: "专为 Hopper 变长序列设计的 MLA 解码算子，以跷跷板调度与分页 KV 榨干 HBM3。",
    accent: colors.compute,
    bright: colors.computeBright,
  },
  {
    day: "Day 2",
    repo: "DeepEP",
    layer: "MoE 专家通信库",
    kpi: "158 GB/s NVL | 58 GB/s IB",
    desc: "支持高吞吐 NVLink+RDMA 非对称转发与纯 IBGDA 零 SM 占用低延迟解码。",
    accent: colors.comm,
    bright: colors.commBright,
  },
  {
    day: "Day 3",
    repo: "DeepGEMM",
    layer: "FP8 矩阵乘 JIT 库",
    kpi: "1,550 TFLOPS | ~300 LOC",
    desc: "运行时即时编译消除模板膨胀，支持密集与 MoE 分组连续/掩码布局细粒度 FP8 算子。",
    accent: colors.copper,
    bright: colors.copperBright,
  },
  {
    day: "Day 4",
    repo: "DualPipe & EPLB",
    layer: "并行流水与负载均衡",
    kpi: "PP16 + EP64 | 256 -> 288",
    desc: "双向微批次交错隐藏跨机 EP 通信；冗余热专家装箱消除 MoE 推理尾延迟。",
    accent: colors.compute,
    bright: colors.computeBright,
  },
  {
    day: "Day 5",
    repo: "3FS & Smallpond",
    layer: "全闪并行存储与数据栈",
    kpi: "6.6 TiB/s | 3.66 TiB/min",
    desc: "CRAQ 强一致 + RDMA 零拷贝随机读，兼顾训练洗牌、Checkpoint 与推理 KV 缓存卸载。",
    accent: colors.comm,
    bright: colors.commBright,
  },
  {
    day: "Day 6",
    repo: "V3/R1 推理系统全景",
    layer: "PD 分离与潮汐调度",
    kpi: "73.7k / 14.8k Tok/s | 545%",
    desc: "EP32 预填充 + EP144 5 阶段解码流水，昼夜算力潮汐复用实现高吞吐低延迟。",
    accent: colors.copper,
    bright: colors.copperBright,
  },
] as const;

const hardwareReflections = [
  {
    idx: "01",
    title: "Tensor Core 高精度累加与块级量化硬件原生化",
    accent: colors.copper,
    bright: colors.copperBright,
    body: "现状痛点：Hopper FP8 WGMMA 累加寄存器仅 14-bit 小数位，迫使软件每 4 次 MMA 转入 CUDA Core 做 FP32 累加。硬件演进建议：在 Tensor Core 内部原生提供 >=32-bit 全精度累加器，并硬化 1x128 / 128x128 细粒度分组动态缩放指令。",
  },
  {
    idx: "02",
    title: "Scale-Up 与 Scale-Out 网络融合及专用通信协处理器",
    accent: colors.comm,
    bright: colors.commBright,
    body: "现状痛点：为打通机内 NVLink 与机间 IB RDMA 并执行 FP8 反量化归约，训练与 Prefill 需牺牲 20–24 个宝贵的计算 SM。硬件演进建议：引入独立的片上网络协处理器 (I/O Die Engine) 统一虚拟寻址，硬件卸载跨平面包转发与在线归约。",
  },
  {
    idx: "03",
    title: "突破访存带宽墙与超大规模集群微秒级自适应容错",
    accent: colors.compute,
    bright: colors.computeBright,
    body: "现状痛点：即便 MLA 削减了 93.3% KV 缓存，长思维链强化学习与超长上下文解码仍高度受限于 HBM 带宽与网络抖动。硬件演进建议：通过 3D 堆叠 DRAM 大幅拉升带宽容量比，在低成本多平面网络中原生集成硬件级包喷洒与拥塞隔离。",
  },
] as const;

/** Slide 14: Closing synthesis of the Open-Source Week stack and ISCA '25 hardware co-design roadmap. */
export function OpenInfraSynthesisPage() {
  return (
    <SlideFrame
      id="open-infra-synthesis"
      dark
      kicker="04 · 全栈开源基础设施全景与下一代 AI 芯片协同设计 / Open-Source Week · ISCA '25"
      title="总结：模型—算子—网络—存储全栈闭环与下一代硬件协同演进"
      takeaway="以软件与算法协同逼近当代硅片极限，反哺下一代 AI 硬件架构定义"
      takeawayTone="copper"
      source="数据来源：DeepSeek 2025.02 Open-Source Week 全系列仓库 · ISCA '25 Insights into DeepSeek-V3 Hardware Reflections"
    >
      {/* Section 1 header */}
      <Text
        style={{
          position: "absolute",
          left: 48,
          top: 106,
          width: 1184,
          height: 22,
          fontSize: 13,
          color: colors.computeBright,
          textWrap: "nowrap",
        }}
      >
        一、DeepSeek Open-Source Week 开源基础设施全栈图谱（从底层 PTX
        算子到万卡集群存储与推理编排）
      </Text>

      {/* 6 columns for Day 1 to Day 6 */}
      {openSourceWeekItems.map((item, idx) => {
        const cardX = 48 + idx * 199;
        return (
          <Shape
            key={item.day}
            style={{
              position: "absolute",
              left: cardX,
              top: 134,
              width: 189,
              height: 214,
              background: colors.obsidianCard,
              border: { color: colors.darkLine, width: 1 },
            }}
          />
        );
      })}
      {openSourceWeekItems.map((item, idx) => {
        const cardX = 48 + idx * 199;
        return (
          <Shape
            key={`bar-${item.day}`}
            style={{
              position: "absolute",
              left: cardX,
              top: 134,
              width: 189,
              height: 3,
              background: item.accent,
            }}
          />
        );
      })}
      {openSourceWeekItems.map((item, idx) => {
        const cardX = 48 + idx * 199;
        return (
          <Shape
            key={`badge-${item.day}`}
            style={{
              position: "absolute",
              left: cardX + 12,
              top: 146,
              width: 54,
              height: 20,
              background: item.accent,
              fontFamily: "metric",
              fontSize: 10,
              fontWeight: 700,
              color: colors.surface,
              align: "center",
              verticalAlign: "middle",
            }}
          >
            {item.day}
          </Shape>
        );
      })}
      {openSourceWeekItems.map((item, idx) => {
        const cardX = 48 + idx * 199;
        return (
          <Text
            key={`layer-${item.day}`}
            style={{
              position: "absolute",
              left: cardX + 72,
              top: 147,
              width: 108,
              height: 18,
              fontSize: 10,
              color: colors.darkMuted,
              textWrap: "nowrap",
            }}
          >
            {item.layer}
          </Text>
        );
      })}
      {openSourceWeekItems.map((item, idx) => {
        const cardX = 48 + idx * 199;
        return (
          <Text
            key={`repo-${item.day}`}
            style={{
              position: "absolute",
              left: cardX + 12,
              top: 174,
              width: 165,
              height: 24,
              fontSize: 14,
              color: colors.surface,
              textWrap: "nowrap",
            }}
          >
            {item.repo}
          </Text>
        );
      })}
      {openSourceWeekItems.map((item, idx) => {
        const cardX = 48 + idx * 199;
        return (
          <Text
            key={`kpi-${item.day}`}
            style={{
              position: "absolute",
              left: cardX + 12,
              top: 202,
              width: 165,
              height: 20,
              fontFamily: "metric",
              fontSize: 11.5,
              fontWeight: 700,
              color: item.bright,
              textWrap: "nowrap",
            }}
          >
            {item.kpi}
          </Text>
        );
      })}
      {openSourceWeekItems.map((item, idx) => {
        const cardX = 48 + idx * 199;
        return (
          <Rule
            key={`rule-${item.day}`}
            x={cardX + 12}
            y={226}
            width={165}
            color={colors.darkLine}
          />
        );
      })}
      {openSourceWeekItems.map((item, idx) => {
        const cardX = 48 + idx * 199;
        return (
          <Text
            key={`desc-${item.day}`}
            style={{
              position: "absolute",
              left: cardX + 12,
              top: 234,
              width: 165,
              height: 102,
              fontSize: 10.5,
              lineHeight: 15.5,
              color: colors.darkMuted,
            }}
          >
            {item.desc}
          </Text>
        );
      })}

      {/* Section 2 header */}
      <Text
        style={{
          position: "absolute",
          left: 48,
          top: 364,
          width: 1184,
          height: 22,
          fontSize: 13,
          color: colors.copperBright,
          textWrap: "nowrap",
        }}
      >
        二、面向下一代 AI 芯片与集群互连的三大硬件协同设计建议（ISCA &apos;25
        论文核心反思）
      </Text>

      {/* 3 columns for Hardware Co-Design Recommendations */}
      {hardwareReflections.map((rec, idx) => {
        const cardX = 48 + idx * 400;
        return (
          <Shape
            key={rec.idx}
            style={{
              position: "absolute",
              left: cardX,
              top: 394,
              width: 384,
              height: 264,
              background: colors.obsidianCard,
              border: { color: colors.darkLine, width: 1 },
            }}
          />
        );
      })}
      {hardwareReflections.map((rec, idx) => {
        const cardX = 48 + idx * 400;
        return (
          <Shape
            key={`left-${rec.idx}`}
            style={{
              position: "absolute",
              left: cardX,
              top: 394,
              width: 4,
              height: 264,
              background: rec.accent,
            }}
          />
        );
      })}
      {hardwareReflections.map((rec, idx) => {
        const cardX = 48 + idx * 400;
        return (
          <Shape
            key={`num-${rec.idx}`}
            style={{
              position: "absolute",
              left: cardX + 18,
              top: 410,
              width: 36,
              height: 24,
              background: rec.accent,
              fontFamily: "metric",
              fontSize: 12,
              fontWeight: 700,
              color: colors.surface,
              align: "center",
              verticalAlign: "middle",
            }}
          >
            {rec.idx}
          </Shape>
        );
      })}
      {hardwareReflections.map((rec, idx) => {
        const cardX = 48 + idx * 400;
        return (
          <Text
            key={`title-${rec.idx}`}
            style={{
              position: "absolute",
              left: cardX + 60,
              top: 412,
              width: 312,
              padding: 0,
              fontSize: 12.5,
              color: rec.bright,
              textWrap: "nowrap",
            }}
          >
            {rec.title}
          </Text>
        );
      })}
      {hardwareReflections.map((rec, idx) => {
        const cardX = 48 + idx * 400;
        return (
          <Rule
            key={`r-${rec.idx}`}
            x={cardX + 18}
            y={446}
            width={348}
            color={colors.darkLine}
          />
        );
      })}
      {hardwareReflections.map((rec, idx) => {
        const cardX = 48 + idx * 400;
        return (
          <Text
            key={`body-${rec.idx}`}
            style={{
              position: "absolute",
              left: cardX + 18,
              top: 458,
              width: 348,
              height: 186,
              fontSize: 11,
              lineHeight: 17,
              color: colors.darkMuted,
            }}
          >
            {rec.body}
          </Text>
        );
      })}
    </SlideFrame>
  );
}
