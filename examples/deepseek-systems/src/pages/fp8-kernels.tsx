import { Chart, Shape, Text } from "@officedk/pptxdk/react";

import {
  MetricCard,
  PanelCard,
  Rule,
  SemanticTag,
} from "../components/primitives.tsx";
import { SlideFrame } from "../templates/slide-frame.tsx";
import { colors } from "../theme.ts";

const kernelCompareCategories = [
  "FlashMLA 访存带宽 (GB/s)",
  "FlashMLA BF16 算力 (TFLOPS)",
  "DeepGEMM FP8 算力 (TFLOPS)",
] as const;

const kernelPerfRate = [89.6, 66.7, 78.3] as const;
const baselinePerfRate = [58.0, 48.0, 56.0] as const;

/** Slide 05: Fine-grained FP8 quantization and FlashMLA / DeepGEMM Hopper kernels. */
export function Fp8KernelsPage() {
  return (
    <SlideFrame
      id="fp8-kernels"
      kicker="02 · 细粒度低比特训练与算子微架构 / FP8 · FlashMLA · DeepGEMM"
      title="细粒度 FP8 混合精度与 Hopper 算子优化：逼近 H800 物理屋顶线"
      takeaway="1x128 激活切片 + 128x128 权重分块量化，相对 BF16 训练误差低于 0.25%"
      takeawayTone="copper"
      source="数据来源：DeepSeek-V3 Technical Report Section 3.3 · FlashMLA / DeepGEMM 开源仓库实测 (H800 SXM5)"
    >
      {/* Left column: Fine-grained FP8 quantization & WGMMA promotion diagram */}
      <PanelCard
        x={48}
        y={108}
        width={616}
        height={552}
        title="细粒度 FP8 量化与 Tensor Core 高精度累加机制"
        subtitle="解决激活异常值 (Outliers) 与 WGMMA 定点累加精度不足两大硬件瓶颈"
        rightBadge="E4M3 动态缩放"
        tone="compute"
      >
        {/* Step 1: Tile & Block Quantization */}
        <Shape
          style={{
            position: "absolute",
            left: 68,
            top: 170,
            width: 576,
            height: 166,
            background: colors.paper,
            border: { color: colors.line, width: 1 },
          }}
        />
        <SemanticTag
          x={84}
          y={184}
          width={156}
          height={22}
          text="1. 在线细粒度分组量化"
          tone="compute"
          solid
        />
        <Text
          style={{
            position: "absolute",
            left: 252,
            top: 186,
            width: 376,
            height: 20,
            fontSize: 11.5,
            color: colors.muted,
            textWrap: "nowrap",
          }}
        >
          以局部缩放因子隔离跨通道离群值，避免全局 Per-Tensor 量化下溢
        </Text>

        {/* Activation 1x128 Tile visual */}
        <Shape
          style={{
            position: "absolute",
            left: 84,
            top: 216,
            width: 272,
            height: 104,
            background: colors.surface,
            border: { color: colors.compute, width: 1 },
          }}
        />
        <Shape
          style={{
            position: "absolute",
            left: 84,
            top: 216,
            width: 272,
            height: 4,
            background: colors.compute,
          }}
        />
        <Text
          style={{
            position: "absolute",
            left: 98,
            top: 226,
            width: 244,
            height: 20,
            fontSize: 12.5,
            color: colors.compute,
            textWrap: "nowrap",
          }}
        >
          激活张量：1x128 连续切片 (Tile)
        </Text>
        {/* Visual grid of 1x128 stripes */}
        {[0, 1, 2].map((row) => (
          <Shape
            key={row}
            style={{
              position: "absolute",
              left: 98,
              top: 250 + row * 20,
              width: 168,
              height: 16,
              background: row === 1 ? colors.copperSoft : colors.computeSoft,
              padding: 0,
              border: {
                color: row === 1 ? colors.copper : colors.compute,
                width: 1,
              },
              fontSize: 9.5,
              color: row === 1 ? colors.copperText : colors.compute,
              align: "center",
              verticalAlign: "middle",
            }}
          >
            {row === 1
              ? "Token 离群切片 (独立 Scale)"
              : "1 x 128 E4M3 Elements"}
          </Shape>
        ))}
        <Text
          style={{
            position: "absolute",
            left: 274,
            top: 252,
            width: 76,
            padding: 0,
            fontSize: 10.5,
            lineHeight: 15,
            color: colors.muted,
          }}
        >
          每 128 通道独立在线计算缩放因子
        </Text>

        {/* Weight 128x128 Block visual */}
        <Shape
          style={{
            position: "absolute",
            left: 372,
            top: 216,
            width: 256,
            height: 104,
            background: colors.surface,
            border: { color: colors.comm, width: 1 },
          }}
        />
        <Shape
          style={{
            position: "absolute",
            left: 372,
            top: 216,
            width: 256,
            height: 4,
            background: colors.comm,
          }}
        />
        <Text
          style={{
            position: "absolute",
            left: 386,
            top: 226,
            width: 228,
            height: 20,
            fontSize: 12.5,
            color: colors.commText,
            textWrap: "nowrap",
          }}
        >
          权重参数：128x128 二维分块 (Block)
        </Text>
        {[0, 1].map((r) =>
          [0, 1].map((c) => (
            <Shape
              key={`${r}-${c}`}
              style={{
                position: "absolute",
                left: 386 + c * 74,
                top: 252 + r * 28,
                width: 68,
                height: 23,
                background: r === 0 && c === 1 ? colors.comm : colors.commSoft,
                fontFamily: "metric",
                fontSize: 10,
                fontWeight: 700,
                color: r === 0 && c === 1 ? colors.surface : colors.commText,
                align: "center",
                verticalAlign: "middle",
              }}
            >
              128x128
            </Shape>
          )),
        )}
        <Text
          style={{
            position: "absolute",
            left: 538,
            top: 252,
            width: 82,
            height: 54,
            fontSize: 10.5,
            lineHeight: 15,
            color: colors.muted,
          }}
        >
          2 的幂次缩放减少双重量化误差
        </Text>

        {/* Step 2: Tensor Core WGMMA -> CUDA Core FP32 Register Promotion */}
        <Shape
          style={{
            position: "absolute",
            left: 68,
            top: 350,
            width: 576,
            height: 178,
            background: colors.paper,
            border: { color: colors.line, width: 1 },
          }}
        />
        <SemanticTag
          x={84}
          y={364}
          width={184}
          height={22}
          text="2. 周期性提升至 FP32 寄存器累加"
          tone="copper"
          solid
        />
        <Text
          style={{
            position: "absolute",
            left: 280,
            top: 366,
            width: 348,
            height: 20,
            fontSize: 11.5,
            color: colors.muted,
            textWrap: "nowrap",
          }}
        >
          弥补 Hopper Tensor Core FP8 仅 14-bit 小数位累加器的截断误差
        </Text>

        {/* Pipeline flow boxes inside Step 2 */}
        <Shape
          style={{
            position: "absolute",
            left: 84,
            top: 400,
            width: 162,
            height: 72,
            background: colors.surface,
            border: { color: colors.compute, width: 1 },
          }}
        />
        <Text
          style={{
            position: "absolute",
            left: 94,
            top: 408,
            width: 142,
            height: 20,
            fontSize: 12,
            color: colors.compute,
            textWrap: "nowrap",
          }}
        >
          FP8 WGMMA 指令流
        </Text>
        <Text
          style={{
            position: "absolute",
            left: 94,
            top: 430,
            width: 142,
            height: 34,
            fontSize: 10.5,
            lineHeight: 15,
            color: colors.muted,
          }}
        >
          沿内积维度连续执行 4 次 MMA (Nc = 128 元素)
        </Text>

        <Shape
          style={{
            position: "absolute",
            left: 252,
            top: 424,
            width: 24,
            height: 24,
            background: colors.slot,
            fontFamily: "metric",
            fontSize: 14,
            fontWeight: 700,
            color: colors.steel,
            align: "center",
            verticalAlign: "middle",
          }}
        >
          +
        </Shape>

        <Shape
          style={{
            position: "absolute",
            left: 282,
            top: 400,
            width: 168,
            height: 72,
            background: colors.copperSoft,
            border: { color: colors.copper, width: 1 },
          }}
        />
        <Text
          style={{
            position: "absolute",
            left: 292,
            top: 408,
            width: 148,
            height: 20,
            fontSize: 12,
            color: colors.copperText,
            textWrap: "nowrap",
          }}
        >
          CUDA Core FP32 提升
        </Text>
        <Text
          style={{
            position: "absolute",
            left: 292,
            top: 430,
            width: 148,
            height: 34,
            fontSize: 10.5,
            lineHeight: 15,
            color: colors.copperText,
          }}
        >
          乘以双向反量化缩放因子并累加至 FP32 主寄存器
        </Text>

        <Shape
          style={{
            position: "absolute",
            left: 456,
            top: 424,
            width: 24,
            height: 24,
            background: colors.slot,
            fontFamily: "metric",
            fontSize: 14,
            fontWeight: 700,
            color: colors.steel,
            align: "center",
            verticalAlign: "middle",
          }}
        >
          =
        </Shape>

        <Shape
          style={{
            position: "absolute",
            left: 486,
            top: 400,
            width: 142,
            height: 72,
            background: colors.commSoft,
            border: { color: colors.comm, width: 1 },
          }}
        />
        <Text
          style={{
            position: "absolute",
            left: 496,
            top: 408,
            width: 122,
            height: 20,
            fontSize: 12,
            color: colors.commText,
            textWrap: "nowrap",
          }}
        >
          WGMMA 流水重叠
        </Text>
        <Text
          style={{
            position: "absolute",
            left: 496,
            top: 430,
            width: 122,
            height: 34,
            fontSize: 10.5,
            lineHeight: 15,
            color: colors.commText,
          }}
        >
          Tensor Core 与 CUDA Core 异构并发无停顿
        </Text>

        {/* Bottom note inside left panel */}
        <Text
          style={{
            position: "absolute",
            left: 84,
            top: 484,
            width: 544,
            height: 36,
            fontSize: 11,
            lineHeight: 16,
            color: colors.steel,
          }}
        >
          通信与激活缓存协同：MoE 专家分发前先量化为 FP8，All-to-All 通信量相比
          BF16 直接减半；低秩缓存激活采用 E5M6 / 自定义 E4M3 格式驻留显存。
        </Text>

        <Rule x={68} y={536} width={576} />
        <Text
          style={{
            position: "absolute",
            left: 68,
            top: 542,
            width: 576,
            padding: 0,
            fontSize: 10.5,
            color: colors.muted,
            textWrap: "nowrap",
          }}
        >
          核心矩阵乘 (Fprop / Dgrad / Wgrad) 全部切换至
          FP8，Embedding、门控与归一化保持高精度
        </Text>
      </PanelCard>

      {/* Right top: Kernel performance chart */}
      <PanelCard
        x={680}
        y={108}
        width={552}
        height={354}
        title="H800 SXM5 开源算子硬件峰值利用率对比 (%)"
        subtitle="FlashMLA 解码核与 DeepGEMM 运行时 JIT 编译库相对 H800 物理上限实测"
        rightBadge="Hopper 架构调优"
        tone="copper"
      >
        <Chart
          id="fp8-kernels-chart"
          legend="top"
          axes={{
            category: {
              kind: "category",
              line: { color: colors.steel, width: 1 },
              style: { fontFamily: "body", fontSize: 10.5, color: colors.ink },
            },
            value: {
              kind: "value",
              min: 0,
              max: 100,
              majorUnit: 25,
              numberFormat: '0"%"',
              gridlines: { color: colors.line, width: 1 },
              line: false,
              style: { fontFamily: "body", fontSize: 10, color: colors.muted },
            },
          }}
          groups={[
            {
              kind: "column",
              data: {
                kind: "shared-categories",
                categories: kernelCompareCategories,
              },
              series: [
                {
                  name: "常规通用实现基准 (%)",
                  values: baselinePerfRate,
                  color: colors.baseline,
                },
                {
                  name: "DeepSeek 开源算子实测 (%)",
                  values: kernelPerfRate,
                  color: colors.copper,
                },
              ],
              dataLabels: {
                value: true,
                position: "outsideEnd",
                numberFormat: '0.0"%"',
                style: { fontFamily: "body", fontSize: 10, color: colors.ink },
              },
            },
          ]}
          style={{
            position: "absolute",
            left: 696,
            top: 168,
            width: 520,
            height: 240,
            fontFamily: "body",
            fontSize: 10.5,
            color: colors.ink,
          }}
        />
        <Text
          style={{
            position: "absolute",
            left: 698,
            top: 416,
            width: 516,
            padding: 0,
            fontSize: 11,
            lineHeight: 16,
            color: colors.muted,
          }}
        >
          注：FlashMLA 使用跷跷板 (Seesaw) 调度与 Paged KV (块大小 64) 榨干
          3,350 GB/s HBM3 带宽；DeepGEMM
          利用运行时即时编译消除模板膨胀，支持普通与分组 MoE 矩阵乘。
        </Text>
      </PanelCard>

      {/* Right bottom: 3 concrete kernel telemetry metric cards */}
      <MetricCard
        x={680}
        y={476}
        width={174}
        height={90}
        label="FlashMLA 显存带宽"
        value="3,000 GB/s"
        caption="达 H800 3.35 TB/s 峰值 89.6%"
        tone="comm"
      />
      <MetricCard
        x={869}
        y={476}
        width={174}
        height={90}
        label="FlashMLA 计算吞吐"
        value="660 TFLOPS"
        caption="BF16 吸收矩阵乘算力利用率"
        tone="compute"
      />
      <MetricCard
        x={1058}
        y={476}
        width={174}
        height={90}
        label="DeepGEMM FP8 峰值"
        value="1,550 TFLOPS"
        caption="仅约 300 行核心 C++ JIT 代码"
        tone="copper"
      />

      {/* Right bottom engineering insight box */}
      <Shape
        style={{
          position: "absolute",
          left: 680,
          top: 578,
          width: 552,
          height: 82,
          background: colors.surface,
          border: { color: colors.line, width: 1 },
        }}
      />
      <Shape
        style={{
          position: "absolute",
          left: 680,
          top: 578,
          width: 4,
          height: 82,
          background: colors.compute,
        }}
      />
      <Text
        style={{
          position: "absolute",
          left: 698,
          top: 586,
          width: 520,
          padding: 0,
          fontSize: 13,
          color: colors.steel,
          textWrap: "nowrap",
        }}
      >
        DeepGEMM 运行时 JIT 编译设计亮点：
      </Text>
      <Text
        style={{
          position: "absolute",
          left: 698,
          top: 608,
          width: 520,
          padding: 0,
          fontSize: 11,
          lineHeight: 16,
          color: colors.muted,
        }}
      >
        把矩阵维度 M/N/K
        与分块大小作为编译期常量动态注入，展开循环并优化共享内存分配；通过 SASS
        指令重排开启 FFMA 重叠比特位，在不依赖重型模板库的前提下跑满 Hopper
        Tensor Core。
      </Text>
    </SlideFrame>
  );
}
