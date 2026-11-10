import { Chart, Shape, Text, wipe } from "@react-pptx/react-pptx";

import { PanelCard, SemanticTag } from "../components/primitives.tsx";
import { SlideFrame } from "../templates/slide-frame.tsx";
import { colors } from "../theme.ts";

export function MlaMoeEfficiencyPage() {
  return (
    <SlideFrame
      slideId="mla-moe-efficiency"
      kicker="01 · 模型架构基座｜MLA 键值压缩与 DeepSeekMoE 算力效率"
      title="MLA 将每 Token KV 缓存压至 70.3 KB，MoE 使单 Token 训练算力降至稠密模型 1/9.8"
      takeaway="ISCA '25 实测对比：KV 缓存较 LLaMA-3.1 405B 缩减 86.4%，单 Token 训练仅 250 GFLOPS"
      takeawayTone="compute"
      source="数据源：Insights into DeepSeek-V3 (ISCA '25, arXiv:2505.09343) Table 1 & Table 2（4096 序列长度，BF16 精度）"
    >
      {/* Left Panel: MLA KV Cache Compression */}
      <PanelCard
        x={48}
        y={108}
        width={580}
        height={554}
        title="表 A：每 Token KV Cache 显存占用对比 (BF16)"
        subtitle="MLA 将键值联合投影至低秩潜在向量，仅需缓存压缩后的潜在状态与解耦 RoPE 向量"
        rightBadge="越低越优 (KB)"
        tone="compute"
      >
        <Chart
          testId="mla-kv-cache-chart"
          kind="bar"
          enter={wipe({ by: "category", from: "left", duration: 460 })}
          x={66}
          y={168}
          width={544}
          height={272}
          style={{ fontFamily: "body", fontSize: 12, color: colors.ink }}
          data={{
            categories: [
              "DeepSeek-V3 (MLA)",
              "Qwen-2.5 72B (GQA)",
              "LLaMA-3.1 405B (GQA)",
            ],
            series: [
              {
                name: "每 Token KV 缓存大小 (KB)",
                values: [70.272, 327.68, 516.096],
                color: colors.baseline,
                points: {
                  0: { color: colors.compute, labelColor: colors.compute },
                  1: { color: colors.baseline, labelColor: colors.steel },
                  2: { color: colors.baseline, labelColor: colors.steel },
                },
              },
            ],
          }}
          legend={false}
          gapWidth={72}
          plotArea={{ left: 146, top: 18, width: 344, height: 204 }}
          dataLabels={{
            value: true,
            position: "outsideEnd",
            numberFormat: '0.0" KB"',
            style: { fontFamily: "body", fontSize: 11.5, color: colors.ink },
          }}
          valueAxis={{
            min: 0,
            max: 600,
            majorUnit: 150,
            numberFormat: '0" KB"',
            gridlines: { color: colors.line, width: 1 },
            line: false,
            style: { fontFamily: "body", fontSize: 11, color: colors.muted },
          }}
          categoryAxis={{
            line: { color: colors.steel, width: 1 },
            style: { fontFamily: "body", fontSize: 12, color: colors.ink },
          }}
        />

        {/* MLA Low-Rank Compression Schematic */}
        <Shape
          x={68}
          y={452}
          width={540}
          height={192}
          fill={colors.paper}
          style={{ borderColor: colors.line, borderWidth: 1 }}
        />
        <Text
          x={84}
          y={462}
          width={508}
          height={20}
          style={{ fontSize: 13, color: colors.steel, textWrap: "nowrap" }}
        >
          MLA（多头潜在注意力）低秩联合压缩机制拆解：
        </Text>

        <SemanticTag
          x={84}
          y={494}
          width={142}
          height={42}
          text="输入隐状态 h_t (7168 维)"
          tone="steel"
        />
        <Shape x={234} y={512} width={24} height={4} fill={colors.compute} />
        <SemanticTag
          x={266}
          y={494}
          width={168}
          height={42}
          text="KV 潜在向量 c_t^KV (512 维) + 解耦 RoPE k_t^R (64 维)"
          tone="compute"
          solid
        />
        <Shape x={442} y={512} width={24} height={4} fill={colors.comm} />
        <SemanticTag
          x={474}
          y={494}
          width={118}
          height={42}
          text="仅缓存 576 维/层 (70.27 KB/Tok)"
          tone="comm"
        />

        <Text
          x={84}
          y={548}
          width={508}
          height={86}
          style={{ fontSize: 12, lineHeight: 19, color: colors.muted }}
        >
          • 传统 GQA 需按键值头数缓存完整 K/V 张量；MLA 将 61 层每层 KV
          缓存收敛为 512 维潜在向量与 64
          维旋转位置编码，推理时将上投影矩阵吸收至 Q 侧，显著缓解长上下文解码的
          HBM 带宽与容量瓶颈。
        </Text>
      </PanelCard>

      {/* Right Panel: DeepSeekMoE Single-Token Compute */}
      <PanelCard
        x={644}
        y={108}
        width={588}
        height={554}
        title="表 B：单 Token 训练计算量对比 (GFLOPS / Token)"
        subtitle="以 37B 激活参数撬动 671B 总参数容量，计算开销低于 72B 稠密模型"
        rightBadge="4K 序列长度"
        tone="comm"
      >
        <Chart
          testId="moe-training-flops-chart"
          kind="column"
          enter={wipe({ by: "category", duration: 460 })}
          x={662}
          y={168}
          width={552}
          height={272}
          style={{ fontFamily: "body", fontSize: 12, color: colors.ink }}
          data={{
            // Keep full model names and parameter counts visible on narrow category axes.
            categories: [
              "DeepSeek-V2\n(21B/236B)",
              "DeepSeek-V3\n(37B/671B)",
              "Qwen-2.5 72B\n(Dense)",
              "LLaMA-3.1 405B\n(Dense)",
            ],
            series: [
              {
                name: "单 Token 训练算力 (GFLOPS)",
                values: [155, 250, 394, 2448],
                color: colors.baseline,
                points: {
                  0: { color: colors.comm, labelColor: colors.commText },
                  1: { color: colors.compute, labelColor: colors.compute },
                  2: { color: colors.baseline, labelColor: colors.steel },
                  3: { color: colors.copper, labelColor: colors.copperText },
                },
              },
            ],
          }}
          legend={false}
          gapWidth={85}
          plotArea={{ left: 58, top: 24, width: 474, height: 194 }}
          dataLabels={{
            value: true,
            position: "outsideEnd",
            numberFormat: '#,##0" G"',
            style: { fontFamily: "body", fontSize: 11.5, color: colors.ink },
          }}
          valueAxis={{
            min: 0,
            max: 2800,
            majorUnit: 700,
            numberFormat: "#,##0",
            gridlines: { color: colors.line, width: 1 },
            line: false,
            style: { fontFamily: "body", fontSize: 11, color: colors.muted },
          }}
          categoryAxis={{
            line: { color: colors.steel, width: 1 },
            style: { fontFamily: "body", fontSize: 11, color: colors.ink },
          }}
        />

        {/* MoE vs Dense Comparison Box */}
        <Shape
          x={664}
          y={452}
          width={548}
          height={192}
          fill={colors.paper}
          style={{ borderColor: colors.line, borderWidth: 1 }}
        />
        <Text
          x={680}
          y={462}
          width={516}
          height={20}
          style={{ fontSize: 13, color: colors.steel, textWrap: "nowrap" }}
        >
          细粒度专家拆分与共享专家隔离（DeepSeekMoE 核心配置）：
        </Text>

        <Shape
          x={680}
          y={492}
          width={250}
          height={64}
          fill={colors.surface}
          style={{ borderColor: colors.compute, borderWidth: 1 }}
        />
        <Text
          x={692}
          y={498}
          width={226}
          height={20}
          style={{ fontSize: 12, color: colors.compute, textWrap: "nowrap" }}
        >
          DeepSeek-V3 (MoE 稀疏激活)
        </Text>
        <Text
          x={692}
          y={520}
          width={226}
          height={30}
          style={{
            fontFamily: "metric",
            fontSize: 18,
            fontWeight: 700,
            color: colors.ink,
            textWrap: "nowrap",
          }}
        >
          250 GFLOPS · 5.5% Act.
        </Text>

        <Shape
          x={946}
          y={492}
          width={250}
          height={64}
          fill={colors.surface}
          style={{ borderColor: colors.baseline, borderWidth: 1 }}
        />
        <Text
          x={958}
          y={498}
          width={226}
          height={20}
          style={{ fontSize: 12, color: colors.steel, textWrap: "nowrap" }}
        >
          LLaMA-3.1 405B (稠密全激活)
        </Text>
        <Text
          x={958}
          y={520}
          width={226}
          height={30}
          style={{
            fontFamily: "metric",
            fontSize: 18,
            fontWeight: 700,
            color: colors.copperText,
            textWrap: "nowrap",
          }}
        >
          2,448 GFLOPS · 9.79x
        </Text>

        <Text
          x={680}
          y={566}
          width={516}
          height={68}
          style={{ fontSize: 12, lineHeight: 19, color: colors.muted }}
        >
          • 除前 3 层采用稠密 FFN 外，其余 58 层均为 MoE 层：每层配置 1
          个常驻共享专家捕获通用知识，256 个细粒度路由专家按 Token 动态激活 8
          个，在扩充模型知识容量的同时严控单步访存与计算开销。
        </Text>
      </PanelCard>
    </SlideFrame>
  );
}
