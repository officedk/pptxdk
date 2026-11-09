import { Shape, Step, Text, fillColor, preset } from "@react-pptx/react-pptx";

import { PanelCard, Rule, SemanticTag } from "../components/primitives.tsx";
import { SlideFrame } from "../templates/slide-frame.tsx";
import { colors } from "../theme.ts";

const routedExperts = [
  { id: "E01", active: true, hot: false },
  { id: "E02", active: false, hot: false },
  { id: "E03", active: true, hot: true },
  { id: "E04", active: false, hot: false },
  { id: "E05", active: true, hot: false },
  { id: "E06", active: false, hot: false },
  { id: "E07", active: true, hot: false },
  { id: "E08", active: false, hot: false },
] as const;

export function AuxLossFreeMtpPage() {
  return (
    <SlideFrame
      slideId="aux-loss-free-mtp"
      kicker="01 · 模型架构基座｜无辅助损失负载均衡与多 Token 预测 (MTP)"
      title="动态偏置路由消除辅助损失干扰，14B MTP 模块兼顾表征强化与 1.8× 投机解码"
      takeaway="核心设计：偏置项仅参与 Top-K 专家选择，门控权重仍取原始亲和度，实现零精度损失均衡"
      takeawayTone="comm"
      source="数据源：DeepSeek-V3 Technical Report (arXiv:2412.19437) Section 2.1.2 & Section 2.2"
    >
      {/* Left Panel: Auxiliary-Loss-Free Load Balancing */}
      <PanelCard
        x={48}
        y={108}
        width={580}
        height={554}
        title="无辅助损失负载均衡（Auxiliary-Loss-Free Load Balancing）"
        subtitle="传统辅助损失强行拉平路由会损害语言建模困惑度；动态偏置项实现解耦调控"
        rightBadge="Top-8 / 256 专家"
        tone="compute"
      >
        {/* Step 1: Router Affinity */}
        <Shape
          x={68}
          y={170}
          width={540}
          height={86}
          fill={colors.paper}
          style={{ borderColor: colors.line, borderWidth: 1 }}
        />
        <SemanticTag
          x={82}
          y={184}
          width={148}
          height={26}
          text="1. 原始亲和度打分"
          tone="compute"
          solid
        />
        <Text
          x={242}
          y={184}
          width={356}
          style={{
            padding: 0,
            fontSize: 11.5,
            color: colors.ink,
            textWrap: "nowrap",
          }}
        >
          采用 Sigmoid 计算亲和度：s_i,t = Sigmoid(u_t^T · e_i)
        </Text>
        <Text
          x={82}
          y={218}
          width={512}
          style={{ padding: 0, fontSize: 11.5, color: colors.muted }}
        >
          归一化门控权重仅由被选中专家的原始亲和度 s_i,t
          决定，不掺入任何人工正则梯度。
        </Text>

        {/* Step 2: Bias Injection */}
        <Shape
          x={68}
          y={268}
          width={540}
          height={98}
          fill={colors.computeSoft}
          style={{ borderColor: colors.compute, borderWidth: 1 }}
        />
        <SemanticTag
          x={82}
          y={280}
          width={148}
          height={26}
          text="2. 偏置项参与选路"
          tone="compute"
          solid
        />
        <Text
          x={242}
          y={282}
          width={356}
          style={{
            padding: 0,
            fontSize: 12.5,
            color: colors.compute,
            textWrap: "nowrap",
          }}
        >
          选路准则：Top-K( s_i,t + b_i )，合成权重仍用原始 s_i,t
        </Text>
        <Text
          x={82}
          y={312}
          width={512}
          style={{
            padding: 0,
            fontSize: 11.5,
            lineHeight: 18,
            color: colors.steel,
          }}
        >
          每个训练步统计各专家实际负载：若专家过载则下调偏置 b_i ← b_i -
          γ；若专家欠载则上调 b_i ← b_i + γ（超参数 γ = 0.001），动态引导后续
          Token 流向空闲专家。
        </Text>

        {/* Step 3: Visual Expert Grid with Click Emphasis */}
        <Shape
          x={68}
          y={378}
          width={540}
          height={176}
          fill={colors.surface}
          style={{ borderColor: colors.line, borderWidth: 1 }}
        />
        <Text
          x={82}
          y={388}
          width={510}
          height={20}
          style={{ fontSize: 12, color: colors.steel, textWrap: "nowrap" }}
        >
          专家层路由示意（1 常驻共享专家 + 256
          路由专家池，点击触发热点偏置调平）：
        </Text>

        <Shape
          x={82}
          y={418}
          width={106}
          height={72}
          fill={colors.comm}
          text={"Shared Expert\n常驻全激活"}
          textStyle={{
            fontSize: 11.5,
            color: colors.surface,
            align: "center",
            verticalAlign: "middle",
          }}
        />

        <Step n={1}>
          {routedExperts.map((expert, idx) => {
            const col = idx % 4;
            const row = Math.floor(idx / 4);
            const exX = 204 + col * 98;
            const exY = 418 + row * 38;
            return (
              <Shape
                key={expert.id}
                testId={`routed-expert-${expert.id}`}
                x={exX}
                y={exY}
                width={90}
                height={32}
                fill={
                  expert.hot
                    ? colors.copperSoft
                    : expert.active
                      ? colors.computeSoft
                      : colors.slot
                }
                {...(expert.hot
                  ? {
                      emphasis: fillColor({
                        from: colors.copperSoft,
                        to: colors.commSoft,
                        duration: 500,
                      }),
                    }
                  : {})}
                style={{
                  borderColor: expert.hot
                    ? colors.copper
                    : expert.active
                      ? colors.compute
                      : colors.line,
                  borderWidth: 1,
                }}
                text={
                  expert.hot
                    ? `${expert.id} 过载 b↓`
                    : expert.active
                      ? `${expert.id} 激活`
                      : `${expert.id} 空闲`
                }
                textStyle={{
                  fontSize: 11,
                  color: expert.hot
                    ? colors.copperText
                    : expert.active
                      ? colors.compute
                      : colors.muted,
                  align: "center",
                  verticalAlign: "middle",
                }}
              />
            );
          })}
        </Step>

        <Text
          x={82}
          y={502}
          width={510}
          height={44}
          style={{ fontSize: 11.5, lineHeight: 18, color: colors.muted }}
        >
          • 节点受限路由（Node-Limited Routing）：256 个专家分为 8 组部署在 8
          个节点上，每个 Token 最多路由至 M=4 个节点，从算法源头限制跨节点 IB
          All-to-All 通信扇出。
        </Text>

        {/* Bottom Summary Strip */}
        <Shape
          x={68}
          y={566}
          width={540}
          height={82}
          fill={colors.commSoft}
          style={{ borderColor: colors.comm, borderWidth: 1 }}
        />
        <Text
          x={82}
          y={576}
          width={512}
          height={62}
          style={{ fontSize: 12, lineHeight: 19, color: colors.commText }}
        >
          工程收益对比：相比传统的序列级辅助损失（Sequence-Wise Auxiliary
          Loss），无辅助损失偏置法在保持各 GPU
          专家计算负载均衡的同时，显著提升了 BBH、MATH 与代码生成基准得分。
        </Text>
      </PanelCard>

      {/* Right Panel: Multi-Token Prediction (MTP) */}
      <PanelCard
        x={644}
        y={108}
        width={588}
        height={554}
        title="多 Token 预测目标（Multi-Token Prediction, MTP）"
        subtitle="通过因果链式顺序预测未来多个 Token，共享词表嵌入与输出头以控制显存开销"
        rightBadge="14B 附加模块"
        tone="copper"
      >
        {/* Architecture Flow Diagram for MTP */}
        <Shape
          x={664}
          y={170}
          width={548}
          height={232}
          fill={colors.paper}
          style={{ borderColor: colors.line, borderWidth: 1 }}
        />
        <Text
          x={680}
          y={180}
          width={516}
          height={20}
          style={{ fontSize: 12.5, color: colors.steel, textWrap: "nowrap" }}
        >
          MTP 因果链式预测结构（主模型 671B + MTP 模组 14B = 仓库总权重 685B）：
        </Text>

        {/* Main Model Column */}
        <Shape
          x={680}
          y={212}
          width={242}
          height={174}
          fill={colors.surface}
          style={{ borderColor: colors.compute, borderWidth: 1 }}
        />
        <Shape x={680} y={212} width={242} height={28} fill={colors.compute} />
        <Text
          x={688}
          y={216}
          width={226}
          height={20}
          style={{
            fontSize: 12,
            color: colors.surface,
            align: "center",
            textWrap: "nowrap",
          }}
        >
          Main Model (Depth k=0 · 61 层)
        </Text>
        <SemanticTag
          x={696}
          y={252}
          width={210}
          height={28}
          text="共享 Embedding 层: 输入 t_i"
          tone="steel"
        />
        <SemanticTag
          x={696}
          y={290}
          width={210}
          height={34}
          text="61× Transformer Blocks (h_i^0)"
          tone="compute"
        />
        <SemanticTag
          x={696}
          y={334}
          width={210}
          height={38}
          text="共享 Output Head → 预测 t_{i+1}"
          tone="comm"
        />

        {/* Connector Arrow */}
        <Shape
          geometry={preset("rightArrow")}
          x={930}
          y={288}
          width={32}
          height={26}
          fill={colors.copper}
        />

        {/* MTP Module Column */}
        <Shape
          x={970}
          y={212}
          width={226}
          height={174}
          fill={colors.surface}
          style={{ borderColor: colors.copper, borderWidth: 1 }}
        />
        <Shape x={970} y={212} width={226} height={28} fill={colors.copper} />
        <Text
          x={978}
          y={216}
          width={210}
          height={20}
          style={{
            fontSize: 12,
            color: colors.surface,
            align: "center",
            textWrap: "nowrap",
          }}
        >
          MTP Module (Depth k=1 · 14B)
        </Text>
        <SemanticTag
          x={984}
          y={252}
          width={198}
          height={28}
          text="拼接 RMSNorm(h_i^0) + Emb(t_{i+1})"
          tone="copper"
        />
        <SemanticTag
          x={984}
          y={290}
          width={198}
          height={34}
          text="线性投影 M_1 + 1× Transformer 层"
          tone="copper"
        />
        <SemanticTag
          x={984}
          y={334}
          width={198}
          height={38}
          text="复用 Output Head → 预测 t_{i+2}"
          tone="comm"
        />

        {/* Dual Lifecycle Roles of MTP */}
        <Shape
          x={664}
          y={416}
          width={548}
          height={110}
          fill={colors.surface}
          style={{ borderColor: colors.line, borderWidth: 1 }}
        />
        <Shape x={664} y={416} width={4} height={110} fill={colors.compute} />
        <Text
          x={682}
          y={426}
          width={514}
          height={22}
          style={{ fontSize: 14, color: colors.compute, textWrap: "nowrap" }}
        >
          阶段一：训练期加密监督信号（可独立丢弃，零推理负担）
        </Text>
        <Text
          x={682}
          y={452}
          width={514}
          height={64}
          style={{ fontSize: 12, lineHeight: 19, color: colors.muted }}
        >
          • 相比独立并行输出头的 MTP 方案，DeepSeek-V3
          保持完整的因果链连接，迫使主模型隐状态同时编码当前与后续 Token
          的上下文规划，显著提升数据利用效率。
        </Text>

        <Shape
          x={664}
          y={538}
          width={548}
          height={110}
          fill={colors.copperSoft}
          style={{ borderColor: colors.copper, borderWidth: 1 }}
        />
        <Shape x={664} y={538} width={4} height={110} fill={colors.copper} />
        <Text
          x={682}
          y={548}
          width={360}
          style={{
            padding: 0,
            fontSize: 14,
            color: colors.copperText,
            textWrap: "nowrap",
          }}
        >
          阶段二：推理期自投机解码（Speculative Decoding）
        </Text>
        <Text
          x={1044}
          y={544}
          width={152}
          style={{
            padding: 0,
            fontFamily: "metric",
            fontSize: 21,
            fontWeight: 700,
            color: colors.copperText,
            align: "right",
            textWrap: "nowrap",
          }}
        >
          ~1.8x Speedup
        </Text>
        <Rule x={682} y={576} width={514} height={1} color={colors.copper} />
        <Text
          x={682}
          y={584}
          width={514}
          style={{
            padding: 0,
            fontSize: 12,
            lineHeight: 19,
            color: colors.ink,
          }}
        >
          • 推理解码时直接利用 MTP 模块前瞻预测下一个 Token，实测第二个 Token
          接受率高达 85%–90%，将单流解码吞吐提升约 1.8 倍。
        </Text>
      </PanelCard>
    </SlideFrame>
  );
}
