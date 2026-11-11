import {
  MotionScope,
  Shape,
  Step,
  Text,
  fade,
  transitions,
  useMotionScope,
  wipe,
  type MorphScope,
} from "@react-pptx/react-pptx";

import {
  EnteredShapeText,
  MetricCard,
  PanelCard,
  Rule,
  SemanticTag,
} from "../components/primitives.tsx";
import { SlideFrame } from "../templates/slide-frame.tsx";
import { colors, motionTiming } from "../theme.ts";

const fiveStages = [
  {
    step: "Stage 1",
    mbA: "MB-A Shared + MLA-1",
    mbB: "MB-B Combine (IBGDA)",
    computeWidth: 184,
  },
  {
    step: "Stage 2",
    mbA: "MB-A MLA-2 (FlashMLA)",
    mbB: "MB-B Shared + MLA-1",
    computeWidth: 184,
  },
  {
    step: "Stage 3",
    mbA: "MB-A Dispatch (IBGDA)",
    mbB: "MB-B MLA-2 (FlashMLA)",
    computeWidth: 184,
  },
  {
    step: "Stage 4",
    mbA: "MB-A MoE FP8 GEMM",
    mbB: "MB-B Dispatch (IBGDA)",
    computeWidth: 184,
  },
  {
    step: "Stage 5",
    mbA: "MB-A Combine (IBGDA)",
    mbB: "MB-B MoE FP8 GEMM",
    computeWidth: 184,
  },
] as const;

/** Slide 10: Decode 5-stage pipeline, Morph SM reallocation, and EPLB load balancing. */
export function DecodePipelinePage({ scope }: { readonly scope: MorphScope }) {
  const stageScope = useMotionScope();

  return (
    <SlideFrame
      id="decode-pipeline"
      kicker="03 · 大规模专家并行 (EP144) 与 5 阶段解码流水线 / IBGDA · EPLB"
      title="Decode 5 阶段双微批次流水与 EPLB：零 SM 通信损耗下的极致低延迟"
      takeaway="IBGDA 纯硬件异步通信释放全部 132 个 SM，EPLB 冗余复制消除尾延迟"
      takeawayTone="comm"
      transition={transitions.morph({
        by: "object",
        duration: motionTiming.morph,
        fallback: transitions.fade({ duration: 500 }),
      })}
      source="数据来源：DeepSeek-V3/R1 Inference System Overview · EPLB (Expert Parallelism Load Balancer) 开源仓库"
    >
      {/* Top Left: Morph Target — H800 132 SMs 100% Compute Allocation in Decode */}
      <PanelCard
        x={48}
        y={108}
        width={684}
        height={232}
        title="Decode 阶段 H800 硬件资源重分配：从 108/24 演进为 132 全计算 SM"
        subtitle="由 IBGDA (InfiniBand GPUDirect Async) + 后台 Hook 轮询接管 RDMA，彻底释放 24 个通信 SM"
        rightBadge="平滑形变 · 100% 算力"
        tone="comm"
      >
        <Shape
          id="morph-sm-frame"
          morphKey={scope.handle("sm-frame")}
          x={68}
          y={170}
          width={644}
          height={112}
          fill={colors.paper}
          style={{ borderColor: colors.comm, borderWidth: 1 }}
        />
        <Shape
          id="morph-sm-mode-tag"
          morphKey={scope.handle("sm-mode-tag")}
          x={84}
          y={182}
          width={272}
          height={24}
          fill={colors.comm}
          text="Decode 模式 · EP144 (18 节点 144 卡)"
          textStyle={{
            fontSize: 11.5,
            color: colors.surface,
            align: "center",
            verticalAlign: "middle",
          }}
        />
        <Text
          morphKey={scope.handle("sm-ratio-caption")}
          x={372}
          y={184}
          width={324}
          height={20}
          style={{
            fontFamily: "metric",
            fontSize: 12,
            fontWeight: 700,
            color: colors.commText,
            align: "right",
            textWrap: "nowrap",
          }}
        >
          132 Compute SMs : 0 Comm SMs (IBGDA)
        </Text>

        {/* 132 Compute SMs expanded across the full bar */}
        <Shape
          id="morph-sm-compute-bar"
          morphKey={scope.handle("sm-compute-bar")}
          x={84}
          y={216}
          width={456}
          height={48}
          fill={colors.compute}
          text="132 计算 SMs (100% 算力投入) · 全速运行 FlashMLA 与 FP8 分组专家 GEMM"
          textStyle={{
            fontSize: 11.5,
            color: colors.surface,
            align: "center",
            verticalAlign: "middle",
          }}
        />

        {/* Morphed 0 Comm SM hardware RDMA offload badge */}
        <Shape
          id="morph-sm-comm-bar"
          morphKey={scope.handle("sm-comm-bar")}
          x={548}
          y={216}
          width={148}
          height={48}
          fill={colors.copper}
          style={{ padding: [2, 4] }}
          text={"0 通信 SM\n网卡 IBGDA 异步直发"}
          textStyle={{
            fontSize: 10.5,
            color: colors.surface,
            align: "center",
            verticalAlign: "middle",
          }}
        />

        <Text
          x={68}
          y={294}
          width={644}
          style={{
            padding: 0,
            fontSize: 11,
            lineHeight: 15.5,
            color: colors.muted,
          }}
        >
          注：由于注意力分为 2
          步计算且需等待共享专家输出，传统双微批次重叠会残留等待空隙；将 MLA
          拆为 Part 1 &amp; Part 2 即可构造 5 阶段无缝流水。
        </Text>
      </PanelCard>

      {/* Top Right: EPLB (Expert Parallelism Load Balancer) Architecture */}
      <PanelCard
        x={748}
        y={108}
        width={484}
        height={232}
        title="EPLB 专家并行动态负载均衡器"
        subtitle="冗余专家复制 + 装箱重排，消除热点专家引发的跨节点等待"
        rightBadge="每 10 分钟热更新"
        tone="copper"
      >
        <MetricCard
          x={768}
          y={170}
          width={214}
          height={76}
          label="Decode 物理专家槽位"
          value="256 -> 288"
          caption="144 卡 x 2 路由槽位 (32 冗余)"
          tone="copper"
        />
        <MetricCard
          x={998}
          y={170}
          width={214}
          height={76}
          label="Prefill 物理专家槽位"
          value="256 -> 320"
          caption="32 卡 x 9 路由槽位 + 1 共享"
          tone="compute"
        />
        <Shape
          x={768}
          y={256}
          width={444}
          height={68}
          fill={colors.paper}
          style={{ borderColor: colors.line, borderWidth: 1 }}
        />
        <Text
          x={780}
          y={262}
          width={420}
          style={{
            padding: 0,
            fontSize: 10.5,
            lineHeight: 15,
            color: colors.steel,
          }}
        >
          • 分层均衡
          (Prefill)：先按节点分组均匀装箱，再在节点内复制高负载专家，降低跨节点流量。
          {"\n"}• 全局均衡 (Decode)：跨 144 卡直接全局复制最热专家，利用 EMA
          预测负载并异步双缓冲替换权重。
        </Text>
      </PanelCard>

      {/* Bottom: 5-Stage Decode Pipeline Diagram with Step + MotionScope */}
      <PanelCard
        x={48}
        y={354}
        width={1184}
        height={306}
        title="Decode 5 阶段双微批次无缝流水线 (5-Stage Dual-Microbatch Pipeline)"
        subtitle="将注意力拆分为 MLA Part 1 与 MLA Part 2，使 Microbatch A 与 Microbatch B 的计算阶段与 IBGDA 通信阶段 100% 严丝合缝咬合"
        rightBadge="点击步进展示 5 阶段"
        tone="compute"
      >
        {/* Row labels */}
        <SemanticTag
          x={68}
          y={426}
          width={140}
          height={28}
          text="流水线时间拍 (Stage)"
          tone="steel"
        />
        <SemanticTag
          x={68}
          y={466}
          width={140}
          height={48}
          text="Microbatch A 轨迹"
          tone="compute"
          solid
        />
        <SemanticTag
          x={68}
          y={526}
          width={140}
          height={48}
          text="Microbatch B 轨迹"
          tone="comm"
          solid
        />

        <Step>
          <MotionScope scope={stageScope} mode="stagger" each={110}>
            {fiveStages.map((s, idx) => {
              const colX = 222 + idx * 198;
              return (
                <EnteredShapeText
                  key={`head-${s.step}`}
                  id={`decode-stage-head-${idx}`}
                  x={colX}
                  y={426}
                  width={186}
                  height={28}
                  fill={colors.slot}
                  enter={fade({ duration: 280 })}
                  text={s.step}
                  textStyle={{
                    fontFamily: "metric",
                    fontSize: 11.5,
                    fontWeight: 700,
                    color: colors.steel,
                    align: "center",
                    verticalAlign: "middle",
                  }}
                />
              );
            })}
            {fiveStages.map((s, idx) => {
              const colX = 222 + idx * 198;
              const mbAIsComm = s.mbA.includes("IBGDA");
              return (
                <EnteredShapeText
                  key={`mba-${s.step}`}
                  id={`decode-stage-mba-${idx}`}
                  x={colX}
                  y={466}
                  width={186}
                  height={48}
                  fill={mbAIsComm ? colors.comm : colors.compute}
                  enter={wipe({ from: "left", duration: 340 })}
                  text={s.mbA}
                  textStyle={{
                    fontSize: 11,
                    color: colors.surface,
                    align: "center",
                    verticalAlign: "middle",
                  }}
                />
              );
            })}
            {fiveStages.map((s, idx) => {
              const colX = 222 + idx * 198;
              const mbBIsComm = s.mbB.includes("IBGDA");
              return (
                <EnteredShapeText
                  key={`mbb-${s.step}`}
                  id={`decode-stage-mbb-${idx}`}
                  x={colX}
                  y={526}
                  width={186}
                  height={48}
                  fill={mbBIsComm ? colors.comm : colors.copper}
                  enter={wipe({ from: "left", duration: 340 })}
                  text={s.mbB}
                  textStyle={{
                    fontSize: 11,
                    color: colors.surface,
                    align: "center",
                    verticalAlign: "middle",
                  }}
                />
              );
            })}
          </MotionScope>
        </Step>

        <Rule x={68} y={590} width={1144} />

        {/* Legend & key engineering takeaways at bottom of panel */}
        <SemanticTag
          x={68}
          y={604}
          width={168}
          height={24}
          text="深钴蓝/赤铜 = 132 SM 计算"
          tone="compute"
        />
        <SemanticTag
          x={248}
          y={604}
          width={168}
          height={24}
          text="青金绿 = 网卡 RDMA 通信"
          tone="comm"
        />
        <Text
          x={430}
          y={602}
          width={782}
          height={44}
          style={{
            fontSize: 11,
            lineHeight: 16,
            color: colors.muted,
          }}
        >
          核心机理：任意时刻（如 Stage 3 与 Stage
          4），当一个微批次在网卡后台执行跨 144 卡的 All-to-All Dispatch 或
          Combine 时，另一个微批次恰好霸占全部 132 个 SM 执行高密度 FlashMLA 或
          MoE FP8 分组矩阵乘，实现 GPU 算力与 IB 网络双向满载。
        </Text>
      </PanelCard>
    </SlideFrame>
  );
}
