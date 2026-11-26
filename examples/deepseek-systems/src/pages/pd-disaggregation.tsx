import {
  Cell,
  P,
  Row,
  Shape,
  Table,
  Text,
  type MorphScope,
} from "@officedk/pptxdk/react";

import { PanelCard, SemanticTag } from "../components/primitives.tsx";
import { SlideFrame } from "../templates/slide-frame.tsx";
import { colors } from "../theme.ts";

const pdTableRows = [
  {
    dim: "最小部署单元与并行策略",
    prefill: "4 节点 32 卡 · EP32 + DP32 (MLA)",
    decode: "18 节点 144 卡 · EP144 + DP144 (MLA)",
    reason:
      "Prefill 序列长、单卡算力饱和；Decode 批大小受限需跨 144 卡平摊访存与专家路由",
  },
  {
    dim: "每卡常驻专家分片配置",
    prefill: "9 个路由专家 + 1 个共享专家 (共 320 槽位)",
    decode: "2 个路由专家 + 1 个共享专家 (共 288 槽位)",
    reason: "预留 32 个冗余槽位给 EPLB 复制高负载热点专家，消除尾延迟木桶效应",
  },
  {
    dim: "H800 (132 SMs) 硬件配比",
    prefill: "108 计算 SMs + 24 通信 SMs (NVLink+RDMA)",
    decode: "132 计算 SMs + 0 通信 SMs (纯 IBGDA 异步)",
    reason:
      "Decode 全面采用 GPUDirect Async 后台轮询，释放全部 132 个 SM 给矩阵与注意力计算",
  },
  {
    dim: "流水重叠与负载均衡模式",
    prefill: "双微批次乒乓交错 (約 50% 通信被掩盖) · 分层 EPLB",
    decode: "5 阶段双微批次无缝流水 · 全局 EPLB 重排",
    reason:
      "Prefill 节点数少采用节点内分层均衡减少跨机流量；Decode 跨 18 节点实施全局热专家复制",
  },
  {
    dim: "线上单节点实测有效吞吐",
    prefill: "73.7k Tokens/s/节点 (含缓存命中综合均值)",
    decode: "14.8k Tokens/s/节点 (高并发长上下文输出)",
    reason: "彻底解耦首字延迟 (TTFT) 与逐字生成延迟 (TPOT) 的硬件冲突",
  },
] as const;

/** Slide 09: Prefill-Decode disaggregation and Prefill dual-microbatch overlap. */
export function PdDisaggregationPage({
  scope,
}: {
  readonly scope: MorphScope;
}) {
  return (
    <SlideFrame
      id="pd-disaggregation"
      kicker="03 · 大规模专家并行 (EP) 与 PD 分离推理架构 / Prefill EP32 · Decode EP144"
      title="Prefill-Decode 分离推理架构：针对算力密集与访存密集实施差异化并行"
      takeaway="Prefill 32 卡高吞吐饱和算力，Decode 144 卡大 EP 释放 100% 计算 SM"
      takeawayTone="compute"
      source="数据来源：DeepSeek-V3/R1 Inference System Overview (Day 6 Open-Source Week) · EPLB 开源仓库"
    >
      {/* Top: Native comparison table between Prefill and Decode */}
      <PanelCard
        x={48}
        y={108}
        width={1184}
        height={304}
        title="Prefill 预填充阶段与 Decode 解码阶段核心系统参数对照"
        subtitle="同一套权重在两类异构部署池中采用截然不同的 EP 规模、通信算子与 SM 分配策略"
        rightBadge="PD 解耦架构"
        tone="compute"
      >
        <Table
          id="pd-comparison-table"
          columns={[196, 286, 286, 380]}
          style={{
            position: "absolute",
            left: 66,
            top: 168,
            width: 1148,
            fontFamily: "body",
            fontSize: 10.5,
            color: colors.ink,
          }}
        >
          <Row
            style={{
              height: 34,
              background: colors.steel,
              color: colors.surface,
              verticalAlign: "middle",
            }}
          >
            <Cell style={{ padding: [4, 8] }}>
              <P>架构维度</P>
            </Cell>
            <Cell
              style={{
                padding: [4, 8],
                background: colors.compute,
                color: colors.surface,
              }}
            >
              <P>Prefill 预填充池 (算力受限 Compute-Bound)</P>
            </Cell>
            <Cell
              style={{
                padding: [4, 8],
                background: colors.comm,
                color: colors.surface,
              }}
            >
              <P>Decode 解码池 (访存与通信延迟受限)</P>
            </Cell>
            <Cell style={{ padding: [4, 8] }}>
              <P>系统设计动因与硬件权衡</P>
            </Cell>
          </Row>

          {pdTableRows.map((row, idx) => {
            const rowBg = idx % 2 === 0 ? colors.surface : colors.paper;
            return (
              <Row
                key={row.dim}
                style={{
                  height: 39,
                  background: rowBg,
                  color: colors.ink,
                  verticalAlign: "middle",
                }}
              >
                <Cell
                  style={{
                    padding: [3, 8],
                    color: colors.steel,
                    border: { color: colors.line, width: 1 },
                  }}
                >
                  <P>{row.dim}</P>
                </Cell>
                <Cell
                  style={{
                    padding: [3, 8],
                    background: idx === 2 ? colors.computeSoft : rowBg,
                    color: idx === 2 ? colors.compute : colors.ink,
                    border: { color: colors.line, width: 1 },
                  }}
                >
                  <P>{row.prefill}</P>
                </Cell>
                <Cell
                  style={{
                    padding: [3, 8],
                    background: idx === 2 ? colors.commSoft : rowBg,
                    color: idx === 2 ? colors.commText : colors.ink,
                    border: { color: colors.line, width: 1 },
                  }}
                >
                  <P>{row.decode}</P>
                </Cell>
                <Cell
                  style={{
                    padding: [3, 8],
                    color: colors.muted,
                    border: { color: colors.line, width: 1 },
                  }}
                >
                  <P>{row.reason}</P>
                </Cell>
              </Row>
            );
          })}
        </Table>
      </PanelCard>

      {/* Bottom Left: H800 132 SMs Partitioning in Prefill (Morph source for Slide 10) */}
      <PanelCard
        x={48}
        y={426}
        width={580}
        height={234}
        title="H800 单芯片 132 个流多处理器 (SM) 硬件资源切分"
        subtitle="Prefill 阶段需划出 24 个 SM 驱动高吞吐 NVLink + RDMA 跨节点转发与 FP8 反量化归约"
        rightBadge="下一页平滑形变"
        tone="compute"
      >
        <Shape
          id="morph-sm-frame"
          morphKey={scope.handle("sm-frame")}
          style={{
            position: "absolute",
            left: 68,
            top: 490,
            width: 540,
            height: 102,
            background: colors.paper,
            border: { color: colors.line, width: 1 },
          }}
        />
        <Shape
          id="morph-sm-mode-tag"
          morphKey={scope.handle("sm-mode-tag")}
          style={{
            position: "absolute",
            left: 84,
            top: 502,
            width: 248,
            height: 24,
            background: colors.compute,
            fontSize: 11.5,
            color: colors.surface,
            align: "center",
            verticalAlign: "middle",
          }}
        >
          Prefill 模式 · EP32 (4 节点 32 卡)
        </Shape>
        <Text
          morphKey={scope.handle("sm-ratio-caption")}
          style={{
            position: "absolute",
            left: 344,
            top: 504,
            width: 250,
            height: 20,
            fontFamily: "metric",
            fontSize: 12,
            fontWeight: 700,
            color: colors.steel,
            align: "right",
            textWrap: "nowrap",
          }}
        >
          108 Compute SMs : 24 Comm SMs
        </Text>

        {/* 108 Compute SMs bar (width 414 of 508) */}
        <Shape
          id="morph-sm-compute-bar"
          morphKey={scope.handle("sm-compute-bar")}
          style={{
            position: "absolute",
            left: 84,
            top: 536,
            width: 410,
            height: 42,
            background: colors.compute,
            fontSize: 11.5,
            color: colors.surface,
            align: "center",
            verticalAlign: "middle",
          }}
        >
          108 计算 SMs (81.8%) · 密集执行 MLA 注意力与 MoE 矩阵乘
        </Shape>
        {/* 24 Comm SMs bar (width 94 of 508) */}
        <Shape
          id="morph-sm-comm-bar"
          morphKey={scope.handle("sm-comm-bar")}
          style={{
            position: "absolute",
            left: 498,
            top: 536,
            width: 96,
            height: 42,
            background: colors.comm,
            fontSize: 11,
            color: colors.surface,
            align: "center",
            verticalAlign: "middle",
          }}
        >
          24 通信 SMs
        </Shape>

        <Text
          style={{
            position: "absolute",
            left: 68,
            top: 604,
            width: 540,
            height: 44,
            fontSize: 11,
            lineHeight: 16,
            color: colors.muted,
          }}
        >
          注：Prefill 单批次 Token 数量大，采用 DeepEP 高吞吐内核（机内 NVLink
          聚合后再发 IB RDMA），24 个通信 SM 可将跨节点流量压缩至每目标节点最多
          1 次转发。
        </Text>
      </PanelCard>

      {/* Bottom Right: Prefill Dual-Microbatch Ping-Pong Timeline */}
      <PanelCard
        x={644}
        y={426}
        width={588}
        height={234}
        title="Prefill 双微批次乒乓交错流水线 (Dual-Microbatch Overlap)"
        subtitle="将单个预填充请求拆分为 Microbatch A 与 Microbatch B，以计算掩盖约 50% 通信开销"
        rightBadge="双流乒乓掩盖"
        tone="comm"
      >
        {/* Compute Stream Row */}
        <SemanticTag
          x={664}
          y={496}
          width={110}
          height={36}
          text="计算流 (108 SMs)"
          tone="compute"
        />
        <Shape
          style={{
            position: "absolute",
            left: 782,
            top: 496,
            width: 118,
            height: 36,
            background: colors.compute,
            fontSize: 10.5,
            color: colors.surface,
            align: "center",
            verticalAlign: "middle",
          }}
        >
          MB-A 注意力计算
        </Shape>
        <Shape
          style={{
            position: "absolute",
            left: 906,
            top: 496,
            width: 118,
            height: 36,
            background: colors.copper,
            fontSize: 10.5,
            color: colors.surface,
            align: "center",
            verticalAlign: "middle",
          }}
        >
          MB-B 注意力计算
        </Shape>
        <Shape
          style={{
            position: "absolute",
            left: 1030,
            top: 496,
            width: 90,
            height: 36,
            background: colors.compute,
            fontSize: 10.5,
            color: colors.surface,
            align: "center",
            verticalAlign: "middle",
          }}
        >
          MB-A 专家乘
        </Shape>
        <Shape
          style={{
            position: "absolute",
            left: 1126,
            top: 496,
            width: 88,
            height: 36,
            background: colors.copper,
            fontSize: 10.5,
            color: colors.surface,
            align: "center",
            verticalAlign: "middle",
          }}
        >
          MB-B 专家乘
        </Shape>

        {/* Comm Stream Row */}
        <SemanticTag
          x={664}
          y={544}
          width={110}
          height={36}
          text="通信流 (24 SMs)"
          tone="comm"
        />
        <Shape
          style={{
            position: "absolute",
            left: 782,
            top: 544,
            width: 118,
            height: 36,
            background: colors.slot,
            border: { color: colors.line, width: 1 },
            fontSize: 10,
            color: colors.muted,
            align: "center",
            verticalAlign: "middle",
          }}
        >
          等待 MB-A 注意力
        </Shape>
        <Shape
          style={{
            position: "absolute",
            left: 906,
            top: 544,
            width: 118,
            height: 36,
            background: colors.comm,
            fontSize: 10.5,
            color: colors.surface,
            align: "center",
            verticalAlign: "middle",
          }}
        >
          MB-A Dispatch
        </Shape>
        <Shape
          style={{
            position: "absolute",
            left: 1030,
            top: 544,
            width: 90,
            height: 36,
            background: colors.comm,
            fontSize: 10.5,
            color: colors.surface,
            align: "center",
            verticalAlign: "middle",
          }}
        >
          MB-B Dispatch
        </Shape>
        <Shape
          style={{
            position: "absolute",
            left: 1126,
            top: 544,
            width: 88,
            height: 36,
            background: colors.comm,
            fontSize: 10.5,
            color: colors.surface,
            align: "center",
            verticalAlign: "middle",
          }}
        >
          MB-A Combine
        </Shape>

        <Text
          style={{
            position: "absolute",
            left: 664,
            top: 596,
            width: 550,
            height: 50,
            fontSize: 11,
            lineHeight: 16,
            color: colors.muted,
          }}
        >
          当计算流执行 Microbatch B 的 MLA 注意力时，通信流并行完成 Microbatch A
          的跨节点专家分发 (Dispatch)；同理，Microbatch A 的专家 GEMM 与
          Microbatch B 的 Dispatch 相互掩盖。
        </Text>
      </PanelCard>
    </SlideFrame>
  );
}
