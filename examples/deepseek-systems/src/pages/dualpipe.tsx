import {
  Cell,
  MotionScope,
  P,
  Row,
  Shape,
  Step,
  Table,
  Text,
  fade,
  useMotionScope,
  wipe,
} from "@react-pptx/react-pptx";

import {
  EnteredShapeText,
  MetricCard,
  PanelCard,
  SemanticTag,
} from "../components/primitives.tsx";
import { SlideFrame } from "../templates/slide-frame.tsx";
import { colors } from "../theme.ts";

const fwdLaneBlocks = [
  { label: "F0", tone: "compute", w: 50 },
  { label: "F1", tone: "compute", w: 50 },
  { label: "F&B (双向重叠)", tone: "copper", w: 124 },
  { label: "F&B (双向重叠)", tone: "copper", w: 124 },
  { label: "B1", tone: "comm", w: 52 },
  { label: "W1", tone: "steel", w: 48 },
] as const;

const revLaneBlocks = [
  { label: "F'0", tone: "compute", w: 50 },
  { label: "F&B (双向重叠)", tone: "copper", w: 124 },
  { label: "F&B (双向重叠)", tone: "copper", w: 124 },
  { label: "B'0", tone: "comm", w: 52 },
  { label: "B'1", tone: "comm", w: 52 },
  { label: "W'0", tone: "steel", w: 46 },
] as const;

const overlapSubStages = [
  {
    lane: "计算流 (208 SMs)",
    items: [
      { label: "前向 MLP / 共享专家", x: 840, w: 112, fill: colors.compute },
      { label: "后向 Attention 梯度", x: 958, w: 112, fill: colors.compute },
      { label: "前向 MLA 注意力", x: 1076, w: 124, fill: colors.compute },
    ],
  },
  {
    lane: "通信流 (24 SMs)",
    items: [
      { label: "All-to-All Combine (B)", x: 840, w: 112, fill: colors.comm },
      { label: "All-to-All Dispatch (F)", x: 958, w: 112, fill: colors.comm },
      { label: "All-to-All Dispatch (B)", x: 1076, w: 124, fill: colors.comm },
    ],
  },
] as const;

/** Slide 06: DualPipe bidirectional pipeline parallelism & F&B overlap. */
export function DualPipePage() {
  const overlapScope = useMotionScope();

  let fwdCursor = 196;
  let revCursor = 196;

  return (
    <SlideFrame
      id="dualpipe"
      kicker="02 · 跨节点训练并行与通信隐藏 / DualPipe · EP64 · 零张量并行"
      title="DualPipe 双向流水线并行：以对称微批次交错隐藏跨节点专家通信"
      takeaway="2,048 卡 H800 集群单日吞吐 272.8B Tokens，零 TP 消除高昂跨卡同步"
      takeawayTone="comm"
      source="数据来源：DeepSeek-V3 Technical Report Section 3.2 · DualPipe 开源仓库 (PP16 · EP64 · ZeRO-1 DP)"
    >
      {/* Top Left: DualPipe Bidirectional Schedule */}
      <PanelCard
        x={48}
        y={108}
        width={628}
        height={252}
        title="DualPipe 对称双向流水线微批次排布示意"
        subtitle="从流水线两端同时喂入正向与镜像微批次，将流水线气泡压缩至 (PP/2 - 1)(F&B + B - 3W)"
        rightBadge="双向对称交错"
        tone="compute"
      >
        {/* Forward stream row */}
        <Text
          x={68}
          y={180}
          width={124}
          style={{
            padding: 0,
            fontSize: 11,
            color: colors.steel,
            textWrap: "nowrap",
          }}
        >
          正向数据流 (Rank 0→P)
        </Text>
        {fwdLaneBlocks.map((b, idx) => {
          const x = fwdCursor;
          fwdCursor += b.w + 6;
          const fill =
            b.tone === "copper"
              ? colors.copper
              : b.tone === "comm"
                ? colors.comm
                : b.tone === "steel"
                  ? colors.steel
                  : colors.compute;
          return (
            <Shape
              key={idx}
              x={x}
              y={172}
              width={b.w}
              height={34}
              fill={fill}
              style={{ padding: [0, 2] }}
              text={b.label}
              textStyle={{
                fontSize: 10.5,
                color: colors.surface,
                align: "center",
                verticalAlign: "middle",
              }}
            />
          );
        })}

        {/* Reverse stream row */}
        <Text
          x={68}
          y={226}
          width={124}
          style={{
            padding: 0,
            fontSize: 11,
            color: colors.steel,
            textWrap: "nowrap",
          }}
        >
          镜像数据流 (Rank P→0)
        </Text>
        {revLaneBlocks.map((b, idx) => {
          const x = revCursor;
          revCursor += b.w + 6;
          const fill =
            b.tone === "copper"
              ? colors.copper
              : b.tone === "comm"
                ? colors.comm
                : b.tone === "steel"
                  ? colors.steel
                  : colors.compute;
          return (
            <Shape
              key={idx}
              x={x}
              y={218}
              width={b.w}
              height={34}
              fill={fill}
              style={{ padding: [0, 2] }}
              text={b.label}
              textStyle={{
                fontSize: 10.5,
                color: colors.surface,
                align: "center",
                verticalAlign: "middle",
              }}
            />
          );
        })}

        {/* Explanation of zero-bubble B = E + W split */}
        <Shape
          x={68}
          y={268}
          width={588}
          height={76}
          fill={colors.paper}
          style={{ borderColor: colors.line, borderWidth: 1 }}
        />
        <SemanticTag
          x={82}
          y={280}
          width={208}
          height={22}
          text="B 拆分为输入梯度 E 与权重梯度 W"
          tone="copper"
        />
        <Text
          x={300}
          y={283}
          width={344}
          style={{
            padding: 0,
            fontSize: 11,
            color: colors.steel,
            textWrap: "nowrap",
          }}
        >
          权重梯度 W 可延后异步填充空隙，使非重叠气泡项减去 3W
        </Text>
        <Text
          x={82}
          y={312}
          width={560}
          style={{
            padding: 0,
            fontSize: 11,
            color: colors.muted,
            textWrap: "nowrap",
          }}
        >
          并行组合：16 路双向流水线 (PP16) + 跨 8 节点 64 路专家并行 (EP64) +
          ZeRO-1 数据并行，全程 0 张量并行 (TP=0)
        </Text>
      </PanelCard>

      {/* Top Right: Fine-grained F&B Compute-Communication Overlap with Step + MotionScope */}
      <PanelCard
        x={692}
        y={108}
        width={540}
        height={252}
        title="F&B 阶段细粒度计算与跨节点通信全重叠"
        subtitle="将一对前向 (F) 与后向 (B) 块切分为 4 个子阶段交叉流水，完全掩盖 EP64 All-to-All"
        rightBadge="点击步进演示重叠"
        tone="comm"
      >
        <Step>
          <MotionScope scope={overlapScope} mode="stagger" each={110}>
            {overlapSubStages.map((row, rIdx) => (
              <EnteredShapeText
                key={row.lane}
                id={`dualpipe-lane-${rIdx}`}
                x={710}
                y={174 + rIdx * 48}
                width={122}
                height={36}
                fill={rIdx === 0 ? colors.computeSoft : colors.commSoft}
                shapeStyle={{
                  borderColor: rIdx === 0 ? colors.compute : colors.comm,
                  borderWidth: 1,
                }}
                enter={fade({ duration: 320 })}
                text={row.lane}
                textStyle={{
                  fontSize: 11,
                  color: rIdx === 0 ? colors.compute : colors.commText,
                  align: "center",
                  verticalAlign: "middle",
                }}
              />
            ))}
            {overlapSubStages.flatMap((row, rIdx) =>
              row.items.map((item, iIdx) => (
                <EnteredShapeText
                  key={`${rIdx}-${iIdx}`}
                  id={`dualpipe-overlap-${rIdx}-${iIdx}`}
                  x={item.x}
                  y={174 + rIdx * 48}
                  width={item.w}
                  height={36}
                  fill={item.fill}
                  enter={wipe({ from: "left", duration: 340 })}
                  text={item.label}
                  textStyle={{
                    fontSize: 10,
                    color: colors.surface,
                    align: "center",
                    verticalAlign: "middle",
                  }}
                />
              )),
            )}
          </MotionScope>
        </Step>

        {/* Dedicated SM Warp specialization callout */}
        <Shape
          x={710}
          y={272}
          width={504}
          height={72}
          fill={colors.paper}
          style={{ borderColor: colors.line, borderWidth: 1 }}
        />
        <Text
          x={724}
          y={282}
          width={476}
          height={20}
          style={{
            fontSize: 12,
            color: colors.commText,
            textWrap: "nowrap",
          }}
        >
          20 个通信 SM 专用化 (Warp Specialization)：
        </Text>
        <Text
          x={724}
          y={304}
          width={476}
          height={34}
          style={{
            fontSize: 10.5,
            lineHeight: 15,
            color: colors.muted,
          }}
        >
          仅划拨 20 个 SM 兼管 IB 跨节点转发、NVLink 聚合与反量化归约，其余 112
          个 SM 满载运行矩阵计算，使跨节点 All-to-All 在计算阴影下等效零开销。
        </Text>
      </PanelCard>

      {/* Bottom Left: Scheduling method comparison table */}
      <PanelCard
        x={48}
        y={374}
        width={828}
        height={286}
        title="流水线并行算法理论开销对比 (PP 为流水阶段数，F/B/W 分别为前向/输入后向/权重后向耗时)"
        subtitle="DualPipe 与 V 形紧凑变体 DualPipeV 在大幅削减流水线气泡的同时控制显存冗余"
        rightBadge="调度复杂度对照"
        tone="steel"
      >
        <Table
          id="dualpipe-table"
          x={66}
          y={438}
          width={792}
          columns={[124, 280, 128, 132, 128]}
          style={{
            fontFamily: "body",
            fontSize: 11.5,
            color: colors.ink,
          }}
        >
          <Row
            style={{
              height: 36,
              backgroundColor: colors.steel,
              color: colors.surface,
              verticalAlign: "middle",
            }}
          >
            <Cell style={{ padding: [6, 8] }}>
              <P>调度机制</P>
            </Cell>
            <Cell style={{ padding: [6, 8] }}>
              <P>流水线气泡时间 (Bubble)</P>
            </Cell>
            <Cell style={{ padding: [6, 8] }}>
              <P style={{ align: "right" }}>单卡参数副本</P>
            </Cell>
            <Cell style={{ padding: [6, 8] }}>
              <P style={{ align: "right" }}>峰值激活显存</P>
            </Cell>
            <Cell style={{ padding: [6, 8] }}>
              <P style={{ align: "right" }}>所需设备数</P>
            </Cell>
          </Row>

          <Row
            style={{
              height: 38,
              backgroundColor: colors.surface,
              color: colors.ink,
              verticalAlign: "middle",
            }}
          >
            <Cell
              style={{
                padding: [6, 8],
                borderColor: colors.line,
                borderWidth: 1,
                color: colors.steel,
              }}
            >
              <P>1F1B 基准</P>
            </Cell>
            <Cell
              style={{
                padding: [6, 8],
                borderColor: colors.line,
                borderWidth: 1,
                fontFamily: "metric",
                color: colors.muted,
              }}
            >
              <P>(PP - 1)(F + B)</P>
            </Cell>
            <Cell
              style={{
                padding: [6, 8],
                borderColor: colors.line,
                borderWidth: 1,
                fontFamily: "metric",
              }}
            >
              <P style={{ align: "right" }}>1x</P>
            </Cell>
            <Cell
              style={{
                padding: [6, 8],
                borderColor: colors.line,
                borderWidth: 1,
                fontFamily: "metric",
              }}
            >
              <P style={{ align: "right" }}>PP</P>
            </Cell>
            <Cell
              style={{
                padding: [6, 8],
                borderColor: colors.line,
                borderWidth: 1,
                fontFamily: "metric",
              }}
            >
              <P style={{ align: "right" }}>PP</P>
            </Cell>
          </Row>

          <Row
            style={{
              height: 38,
              backgroundColor: colors.paper,
              color: colors.ink,
              verticalAlign: "middle",
            }}
          >
            <Cell
              style={{
                padding: [6, 8],
                borderColor: colors.line,
                borderWidth: 1,
                color: colors.steel,
              }}
            >
              <P>ZB1P 零气泡</P>
            </Cell>
            <Cell
              style={{
                padding: [6, 8],
                borderColor: colors.line,
                borderWidth: 1,
                fontFamily: "metric",
                color: colors.muted,
              }}
            >
              <P>(PP - 1)(F + B - 2W)</P>
            </Cell>
            <Cell
              style={{
                padding: [6, 8],
                borderColor: colors.line,
                borderWidth: 1,
                fontFamily: "metric",
              }}
            >
              <P style={{ align: "right" }}>1x</P>
            </Cell>
            <Cell
              style={{
                padding: [6, 8],
                borderColor: colors.line,
                borderWidth: 1,
                fontFamily: "metric",
              }}
            >
              <P style={{ align: "right" }}>PP</P>
            </Cell>
            <Cell
              style={{
                padding: [6, 8],
                borderColor: colors.line,
                borderWidth: 1,
                fontFamily: "metric",
              }}
            >
              <P style={{ align: "right" }}>PP</P>
            </Cell>
          </Row>

          <Row
            style={{
              height: 42,
              backgroundColor: colors.computeSoft,
              color: colors.compute,
              verticalAlign: "middle",
            }}
          >
            <Cell
              style={{
                padding: [6, 8],
                borderColor: colors.line,
                borderWidth: 1,
              }}
            >
              <P>DualPipe (V3)</P>
            </Cell>
            <Cell
              style={{
                padding: [6, 8],
                borderColor: colors.line,
                borderWidth: 1,
                fontFamily: "metric",
                fontWeight: 700,
              }}
            >
              <P>(PP/2 - 1)(F&amp;B + B - 3W)</P>
            </Cell>
            <Cell
              style={{
                padding: [6, 8],
                borderColor: colors.line,
                borderWidth: 1,
                fontFamily: "metric",
                fontWeight: 700,
              }}
            >
              <P style={{ align: "right" }}>2x</P>
            </Cell>
            <Cell
              style={{
                padding: [6, 8],
                borderColor: colors.line,
                borderWidth: 1,
                fontFamily: "metric",
                fontWeight: 700,
              }}
            >
              <P style={{ align: "right" }}>PP + 1</P>
            </Cell>
            <Cell
              style={{
                padding: [6, 8],
                borderColor: colors.line,
                borderWidth: 1,
                fontFamily: "metric",
                fontWeight: 700,
              }}
            >
              <P style={{ align: "right" }}>PP</P>
            </Cell>
          </Row>

          <Row
            style={{
              height: 42,
              backgroundColor: colors.commSoft,
              color: colors.commText,
              verticalAlign: "middle",
            }}
          >
            <Cell
              style={{
                padding: [6, 8],
                borderColor: colors.line,
                borderWidth: 1,
              }}
            >
              <P>DualPipeV (V 形)</P>
            </Cell>
            <Cell
              style={{
                padding: [6, 8],
                borderColor: colors.line,
                borderWidth: 1,
                fontFamily: "metric",
                fontWeight: 700,
              }}
            >
              <P>(PP/2 - 1)(F&amp;B + B - 3W)</P>
            </Cell>
            <Cell
              style={{
                padding: [6, 8],
                borderColor: colors.line,
                borderWidth: 1,
                fontFamily: "metric",
                fontWeight: 700,
              }}
            >
              <P style={{ align: "right" }}>1x</P>
            </Cell>
            <Cell
              style={{
                padding: [6, 8],
                borderColor: colors.line,
                borderWidth: 1,
                fontFamily: "metric",
                fontWeight: 700,
              }}
            >
              <P style={{ align: "right" }}>PP + 1</P>
            </Cell>
            <Cell
              style={{
                padding: [6, 8],
                borderColor: colors.line,
                borderWidth: 1,
                fontFamily: "metric",
                fontWeight: 700,
              }}
            >
              <P style={{ align: "right" }}>PP / 2</P>
            </Cell>
          </Row>
        </Table>
      </PanelCard>

      {/* Bottom Right: 2,048 H800 cluster real throughput metrics */}
      <MetricCard
        x={892}
        y={374}
        width={340}
        height={90}
        label="2,048 卡 H800 集群单日总训练吞吐"
        value="272.80B Tokens/d"
        caption="4K 序列长度实测 · 每节点日均吞吐 1.07B Tokens"
        tone="copper"
      />
      <MetricCard
        x={892}
        y={472}
        width={164}
        height={90}
        label="Non-Causal MFU"
        value="43.73%"
        caption="基于 BF16 峰值算力口径"
        tone="compute"
      />
      <MetricCard
        x={1068}
        y={472}
        width={164}
        height={90}
        label="Causal MFU"
        value="38.94%"
        caption="扣除因果掩码有效算力"
        tone="comm"
      />
      <Shape
        x={892}
        y={570}
        width={340}
        height={90}
        fill={colors.surface}
        style={{ borderColor: colors.line, borderWidth: 1 }}
      />
      <Shape x={892} y={570} width={4} height={90} fill={colors.steel} />
      <Text
        x={908}
        y={580}
        width={312}
        height={20}
        style={{
          fontSize: 12.5,
          color: colors.steel,
          textWrap: "nowrap",
        }}
      >
        为什么训练过程能做到零张量并行 (TP = 0)？
      </Text>
      <Text
        x={908}
        y={602}
        width={312}
        height={48}
        style={{
          fontSize: 11,
          lineHeight: 15.5,
          color: colors.muted,
        }}
      >
        MLA 削减激活体积 + 细粒度 MoE 降低单卡计算密度 + DualPipe 全隐藏 EP64
        通信，无需切分单层矩阵即可装入显存。
      </Text>
    </SlideFrame>
  );
}
