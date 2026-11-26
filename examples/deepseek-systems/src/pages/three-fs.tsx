import { Chart, Shape, Text } from "@officedk/pptxdk/react";

import {
  MetricCard,
  PanelCard,
  Rule,
  SemanticTag,
} from "../components/primitives.tsx";
import { SlideFrame } from "../templates/slide-frame.tsx";
import { colors } from "../theme.ts";

const storageCategories = [
  "单客户端 KVCache 读取峰值 (GiB/s)",
  "Smallpond GraySort 排序 (GiB/s)",
  "180 节点单节点平均读取 (GiB/s)",
  "单存储节点 2x200G 网卡上限 (GiB/s)",
] as const;

const storageMeasured = [40.0, 62.5, 37.5, 46.5] as const;

const archLayers = [
  {
    title: "上层计算与数据处理生态 (Smallpond / 训练与推理引擎)",
    badge: "DuckDB + Ray 分布式引擎",
    tone: "compute" as const,
    detail:
      "基于 DuckDB + 3FS 构建轻量级 PB 级数据处理框架 Smallpond（无常驻服务）；支持训练样本无缓存随机洗牌读取、并行 Checkpoint 写入与推理 KVCache 落盘卸载。",
  },
  {
    title: "零拷贝用户态通信层 (hf3fs_usrbio & FUSE 挂载)",
    badge: "共享内存环 + RDMA 直传",
    tone: "comm" as const,
    detail:
      "针对 Python 大并发 I/O 绕过内核态 FUSE 拷贝开销，通过共享内存环形缓冲区 (Ring Buffer) 与 InfiniBand RDMA Write/Read 直接打通客户端内存与远端 NVMe。",
  },
  {
    title: "元数据与一致性控制面 (Cluster Manager + Meta Service)",
    badge: "CRAQ 链式复制 + FoundationDB",
    tone: "copper" as const,
    detail:
      "采用链式复制与分摊查询 (CRAQ) 实现写全链、读任意节点的强一致性 (Write-All-Read-Any)，目录树与分块位置布局持久化于分布式事务数据库 FoundationDB。",
  },
  {
    title: "底层全闪存存储节点集群 (180 Storage Nodes)",
    badge: "2,880 块 14TiB NVMe SSD",
    tone: "steel" as const,
    detail:
      "每台存储节点配备 16 块 14TiB PCIe NVMe SSD 与 2 张 200Gbps InfiniBand 网卡，将块随机请求打散至数千张物理盘，彻底消除局部 I/O 热点。",
  },
] as const;

/** Slide 08: 3FS Fire-Flyer File System and Smallpond data processing framework. */
export function ThreeFsPage() {
  return (
    <SlideFrame
      id="three-fs"
      kicker="02 · 存算分离全闪并行文件系统 / 3FS · CRAQ · Smallpond"
      title="3FS 分布式文件系统与 Smallpond：聚合 6.6 TiB/s 随机读与推理 KV 卸载"
      takeaway="180 存储节点聚合读取达 6.6 TiB/s，110.5 TiB GraySort 仅耗时 30 分 14 秒"
      takeawayTone="comm"
      source="数据来源：3FS (Fire-Flyer File System) 与 Smallpond 开源仓库实测报告 (Day 5 Open-Source Week)"
    >
      {/* Left column: 3FS 4-layer architecture stack */}
      <PanelCard
        x={48}
        y={108}
        width={616}
        height={552}
        title="3FS 存算分离全闪存架构栈"
        subtitle="摒弃传统基于局部性缓存的文件系统假设，专为 AI 随机样本读取与 KVCache 大吞吐卸载设计"
        rightBadge="RDMA + NVMe 原生"
        tone="comm"
      >
        {archLayers.map((layer, idx) => {
          const y = 172 + idx * 118;
          return (
            <Shape
              key={layer.title}
              style={{
                position: "absolute",
                left: 68,
                top: y,
                width: 576,
                height: 106,
                background: colors.paper,
                border: { color: colors.line, width: 1 },
              }}
            />
          );
        })}
        {archLayers.map((layer, idx) => {
          const y = 172 + idx * 118;
          const accent =
            layer.tone === "compute"
              ? colors.compute
              : layer.tone === "comm"
                ? colors.comm
                : layer.tone === "copper"
                  ? colors.copper
                  : colors.steel;
          return (
            <Shape
              key={`accent-${idx}`}
              style={{
                position: "absolute",
                left: 68,
                top: y,
                width: 4,
                height: 106,
                background: accent,
              }}
            />
          );
        })}
        {archLayers.map((layer, idx) => {
          const y = 172 + idx * 118;
          return (
            <Text
              key={`title-${idx}`}
              style={{
                position: "absolute",
                left: 84,
                top: y + 10,
                width: 368,
                height: 22,
                fontSize: 12.5,
                color: colors.ink,
                textWrap: "nowrap",
              }}
            >
              {layer.title}
            </Text>
          );
        })}
        {archLayers.map((layer, idx) => {
          const y = 172 + idx * 118;
          return (
            <SemanticTag
              key={`tag-${idx}`}
              x={460}
              y={y + 10}
              width={170}
              height={22}
              text={layer.badge}
              tone={layer.tone}
            />
          );
        })}
        {archLayers.map((layer, idx) => {
          const y = 172 + idx * 118;
          return (
            <Text
              key={`detail-${idx}`}
              style={{
                position: "absolute",
                left: 84,
                top: y + 38,
                width: 544,
                height: 60,
                fontSize: 11,
                lineHeight: 16,
                color: colors.muted,
              }}
            >
              {layer.detail}
            </Text>
          );
        })}
      </PanelCard>

      {/* Right top: 3 KPI metric cards */}
      <MetricCard
        x={680}
        y={108}
        width={174}
        height={92}
        label="集群聚合读取吞吐"
        value="6.6 TiB/s"
        caption="180 存储节点 + 500+ 客户端"
        tone="comm"
      />
      <MetricCard
        x={869}
        y={108}
        width={174}
        height={92}
        label="GraySort 排序速率"
        value="3.66 TiB/min"
        caption="110.5 TiB 数据耗时 30m14s"
        tone="copper"
      />
      <MetricCard
        x={1058}
        y={108}
        width={174}
        height={92}
        label="KVCache 客户端峰值"
        value="40+ GiB/s"
        caption="大幅替代高成本 DRAM 缓存"
        tone="compute"
      />

      {/* Right middle: Horizontal bar chart of per-node/per-client storage bandwidth */}
      <PanelCard
        x={680}
        y={214}
        width={552}
        height={320}
        title="3FS 关键链路单节点/单客户端实测吞吐速率 (GiB/s)"
        subtitle="在携载训练后台流量压力的同时，单存储节点读取速率逼近 2x200Gbps 网卡物理上限"
        rightBadge="线速 I/O 利用率"
        tone="comm"
      >
        <Chart
          id="three-fs-chart"
          legend={false}
          axes={{
            category: {
              kind: "category",
              line: { color: colors.steel, width: 1 },
              style: { fontFamily: "body", fontSize: 10.5, color: colors.ink },
            },
            value: {
              kind: "value",
              min: 0,
              max: 70,
              majorUnit: 14,
              numberFormat: "0.0",
              gridlines: { color: colors.line, width: 1 },
              line: false,
              style: { fontFamily: "body", fontSize: 10, color: colors.muted },
            },
          }}
          groups={[
            {
              kind: "bar",
              data: {
                kind: "shared-categories",
                categories: storageCategories,
              },
              series: [
                {
                  name: "实测持续吞吐 / 硬件参考线 (GiB/s)",
                  values: storageMeasured,
                  color: colors.comm,
                  points: {
                    0: { color: colors.compute, labelColor: colors.compute },
                    1: { color: colors.copper, labelColor: colors.copperText },
                    2: { color: colors.comm, labelColor: colors.commText },
                    3: { color: colors.baseline, labelColor: colors.steel },
                  },
                },
              ],
              dataLabels: {
                value: true,
                position: "outsideEnd",
                numberFormat: "0.0",
                style: { fontFamily: "body", fontSize: 10, color: colors.ink },
              },
            },
          ]}
          style={{
            position: "absolute",
            left: 696,
            top: 274,
            width: 520,
            height: 212,
            fontFamily: "body",
            fontSize: 10.5,
            color: colors.ink,
          }}
        />
        <Rule x={696} y={492} width={520} />
        <Text
          style={{
            position: "absolute",
            left: 696,
            top: 498,
            width: 520,
            height: 28,
            fontSize: 10.5,
            color: colors.muted,
            textWrap: "nowrap",
          }}
        >
          注：3.66 TiB/min 折合每秒 62.5 GiB/s 全集群端到端洗牌排序吞吐（25
          存储节点 + 50 计算节点）。
        </Text>
      </PanelCard>

      {/* Right bottom: Inference KVCache on 3FS cost-benefit summary */}
      <Shape
        style={{
          position: "absolute",
          left: 680,
          top: 548,
          width: 552,
          height: 112,
          background: colors.surface,
          border: { color: colors.line, width: 1 },
        }}
      />
      <Shape
        style={{
          position: "absolute",
          left: 680,
          top: 548,
          width: 4,
          height: 112,
          background: colors.copper,
        }}
      />
      <Text
        style={{
          position: "absolute",
          left: 698,
          top: 558,
          width: 520,
          height: 22,
          fontSize: 13,
          color: colors.steel,
          textWrap: "nowrap",
        }}
      >
        为什么推理服务将 3FS 作为多轮上下文 KVCache 共享层？
      </Text>
      <Text
        style={{
          position: "absolute",
          left: 698,
          top: 584,
          width: 520,
          height: 64,
          fontSize: 11,
          lineHeight: 16,
          color: colors.muted,
        }}
      >
        在长文本与多轮对话场景下，单机 HBM 与主机 DRAM 容量受限且造价高昂。3FS
        凭借 CRAQ 读任意副本能力与每客户端 40+ GiB/s 读取带宽，使所有 Prefill
        节点能够跨机零拷贝复用历史 MLA 压缩向量（线上实测硬盘缓存命中率达
        56.3%）。
      </Text>
    </SlideFrame>
  );
}
