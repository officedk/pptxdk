import { readFile } from "node:fs/promises";

import {
  Box,
  Cell,
  Chart,
  Deck,
  defineTheme,
  Row,
  Slide,
  Table,
  Text,
} from "@react-pptx/react-pptx";
import { Suspense, use, useEffect, useState } from "react";

export type SalesSection = "region" | "product";

export type SalesRow = {
  readonly segment: string;
  readonly sales: string;
  readonly comparison: string;
  readonly result: string;
};

export type SalesDataSource = {
  load(
    section: SalesSection,
    signal?: AbortSignal,
  ): Promise<readonly SalesRow[]>;
};

export type SalesOverview = {
  readonly period: string;
  readonly totalSales: string;
  readonly growth: string;
  readonly orders: string;
  readonly averageOrderValue: string;
};

export type SalesReportOptions = {
  readonly dataSource?: SalesDataSource;
  /** 调用方可先准备概览，再与表格组件内部读取的明细组合。 */
  readonly overview: SalesOverview;
  readonly signal?: AbortSignal;
};

type SalesTableResult = {
  readonly rows: readonly SalesRow[];
  readonly usesPreviousPeriod: boolean;
};

const reportTheme = defineTheme({
  name: "季度经营简报",
  colors: {
    accent1: "#167D91",
    accent2: "#E28B49",
    accent3: "#5F9E73",
    ink: "#17324D",
    muted: "#60758A",
    surface: "#FFFFFF",
    canvas: "#F3F6F8",
    border: "#DCE4EA",
    white: "#FFFFFF",
    positive: "#277A5B",
  },
  fonts: { display: "Aptos Display", body: "Aptos" },
});

const previousPeriodRegionRows: readonly SalesRow[] = [
  {
    segment: "北方区",
    sales: "¥1,395万",
    comparison: "同比 +11%",
    result: "目标 101%",
  },
  {
    segment: "华东区",
    sales: "¥1,240万",
    comparison: "同比 +8%",
    result: "目标 100%",
  },
  {
    segment: "南方区",
    sales: "¥1,050万",
    comparison: "同比 +9%",
    result: "目标 99%",
  },
  {
    segment: "西部区",
    sales: "¥930万",
    comparison: "同比 +5%",
    result: "目标 96%",
  },
];

function isSalesRow(value: unknown): value is SalesRow {
  if (typeof value !== "object" || value === null) return false;
  const row = value as Record<string, unknown>;
  return ["segment", "sales", "comparison", "result"].every(
    (field) => typeof row[field] === "string",
  );
}

function readSalesRows(value: unknown): readonly SalesRow[] {
  if (!Array.isArray(value) || !value.every(isSalesRow)) {
    throw new Error("销售数据格式不正确");
  }
  return value;
}

function readSalesOverview(value: unknown): SalesOverview {
  if (typeof value !== "object" || value === null) {
    throw new Error("销售概览格式不正确");
  }
  const overview = value as Record<string, unknown>;
  const fields = [
    "period",
    "totalSales",
    "growth",
    "orders",
    "averageOrderValue",
  ] as const;
  if (!fields.every((field) => typeof overview[field] === "string")) {
    throw new Error("销售概览字段不完整");
  }
  return Object.fromEntries(
    fields.map((field) => [field, overview[field]]),
  ) as SalesOverview;
}

async function readBundledData(signal?: AbortSignal) {
  const content = await readFile(
    new URL("./assets/sales.json", import.meta.url),
    { encoding: "utf8", signal },
  );
  const data: unknown = JSON.parse(content);
  if (typeof data !== "object" || data === null) {
    throw new Error("销售数据文件格式不正确");
  }
  return data as Record<string, unknown>;
}

/** 入口预先读取概览，供同步内容和组件内部异步表格组合使用。 */
export async function loadBundledOverview(): Promise<SalesOverview> {
  return readSalesOverview((await readBundledData())["overview"]);
}

function createBundledSalesSource(): SalesDataSource {
  return {
    async load(section, signal) {
      const data = await readBundledData(signal);
      return readSalesRows(data[section]);
    },
  };
}

/** 创建指向调用方可控 HTTP 服务的数据源，报告本身仍负责发起异步读取。 */
export function createHttpSalesSource(baseUrl: string): SalesDataSource {
  const base = baseUrl.endsWith("/") ? baseUrl : `${baseUrl}/`;
  return {
    async load(section, signal) {
      const response = await fetch(
        new URL(section, base),
        signal === undefined ? undefined : { signal },
      );
      if (!response.ok) {
        throw new Error(`销售数据暂不可用：${section}（${response.status}）`);
      }
      return readSalesRows(await response.json());
    },
  };
}

function MetricCard({ label, value }: { label: string; value: string }) {
  return (
    <Box
      style={{
        display: "flex",
        flexDirection: "column",
        justifyContent: "center",
        gap: 5,
        width: 208,
        height: 88,
        padding: 12,
        backgroundColor: "surface",
        borderColor: "border",
        borderWidth: 1,
        borderRadius: 8,
      }}
    >
      <Text
        height={18}
        style={{ color: "muted", fontFamily: "body", fontSize: 12 }}
      >
        {label}
      </Text>
      <Text
        height={30}
        style={{
          color: "ink",
          fontFamily: "display",
          fontSize: 22,
          fontWeight: "bold",
        }}
      >
        {value}
      </Text>
    </Box>
  );
}

function SalesTableData({
  promise,
  section,
}: {
  readonly promise: Promise<SalesTableResult>;
  readonly section: SalesSection;
}) {
  const result = use(promise);
  return (
    <>
      {result.usesPreviousPeriod ? (
        <Text
          height={22}
          style={{ color: "accent2", fontFamily: "body", fontSize: 11 }}
        >
          本期数据尚未完整，以下为最近一期记录
        </Text>
      ) : null}
      <Table
        testId={`${section}-sales-table`}
        width={396}
        columns={[88, 92, 116, 100]}
        style={{ fontFamily: "body", fontSize: 11, color: "ink" }}
      >
        <Row style={{ height: 34, backgroundColor: "accent1", color: "white" }}>
          <Cell>业务板块</Cell>
          <Cell>销售额</Cell>
          <Cell>经营指标</Cell>
          <Cell>完成情况</Cell>
        </Row>
        {result.rows.map((row) => (
          <Row key={`${section}-${row.segment}`} style={{ height: 38 }}>
            <Cell>{row.segment}</Cell>
            <Cell>{row.sales}</Cell>
            <Cell>{row.comparison}</Cell>
            <Cell>{row.result}</Cell>
          </Row>
        ))}
      </Table>
    </>
  );
}

/** 每个表格实例自行发起读取并持有自己的 Promise，避免表格或构建间共享状态。 */
export function SalesTable({
  section,
  title,
  description,
  source,
  signal,
  fallbackRows,
}: {
  readonly section: SalesSection;
  readonly title: string;
  readonly description: string;
  readonly source: SalesDataSource;
  readonly signal?: AbortSignal;
  readonly fallbackRows?: readonly SalesRow[];
}) {
  const [promise, setPromise] = useState<Promise<SalesTableResult> | null>(
    null,
  );

  useEffect(() => {
    const loading = source.load(section, signal).then((rows) => ({
      rows,
      usesPreviousPeriod: false,
    }));
    const settled =
      fallbackRows === undefined
        ? loading
        : loading.catch((error: unknown) => {
            if (signal?.aborted) throw signal.reason ?? error;
            return {
              rows: fallbackRows,
              usesPreviousPeriod: true,
            };
          });
    setPromise(settled);
  }, [fallbackRows, section, signal, source]);

  return (
    <Box
      style={{
        display: "flex",
        flexDirection: "column",
        gap: 9,
        width: 430,
        height: 334,
        padding: 14,
        backgroundColor: "surface",
        borderColor: "border",
        borderWidth: 1,
        borderRadius: 8,
      }}
    >
      <Text
        height={24}
        style={{
          color: "ink",
          fontFamily: "display",
          fontSize: 16,
          fontWeight: "bold",
        }}
      >
        {title}
      </Text>
      <Text
        height={20}
        style={{ color: "muted", fontFamily: "body", fontSize: 11 }}
      >
        {description}
      </Text>
      {promise === null ? (
        <Text
          height={22}
          style={{ color: "muted", fontFamily: "body", fontSize: 11 }}
        >
          正在整理本期数据……
        </Text>
      ) : (
        <Suspense
          fallback={
            <Text
              height={22}
              style={{ color: "muted", fontFamily: "body", fontSize: 11 }}
            >
              正在整理本期数据……
            </Text>
          }
        >
          <SalesTableData promise={promise} section={section} />
        </Suspense>
      )}
    </Box>
  );
}

export function createSalesReport(options: SalesReportOptions) {
  return <SalesReport {...options} />;
}

export default function SalesReport({
  dataSource = createBundledSalesSource(),
  overview,
  signal,
}: SalesReportOptions) {
  return (
    <Deck
      width={960}
      height={540}
      title="季度销售经营报告"
      theme={reportTheme}
      style={{ color: "ink", fontFamily: "body", fontSize: 14 }}
    >
      <Slide slideId="executive-summary" background="canvas">
        <Text
          x={40}
          y={28}
          width={700}
          height={36}
          style={{
            color: "ink",
            fontFamily: "display",
            fontSize: 26,
            fontWeight: "bold",
          }}
        >
          季度销售经营报告
        </Text>
        <Text
          x={40}
          y={70}
          width={700}
          height={22}
          style={{ color: "muted", fontFamily: "body", fontSize: 13 }}
        >
          {overview.period} · 经营概览与增长趋势
        </Text>

        <Box
          testId="sales-metrics"
          style={{
            position: "absolute",
            left: 40,
            top: 112,
            display: "flex",
            gap: 16,
            width: 880,
            height: 88,
          }}
        >
          <MetricCard label="季度销售额" value={overview.totalSales} />
          <MetricCard label="同比增长" value={overview.growth} />
          <MetricCard label="有效订单" value={overview.orders} />
          <MetricCard label="平均客单价" value={overview.averageOrderValue} />
        </Box>

        <Box
          testId="sales-trend-layout"
          style={{
            position: "absolute",
            left: 40,
            top: 220,
            display: "flex",
            gap: 20,
            width: 880,
            height: 286,
          }}
        >
          <Box
            style={{
              display: "flex",
              flexDirection: "column",
              gap: 8,
              width: 574,
              height: 286,
              padding: 16,
              backgroundColor: "surface",
              borderColor: "border",
              borderWidth: 1,
              borderRadius: 8,
            }}
          >
            <Text
              height={22}
              style={{
                color: "ink",
                fontFamily: "display",
                fontSize: 15,
                fontWeight: "bold",
              }}
            >
              销售额走势（百万元）
            </Text>
            <Chart
              testId="quarterly-sales-trend"
              kind="line"
              style={{
                width: 540,
                height: 218,
                fontFamily: "body",
                fontSize: 11,
              }}
              data={{
                categories: ["一季度", "二季度", "三季度"],
                series: [
                  {
                    name: "本年",
                    values: [39.4, 44.2, 48.6],
                    color: "accent1",
                  },
                  {
                    name: "上年",
                    values: [35.1, 39.2, 43.1],
                    color: "accent2",
                  },
                ],
              }}
              legend="bottom"
              categoryAxis={{ style: { fontFamily: "body", fontSize: 10 } }}
              valueAxis={{
                min: 30,
                max: 55,
                majorUnit: 5,
                numberFormat: "0.0",
                gridlines: { color: "border", width: 1 },
                style: { fontFamily: "body", fontSize: 10 },
              }}
            />
          </Box>
          <Box
            style={{
              display: "flex",
              flexDirection: "column",
              gap: 10,
              width: 286,
              height: 286,
              padding: 16,
              backgroundColor: "surface",
              borderColor: "border",
              borderWidth: 1,
              borderRadius: 8,
            }}
          >
            <Text
              height={24}
              style={{
                color: "ink",
                fontFamily: "display",
                fontSize: 15,
                fontWeight: "bold",
              }}
            >
              本季经营结论
            </Text>
            <Text
              height={48}
              style={{
                color: "positive",
                fontFamily: "body",
                fontSize: 14,
                fontWeight: "bold",
              }}
            >
              收入连续三个季度增长，整体高于上年同期
            </Text>
            <Text
              height={54}
              style={{ color: "muted", fontFamily: "body", fontSize: 12 }}
            >
              北方区增幅领先；云端服务贡献接近总销售额的一半。
            </Text>
            <Text
              height={54}
              style={{ color: "muted", fontFamily: "body", fontSize: 12 }}
            >
              建议保持高增长区域的交付节奏，并继续提升智能终端毛利率。
            </Text>
          </Box>
        </Box>
      </Slide>

      <Slide slideId="sales-details" background="canvas">
        <Text
          x={40}
          y={28}
          width={700}
          height={36}
          style={{
            color: "ink",
            fontFamily: "display",
            fontSize: 24,
            fontWeight: "bold",
          }}
        >
          区域与产品销售明细
        </Text>
        <Text
          x={40}
          y={70}
          width={800}
          height={22}
          style={{ color: "muted", fontFamily: "body", fontSize: 13 }}
        >
          区域表现均衡，核心产品保持稳定贡献
        </Text>
        <Box
          testId="sales-tables-layout"
          style={{
            position: "absolute",
            left: 40,
            top: 112,
            display: "flex",
            gap: 20,
            width: 880,
            height: 354,
          }}
        >
          <SalesTable
            section="region"
            title="区域销售表现"
            description="销售额、同比变化与季度目标完成情况"
            source={dataSource}
            {...(signal === undefined ? {} : { signal })}
            fallbackRows={previousPeriodRegionRows}
          />
          <SalesTable
            section="product"
            title="重点产品表现"
            description="产品线收入、订单规模与毛利水平"
            source={dataSource}
            {...(signal === undefined ? {} : { signal })}
          />
        </Box>
        <Text
          x={40}
          y={490}
          width={880}
          height={20}
          style={{ color: "muted", fontFamily: "body", fontSize: 10 }}
        >
          数据口径：截至 2026 年 9 月 30
          日；本期数据暂缺时参考最近一期完整记录。
        </Text>
      </Slide>
    </Deck>
  );
}
