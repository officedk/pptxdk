import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdir, writeFile } from "node:fs/promises";
import { createServer, type ServerResponse } from "node:http";
import { join } from "node:path";

import { BuildError, build } from "@react-pptx/react-pptx";
import { buildFromSource } from "@react-pptx/react-pptx/node";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
import { describe, it } from "vitest";

import {
  createHttpSalesSource,
  createSalesReport,
  type SalesOverview,
  type SalesRow,
} from "../src/sales-report.tsx";

const testOverview: SalesOverview = {
  period: "2026 年第三季度",
  totalSales: "¥4,860 万",
  growth: "+12.8%",
  orders: "8,420 单",
  averageOrderValue: "¥5,771",
};

type SalesServer = {
  readonly origin: string;
  readonly allRequested: Promise<void>;
  respond(path: string, rows: readonly SalesRow[], status?: number): void;
  responseClosed(path: string): Promise<void>;
  close(): Promise<void>;
};

async function startSalesServer(
  expectedPaths: readonly string[],
): Promise<SalesServer> {
  const requests = new Map<string, ServerResponse>();
  const closed = new Map<string, PromiseWithResolvers<void>>();
  const arrived = Promise.withResolvers<void>();
  const server = createServer((request, response) => {
    const path = request.url;
    if (path === undefined || !expectedPaths.includes(path)) {
      response.writeHead(404).end();
      return;
    }
    requests.set(path, response);
    const signal = Promise.withResolvers<void>();
    closed.set(path, signal);
    response.once("close", signal.resolve);
    if (expectedPaths.every((expected) => requests.has(expected))) {
      arrived.resolve();
    }
  });
  await new Promise<void>((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", resolve);
  });
  const address = server.address();
  assert.ok(address && typeof address !== "string");

  return {
    origin: `http://127.0.0.1:${address.port}`,
    allRequested: arrived.promise,
    respond(path, rows, status = 200) {
      const response = requests.get(path);
      assert.ok(response, `预期收到 ${path} 请求`);
      response.writeHead(status, { "Content-Type": "application/json" });
      response.end(JSON.stringify(rows));
    },
    responseClosed(path) {
      const signal = closed.get(path);
      assert.ok(signal, `预期收到 ${path} 请求`);
      return signal.promise;
    },
    close() {
      return new Promise<void>((resolve, reject) => {
        server.close((error) =>
          error === undefined ? resolve() : reject(error),
        );
      });
    },
  };
}

function tableValues(
  file: Awaited<ReturnType<typeof validatePptx>>,
  label: string,
): readonly string[] {
  const table = file.slides[1]?.objects.find(({ table }) =>
    table?.cells.some((row) => row.includes(label)),
  );
  assert.ok(table?.table);
  return table.table.cells.flat();
}

describe("销售报告示例", () => {
  it("从正式源码生成并独立读取包含表格、趋势图与主题的报告", async () => {
    const result = await buildFromSource(
      new URL("../src/index.ts", import.meta.url),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(file.presentation.slideCount, 2);

    const firstSlide = file.slides[0];
    const detailSlide = file.slides[1];
    assert.ok(firstSlide);
    assert.ok(detailSlide);
    assert.equal(result.report.evidence, "generation");
    assert.ok(
      firstSlide.objects.some(({ text }) =>
        text?.text.includes("2026 年第三季度"),
      ),
    );
    assert.ok(
      firstSlide.objects.some(({ text }) => text?.text.includes("¥4,860 万")),
    );
    const trend = firstSlide.objects.find(({ chart }) =>
      chart?.types.includes("lineChart"),
    );
    assert.ok(trend?.chart);
    assert.deepEqual(
      trend.chart.series.map(({ name, values }) => [name, values]),
      [
        ["本年", [39.4, 44.2, 48.6]],
        ["上年", [35.1, 39.2, 43.1]],
      ],
    );
    assert.deepEqual(trend.chart.embeddedWorkbook?.series, trend.chart.series);
    assert.ok(
      result.report.slides[0]?.layouts?.some(
        ({ testId, box }) =>
          testId === "sales-trend-layout" &&
          box.width === 880 &&
          box.height === 286,
      ),
    );

    const tables = detailSlide.objects.filter(
      ({ table }) => table !== undefined,
    );
    assert.equal(tables.length, 2);
    assert.ok(tables[0]?.table?.cells.some((row) => row.includes("北方区")));
    assert.ok(tables[1]?.table?.cells.some((row) => row.includes("云端服务")));
    assert.ok(
      file.slides.every((slide) =>
        slide.objects.every(
          ({ text }) => !text?.text.includes("正在整理本期数据"),
        ),
      ),
    );

    const theme = unzipSync(result.pptx)["ppt/theme/theme1.xml"];
    assert.ok(theme);
    assert.match(strFromU8(theme), /<a:srgbClr val="167D91"\/>/u);

    const directory = join(import.meta.dirname, "../test-results/sales-report");
    await mkdir(directory, { recursive: true });
    await writeFile(join(directory, "sales-report.pptx"), result.pptx);
    await writeFile(
      join(directory, "sales-report-report.json"),
      `${JSON.stringify(result.report, null, 2)}\n`,
    );
    await writeFile(
      join(directory, "sales-report-reader.json"),
      `${JSON.stringify(file, null, 2)}\n`,
    );
    await writeFile(
      join(directory, "sales-report.sha256"),
      `${createHash("sha256").update(result.pptx).digest("hex")}  sales-report.pptx\n`,
    );
  });

  it("编译后的示例仍从包内数据资源读取两张销售表", async () => {
    const result = await buildFromSource(
      new URL("../dist/index.js", import.meta.url),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.ok(tableValues(file, "北方区").includes("北方区"));
    assert.ok(tableValues(file, "云端服务").includes("云端服务"));
  });

  it("并发报告各自读取两个表格实例并保留入口预取概览", async () => {
    const server = await startSalesServer([
      "/first/region",
      "/first/product",
      "/second/region",
      "/second/product",
    ]);
    try {
      const firstBuild = build(
        createSalesReport({
          dataSource: createHttpSalesSource(`${server.origin}/first`),
          overview: {
            period: "甲报告预取期间",
            totalSales: "甲预取 ¥4,860 万",
            growth: "+12.8%",
            orders: "8,420 单",
            averageOrderValue: "¥5,771",
          },
        }),
      );
      const secondBuild = build(
        createSalesReport({
          dataSource: createHttpSalesSource(`${server.origin}/second`),
          overview: {
            period: "乙报告预取期间",
            totalSales: "乙预取 ¥2,740 万",
            growth: "+8.1%",
            orders: "5,120 单",
            averageOrderValue: "¥5,352",
          },
        }),
      );
      let firstStatus = "pending";
      void firstBuild.then(
        () => {
          firstStatus = "done";
        },
        () => {
          firstStatus = "failed";
        },
      );

      await server.allRequested;
      server.respond("/first/product", [
        {
          segment: "甲云服务",
          sales: "¥2,120万",
          comparison: "订单 1,900",
          result: "毛利率 53%",
        },
      ]);
      await server.responseClosed("/first/product");
      server.respond("/second/region", [
        {
          segment: "乙东部区",
          sales: "¥1,460万",
          comparison: "同比 +9%",
          result: "目标 102%",
        },
      ]);
      server.respond("/second/product", [
        {
          segment: "乙终端产品",
          sales: "¥1,280万",
          comparison: "订单 2,100",
          result: "毛利率 41%",
        },
      ]);

      const secondResult = await secondBuild;
      const secondFile = await validatePptx(secondResult.pptx);
      assert.deepEqual(secondFile.errors, []);
      assert.ok(
        secondFile.slides[0]?.objects.some(({ text }) =>
          text?.text.includes("乙预取 ¥2,740 万"),
        ),
      );
      assert.ok(tableValues(secondFile, "乙东部区").includes("乙东部区"));
      assert.ok(tableValues(secondFile, "乙终端产品").includes("乙终端产品"));
      assert.equal(firstStatus, "pending");

      server.respond("/first/region", [
        {
          segment: "甲北方区",
          sales: "¥1,510万",
          comparison: "同比 +17%",
          result: "目标 105%",
        },
      ]);
      const firstResult = await firstBuild;
      const firstFile = await validatePptx(firstResult.pptx);
      assert.deepEqual(firstFile.errors, []);
      assert.ok(
        firstFile.slides[0]?.objects.some(({ text }) =>
          text?.text.includes("甲预取 ¥4,860 万"),
        ),
      );
      assert.ok(tableValues(firstFile, "甲北方区").includes("甲北方区"));
      assert.ok(tableValues(firstFile, "甲云服务").includes("甲云服务"));
      assert.ok(!tableValues(firstFile, "甲北方区").includes("乙东部区"));
      assert.ok(!tableValues(firstFile, "甲云服务").includes("乙终端产品"));
    } finally {
      await server.close();
    }
  });

  it("区域请求失败时使用最近一期业务数据并保留另一表格的本期结果", async () => {
    const server = await startSalesServer([
      "/fallback/region",
      "/fallback/product",
    ]);
    try {
      const pending = build(
        createSalesReport({
          dataSource: createHttpSalesSource(`${server.origin}/fallback`),
          overview: testOverview,
        }),
      );
      await server.allRequested;
      server.respond("/fallback/region", [], 503);
      server.respond("/fallback/product", [
        {
          segment: "现场协作平台",
          sales: "¥1,060万",
          comparison: "订单 1,320",
          result: "毛利率 49%",
        },
      ]);

      const result = await pending;
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      assert.ok(
        file.slides[1]?.objects.some(({ text }) =>
          text?.text.includes("本期数据尚未完整，以下为最近一期记录"),
        ),
      );
      assert.ok(tableValues(file, "北方区").includes("¥1,395万"));
      assert.ok(tableValues(file, "现场协作平台").includes("现场协作平台"));
    } finally {
      await server.close();
    }
  });

  it("没有业务回退的产品请求失败时拒绝生成报告", async () => {
    const server = await startSalesServer([
      "/failure/region",
      "/failure/product",
    ]);
    try {
      const pending = build(
        createSalesReport({
          dataSource: createHttpSalesSource(`${server.origin}/failure`),
          overview: testOverview,
        }),
      );
      await server.allRequested;
      server.respond("/failure/region", [
        {
          segment: "正常区域",
          sales: "¥1,000万",
          comparison: "同比 +5%",
          result: "目标 100%",
        },
      ]);
      server.respond("/failure/product", [], 503);

      await assert.rejects(pending, BuildError);
    } finally {
      await server.close();
    }
  });

  it("取消时拒绝报告并关闭仍在等待的 HTTP 响应", async () => {
    const server = await startSalesServer([
      "/cancel/region",
      "/cancel/product",
    ]);
    const controller = new AbortController();
    const reason = new Error("不再需要这份销售报告");
    try {
      const pending = build(
        createSalesReport({
          dataSource: createHttpSalesSource(`${server.origin}/cancel`),
          overview: testOverview,
          signal: controller.signal,
        }),
        { signal: controller.signal },
      );
      const rejected = assert.rejects(
        pending,
        (error: unknown) => error === reason,
      );
      await server.allRequested;
      server.respond("/cancel/product", [
        {
          segment: "已就绪产品",
          sales: "¥900万",
          comparison: "订单 1,100",
          result: "毛利率 46%",
        },
      ]);
      await server.responseClosed("/cancel/product");
      controller.abort(reason);
      await rejected;
      await server.responseClosed("/cancel/region");
    } finally {
      await server.close();
    }
  });
});
