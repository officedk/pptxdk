import { createSalesReport, loadBundledOverview } from "./sales-report.tsx";

export {
  createHttpSalesSource,
  createSalesReport,
  SalesTable,
} from "./sales-report.tsx";
export type {
  SalesDataSource,
  SalesOverview,
  SalesReportOptions,
  SalesRow,
  SalesSection,
} from "./sales-report.tsx";

// 概览在入口先读取，表格明细则由两个组件各自异步准备。
const overview = await loadBundledOverview();

export default createSalesReport({ overview });
