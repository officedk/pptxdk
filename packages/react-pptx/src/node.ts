import { readFileSync } from "node:fs";
import { registerHooks } from "node:module";
import { extname, sep } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

import { transformSync, type Loader } from "esbuild";
import {
  createElement,
  isValidElement,
  type ComponentType,
  type ReactNode,
} from "react";

import { build } from "./index.ts";
import type { BuildOptions, BuildResult } from "./index.ts";

const sourceLoaders: Readonly<Record<string, Loader>> = {
  ".ts": "ts",
  ".tsx": "tsx",
  ".mts": "ts",
  ".jsx": "jsx",
};

export type BuildFromSourceOptions = BuildOptions & {
  /** 用入口模块的具名导出作为文稿；默认读取 `default`。 */
  readonly exportName?: string;
};

/** 加载作者源码失败时抛出的错误。 */
export class SourceLoadError extends Error {
  constructor(message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = "SourceLoadError";
  }
}

function registerSourceLoader() {
  // 延迟导入也需要源码转换，因此在整次源码构建期间持有加载钩子。
  return registerHooks({
    load(url, context, nextLoad) {
      if (!url.startsWith("file:")) return nextLoad(url, context);

      const file = fileURLToPath(url);
      if (file.split(sep).includes("node_modules")) {
        return nextLoad(url, context);
      }

      const loader = sourceLoaders[extname(file)];
      if (loader === undefined) return nextLoad(url, context);

      const result = transformSync(readFileSync(file, "utf8"), {
        loader,
        format: "esm",
        jsx: "automatic",
        jsxImportSource: "react",
        sourcemap: "inline",
        sourcefile: file,
        target: "node24",
      });
      return { format: "module", source: result.code, shortCircuit: true };
    },
  });
}

function sourceUrl(entry: string | URL): URL {
  if (entry instanceof URL) {
    if (entry.protocol !== "file:") {
      throw new SourceLoadError("文稿入口必须是本地文件路径或 file URL");
    }
    return entry;
  }
  return pathToFileURL(entry);
}

async function loadDeck(url: URL, exportName: string): Promise<ReactNode> {
  let deckModule: Record<string, unknown>;
  try {
    deckModule = (await import(url.href)) as Record<string, unknown>;
  } catch (cause) {
    const details =
      cause instanceof Error ? (cause.stack ?? cause.message) : String(cause);
    throw new SourceLoadError(
      `无法加载文稿源码 ${fileURLToPath(url)}：\n${details}`,
      {
        cause,
      },
    );
  }

  const exported = deckModule[exportName];
  if (isValidElement(exported)) return exported;
  if (typeof exported === "function") {
    return createElement(exported as ComponentType);
  }

  throw new SourceLoadError(
    `文稿源码 ${fileURLToPath(url)} 必须导出名为 ${JSON.stringify(exportName)} 的 React 元素或组件`,
  );
}

/**
 * 从 Node.js 本地 ESM 源码入口执行正式构建。TS、TSX、MTS 和 JSX 由本入口加载；
 * 入口及其源码依赖仍按 Node 的文件位置解析。返回值与 `build` 完全相同。
 */
export async function buildFromSource(
  entry: string | URL,
  options: BuildFromSourceOptions = {},
): Promise<BuildResult> {
  const { exportName = "default", ...buildOptions } = options;
  const url = sourceUrl(entry);
  if (buildOptions.signal?.aborted) return build(null, buildOptions);

  const hooks = registerSourceLoader();
  try {
    const deck = await loadDeck(url, exportName);
    return await build(deck, buildOptions);
  } finally {
    hooks.deregister();
  }
}
