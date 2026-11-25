import { basename } from "node:path";

import { defineConfig } from "rolldown";
import { dts } from "rolldown-plugin-dts";

import manifest from "./package.json" with { type: "json" };

const publicEntries = Object.fromEntries(
  Object.values(manifest.exports).map(({ source }) => [
    basename(source, ".ts"),
    source,
  ]),
);
const dependencies = [manifest.name, ...Object.keys(manifest.dependencies)];
const external = (id: string) =>
  dependencies.some((name) => id === name || id.startsWith(name + "/"));

export default defineConfig([
  {
    input: { ...publicEntries, cli: "src/cli.ts" },
    logLevel: "warn",
    platform: "node",
    external,
    output: {
      dir: "dist",
      format: "esm",
      // Capability selection and deferred backends rely on the original module boundaries.
      preserveModules: true,
    },
  },
  {
    input: Object.fromEntries(
      Object.keys(publicEntries).map((name) => [
        name,
        `dist/.declarations/${name}.d.ts`,
      ]),
    ),
    logLevel: "warn",
    platform: "node",
    external,
    plugins: [
      dts({ dtsOnly: true, generator: "oxc", tsconfig: "tsconfig.build.json" }),
    ],
    output: { dir: "dist", format: "esm" },
  },
]);
