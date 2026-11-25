#!/usr/bin/env node
import { nodeCliOutput } from "@ykdz/cli-contract/node";

import { runCli } from "./index.ts";

const interrupt = new AbortController();
const onInterrupt = () => interrupt.abort();
process.once("SIGINT", onInterrupt);

try {
  process.exitCode = await runCli({
    argv: process.argv,
    signal: interrupt.signal,
    write: nodeCliOutput({ stdout: process.stdout, stderr: process.stderr }),
  });
} finally {
  process.removeListener("SIGINT", onInterrupt);
}
