#!/usr/bin/env node
import { runCli } from "./index.ts";

const interrupt = new AbortController();
const onInterrupt = () => interrupt.abort();
process.once("SIGINT", onInterrupt);

try {
  process.exitCode = await runCli({
    argv: process.argv,
    signal: interrupt.signal,
    write: ({ destination, chunk }) => {
      const stream = destination === "stdout" ? process.stdout : process.stderr;
      stream.write(chunk);
    },
  });
} finally {
  process.removeListener("SIGINT", onInterrupt);
}
