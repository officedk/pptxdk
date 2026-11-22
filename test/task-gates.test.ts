import assert from "node:assert/strict";

import { describe, it } from "vitest";

import {
  inspectTaskGate,
  taskGateSchema,
  type TaskGate,
} from "../scripts/task-gates.ts";

const entryCommand = "node --conditions=source scripts/task-gates.ts";

type MockScenario = {
  readonly query: string;
  readonly dryRuns?: Readonly<Record<string, string>>;
  readonly defaultDryRun?: string;
};

const mockTurboScript = `
const scenario = JSON.parse(process.argv[1]);
const args = process.argv.slice(2);
if (args[0] === "query") {
  process.stdout.write(scenario.query);
  process.exit(0);
}
if (args[0] === "run") {
  const tasks = args.slice(1, -1).join(",");
  const output = scenario.dryRuns?.[tasks] ?? scenario.defaultDryRun ?? JSON.stringify({ tasks: [] });
  process.stdout.write(output);
  process.exit(0);
}
process.stderr.write("Unexpected turbo command");
process.exit(1);
`;

function mockTurbo(scenario: MockScenario): readonly [string, ...string[]] {
  return [process.execPath, "-e", mockTurboScript, JSON.stringify(scenario)];
}

function makeRootTasks(
  extraRootTasks: readonly Record<string, unknown>[] = [],
): readonly Record<string, unknown>[] {
  return [
    {
      name: "check",
      fullName: "//#check",
      tags: ["quality:entry"],
      script: `${entryCommand} check`,
      command: `${entryCommand} check`,
    },
    {
      name: "check:heavy",
      fullName: "//#check:heavy",
      tags: ["quality:entry"],
      script: `${entryCommand} check:heavy`,
      command: `${entryCommand} check:heavy`,
    },
    {
      name: "check:full",
      fullName: "//#check:full",
      tags: ["quality:entry"],
      script: `${entryCommand} check:full`,
      command: `${entryCommand} check:full`,
    },
    {
      name: "fix",
      fullName: "//#fix",
      tags: ["work:repair"],
      script: `${entryCommand} fix`,
      command: `${entryCommand} fix`,
    },
    {
      name: "lint",
      fullName: "//#lint",
      tags: ["quality:light"],
      script: "oxlint .",
      command: "oxlint .",
    },
    ...extraRootTasks,
  ];
}

function makeBasePackages(
  extraPackages: readonly Record<string, unknown>[] = [],
  extraRootTasks: readonly Record<string, unknown>[] = [],
): readonly Record<string, unknown>[] {
  return [
    {
      name: "//",
      tasks: {
        items: makeRootTasks(extraRootTasks),
      },
    },
    ...extraPackages,
  ];
}

function makeQueryPayload(
  packages: readonly Record<string, unknown>[],
): Record<string, unknown> {
  return {
    version: "2.11.10",
    data: {
      packages: {
        items: packages,
      },
    },
  };
}

describe("task-gates schema validation and task graph rules", () => {
  it("exports taskGateSchema and accepts valid turbo query and dry-run payloads including null script/command", () => {
    const validGate: TaskGate = taskGateSchema.parse("check:full");
    assert.equal(validGate, "check:full");
    assert.equal(taskGateSchema.safeParse("unknown-gate").success, false);

    const queryPayload = makeQueryPayload(
      makeBasePackages([
        {
          name: "@pptxdk/tooling-config",
          tasks: {
            items: [
              {
                name: "build",
                fullName: "@pptxdk/tooling-config#build",
                tags: [],
                script: null,
                command: null,
              },
            ],
          },
        },
      ]),
    );

    const result = inspectTaskGate(
      validGate,
      process.cwd(),
      entryCommand,
      mockTurbo({
        query: JSON.stringify(queryPayload),
        dryRuns: {
          "//#lint": JSON.stringify({
            id: "dry-run-1",
            packages: ["//"],
            tasks: [{ taskId: "//#lint", hash: "abc" }],
          }),
        },
      }),
    );

    assert.deepEqual(result.diagnostics, []);
    assert.equal(result.plans.length, 2);
  });

  it("reports malformed JSON from turbo query and dry-run responses", () => {
    const badQuery = inspectTaskGate(
      "check",
      process.cwd(),
      entryCommand,
      mockTurbo({ query: "{not-json" }),
    );
    assert.equal(badQuery.diagnostics.length, 1);
    assert.match(
      badQuery.diagnostics[0]!,
      /Turbo native task directory is not valid JSON/,
    );

    const badDryRun = inspectTaskGate(
      "check",
      process.cwd(),
      entryCommand,
      mockTurbo({
        query: JSON.stringify(makeQueryPayload(makeBasePackages())),
        dryRuns: { "//#lint": "{not-json" },
      }),
    );
    assert.ok(
      badDryRun.diagnostics.some((d) =>
        /Turbo task graph dry run is not valid JSON/.test(d),
      ),
    );
  });

  it("reports localized field path diagnostics for missing fields and type errors in turbo query and dry-run", () => {
    const missingPackagesItems = inspectTaskGate(
      "check",
      process.cwd(),
      entryCommand,
      mockTurbo({ query: JSON.stringify({ data: { packages: {} } }) }),
    );
    assert.ok(
      missingPackagesItems.diagnostics.some((d) =>
        d.includes("data.packages.items"),
      ),
    );

    const missingTasksItems = inspectTaskGate(
      "check",
      process.cwd(),
      entryCommand,
      mockTurbo({
        query: JSON.stringify({
          data: { packages: { items: [{ name: "//", tasks: {} }] } },
        }),
      }),
    );
    assert.ok(
      missingTasksItems.diagnostics.some((d) =>
        d.includes("data.packages.items[0].tasks.items"),
      ),
    );

    const wrongTaskFieldTypes = inspectTaskGate(
      "check",
      process.cwd(),
      entryCommand,
      mockTurbo({
        query: JSON.stringify(
          makeQueryPayload([
            {
              name: "@pptxdk/core",
              tasks: {
                items: [
                  {
                    name: "",
                    fullName: "@pptxdk/core#test",
                    tags: [123],
                    script: 456,
                    command: null,
                  },
                ],
              },
            },
          ]),
        ),
      }),
    );
    assert.ok(
      wrongTaskFieldTypes.diagnostics.some((d) =>
        d.includes("data.packages.items[0].tasks.items[0].name"),
      ),
    );
    assert.ok(
      wrongTaskFieldTypes.diagnostics.some((d) =>
        d.includes("data.packages.items[0].tasks.items[0].tags[0]"),
      ),
    );
    assert.ok(
      wrongTaskFieldTypes.diagnostics.some((d) =>
        d.includes("data.packages.items[0].tasks.items[0].script"),
      ),
    );

    const missingDryRunTasks = inspectTaskGate(
      "check",
      process.cwd(),
      entryCommand,
      mockTurbo({
        query: JSON.stringify(makeQueryPayload(makeBasePackages())),
        dryRuns: { "//#lint": JSON.stringify({}) },
      }),
    );
    assert.ok(
      missingDryRunTasks.diagnostics.some((d) => d.includes("at tasks:")),
    );

    const missingTaskId = inspectTaskGate(
      "check",
      process.cwd(),
      entryCommand,
      mockTurbo({
        query: JSON.stringify(makeQueryPayload(makeBasePackages())),
        dryRuns: { "//#lint": JSON.stringify({ tasks: [{}] }) },
      }),
    );
    assert.ok(
      missingTaskId.diagnostics.some((d) => d.includes("tasks[0].taskId")),
    );

    const invalidTaskIdType = inspectTaskGate(
      "check",
      process.cwd(),
      entryCommand,
      mockTurbo({
        query: JSON.stringify(makeQueryPayload(makeBasePackages())),
        dryRuns: { "//#lint": JSON.stringify({ tasks: [{ taskId: 42 }] }) },
      }),
    );
    assert.ok(
      invalidTaskIdType.diagnostics.some((d) => d.includes("tasks[0].taskId")),
    );
  });

  it("enforces task role, protected verification, and tampered entrypoint rules", () => {
    const queryPayload = makeQueryPayload([
      {
        name: "//",
        tasks: {
          items: [
            {
              name: "check",
              fullName: "//#check",
              tags: ["quality:entry"],
              script: "tampered check",
              command: "tampered check",
            },
            {
              name: "check:heavy",
              fullName: "//#check:heavy",
              tags: ["quality:entry"],
              script: `${entryCommand} check:heavy`,
              command: `${entryCommand} check:heavy`,
            },
            {
              name: "check:full",
              fullName: "//#check:full",
              tags: ["quality:entry"],
              script: `${entryCommand} check:full`,
              command: `${entryCommand} check:full`,
            },
            {
              name: "fix",
              fullName: "//#fix",
              tags: ["work:repair"],
              script: `${entryCommand} fix`,
              command: `${entryCommand} fix`,
            },
          ],
        },
      },
      {
        name: "@pptxdk/core",
        tasks: {
          items: [
            {
              name: "build",
              fullName: "@pptxdk/core#build",
              tags: [],
              script: "tsc -b",
              command: "tsc -b",
            },
            {
              name: "typecheck",
              fullName: "@pptxdk/core#typecheck",
              tags: ["quality:light", "quality:heavy"],
              script: "tsc --noEmit",
              command: "tsc --noEmit",
            },
            {
              name: "format:check",
              fullName: "@pptxdk/core#format:check",
              tags: ["quality:bogus"],
              script: "oxfmt --list-different .",
              command: "oxfmt --list-different .",
            },
            {
              name: "test",
              fullName: "@pptxdk/core#test",
              tags: ["work:on-demand"],
              script: "vitest run",
              command: "vitest run",
            },
          ],
        },
      },
    ]);

    const result = inspectTaskGate(
      "check",
      process.cwd(),
      entryCommand,
      mockTurbo({ query: JSON.stringify(queryPayload) }),
    );

    assert.ok(
      result.diagnostics.some((d) =>
        d.includes("//#check public entrypoint wiring error"),
      ),
    );
    assert.ok(
      result.diagnostics.some((d) =>
        d.includes("@pptxdk/core#build is missing a unique role tag"),
      ),
    );
    assert.ok(
      result.diagnostics.some((d) =>
        d.includes("@pptxdk/core#typecheck has multiple role tags"),
      ),
    );
    assert.ok(
      result.diagnostics.some((d) =>
        d.includes("@pptxdk/core#format:check uses unknown role tag"),
      ),
    );
    assert.ok(
      result.diagnostics.some((d) =>
        d.includes("@pptxdk/core#test is a protected verification task"),
      ),
    );
  });

  it("enforces dangling commandless role nodes and layer closure rules", () => {
    const queryPayload = makeQueryPayload(
      makeBasePackages([
        {
          name: "@pptxdk/core",
          tasks: {
            items: [
              {
                name: "aggregate",
                fullName: "@pptxdk/core#aggregate",
                tags: ["quality:light"],
                script: null,
                command: null,
              },
              {
                name: "test",
                fullName: "@pptxdk/core#test",
                tags: ["quality:heavy"],
                script: "vitest run",
                command: "vitest run",
              },
              {
                name: "manual",
                fullName: "@pptxdk/core#manual",
                tags: ["quality:conditional"],
                script: "node manual.ts",
                command: "node manual.ts",
              },
            ],
          },
        },
      ]),
    );

    const result = inspectTaskGate(
      "check",
      process.cwd(),
      entryCommand,
      mockTurbo({
        query: JSON.stringify(queryPayload),
        dryRuns: {
          "@pptxdk/core#aggregate": JSON.stringify({
            tasks: [{ taskId: "@pptxdk/core#aggregate" }],
          }),
          "//#lint,@pptxdk/core#aggregate": JSON.stringify({
            tasks: [
              { taskId: "//#lint" },
              { taskId: "@pptxdk/core#test" },
              { taskId: "@pptxdk/core#manual" },
              { taskId: "//#check:heavy" },
            ],
          }),
        },
      }),
    );

    assert.ok(
      result.diagnostics.some((d) =>
        d.includes(
          "@pptxdk/core#aggregate declares dangling role tag quality:light",
        ),
      ),
    );
    assert.ok(
      result.diagnostics.some((d) =>
        d.includes("light check cannot indirectly trigger heavy verification"),
      ),
    );
    assert.ok(
      result.diagnostics.some((d) =>
        d.includes(
          "standard verification cannot indirectly trigger conditional verification",
        ),
      ),
    );
    assert.ok(
      result.diagnostics.some((d) =>
        d.includes("leaf task cannot depend on public aggregate entrypoint"),
      ),
    );
  });

  it("validates the live workspace check:full task graph with zero diagnostics", () => {
    const live = inspectTaskGate("check:full");
    assert.deepEqual(live.diagnostics, []);
    assert.equal(live.plans.length, 2);
    assert.ok(
      live.plans[0]!.tasks.some((task) => task.fullName === "//#test"),
      "Expected //#test to be registered in quality:light plan",
    );
  });
});
