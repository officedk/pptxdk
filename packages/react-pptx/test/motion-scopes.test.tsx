import assert from "node:assert/strict";

import {
  after,
  Box,
  breathe,
  build,
  BuildError,
  Deck,
  fade,
  Group,
  MotionScope,
  Shape,
  Slide,
  spin,
  Step,
  Text,
  useMotionScope,
  repeat,
  type MotionObservation,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createElement, Fragment, type ReactNode } from "react";
import { describe, it } from "vitest";

function movingShape(testId: string, duration: number) {
  return createElement(Shape, {
    testId,
    x: 0,
    y: 0,
    width: 20,
    height: 20,
    enter: fade({ duration }),
  });
}

function observedShape(
  testId: string,
  text: string,
  y: number,
  fill: string,
  duration: number,
  delay = 0,
) {
  return createElement(Shape, {
    testId,
    x: 10,
    y,
    width: 280,
    height: 24,
    fill,
    text,
    enter: fade({ duration, delay }),
  });
}

function unknownLetterText(testId: string, text: string, y: number) {
  return createElement(
    Text,
    {
      testId,
      x: 10,
      y,
      width: 280,
      height: 32,
      enter: fade({ by: "letter", duration: 1000, intervalMs: 500 }),
    },
    text,
  );
}

function SingleSourceUnknownStartAliases({
  trigger,
  sourceCompletion = "unknown",
}: {
  readonly trigger: "auto" | "click";
  readonly sourceCompletion?: "known" | "unknown" | "no-natural-completion";
}) {
  const source = useMotionScope();
  const parent = useMotionScope();
  const alias = useMotionScope();
  const secondAlias = useMotionScope();
  const afterParent = useMotionScope();
  const sourceContent =
    sourceCompletion === "unknown"
      ? unknownLetterText("single-alias-source", "TT", 8)
      : sourceCompletion === "known"
        ? movingShape("single-alias-source", 100)
        : createElement(Shape, {
            testId: "single-alias-source",
            x: 10,
            y: 8,
            width: 20,
            height: 20,
            fill: "#164E3D",
            emphasis: repeat(
              breathe({ period: 1000, from: "#164E3D", to: "#4ADE80" }),
              "forever",
            ),
          });
  const content = createElement(
    Fragment,
    null,
    createElement(MotionScope, { scope: source }, sourceContent),
    createElement(
      MotionScope,
      { scope: parent, startAt: after(source.end, 2500) },
      movingShape("single-alias-parent", 100),
    ),
    createElement(
      MotionScope,
      { scope: alias, startAt: after(parent.start, 19) },
      movingShape("single-alias-first", 100),
    ),
    createElement(
      MotionScope,
      {
        scope: secondAlias,
        startAt: after(alias.start, 13),
        delay: 7,
      },
      movingShape("single-alias-second", 100),
    ),
    createElement(
      MotionScope,
      { scope: afterParent, startAt: parent.end },
      movingShape("single-alias-after-parent-end", 50),
    ),
  );
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: `single-source-start-alias-${trigger}` },
      trigger === "click" ? createElement(Step, { n: 1 }, content) : content,
    ),
  );
}

function SingleSourceAliasUnderLaterParent() {
  const source = useMotionScope();
  const parent = useMotionScope();
  const lateParent = useMotionScope();
  const alias = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "single-source-alias-under-later-parent" },
      createElement(
        MotionScope,
        { scope: source },
        unknownLetterText("later-alias-source", "TT", 8),
      ),
      createElement(
        MotionScope,
        { scope: parent, startAt: after(source.end, 2500) },
        movingShape("later-alias-parent", 100),
      ),
      createElement(
        MotionScope,
        { scope: lateParent, startAt: after(parent.end, 200) },
        createElement(
          MotionScope,
          { scope: alias, startAt: after(parent.start, 19) },
          movingShape("later-alias-consumer", 100),
        ),
      ),
    ),
  );
}

type UnknownStartSequenceOptions = {
  readonly mode?: "parallel" | "sequence" | "stagger";
  readonly nested?: boolean;
  readonly explicitStep?: boolean;
};

function UnknownStartSequence({
  mode = "sequence",
  nested = false,
  explicitStep = false,
}: UnknownStartSequenceOptions) {
  const source = useMotionScope();
  const follow = useMotionScope();
  const nestedSequence = useMotionScope();
  const follower = useMotionScope();
  const sourceScope = createElement(
    MotionScope,
    { scope: source },
    unknownLetterText("unknown-source", "ABC", 8),
  );
  const first = observedShape("sequence-first", "First", 48, "#ED7D31", 100);
  const second = observedShape("sequence-second", "Second", 78, "#4472C4", 100);
  const followContents = nested
    ? createElement(
        MotionScope,
        { scope: nestedSequence, mode: "sequence" },
        first,
        second,
      )
    : [first, second];
  const followScope = createElement(
    MotionScope,
    {
      scope: follow,
      mode: nested ? "parallel" : mode,
      ...(mode === "stagger" && !nested ? { each: 100 } : {}),
      startAt: after(source.end, 25),
    },
    followContents,
  );
  const followerScope = createElement(
    MotionScope,
    { scope: follower, startAt: after(follow.end, 10) },
    observedShape("sequence-follower", "Follower", 108, "#70AD47", 50),
  );
  const content = explicitStep
    ? createElement(Step, { n: 2 }, sourceScope, followScope, followerScope)
    : createElement(Fragment, null, sourceScope, followScope, followerScope);
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(Slide, { slideId: "unknown-start-sequence" }, content),
  );
}

function SelfStartedNestedPredecessor() {
  const source = useMotionScope();
  const follow = useMotionScope();
  const predecessor = useMotionScope();
  const follower = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "self-started-nested-predecessor" },
      createElement(
        MotionScope,
        { scope: source },
        unknownLetterText("self-start-source", "ABC", 8),
      ),
      createElement(
        MotionScope,
        {
          scope: follow,
          mode: "sequence",
          startAt: after(source.end, 25),
        },
        createElement(
          MotionScope,
          {
            scope: predecessor,
            startAt: after(follow.start, 40),
          },
          unknownLetterText("predecessor-long", "ABCDEFG", 48),
          unknownLetterText("predecessor-short", "AB", 88),
        ),
        observedShape("after-predecessor", "D", 128, "#ED7D31", 100),
      ),
      createElement(
        MotionScope,
        { scope: follower, startAt: after(follow.end, 10) },
        observedShape(
          "after-self-started-sequence",
          "Follower",
          158,
          "#70AD47",
          50,
        ),
      ),
    ),
  );
}

type SequenceStartCausalityOptions = {
  readonly sourceUnknown: boolean;
  readonly firstUnknown: boolean;
  readonly offsetMs: number;
  readonly explicitStep?: boolean;
  readonly automaticClickStep?: boolean;
  readonly wrapFirstInScope?: boolean;
};

function SequenceStartCausality({
  sourceUnknown,
  firstUnknown,
  offsetMs,
  explicitStep = false,
  automaticClickStep = false,
  wrapFirstInScope = false,
}: SequenceStartCausalityOptions) {
  const source = useMotionScope();
  const follow = useMotionScope();
  const firstScope = useMotionScope();
  const child = useMotionScope();
  const follower = useMotionScope();
  const sourceContent = sourceUnknown
    ? createElement(
        Text,
        {
          testId: "causality-source",
          x: 10,
          y: 8,
          width: 280,
          height: 32,
          enter: fade({ by: "letter", duration: 733, intervalMs: 211 }),
        },
        "AB",
      )
    : movingShape("causality-source", 137);
  const first = firstUnknown
    ? createElement(
        Text,
        {
          testId: "causality-first",
          x: 10,
          y: 48,
          width: 280,
          height: 32,
          enter: fade({ by: "letter", duration: 733, intervalMs: 211 }),
        },
        "AB",
      )
    : movingShape("causality-first", 137);
  const firstMember = wrapFirstInScope
    ? createElement(MotionScope, { scope: firstScope }, first)
    : first;
  const content = createElement(
    Fragment,
    null,
    createElement(MotionScope, { scope: source }, sourceContent),
    createElement(
      MotionScope,
      { scope: follow, mode: "sequence", startAt: after(source.end, 25) },
      firstMember,
      createElement(
        MotionScope,
        { scope: child, startAt: after(follow.start, offsetMs) },
        movingShape("causality-second", 137),
      ),
    ),
    createElement(
      MotionScope,
      { scope: follower, startAt: after(follow.end, 11) },
      movingShape("causality-follower", 50),
    ),
  );
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "sequence-start-causality" },
      explicitStep
        ? createElement(Step, { n: 2 }, content)
        : automaticClickStep
          ? createElement(Step, null, content)
          : content,
    ),
  );
}

type ForwardSiblingOptions = { readonly sourceFirst: boolean };

function ForwardSibling({ sourceFirst }: ForwardSiblingOptions) {
  const activator = useMotionScope();
  const parent = useMotionScope();
  const source = useMotionScope();
  const waiter = useMotionScope();
  const sourceScope = createElement(
    MotionScope,
    { scope: source },
    movingShape("forward-source", 200),
  );
  const waitingScope = createElement(
    MotionScope,
    { scope: waiter, startAt: source.end },
    movingShape("forward-waiter", 100),
  );
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "forward-sibling" },
      createElement(
        MotionScope,
        { scope: activator },
        unknownLetterText("forward-activator", "ABC", 8),
      ),
      createElement(
        MotionScope,
        {
          scope: parent,
          mode: "parallel",
          startAt: after(activator.end, 10),
        },
        ...(sourceFirst
          ? [sourceScope, waitingScope]
          : [waitingScope, sourceScope]),
      ),
    ),
  );
}

function IncomparableUnknownStarts() {
  const firstSource = useMotionScope();
  const secondSource = useMotionScope();
  const parent = useMotionScope();
  const child = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "incomparable-unknown-starts" },
      createElement(
        MotionScope,
        { scope: firstSource },
        unknownLetterText("incomparable-first", "ABC", 8),
      ),
      createElement(
        MotionScope,
        { scope: secondSource },
        unknownLetterText("incomparable-second", "DEF", 48),
      ),
      createElement(
        MotionScope,
        { scope: parent, startAt: firstSource.end },
        createElement(
          MotionScope,
          { scope: child, startAt: secondSource.end },
          movingShape("incomparable-child", 100),
        ),
      ),
    ),
  );
}

function ReaderCycleDeck() {
  const first = useMotionScope();
  const second = useMotionScope();
  const firstWaiter = useMotionScope();
  const secondWaiter = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "reader-completion-cycle" },
      createElement(
        MotionScope,
        { scope: first },
        unknownLetterText("cycle-first", "AB", 8),
      ),
      createElement(
        MotionScope,
        { scope: second },
        unknownLetterText("cycle-second", "CD", 48),
      ),
      createElement(
        MotionScope,
        { scope: firstWaiter, startAt: after(first.end, 0) },
        movingShape("after-cycle-first", 100),
      ),
      createElement(
        MotionScope,
        { scope: secondWaiter, startAt: after(second.end, 0) },
        movingShape("after-cycle-second", 100),
      ),
    ),
  );
}

function startAfterCondition(
  xml: string,
  timingId: number,
  sourceTimingId: number,
): string {
  const marker = `<p:cTn id="${timingId}"`;
  const nodeStart = xml.indexOf(marker);
  assert.notEqual(nodeStart, -1);
  const conditionsStart = xml.indexOf("<p:stCondLst>", nodeStart);
  const conditionsEnd = xml.indexOf("</p:stCondLst>", conditionsStart);
  assert.ok(conditionsStart > nodeStart && conditionsEnd > conditionsStart);
  const conditions = xml.slice(conditionsStart, conditionsEnd);
  const rewritten = conditions.replace(
    /<p:cond\b[^>]*>[\s\S]*?<\/p:cond>/u,
    `<p:cond evt="end" delay="0"><p:tn val="${sourceTimingId}"></p:tn></p:cond>`,
  );
  assert.notEqual(rewritten, conditions);
  return xml.slice(0, conditionsStart) + rewritten + xml.slice(conditionsEnd);
}

function startAtDelay(xml: string, timingId: number, delayMs: number): string {
  const marker = `<p:cTn id="${timingId}"`;
  const nodeStart = xml.indexOf(marker);
  assert.notEqual(nodeStart, -1);
  const conditionsStart = xml.indexOf("<p:stCondLst>", nodeStart);
  const conditionsEnd = xml.indexOf("</p:stCondLst>", conditionsStart);
  assert.ok(conditionsStart > nodeStart && conditionsEnd > conditionsStart);
  const conditions = xml.slice(conditionsStart, conditionsEnd);
  const rewritten = conditions.replace(
    /<p:cond\b[^>]*>[\s\S]*?<\/p:cond>/u,
    `<p:cond delay="${delayMs}"></p:cond>`,
  );
  assert.notEqual(rewritten, conditions);
  return xml.slice(0, conditionsStart) + rewritten + xml.slice(conditionsEnd);
}

function startConditions(xml: string, timingId: number): string {
  const marker = `<p:cTn id="${timingId}"`;
  const nodeStart = xml.indexOf(marker);
  assert.notEqual(nodeStart, -1);
  const conditionsStart = xml.indexOf("<p:stCondLst>", nodeStart);
  const conditionsEnd = xml.indexOf("</p:stCondLst>", conditionsStart);
  assert.ok(conditionsStart > nodeStart && conditionsEnd > conditionsStart);
  return xml.slice(conditionsStart, conditionsEnd);
}

function assertTimingReferencesFollowDefinitions(xml: string): void {
  const definitions = new Map<string, number>();
  for (const match of xml.matchAll(/<p:cTn\b[^>]*\bid="(\d+)"/gu)) {
    const id = match[1];
    assert.ok(id);
    assert.equal(definitions.has(id), false, `重复的时间节点定义 ${id}`);
    definitions.set(id, match.index);
  }
  for (const match of xml.matchAll(/<p:tn\b[^>]*\bval="(\d+)"/gu)) {
    const id = match[1];
    assert.ok(id);
    const definition = definitions.get(id);
    assert.ok(definition !== undefined, `时间引用 ${id} 没有节点定义`);
    assert.ok(definition < match.index, `时间引用 ${id} 出现在节点定义之前`);
  }
}

function directTimingChildren(xml: string): string[] {
  const tags = /<p:par\b[^>]*>|<\/p:par>/gu;
  const children: string[] = [];
  let depth = 0;
  let childStart = -1;
  for (const tag of xml.matchAll(tags)) {
    const value = tag[0];
    const start = tag.index;
    if (value === "</p:par>") {
      depth -= 1;
      if (depth === 0 && childStart >= 0) {
        children.push(xml.slice(childStart, start + value.length));
        childStart = -1;
      }
    } else {
      if (depth === 0) childStart = start;
      depth += 1;
    }
  }
  return children;
}

function firstMainSequenceEntryNodeType(
  xml: string,
  stepIndex: number,
): string | undefined {
  const sequenceStart = xml.indexOf("<p:seq");
  const sequenceEnd = xml.indexOf("</p:seq>", sequenceStart);
  assert.ok(sequenceStart >= 0 && sequenceEnd > sequenceStart);
  const sequence = xml.slice(sequenceStart, sequenceEnd);
  const mainSequenceStart = sequence.indexOf('nodeType="mainSeq"');
  const stepsStart = sequence.indexOf("<p:childTnLst>", mainSequenceStart);
  assert.ok(mainSequenceStart >= 0 && stepsStart > mainSequenceStart);
  const listLength = "<p:childTnLst>".length;
  const steps = directTimingChildren(sequence.slice(stepsStart + listLength));
  const step = steps[stepIndex];
  assert.ok(step);
  const childListStart = step.indexOf("<p:childTnLst>");
  assert.ok(childListStart >= 0);
  const nodes = directTimingChildren(step.slice(childListStart + listLength));
  const withNode = nodes[0];
  assert.ok(withNode);
  const withChildrenStart = withNode.indexOf("<p:childTnLst>");
  assert.ok(withChildrenStart >= 0);
  // LibreOffice fixMainSequenceTiming reads the first child of this with-node.
  const entryNodes = directTimingChildren(
    withNode.slice(withChildrenStart + listLength),
  );
  const entryNode = entryNodes[0];
  assert.ok(entryNode);
  const attributes = entryNode.match(/<p:cTn\b([^>]*)>/u)?.[1];
  return attributes?.match(/(?:^|\s)nodeType="([^"]+)"/u)?.[1];
}

function StaggeredSource() {
  const source = useMotionScope();
  const namedCompletion = source.cue("finished");
  const directFollower = useMotionScope();
  const namedFollower = useMotionScope();

  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "scope-order" },
      createElement(
        MotionScope,
        { scope: directFollower, startAt: after(source.end, 50) },
        movingShape("after-end", 300),
      ),
      createElement(
        MotionScope,
        { scope: namedFollower, startAt: after(namedCompletion, 25) },
        movingShape("after-named-end", 300),
      ),
      createElement(
        MotionScope,
        {
          scope: source,
          mode: "stagger",
          each: 100,
          publishEnd: namedCompletion,
        },
        movingShape("long-first", 1000),
        movingShape("medium-second", 800),
        movingShape("short-last", 50),
      ),
    ),
  );
}

function EmptyScopeDeck({
  includeSource,
  sourceMode = "empty",
}: {
  readonly includeSource: boolean;
  readonly sourceMode?: "empty" | "static" | "disabled";
}) {
  const empty = useMotionScope();
  const emptyStart = empty.cue("empty-start");
  const emptyEnd = empty.cue("empty-end");
  const fromStart = useMotionScope();
  const fromEnd = useMotionScope();
  const fromNamedStart = useMotionScope();
  const fromNamedEnd = useMotionScope();
  const sourceChildren =
    sourceMode === "empty"
      ? null
      : sourceMode === "static"
        ? createElement(Shape, {
            testId: "static-source",
            x: 0,
            y: 0,
            width: 20,
            height: 20,
          })
        : movingShape("disabled-source", 100);

  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "empty-scope" },
      includeSource
        ? createElement(
            MotionScope,
            {
              scope: empty,
              publishStart: emptyStart,
              publishEnd: emptyEnd,
              motion: sourceMode !== "disabled",
            },
            sourceChildren,
          )
        : null,
      createElement(
        MotionScope,
        { scope: fromStart, startAt: after(empty.start, 50) },
        movingShape("from-start", 100),
      ),
      createElement(
        MotionScope,
        { scope: fromEnd, startAt: after(empty.end, 80) },
        movingShape("from-end", 100),
      ),
      createElement(
        MotionScope,
        { scope: fromNamedStart, startAt: after(emptyStart, 30) },
        movingShape("from-named-start", 100),
      ),
      createElement(
        MotionScope,
        { scope: fromNamedEnd, startAt: after(emptyEnd, 40) },
        movingShape("from-named-end", 100),
      ),
    ),
  );
}

function CyclicScopes() {
  const first = useMotionScope();
  const second = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "cycle" },
      createElement(
        MotionScope,
        { scope: first, startAt: second.end },
        movingShape("first", 100),
      ),
      createElement(
        MotionScope,
        { scope: second, startAt: first.end },
        movingShape("second", 100),
      ),
    ),
  );
}

function CrossStepScopes() {
  const first = useMotionScope();
  const second = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "cross-step" },
      createElement(
        Step,
        { n: 1 },
        createElement(
          MotionScope,
          { scope: first },
          movingShape("first-step", 100),
        ),
      ),
      createElement(
        Step,
        { n: 2 },
        createElement(
          MotionScope,
          { scope: second, startAt: first.end },
          movingShape("second-step", 100),
        ),
      ),
    ),
  );
}

function SequencedScopes() {
  const sequence = useMotionScope();
  const parallel = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "sequence" },
      createElement(
        MotionScope,
        { scope: sequence, mode: "sequence" },
        movingShape("first", 200),
        createElement(
          MotionScope,
          { scope: parallel },
          movingShape("parallel-short", 100),
          movingShape("parallel-long", 700),
        ),
        movingShape("last", 100),
      ),
    ),
  );
}

function UnknownCompletionTree() {
  const outer = useMotionScope();
  const chain = useMotionScope();
  const inner = useMotionScope();
  const selfStarted = useMotionScope();
  const follower = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "unknown-completion-tree" },
      createElement(
        MotionScope,
        { scope: outer },
        createElement(
          MotionScope,
          { scope: chain, mode: "sequence" },
          createElement(
            MotionScope,
            { scope: inner, mode: "stagger", each: 1000 },
            createElement(
              Text,
              {
                testId: "long-first",
                x: 10,
                y: 8,
                width: 300,
                height: 40,
                enter: fade({ by: "letter", duration: 1000, intervalMs: 500 }),
              },
              "ABCDEFGHIJKLMN",
            ),
            createElement(
              Text,
              {
                testId: "short-late",
                x: 10,
                y: 45,
                width: 300,
                height: 40,
                enter: fade({ by: "letter", duration: 1000, intervalMs: 500 }),
              },
              "AB",
            ),
          ),
          observedShape(
            "finite-after-unknown",
            "D: starts after inner.end",
            122,
            "#ED7D31",
            100,
          ),
        ),
        createElement(
          MotionScope,
          { scope: selfStarted, startAt: after(chain.start, 4000) },
          observedShape(
            "self-started-finite",
            "Self start +4000 ms (chain.start)",
            92,
            "#4472C4",
            100,
          ),
        ),
      ),
      createElement(
        MotionScope,
        { scope: follower, startAt: outer.end },
        observedShape(
          "after-unknown-join",
          "Outer follower: after outer.end",
          150,
          "#70AD47",
          50,
        ),
      ),
    ),
  );
}

type IncomparableJoinOptions = {
  readonly firstText: string;
  readonly trigger: "auto" | "click" | "explicit";
  readonly unknownParentStart?: boolean;
  readonly parentStartGate?: boolean;
  readonly scopeDelayMs?: number;
  readonly effectDelayMs?: number;
};

function ConditionJoinStartAlias({
  parentStartGate,
  trigger = "auto",
}: {
  readonly parentStartGate: boolean;
  readonly trigger?: "auto" | "click";
}) {
  const chain = useMotionScope();
  const inner = useMotionScope();
  const delayed = useMotionScope();
  const joined = useMotionScope();
  const alias = useMotionScope();
  const parentAlias = useMotionScope();
  const innerScope = createElement(
    MotionScope,
    { scope: inner },
    unknownLetterText("alias-first", "AB", 8),
    createElement(
      MotionScope,
      { scope: delayed, delay: 1000 },
      unknownLetterText("alias-late", "XY", 42),
    ),
  );
  const joinedScope = createElement(
    MotionScope,
    {
      scope: joined,
      startAt: parentStartGate
        ? after(chain.start, 5000)
        : after(delayed.end, 11000),
    },
    observedShape("alias-join-consumer", "Join consumer", 76, "#ED7D31", 100),
    createElement(
      MotionScope,
      { scope: parentAlias, startAt: after(joined.start, 13), delay: 7 },
      observedShape(
        "parent-alias-consumer",
        "Parent alias",
        100,
        "#4472C4",
        100,
        5,
      ),
    ),
  );
  const content = createElement(
    Fragment,
    null,
    createElement(
      MotionScope,
      { scope: chain, mode: "sequence" },
      innerScope,
      joinedScope,
    ),
    createElement(
      MotionScope,
      { scope: alias, startAt: after(joined.start, 19), delay: 23 },
      observedShape(
        "alias-consumer",
        "Alias consumer",
        122,
        "#C00000",
        100,
        31,
      ),
    ),
  );
  return createElement(
    Deck,
    { width: 320, height: 190 },
    createElement(
      Slide,
      { slideId: `condition-join-start-alias-${trigger}` },
      trigger === "click" ? createElement(Step, { n: 1 }, content) : content,
    ),
  );
}

function LateParentRetainedAlias() {
  const host = useMotionScope();
  const first = useMotionScope();
  const delayed = useMotionScope();
  const joined = useMotionScope();
  const parent = useMotionScope();
  const alias = useMotionScope();
  const afterParent = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 200 },
    createElement(
      Slide,
      { slideId: "late-parent-retained-alias" },
      createElement(
        MotionScope,
        { scope: host, mode: "sequence" },
        createElement(
          MotionScope,
          { scope: first },
          unknownLetterText(
            "late-alias-first",
            "ABCDEFGHIJKLMNOPQRSTUVWXYZ123456",
            8,
          ),
          createElement(
            MotionScope,
            { scope: delayed, delay: 1000 },
            unknownLetterText("late-alias-delayed", "XY", 42),
          ),
        ),
        createElement(
          MotionScope,
          { scope: joined, startAt: after(delayed.end, 11000) },
          observedShape("late-alias-join", "Join", 76, "#ED7D31", 100),
        ),
      ),
      createElement(
        MotionScope,
        { scope: parent, startAt: after(first.end, 10) },
        createElement(
          MotionScope,
          { scope: alias, startAt: after(joined.start, 19) },
          observedShape("late-alias-consumer", "Alias", 108, "#C00000", 100),
        ),
      ),
      createElement(
        MotionScope,
        { scope: afterParent, startAt: parent.end },
        observedShape(
          "late-alias-after-parent",
          "After parent",
          140,
          "#70AD47",
          100,
        ),
      ),
    ),
  );
}

function SameJoinParentAlias() {
  const chain = useMotionScope();
  const first = useMotionScope();
  const delayed = useMotionScope();
  const joined = useMotionScope();
  const parent = useMotionScope();
  const alias = useMotionScope();
  const afterParent = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "same-join-parent-alias" },
      createElement(
        Fragment,
        null,
        createElement(
          MotionScope,
          { scope: chain, mode: "sequence" },
          createElement(
            MotionScope,
            { scope: first },
            unknownLetterText(
              "same-parent-first",
              "ABCDEFGHIJKLMNOPQRSTUVWXYZ123456",
              8,
            ),
            createElement(
              MotionScope,
              { scope: delayed, delay: 1000 },
              unknownLetterText("same-parent-delayed", "XY", 42),
            ),
          ),
          createElement(
            MotionScope,
            { scope: joined, startAt: after(delayed.end, 11000) },
            observedShape("same-parent-join", "Join", 76, "#ED7D31", 100),
          ),
        ),
        createElement(
          MotionScope,
          { scope: parent, startAt: after(joined.start, 10) },
          createElement(
            MotionScope,
            { scope: alias, startAt: after(joined.start, 19) },
            observedShape("same-parent-alias", "Alias", 108, "#C00000", 100),
          ),
        ),
        createElement(
          MotionScope,
          { scope: afterParent, startAt: parent.end },
          observedShape(
            "same-parent-after",
            "After parent",
            140,
            "#70AD47",
            100,
          ),
        ),
      ),
    ),
  );
}

function ForwardConditionJoinAlias() {
  const host = useMotionScope();
  const first = useMotionScope();
  const delayed = useMotionScope();
  const joined = useMotionScope();
  const alias = useMotionScope();
  const follower = useMotionScope();
  const source = createElement(
    MotionScope,
    { scope: host, mode: "sequence" },
    createElement(
      MotionScope,
      { scope: first },
      unknownLetterText("forward-alias-first", "AB", 8),
      createElement(
        MotionScope,
        { scope: delayed, delay: 1000 },
        unknownLetterText("forward-alias-delayed", "XY", 42),
      ),
    ),
    createElement(
      MotionScope,
      { scope: joined, startAt: after(delayed.end, 11000) },
      observedShape("forward-alias-join", "Join", 76, "#ED7D31", 100),
    ),
  );
  return createElement(
    Deck,
    { width: 320, height: 200 },
    createElement(
      Slide,
      { slideId: "forward-condition-join-alias" },
      createElement(
        Fragment,
        null,
        createElement(
          MotionScope,
          { scope: alias, startAt: after(joined.start, 19) },
          observedShape("forward-alias-consumer", "Alias", 108, "#C00000", 100),
        ),
        source,
        createElement(
          MotionScope,
          { scope: follower, startAt: host.end },
          observedShape(
            "forward-alias-follower",
            "Follower",
            140,
            "#70AD47",
            100,
          ),
        ),
      ),
    ),
  );
}

type RootEndChainRoot =
  | "alias"
  | "afterAlias"
  | "afterAfterAlias"
  | "host"
  | "follower"
  | "targetGroup";

type RootEndChainOptions = {
  readonly rootOrder: readonly RootEndChainRoot[];
  readonly includeAfterAfterAlias?: boolean;
  readonly includeTargetGroup?: boolean;
  readonly explicitStep?: boolean;
};

function ForwardConditionJoinOrdinaryEndChain({
  rootOrder,
  includeAfterAfterAlias = false,
  includeTargetGroup = false,
  explicitStep = false,
}: RootEndChainOptions) {
  const host = useMotionScope();
  const first = useMotionScope();
  const delayed = useMotionScope();
  const joined = useMotionScope();
  const alias = useMotionScope();
  const afterAlias = useMotionScope();
  const afterAfterAlias = useMotionScope();
  const follower = useMotionScope();
  const targetGroupSequence = useMotionScope();
  const targetGroupFollower = useMotionScope();
  const source = createElement(
    MotionScope,
    { scope: host, mode: "sequence" },
    createElement(
      MotionScope,
      { scope: first },
      unknownLetterText("chain-first", "AB", 8),
      createElement(
        MotionScope,
        { scope: delayed, delay: 1000 },
        unknownLetterText("chain-delayed", "XY", 42),
      ),
    ),
    createElement(
      MotionScope,
      { scope: joined, startAt: after(delayed.end, 11000) },
      observedShape("chain-join", "Join", 76, "#ED7D31", 100),
    ),
  );
  const targetGroupSource = createElement(
    MotionScope,
    { scope: targetGroupSequence, mode: "sequence" },
    createElement(Shape, {
      testId: "chain-target-group-source",
      x: 10,
      y: 220,
      width: 280,
      height: 24,
      fill: "#7030A0",
      text: "AB",
      enter: fade({ by: "letter", duration: 100, intervalMs: 50 }),
      emphasis: spin({ duration: 200 }),
    }),
    createElement(
      MotionScope,
      { scope: targetGroupFollower },
      observedShape(
        "chain-target-group-follower",
        "Group follower",
        248,
        "#7030A0",
        100,
      ),
    ),
  );
  const rootNodes: Record<RootEndChainRoot, ReactNode> = {
    alias: createElement(
      MotionScope,
      { scope: alias, startAt: after(joined.start, 19) },
      observedShape("chain-alias", "Alias", 108, "#C00000", 100),
    ),
    afterAlias: createElement(
      MotionScope,
      { scope: afterAlias, startAt: alias.end },
      observedShape("chain-after-alias", "After alias", 136, "#4472C4", 100),
    ),
    afterAfterAlias: includeAfterAfterAlias
      ? createElement(
          MotionScope,
          { scope: afterAfterAlias, startAt: after(afterAlias.end, 7) },
          observedShape(
            "chain-after-after-alias",
            "After after alias",
            164,
            "#A5A5A5",
            100,
          ),
        )
      : null,
    host: source,
    follower: createElement(
      MotionScope,
      { scope: follower, startAt: host.end },
      observedShape("chain-follower", "Follower", 192, "#70AD47", 100),
    ),
    targetGroup: includeTargetGroup ? targetGroupSource : null,
  };
  const roots = createElement(
    Fragment,
    null,
    ...rootOrder.flatMap((root) =>
      rootNodes[root] === null ? [] : [rootNodes[root]],
    ),
  );
  return createElement(
    Deck,
    { width: 320, height: 280 },
    createElement(
      Slide,
      { slideId: "forward-condition-join-ordinary-end-chain" },
      explicitStep ? createElement(Step, { n: 2 }, roots) : roots,
    ),
  );
}

type EarlyHostParallelJoinOptions = {
  readonly longFirst?: boolean;
  readonly sourceOrder?: "first-late" | "late-first";
  readonly trigger?: "auto" | "click";
};

function EarlyHostParallelJoin({
  longFirst = false,
  sourceOrder = "first-late",
  trigger = "auto",
}: EarlyHostParallelJoinOptions) {
  const host = useMotionScope();
  const inner = useMotionScope();
  const late = useMotionScope();
  const parent = useMotionScope();
  const delayedConsumer = useMotionScope();
  const follower = useMotionScope();
  const firstSource = createElement(
    Text,
    {
      testId: "early-host-first",
      x: 10,
      y: 8,
      width: 300,
      height: 24,
      style: { fontSize: 12 },
      enter: fade({ by: "letter", duration: 1000, intervalMs: 500 }),
    },
    longFirst ? "ABCDEFGHIJKLMNOPQRSTUVWXYZ123456" : "AB",
  );
  const lateSource = createElement(
    MotionScope,
    { scope: late, delay: 1000 },
    createElement(
      Text,
      {
        testId: "early-host-late",
        x: 10,
        y: 40,
        width: 300,
        height: 24,
        style: { fontSize: 12 },
        enter: fade({ by: "letter", duration: 1000, intervalMs: 500 }),
      },
      "XY",
    ),
  );
  const sourceMembers =
    sourceOrder === "first-late"
      ? [firstSource, lateSource]
      : [lateSource, firstSource];
  const sourceTree = createElement(
    MotionScope,
    { scope: inner },
    ...sourceMembers,
  );
  const consumerTree = createElement(
    MotionScope,
    { scope: parent, startAt: inner.end },
    createElement(
      MotionScope,
      {
        scope: delayedConsumer,
        startAt: after(late.end, 11000),
        delay: 1000,
      },
      observedShape("early-host-consumer", "D", 76, "#ED7D31", 100),
    ),
  );
  const content = createElement(
    Fragment,
    null,
    createElement(
      MotionScope,
      { scope: host, mode: "parallel" },
      sourceTree,
      consumerTree,
    ),
    createElement(
      MotionScope,
      { scope: follower, startAt: host.end },
      observedShape("early-host-follower", "Follower", 112, "#70AD47", 50),
    ),
  );
  return createElement(
    Deck,
    { width: 320, height: 170 },
    createElement(
      Slide,
      { slideId: `early-host-${longFirst ? "long" : "short"}-${trigger}` },
      trigger === "click" ? createElement(Step, { n: 2 }, content) : content,
    ),
  );
}

type CrossBranchJoinOptions = {
  readonly kind: "cross-forward" | "cross-reverse" | "nonsequence";
  readonly trigger: "auto" | "click";
};

function CrossBranchConditionJoin({ kind, trigger }: CrossBranchJoinOptions) {
  const host = useMotionScope();
  const first = useMotionScope();
  const late = useMotionScope();
  const parent = useMotionScope();
  const consumer = useMotionScope();
  const follower = useMotionScope();
  const firstSource = createElement(
    MotionScope,
    { scope: first },
    unknownLetterText("cross-first", "AB", 8),
  );
  const lateSource = createElement(
    MotionScope,
    { scope: late, delay: 1000 },
    unknownLetterText("cross-late", "XY", 40),
  );
  const consumerMember = createElement(
    MotionScope,
    {
      scope: consumer,
      ...(kind === "nonsequence"
        ? { startAt: after(late.end, 11000), delay: 1000 }
        : { startAt: after(first.end, 11000) }),
    },
    observedShape("cross-consumer", "D", 76, "#ED7D31", 100),
  );
  const parentTree = createElement(
    MotionScope,
    {
      scope: parent,
      ...(kind === "nonsequence" ? { startAt: first.end } : {}),
      ...(kind === "nonsequence" ? {} : { mode: "sequence" as const }),
    },
    ...(kind === "nonsequence"
      ? [consumerMember]
      : [lateSource, consumerMember]),
  );
  const children =
    kind === "nonsequence"
      ? [firstSource, lateSource, parentTree]
      : kind === "cross-forward"
        ? [firstSource, parentTree]
        : [parentTree, firstSource];
  const content = createElement(
    Fragment,
    null,
    createElement(MotionScope, { scope: host, mode: "parallel" }, ...children),
    createElement(
      MotionScope,
      { scope: follower, startAt: host.end },
      observedShape("cross-follower", "Follower", 112, "#70AD47", 50),
    ),
  );
  return createElement(
    Deck,
    { width: 320, height: 170 },
    createElement(
      Slide,
      { slideId: `${kind}-${trigger}` },
      trigger === "click" ? createElement(Step, { n: 2 }, content) : content,
    ),
  );
}

type NestedMaxConditionJoinOptions = {
  readonly reverse: boolean;
  readonly trigger: "auto" | "click";
  readonly control?: boolean;
  readonly partialCoverage?: boolean;
};

function NestedMaxConditionJoin({
  reverse,
  trigger,
  control = false,
  partialCoverage = false,
}: NestedMaxConditionJoinOptions) {
  const host = useMotionScope();
  const sources = useMotionScope();
  const first = useMotionScope();
  const delayed = useMotionScope();
  const parent = useMotionScope();
  const child = useMotionScope();
  const firstSource = createElement(
    MotionScope,
    { scope: first },
    unknownLetterText("nested-max-first", "AB", 8),
  );
  const delayedSource = createElement(
    MotionScope,
    { scope: delayed, delay: 1000 },
    unknownLetterText("nested-max-delayed", "XY", 40),
  );
  const childScope = createElement(
    MotionScope,
    {
      scope: child,
      startAt: control
        ? after(parent.start, 1)
        : partialCoverage
          ? after(sources.end, 1)
          : after(delayed.end, 12000),
      ...(control || partialCoverage ? {} : { delay: 1 }),
    },
    observedShape("nested-max-consumer", "D", 76, "#ED7D31", 100),
  );
  const parentScope = createElement(
    MotionScope,
    { scope: parent, startAt: after(delayed.end, 11000) },
    childScope,
  );
  const sourceTree = createElement(
    MotionScope,
    { scope: sources, mode: "parallel" },
    ...(reverse ? [delayedSource, firstSource] : [firstSource, delayedSource]),
  );
  const content = createElement(
    MotionScope,
    { scope: host, mode: "sequence" },
    sourceTree,
    parentScope,
  );
  return createElement(
    Deck,
    { width: 320, height: 170 },
    createElement(
      Slide,
      { slideId: `nested-max-${reverse ? "reverse" : "forward"}-${trigger}` },
      trigger === "click" ? createElement(Step, { n: 1 }, content) : content,
    ),
  );
}

type DfsEndCollectorOptions = {
  readonly longFirst?: boolean;
  readonly branchOrder?: "p-first" | "q-first";
  readonly consumerOrder?: "p-first" | "q-first";
  readonly trigger?: "auto" | "click";
  readonly wordSources?: boolean;
  readonly firstText?: string;
  readonly secondText?: string;
  readonly emptyFirstLetter?: boolean;
  readonly pDelayMs?: number;
  readonly pStartAlias?: boolean;
};

function DfsEndCollectors({
  longFirst = false,
  branchOrder = "p-first",
  consumerOrder = "p-first",
  trigger = "auto",
  wordSources = false,
  firstText,
  secondText = "GHIJKLMN",
  emptyFirstLetter = false,
  pDelayMs,
  pStartAlias = false,
}: DfsEndCollectorOptions) {
  const host = useMotionScope();
  const p = useMotionScope();
  const q = useMotionScope();
  const a = useMotionScope();
  const b = useMotionScope();
  const d = useMotionScope();
  const e = useMotionScope();
  const afterP = useMotionScope();
  const afterQ = useMotionScope();
  const follower = useMotionScope();
  const pStartFollower = useMotionScope();
  const firstSourceText =
    firstText ?? (longFirst ? "ABCDEFGHIJKLMNOPQRSTUVWXYZ123456" : "AB");
  const secondSourceText = secondText;
  const firstSource = wordSources
    ? createElement(
        Text,
        {
          testId: "dfs-first",
          x: 8,
          y: 8,
          width: 240,
          height: 24,
          style: { fontSize: 12 },
          enter: fade({ by: "word", duration: 600, intervalMs: 200 }),
        },
        firstSourceText,
      )
    : unknownLetterText(
        "dfs-first",
        emptyFirstLetter ? "" : firstSourceText,
        8,
      );
  const secondSource = wordSources
    ? createElement(
        Text,
        {
          testId: "dfs-second",
          x: 40,
          y: 38,
          width: 240,
          height: 24,
          style: { fontSize: 12 },
          enter: fade({ by: "word", duration: 600, intervalMs: 200 }),
        },
        secondSourceText,
      )
    : unknownLetterText("dfs-second", secondSourceText, 40);

  const firstBranch = createElement(
    MotionScope,
    {
      scope: p,
      mode: "sequence",
      ...(pDelayMs === undefined ? {} : { delay: pDelayMs }),
    },
    createElement(MotionScope, { scope: a }, firstSource),
    createElement(
      MotionScope,
      { scope: d, startAt: after(b.end, 11000) },
      observedShape("dfs-d", "D", 76, "#ED7D31", 100),
    ),
  );
  const secondBranch = createElement(
    MotionScope,
    { scope: q, mode: "sequence" },
    createElement(MotionScope, { scope: b }, secondSource),
    createElement(
      MotionScope,
      { scope: e, startAt: after(a.end, 11000) },
      observedShape("dfs-e", "E", 104, "#4472C4", 100),
    ),
  );
  const branches =
    branchOrder === "p-first"
      ? [firstBranch, secondBranch]
      : [secondBranch, firstBranch];
  const pConsumer = createElement(
    MotionScope,
    { scope: afterP, startAt: p.end },
    observedShape("dfs-after-p", "After P", 132, "#70AD47", 100),
  );
  const qConsumer = createElement(
    MotionScope,
    { scope: afterQ, startAt: q.end },
    observedShape("dfs-after-q", "After Q", 160, "#A5A5A5", 100),
  );
  const consumers =
    consumerOrder === "p-first"
      ? [pConsumer, qConsumer]
      : [qConsumer, pConsumer];
  const startAliasConsumer = pStartAlias
    ? createElement(
        MotionScope,
        { scope: pStartFollower, startAt: p.start },
        observedShape("dfs-p-start-alias", "P start", 216, "#5B9BD5", 50),
      )
    : undefined;
  const content = createElement(
    Fragment,
    null,
    createElement(MotionScope, { scope: host, mode: "parallel" }, ...branches),
    ...consumers,
    startAliasConsumer,
    createElement(
      MotionScope,
      { scope: follower, startAt: host.end },
      observedShape("dfs-host-follower", "Host follower", 188, "#5B9BD5", 50),
    ),
  );
  return createElement(
    Deck,
    { width: 320, height: 240 },
    createElement(
      Slide,
      {
        slideId: `dfs-end-collectors-${longFirst ? "long" : "short"}-${branchOrder}-${consumerOrder}-${trigger}-${wordSources ? "word" : "letter"}`,
      },
      trigger === "click" ? createElement(Step, { n: 2 }, content) : content,
    ),
  );
}

function IncomparableUnknownJoin({
  firstText,
  trigger,
  unknownParentStart = false,
  parentStartGate = false,
  scopeDelayMs = 0,
  effectDelayMs = 0,
}: IncomparableJoinOptions) {
  const parentSource = useMotionScope();
  const outer = useMotionScope();
  const chain = useMotionScope();
  const inner = useMotionScope();
  const delayed = useMotionScope();
  const afterJoin = useMotionScope();
  const selfStarted = useMotionScope();
  const follower = useMotionScope();
  const source = createElement(
    Text,
    {
      testId: "join-parent-source",
      x: 10,
      y: 8,
      width: 300,
      height: 24,
      style: { fontSize: 12 },
      enter: fade({ by: "letter", duration: 1000, intervalMs: 500 }),
    },
    "AB",
  );
  const first = createElement(
    Text,
    {
      testId: "join-first",
      x: 10,
      y: 8,
      width: 300,
      height: 24,
      style: { fontSize: 12 },
      enter: fade({ by: "letter", duration: 1000, intervalMs: 500 }),
    },
    firstText,
  );
  const late = createElement(
    Text,
    {
      testId: "join-late",
      x: 10,
      y: 42,
      width: 300,
      height: 24,
      style: { fontSize: 12 },
      enter: fade({ by: "letter", duration: 1000, intervalMs: 500 }),
    },
    "XY",
  );
  const chainContents = createElement(
    MotionScope,
    { scope: inner },
    first,
    createElement(MotionScope, { scope: delayed, delay: 1000 }, late),
  );
  const orderedChain = createElement(
    MotionScope,
    { scope: chain, mode: "sequence" },
    chainContents,
    createElement(
      MotionScope,
      {
        scope: afterJoin,
        startAt: parentStartGate
          ? after(chain.start, 5000)
          : after(delayed.end, 11000),
        ...(scopeDelayMs === 0 ? {} : { delay: scopeDelayMs }),
      },
      observedShape("join-after", "D", 122, "#ED7D31", 100, effectDelayMs),
    ),
  );
  const outerScope = createElement(
    MotionScope,
    {
      scope: outer,
      ...(unknownParentStart ? { startAt: after(parentSource.end, 25) } : {}),
    },
    orderedChain,
    createElement(
      MotionScope,
      { scope: selfStarted, startAt: after(chain.start, 4000) },
      observedShape("join-independent", "C", 88, "#4472C4", 100),
    ),
  );
  const content = createElement(
    Fragment,
    null,
    unknownParentStart
      ? createElement(MotionScope, { scope: parentSource }, source)
      : null,
    outerScope,
    createElement(
      MotionScope,
      { scope: follower, startAt: outer.end },
      observedShape("join-follower", "E", 156, "#70AD47", 100),
    ),
  );
  const stepped =
    trigger === "explicit"
      ? createElement(Step, { n: 2 }, content)
      : trigger === "click"
        ? createElement(Step, null, content)
        : content;
  return createElement(
    Deck,
    { width: 320, height: 190 },
    createElement(Slide, { slideId: "incomparable-unknown-join" }, stepped),
  );
}

function UnknownCompletionEntryDeck() {
  const autoCollection = useMotionScope();
  const autoSource = useMotionScope();
  const autoFollower = useMotionScope();
  const clickCollection = useMotionScope();
  const clickSource = useMotionScope();
  const clickFollower = useMotionScope();

  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "unknown-completion-entry" },
      createElement(
        MotionScope,
        { scope: autoCollection },
        createElement(
          MotionScope,
          { scope: autoSource, mode: "stagger", each: 100 },
          unknownLetterText("auto-unknown-long", "ABCD", 8),
          unknownLetterText("auto-unknown-short", "AB", 48),
        ),
        createElement(
          MotionScope,
          { scope: autoFollower, startAt: after(autoSource.end, 25) },
          movingShape("auto-after-unknown", 80),
        ),
      ),
      createElement(
        Step,
        { n: 1 },
        createElement(
          MotionScope,
          { scope: clickCollection },
          createElement(
            MotionScope,
            { scope: clickSource, mode: "stagger", each: 100 },
            unknownLetterText("click-unknown-long", "WXYZ", 88),
            unknownLetterText("click-unknown-short", "WX", 128),
          ),
          createElement(
            MotionScope,
            {
              scope: clickFollower,
              startAt: after(clickSource.end, 30),
            },
            movingShape("click-after-unknown", 80),
          ),
        ),
      ),
    ),
  );
}

function EmptySequenceMember() {
  const sequence = useMotionScope();
  const empty = useMotionScope();
  const emptyCompletion = empty.cue("done");
  const follower = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "empty-sequence-member" },
      createElement(
        MotionScope,
        { scope: sequence, mode: "sequence" },
        movingShape("before-empty", 200),
        createElement(
          MotionScope,
          { scope: empty, publishEnd: emptyCompletion },
          null,
        ),
        movingShape("after-empty", 100),
      ),
      createElement(
        MotionScope,
        { scope: follower, startAt: after(emptyCompletion, 10) },
        movingShape("after-empty-completion", 100),
      ),
    ),
  );
}

function DelayedInactiveSequenceMember({
  sourceMode,
}: {
  readonly sourceMode: "empty" | "static" | "disabled";
}) {
  const sequence = useMotionScope();
  const inactive = useMotionScope();
  const afterSequence = useMotionScope();
  const afterInactive = useMotionScope();
  const inactiveContent =
    sourceMode === "empty"
      ? null
      : sourceMode === "static"
        ? createElement(Shape, {
            testId: "static-source",
            x: 0,
            y: 0,
            width: 20,
            height: 20,
          })
        : movingShape("disabled-source", 100);

  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "delayed-inactive-sequence-member" },
      createElement(
        MotionScope,
        { scope: sequence, mode: "sequence" },
        movingShape("first", 100),
        createElement(
          MotionScope,
          {
            scope: inactive,
            delay: 1000,
            motion: sourceMode !== "disabled",
          },
          inactiveContent,
        ),
        movingShape("second", 100),
      ),
      createElement(
        MotionScope,
        { scope: afterSequence, startAt: sequence.end },
        movingShape("after-sequence", 100),
      ),
      createElement(
        MotionScope,
        { scope: afterInactive, startAt: after(inactive.end, 25) },
        movingShape("after-inactive", 100),
      ),
    ),
  );
}

function EmptyStaggerMember() {
  const stagger = useMotionScope();
  const empty = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "empty-stagger-member" },
      createElement(
        MotionScope,
        { scope: stagger, mode: "stagger", each: 100 },
        movingShape("stagger-before-empty", 100),
        createElement(MotionScope, { scope: empty }, null),
        movingShape("stagger-after-empty", 100),
      ),
    ),
  );
}

function SelfStartedMember() {
  const group = useMotionScope();
  const leader = useMotionScope();
  const later = useMotionScope();
  const afterGroup = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "self-started-member" },
      createElement(
        MotionScope,
        { scope: group },
        createElement(
          MotionScope,
          { scope: leader },
          movingShape("leader", 100),
        ),
        createElement(
          MotionScope,
          { scope: later, startAt: after(leader.end, 20) },
          movingShape("self-started", 50),
        ),
      ),
      createElement(
        MotionScope,
        { scope: afterGroup, startAt: after(group.end, 10) },
        movingShape("after-group", 50),
      ),
    ),
  );
}

function BoxAndDisabledMembers() {
  const stagger = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "box-motion" },
      createElement(
        Box,
        { enter: fade({ delay: 800 }) },
        createElement(Shape, {
          testId: "box-default",
          x: 0,
          y: 0,
          width: 20,
          height: 20,
        }),
        createElement(Shape, {
          testId: "own-entry",
          x: 24,
          y: 0,
          width: 20,
          height: 20,
          enter: fade({ duration: 200 }),
        }),
      ),
      createElement(
        MotionScope,
        { scope: stagger, mode: "stagger", each: 100 },
        movingShape("stagger-first", 100),
        createElement(Shape, {
          testId: "static",
          x: 48,
          y: 0,
          width: 20,
          height: 20,
        }),
        createElement(
          Box,
          { motion: false },
          movingShape("disabled-descendant", 100),
        ),
        movingShape("stagger-second", 100),
      ),
    ),
  );
}

function ReusedNamedScope() {
  const source = useMotionScope();
  const done = source.cue("done");
  const follower = useMotionScope();
  return createElement(
    Fragment,
    null,
    createElement(
      MotionScope,
      { scope: source, publishEnd: done },
      movingShape("source", 100),
    ),
    createElement(
      MotionScope,
      { scope: follower, startAt: after(done, 10) },
      movingShape("follower", 100),
    ),
  );
}

function OrdinaryComposition() {
  return createElement(
    Fragment,
    null,
    movingShape("function-first", 100),
    createElement(
      Group,
      { testId: "native-group" },
      movingShape("group-first", 100),
      movingShape("group-second", 100),
    ),
  );
}

describe("组件编排作用域", () => {
  it("等待交错作用域的全部成员，并允许前向引用完成与命名时刻", async () => {
    const result = await build(createElement(StaggeredSource));
    const slide = result.report.slides[0];
    const motion = slide?.capabilities?.["motion"] as MotionObservation;
    const times = motion.steps[0]?.effects.map(({ target, startMs, endMs }) => [
      target.testId,
      startMs,
      endMs,
    ]);

    assert.deepEqual(times, [
      ["long-first", 0, 1000],
      ["medium-second", 100, 900],
      ["short-last", 200, 250],
      ["after-named-end", 1025, 1325],
      ["after-end", 1050, 1350],
    ]);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.animations?.[0]?.effects.map(({ targetId, delayMs }) => [
        slide?.objects.find(({ id }) => id === targetId)?.testId,
        delayMs,
      ]),
      [
        ["long-first", 0],
        ["medium-second", 100],
        ["short-last", 200],
        ["after-named-end", 1025],
        ["after-end", 1050],
      ],
    );
  });

  it("保留空作用域及其边界时刻，并把整段条件移除报告为缺失引用", async () => {
    for (const sourceMode of ["empty", "static", "disabled"] as const) {
      const result = await build(
        createElement(EmptyScopeDeck, { includeSource: true, sourceMode }),
      );
      const slide = result.report.slides[0];
      const motion = slide?.capabilities?.["motion"] as MotionObservation;
      assert.deepEqual(motion.scopes?.[0], {
        id: "scope-1",
        mode: "parallel",
        delayMs: 0,
        startMs: 0,
        endMs: 0,
        completion: "known",
        members: [],
      });
      assert.deepEqual(
        motion.steps[0]?.effects.map(({ target, startMs }) => [
          target.testId,
          startMs,
        ]),
        [
          ["from-named-start", 30],
          ["from-named-end", 40],
          ["from-start", 50],
          ["from-end", 80],
        ],
      );
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      assert.equal(slide?.objects.length, sourceMode === "empty" ? 4 : 5);
      assert.equal(
        file.slides[0]?.objects.length,
        sourceMode === "empty" ? 4 : 5,
      );
    }

    await assert.rejects(
      build(createElement(EmptyScopeDeck, { includeSource: false })),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "missing-cue");
        return true;
      },
    );
  });

  it("拒绝依赖环和明确跨点击步的作用域起止关系", async () => {
    await assert.rejects(
      build(createElement(CyclicScopes)),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "motion-cycle");
        return true;
      },
    );

    await assert.rejects(
      build(createElement(CrossStepScopes)),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "cross-step-motion");
        return true;
      },
    );
  });

  it("拒绝没有原生合并形式的不可比较未知起点", async () => {
    await assert.rejects(
      build(createElement(IncomparableUnknownStarts)),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "unknown-motion-start");
        assert.match(
          error.diagnostics[0]?.message ?? "",
          /不可比较的未知起点/u,
        );
        return true;
      },
    );
  });

  it("顺序作用域等待嵌套并行作用域的全部成员", async () => {
    const result = await build(createElement(SequencedScopes));
    const slide = result.report.slides[0];
    const motion = slide?.capabilities?.["motion"] as MotionObservation;
    assert.deepEqual(
      motion.steps[0]?.effects.map(({ target, startMs, endMs }) => [
        target.testId,
        startMs,
        endMs,
      ]),
      [
        ["first", 0, 200],
        ["parallel-short", 200, 300],
        ["parallel-long", 200, 900],
        ["last", 900, 1000],
      ],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.animations?.[0]?.effects.map(({ targetId, delayMs }) => [
        slide?.objects.find(({ id }) => id === targetId)?.testId,
        delayMs,
      ]),
      [
        ["first", 0],
        ["parallel-short", 200],
        ["parallel-long", 200],
        ["last", 900],
      ],
    );
  });

  it("未知时长的嵌套成员完成后才启动后继，并让外层等待全部成员", async () => {
    const result = await build(createElement(UnknownCompletionTree));
    const slide = result.report.slides[0];
    const motion = slide?.capabilities?.["motion"] as MotionObservation;

    assert.deepEqual(
      motion.steps[0]?.effects.map(({ target, startMs, endMs }) => [
        target.testId,
        startMs,
        endMs,
      ]),
      [
        ["long-first", 0, undefined],
        ["short-late", 1000, undefined],
        ["self-started-finite", 4000, 4100],
        ["finite-after-unknown", undefined, undefined],
        ["after-unknown-join", undefined, undefined],
      ],
    );
    const objectId = (testId: string) =>
      slide?.objects.find((object) => object.testId === testId)?.id;
    const containsObject = (scopeId: string | undefined, testId: string) =>
      motion.scopes
        ?.find(({ id }) => id === scopeId)
        ?.members.some(
          (member) =>
            member.kind === "object" && member.id === objectId(testId),
        );
    const innerScope = motion.scopes?.find(
      ({ members }) =>
        members.some(
          (member) =>
            member.kind === "object" && member.id === objectId("long-first"),
        ) &&
        members.some(
          (member) =>
            member.kind === "object" && member.id === objectId("short-late"),
        ),
    );
    const chainScope = motion.scopes?.find(
      ({ members }) =>
        members.some(
          (member) => member.kind === "scope" && member.id === innerScope?.id,
        ) &&
        members.some(
          (member) =>
            member.kind === "object" &&
            member.id === objectId("finite-after-unknown"),
        ),
    );
    const outerScope = motion.scopes?.find(
      ({ members }) =>
        members.some(
          (member) => member.kind === "scope" && member.id === chainScope?.id,
        ) &&
        members.some(
          (member) =>
            member.kind === "scope" &&
            containsObject(member.id, "self-started-finite"),
        ),
    );
    assert.deepEqual(
      [
        innerScope?.completion,
        innerScope?.endMs,
        chainScope?.completion,
        chainScope?.endMs,
      ],
      ["unknown", undefined, "unknown", undefined],
    );
    const finiteFollower = motion.steps[0]?.effects.find(
      ({ target }) => target.testId === "finite-after-unknown",
    );
    const outerWaiter = motion.steps[0]?.effects.find(
      ({ target }) => target.testId === "after-unknown-join",
    );
    assert.equal(
      finiteFollower?.startAfter &&
        "scopeId" in finiteFollower.startAfter &&
        finiteFollower.startAfter.scopeId,
      innerScope?.id,
    );
    assert.deepEqual(
      [outerScope?.completion, outerScope?.endMs],
      ["unknown", undefined],
    );
    assert.equal(
      outerWaiter?.startAfter &&
        "scopeId" in outerWaiter.startAfter &&
        outerWaiter.startAfter.scopeId,
      outerScope?.id,
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.extensions, []);
    const fileSlide = file.slides[0];
    const fileEffects = fileSlide?.animations?.[0]?.effects ?? [];
    assert.equal(fileEffects.length, 5);
    const fileEffect = (testId: string) =>
      fileEffects.find(({ targetId }) => targetId === objectId(testId));
    for (const testId of ["long-first", "short-late"]) {
      assert.equal(fileEffect(testId)?.by, "letter");
      assert.equal(fileEffect(testId)?.completion, "unknown");
    }

    const longEffect = fileEffect("long-first");
    const shortEffect = fileEffect("short-late");
    const finiteAfterUnknown = fileEffect("finite-after-unknown");
    const selfStarted = fileEffect("self-started-finite");
    const outerFollower = fileEffect("after-unknown-join");
    const containers = fileSlide?.animationContainers ?? [];
    const inner = containers.find(
      ({ children, endSync }) =>
        endSync === "all" &&
        children.some(
          ({ kind, timingId }) =>
            kind === "effect" && timingId === longEffect?.timingId,
        ) &&
        children.some(
          ({ kind, timingId }) =>
            kind === "effect" && timingId === shortEffect?.timingId,
        ),
    );
    assert.equal(inner?.children.length, 2);
    assert.equal(inner?.durationMs, undefined);
    assert.equal(finiteAfterUnknown?.startAfter?.timingId, inner?.timingId);
    const outer = containers.find(
      ({ timingId }) => timingId === outerFollower?.startAfter?.timingId,
    );
    assert.equal(outer?.endSync, "all");
    assert.equal(outer?.durationMs, undefined);
    assert.ok(
      outer?.children.some(
        ({ kind, timingId }) =>
          kind === "container" && timingId === inner?.timingId,
      ),
    );
    assert.ok(
      outer?.children.some(
        ({ kind, timingId }) =>
          kind === "effect" && timingId === finiteAfterUnknown?.timingId,
      ),
    );
    assert.ok(
      outer?.children.some(
        ({ kind, timingId }) =>
          kind === "effect" && timingId === selfStarted?.timingId,
      ),
    );
    assert.equal(outerFollower?.startAfter?.timingId, outer?.timingId);

    const archive = unzipSync(result.pptx);
    const slidePart = archive["ppt/slides/slide1.xml"];
    assert.ok(slidePart);
    const xml = strFromU8(slidePart);
    const innerStart = xml.indexOf(`<p:cTn id="${inner?.timingId}"`);
    const returnStart = xml.indexOf('<p:rtn val="all"></p:rtn>', innerStart);
    assert.ok(innerStart >= 0 && returnStart >= 0);
    const invalidReturnXml =
      xml.slice(0, returnStart) +
      '<p:rtn val="first"></p:rtn>' +
      xml.slice(returnStart + '<p:rtn val="all"></p:rtn>'.length);
    archive["ppt/slides/slide1.xml"] = strToU8(invalidReturnXml);
    const invalidRead = await validatePptx(zipSync(archive));
    assert.ok(
      invalidRead.errors.some(
        ({ code }) => code === "invalid-animation-end-sync",
      ),
    );
    assert.ok(
      invalidRead.errors.some(
        ({ code }) => code === "invalid-animation-end-dependency",
      ),
    );
  });

  it("单一未知 END 经连续起点别名保留真实来源、偏移及点击步入口", async () => {
    for (const trigger of ["auto", "click"] as const) {
      const result = await build(
        createElement(SingleSourceUnknownStartAliases, { trigger }),
      );
      const slide = result.report.slides[0];
      const motion = slide?.capabilities?.["motion"] as MotionObservation;
      const objectId = (testId: string) =>
        slide?.objects.find((object) => object.testId === testId)?.id;
      const scopeContaining = (testId: string) =>
        motion.scopes?.find(({ members }) =>
          members.some(
            (member) =>
              member.kind === "object" && member.id === objectId(testId),
          ),
        );
      const sourceScope = scopeContaining("single-alias-source");
      const parentScope = scopeContaining("single-alias-parent");
      const aliasScope = scopeContaining("single-alias-first");
      const secondAliasScope = scopeContaining("single-alias-second");
      const afterParentScope = scopeContaining("single-alias-after-parent-end");
      assert.ok(
        sourceScope &&
          parentScope &&
          aliasScope &&
          secondAliasScope &&
          afterParentScope,
      );
      assert.deepEqual(parentScope.startAfter, {
        scopeId: sourceScope.id,
        point: "end",
        offsetMs: 2500,
      });
      assert.deepEqual(aliasScope.startAfter, {
        scopeId: sourceScope.id,
        point: "end",
        offsetMs: 2519,
      });
      assert.deepEqual(secondAliasScope.startAfter, {
        scopeId: sourceScope.id,
        point: "end",
        offsetMs: 2539,
      });
      assert.deepEqual(afterParentScope.startAfter, {
        scopeId: parentScope.id,
        point: "end",
        offsetMs: 0,
      });
      assert.equal(sourceScope.startMs, 0);
      assert.equal(sourceScope.completion, "unknown");
      assert.equal(sourceScope.endMs, undefined);
      for (const scope of [parentScope, aliasScope, secondAliasScope]) {
        assert.equal(scope.completion, "unknown");
        assert.equal(scope.startMs, undefined);
        assert.equal(scope.endMs, undefined);
      }

      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      assert.deepEqual(file.extensions, []);
      const fileSlide = file.slides[0];
      const effects = fileSlide?.animations?.[0]?.effects ?? [];
      assert.equal(fileSlide?.animations?.[0]?.trigger, trigger);
      assert.equal(effects.length, 5);
      assert.equal(new Set(effects.map(({ targetId }) => targetId)).size, 5);
      assert.equal(fileSlide?.objects.length, 5);
      const effectFor = (testId: string) =>
        effects.find(({ targetId }) => targetId === objectId(testId));
      const sourceEffect = effectFor("single-alias-source");
      const parentEffect = effectFor("single-alias-parent");
      const aliasEffect = effectFor("single-alias-first");
      const secondAliasEffect = effectFor("single-alias-second");
      const afterParentEffect = effectFor("single-alias-after-parent-end");
      assert.ok(
        sourceEffect &&
          parentEffect &&
          aliasEffect &&
          secondAliasEffect &&
          afterParentEffect,
      );
      const containers = fileSlide?.animationContainers ?? [];
      const effectContainer = (timingId: number) =>
        containers.find(({ children }) =>
          children.some(
            ({ kind, timingId: childId }) =>
              kind === "effect" && childId === timingId,
          ),
        );
      const sourceContainer = effectContainer(sourceEffect.timingId);
      const parentContainer = effectContainer(parentEffect.timingId);
      assert.ok(sourceContainer && parentContainer);
      assert.deepEqual(
        [aliasEffect.startAfter?.timingId, aliasEffect.startAfter?.offsetMs],
        [sourceContainer.timingId, 2519],
      );
      assert.deepEqual(
        [
          secondAliasEffect.startAfter?.timingId,
          secondAliasEffect.startAfter?.offsetMs,
        ],
        [sourceContainer.timingId, 2539],
      );
      assert.deepEqual(
        [
          afterParentEffect.startAfter?.timingId,
          afterParentEffect.startAfter?.offsetMs,
        ],
        [parentContainer.timingId, 0],
      );
      assert.equal(sourceContainer.durationMs, undefined);
      assert.equal(parentContainer.durationMs, undefined);
      assert.equal(
        effectContainer(aliasEffect.timingId)?.durationMs,
        undefined,
      );
      assert.equal(
        containers.filter(
          ({ children, durationMs, endSync }) =>
            endSync === "all" &&
            durationMs === undefined &&
            children.length === 2 &&
            children.every(({ kind, timingId }) => {
              const gate = containers.find(
                ({ timingId: id }) => id === timingId,
              );
              return (
                kind === "container" &&
                gate?.durationMs === 0 &&
                gate.children.length === 0
              );
            }),
        ).length,
        0,
      );

      const slidePart = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
      assert.ok(slidePart);
      const xml = strFromU8(slidePart);
      assertTimingReferencesFollowDefinitions(xml);
      for (const [effect, delayMs] of [
        [aliasEffect, 2519],
        [secondAliasEffect, 2539],
      ] as const) {
        assert.match(
          startConditions(xml, effect.timingId),
          new RegExp(
            `<p:cond evt="end" delay="${delayMs}"><p:tn val="${sourceContainer.timingId}"><\\/p:tn><\\/p:cond>`,
            "u",
          ),
        );
      }
      assert.match(
        startConditions(xml, afterParentEffect.timingId),
        new RegExp(
          `<p:cond evt="end" delay="0"><p:tn val="${parentContainer.timingId}"><\\/p:tn><\\/p:cond>`,
          "u",
        ),
      );
    }
  });

  it("已知起点别名仍保留数值，无自然完成来源仍不可等待", async () => {
    const result = await build(
      createElement(SingleSourceUnknownStartAliases, {
        trigger: "auto",
        sourceCompletion: "known",
      }),
    );
    const slide = result.report.slides[0];
    const motion = slide?.capabilities?.["motion"] as MotionObservation;
    const objectId = (testId: string) =>
      slide?.objects.find((object) => object.testId === testId)?.id;
    const scopeContaining = (testId: string) =>
      motion.scopes?.find(({ members }) =>
        members.some(
          (member) =>
            member.kind === "object" && member.id === objectId(testId),
        ),
      );
    const sourceScope = scopeContaining("single-alias-source");
    const parentScope = scopeContaining("single-alias-parent");
    const aliasScope = scopeContaining("single-alias-first");
    const secondAliasScope = scopeContaining("single-alias-second");
    const afterParentScope = scopeContaining("single-alias-after-parent-end");
    assert.ok(
      sourceScope &&
        parentScope &&
        aliasScope &&
        secondAliasScope &&
        afterParentScope,
    );
    assert.deepEqual(
      [
        sourceScope.startMs,
        sourceScope.endMs,
        sourceScope.completion,
        parentScope.startMs,
        parentScope.endMs,
        parentScope.completion,
        aliasScope.startMs,
        aliasScope.endMs,
        aliasScope.completion,
        secondAliasScope.startMs,
        secondAliasScope.endMs,
        secondAliasScope.completion,
        afterParentScope.startMs,
      ],
      [
        0,
        100,
        "known",
        2600,
        2700,
        "known",
        2619,
        2719,
        "known",
        2639,
        2739,
        "known",
        2700,
      ],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.extensions, []);
    const fileSlide = file.slides[0];
    const effects = fileSlide?.animations?.[0]?.effects ?? [];
    assert.equal(effects.length, 5);
    assert.equal(fileSlide?.objects.length, 5);
    const aliasEffect = effects.find(
      ({ targetId }) => targetId === objectId("single-alias-first"),
    );
    const secondAliasEffect = effects.find(
      ({ targetId }) => targetId === objectId("single-alias-second"),
    );
    assert.ok(aliasEffect && secondAliasEffect);
    assert.deepEqual(
      [aliasEffect.delayMs, aliasEffect.startAfter],
      [2619, undefined],
    );
    assert.deepEqual(
      [secondAliasEffect.delayMs, secondAliasEffect.startAfter],
      [2639, undefined],
    );

    await assert.rejects(
      build(
        createElement(SingleSourceUnknownStartAliases, {
          trigger: "auto",
          sourceCompletion: "no-natural-completion",
        }),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "no-natural-completion");
        return true;
      },
    );
  });

  it("晚父起点主导别名后仍保留真实父门槛", async () => {
    const result = await build(
      createElement(SingleSourceAliasUnderLaterParent),
    );
    const slide = result.report.slides[0];
    const motion = slide?.capabilities?.["motion"] as MotionObservation;
    const objectId = (testId: string) =>
      slide?.objects.find(({ testId: id }) => id === testId)?.id;
    const scopeContaining = (testId: string) =>
      motion.scopes?.find(({ members }) =>
        members.some(
          (member) =>
            member.kind === "object" && member.id === objectId(testId),
        ),
      );
    const sourceScope = scopeContaining("later-alias-source");
    const parentScope = scopeContaining("later-alias-parent");
    const aliasScope = scopeContaining("later-alias-consumer");
    const lateParentScope = motion.scopes?.find(({ members }) =>
      members.some(
        (member) => member.kind === "scope" && member.id === aliasScope?.id,
      ),
    );
    assert.ok(sourceScope && parentScope && aliasScope && lateParentScope);
    assert.deepEqual(parentScope.startAfter, {
      scopeId: sourceScope.id,
      point: "end",
      offsetMs: 2500,
    });
    const lateGate = {
      scopeId: parentScope.id,
      point: "end" as const,
      offsetMs: 200,
    };
    assert.deepEqual(lateParentScope.startAfter, lateGate);
    assert.deepEqual(aliasScope.startAfter, lateGate);
    assert.equal(aliasScope.startMs, undefined);
    assert.equal(aliasScope.completion, "unknown");
    assert.equal(aliasScope.endMs, undefined);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.extensions, []);
    const fileSlide = file.slides[0];
    assert.equal(fileSlide?.objects.length, 3);
    const effects = fileSlide?.animations?.[0]?.effects ?? [];
    assert.equal(effects.length, 3);
    assert.equal(new Set(effects.map(({ targetId }) => targetId)).size, 3);
    const parentEffect = effects.find(
      ({ targetId }) => targetId === objectId("later-alias-parent"),
    );
    const aliasEffect = effects.find(
      ({ targetId }) => targetId === objectId("later-alias-consumer"),
    );
    assert.ok(parentEffect && aliasEffect);
    const containers = fileSlide?.animationContainers ?? [];
    const parentContainer = containers.find(({ children }) =>
      children.some(
        ({ kind, timingId }) =>
          kind === "effect" && timingId === parentEffect.timingId,
      ),
    );
    assert.ok(parentContainer);
    assert.deepEqual(
      [aliasEffect.startAfter?.timingId, aliasEffect.startAfter?.offsetMs],
      [parentContainer.timingId, 200],
    );
    const slidePart = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
    assert.ok(slidePart);
    const xml = strFromU8(slidePart);
    assertTimingReferencesFollowDefinitions(xml);
    assert.match(
      startConditions(xml, aliasEffect.timingId),
      new RegExp(
        `<p:cond evt="end" delay="200"><p:tn val="${parentContainer.timingId}"><\\/p:tn><\\/p:cond>`,
        "u",
      ),
    );
  });

  it("作用域起点可别名到已有条件集合完成并保留来源偏移", async () => {
    for (const parentStartGate of [false, true]) {
      const result = await build(
        createElement(ConditionJoinStartAlias, { parentStartGate }),
      );
      const slide = result.report.slides[0];
      const motion = slide?.capabilities?.["motion"] as MotionObservation;
      const objectId = (testId: string) =>
        slide?.objects.find(({ testId: id }) => id === testId)?.id;
      const scopeContaining = (testId: string) =>
        motion.scopes?.find(({ members }) =>
          members.some(
            (member) =>
              member.kind === "object" && member.id === objectId(testId),
          ),
        );
      const delayedScope = scopeContaining("alias-late");
      const innerScope = motion.scopes?.find(
        ({ members }) =>
          members.some(
            (member) =>
              member.kind === "object" && member.id === objectId("alias-first"),
          ) &&
          members.some(
            (member) =>
              member.kind === "scope" && member.id === delayedScope?.id,
          ),
      );
      const joinedScope = scopeContaining("alias-join-consumer");
      const aliasScope = scopeContaining("alias-consumer");
      const parentAliasScope = scopeContaining("parent-alias-consumer");
      const aliasEffect = motion.steps[0]?.effects.find(
        ({ target }) => target.testId === "alias-consumer",
      );
      const parentAliasEffect = motion.steps[0]?.effects.find(
        ({ target }) => target.testId === "parent-alias-consumer",
      );
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      assert.deepEqual(file.extensions, []);
      const fileSlide = file.slides[0];
      const fileEffect = fileSlide?.animations?.[0]?.effects.find(
        ({ targetId }) => targetId === objectId("alias-consumer"),
      );
      const parentFileEffect = fileSlide?.animations?.[0]?.effects.find(
        ({ targetId }) => targetId === objectId("parent-alias-consumer"),
      );
      const containers = fileSlide?.animationContainers ?? [];
      const childrenOf = (parentId: number | undefined) =>
        containers.find(({ timingId }) => timingId === parentId);
      const joins = containers.filter(
        ({ children, durationMs, endSync }) =>
          endSync === "all" &&
          durationMs === undefined &&
          children.length === 2 &&
          children.every(({ timingId, kind }) => {
            const gate = childrenOf(timingId);
            return (
              kind === "container" &&
              gate?.durationMs === 0 &&
              gate.children.length === 0
            );
          }),
      );
      const join = joins[0];
      assert.ok(
        delayedScope &&
          innerScope &&
          joinedScope &&
          aliasScope &&
          parentAliasScope &&
          aliasEffect &&
          parentAliasEffect &&
          fileEffect &&
          parentFileEffect &&
          join,
      );
      assert.equal(joins.length, 1);
      const chainScope = motion.scopes?.find(
        ({ members }) =>
          members.some(
            (member) => member.kind === "scope" && member.id === innerScope.id,
          ) &&
          members.some(
            (member) => member.kind === "scope" && member.id === joinedScope.id,
          ),
      );
      assert.ok(chainScope);
      const conditionSources = [
        { scopeId: innerScope.id, point: "end" as const, offsetMs: 0 },
        parentStartGate
          ? {
              scopeId: chainScope.id,
              point: "start" as const,
              offsetMs: 5000,
            }
          : {
              scopeId: delayedScope.id,
              point: "end" as const,
              offsetMs: 11000,
            },
      ];
      const shift = (offsetMs: number) =>
        conditionSources.map((dependency) => ({
          ...dependency,
          offsetMs: dependency.offsetMs + offsetMs,
        }));
      assert.deepEqual(joinedScope.startAfterAll, conditionSources);
      assert.deepEqual(aliasScope.startAfterAll, [...shift(42)]);
      assert.deepEqual(aliasEffect.startAfterAll, [...shift(73)]);
      assert.deepEqual(parentAliasScope.startAfterAll, shift(20));
      assert.deepEqual(parentAliasEffect.startAfterAll, shift(25));
      assert.deepEqual(
        [aliasScope.startMs, aliasScope.endMs, aliasScope.completion],
        [undefined, undefined, "unknown"],
      );
      assert.deepEqual(
        [
          parentAliasScope.startMs,
          parentAliasScope.endMs,
          parentAliasScope.completion,
        ],
        [undefined, undefined, "unknown"],
      );
      assert.deepEqual(
        [fileEffect.startAfter?.timingId, fileEffect.startAfter?.offsetMs],
        [join.timingId, 73],
      );
      assert.deepEqual(
        [
          parentFileEffect.startAfter?.timingId,
          parentFileEffect.startAfter?.offsetMs,
        ],
        [join.timingId, 25],
      );
      assert.equal(
        fileSlide?.objects.length,
        5,
        "别名复用不得创建用于计时的 Shape",
      );

      const archive = unzipSync(result.pptx);
      const slidePart = archive["ppt/slides/slide1.xml"];
      assert.ok(slidePart);
      const xml = strFromU8(slidePart);
      const timingNodeIndex = (timingId: number | undefined) =>
        xml.indexOf(`<p:cTn id="${timingId}"`);
      const joinIndex = timingNodeIndex(join.timingId);
      const gateIndexes = join.children.flatMap((gate) => {
        const sourceId = childrenOf(gate.timingId)?.startAfter?.timingId;
        return sourceId === undefined ? [] : [timingNodeIndex(sourceId)];
      });
      assert.ok(joinIndex >= 0);
      assert.ok(gateIndexes.every((index) => index >= 0 && index < joinIndex));
      if (parentStartGate) {
        const numericGate = join.children.find(
          (gate) => childrenOf(gate.timingId)?.startAfter === undefined,
        );
        assert.ok(numericGate);
        assert.match(
          startConditions(xml, numericGate.timingId),
          /<p:cond delay="5000"><\/p:cond>/u,
        );
      }
      assert.ok(timingNodeIndex(fileEffect.timingId) > joinIndex);
      assert.match(
        startConditions(xml, fileEffect.timingId),
        new RegExp(
          `<p:cond evt="end" delay="73"><p:tn val="${join.timingId}"><\\/p:tn><\\/p:cond>`,
          "u",
        ),
      );
      assert.match(
        startConditions(xml, parentFileEffect.timingId),
        new RegExp(
          `<p:cond evt="end" delay="25"><p:tn val="${join.timingId}"><\\/p:tn><\\/p:cond>`,
          "u",
        ),
      );
    }
  });

  it("点击步中的多来源起点别名仍复用唯一条件集合", async () => {
    const result = await build(
      createElement(ConditionJoinStartAlias, {
        parentStartGate: false,
        trigger: "click",
      }),
    );
    const slide = result.report.slides[0];
    const motion = slide?.capabilities?.["motion"] as MotionObservation;
    const objectId = (testId: string) =>
      slide?.objects.find(({ testId: id }) => id === testId)?.id;
    const scopeContaining = (testId: string) =>
      motion.scopes?.find(({ members }) =>
        members.some(
          (member) =>
            member.kind === "object" && member.id === objectId(testId),
        ),
      );
    const delayedScope = scopeContaining("alias-late");
    const joinedScope = scopeContaining("alias-join-consumer");
    const aliasScope = scopeContaining("alias-consumer");
    const innerScope = motion.scopes?.find(
      ({ members }) =>
        members.some(
          (member) =>
            member.kind === "object" && member.id === objectId("alias-first"),
        ) &&
        members.some(
          (member) => member.kind === "scope" && member.id === delayedScope?.id,
        ),
    );
    assert.ok(delayedScope && joinedScope && aliasScope && innerScope);
    const sources = [
      { scopeId: innerScope.id, point: "end" as const, offsetMs: 0 },
      { scopeId: delayedScope.id, point: "end" as const, offsetMs: 11000 },
    ];
    assert.deepEqual(joinedScope.startAfterAll, sources);
    assert.deepEqual(
      aliasScope.startAfterAll,
      sources.map((source) => ({ ...source, offsetMs: source.offsetMs + 42 })),
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.extensions, []);
    const fileSlide = file.slides[0];
    assert.equal(fileSlide?.animations?.[0]?.trigger, "click");
    assert.equal(fileSlide?.objects.length, 5);
    const effects = fileSlide?.animations?.[0]?.effects ?? [];
    assert.equal(effects.length, 5);
    assert.equal(new Set(effects.map(({ targetId }) => targetId)).size, 5);
    const aliasEffect = effects.find(
      ({ targetId }) => targetId === objectId("alias-consumer"),
    );
    assert.ok(aliasEffect);
    const containers = fileSlide?.animationContainers ?? [];
    const joins = containers.filter(
      ({ children, durationMs, endSync }) =>
        endSync === "all" &&
        durationMs === undefined &&
        children.length === 2 &&
        children.every(({ kind, timingId }) => {
          const gate = containers.find(({ timingId: id }) => id === timingId);
          return (
            kind === "container" &&
            gate?.durationMs === 0 &&
            gate.children.length === 0
          );
        }),
    );
    assert.equal(joins.length, 1);
    assert.deepEqual(
      [aliasEffect.startAfter?.timingId, aliasEffect.startAfter?.offsetMs],
      [joins[0]?.timingId, 73],
    );
    const slidePart = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
    assert.ok(slidePart);
    const xml = strFromU8(slidePart);
    assertTimingReferencesFollowDefinitions(xml);
    assert.match(
      startConditions(xml, aliasEffect.timingId),
      new RegExp(
        `<p:cond evt="end" delay="73"><p:tn val="${joins[0]?.timingId}"><\\/p:tn><\\/p:cond>`,
        "u",
      ),
    );
  });

  it("父作用域晚于条件集合完成时保留未知起点错误", async () => {
    await assert.rejects(
      build(createElement(LateParentRetainedAlias)),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.ok(
          error.diagnostics.some(({ code }) => code === "unknown-motion-start"),
        );
        return true;
      },
    );
  });

  it("父作用域比同一条件集合晚 10 毫秒时把 alias 精确化为 9 毫秒", async () => {
    const result = await build(createElement(SameJoinParentAlias));
    const slide = result.report.slides[0];
    const motion = slide?.capabilities?.["motion"] as MotionObservation;
    const objectId = (testId: string) =>
      slide?.objects.find(({ testId: id }) => id === testId)?.id;
    const scopeContaining = (testId: string) =>
      motion.scopes?.find(({ members }) =>
        members.some(
          (member) =>
            member.kind === "object" && member.id === objectId(testId),
        ),
      );
    const firstScope = scopeContaining("same-parent-first");
    const delayedScope = scopeContaining("same-parent-delayed");
    const aliasScope = scopeContaining("same-parent-alias");
    const parentScope = motion.scopes?.find(({ members }) =>
      members.some(
        (member) => member.kind === "scope" && member.id === aliasScope?.id,
      ),
    );
    const afterParentScope = scopeContaining("same-parent-after");
    assert.ok(
      firstScope &&
        delayedScope &&
        parentScope &&
        aliasScope &&
        afterParentScope,
    );
    const conditionSources = [
      { scopeId: firstScope.id, point: "end" as const, offsetMs: 0 },
      { scopeId: delayedScope.id, point: "end" as const, offsetMs: 11000 },
    ];
    assert.deepEqual(parentScope.startAfterAll, [
      { ...conditionSources[0], offsetMs: 10 },
      { ...conditionSources[1], offsetMs: 11010 },
    ]);
    assert.deepEqual(aliasScope.startAfterAll, [
      { ...conditionSources[0], offsetMs: 19 },
      { ...conditionSources[1], offsetMs: 11019 },
    ]);
    assert.deepEqual(
      [aliasScope.startMs, aliasScope.endMs, aliasScope.completion],
      [undefined, undefined, "unknown"],
    );
    assert.deepEqual(afterParentScope.startAfter, {
      scopeId: parentScope.id,
      point: "end",
      offsetMs: 0,
    });

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.extensions, []);
    const fileSlide = file.slides[0];
    const fileEffects = fileSlide?.animations?.[0]?.effects ?? [];
    const aliasEffect = fileEffects.find(
      ({ targetId }) => targetId === objectId("same-parent-alias"),
    );
    const afterParentEffect = fileEffects.find(
      ({ targetId }) => targetId === objectId("same-parent-after"),
    );
    assert.ok(aliasEffect && afterParentEffect);
    const parent = fileSlide?.animationContainers?.find(
      ({ timingId }) => timingId === afterParentEffect.startAfter?.timingId,
    );
    assert.ok(parent);
    assert.equal(parent.startAfter?.offsetMs, 10);
    assert.equal(parent.startAfter?.source, "container");
    assert.equal(aliasEffect.delayMs, 19);
    assert.deepEqual(
      [aliasEffect.startAfter?.timingId, aliasEffect.startAfter?.offsetMs],
      [parent.startAfter?.timingId, 10],
    );

    const slidePart = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
    assert.ok(slidePart);
    const xml = strFromU8(slidePart);
    assertTimingReferencesFollowDefinitions(xml);
    assert.match(
      startConditions(xml, aliasEffect.timingId),
      /<p:cond delay="9"><\/p:cond>/u,
    );
  });

  it("重排独立起点的计划兄弟以先定义 alias 的真实条件来源", async () => {
    const result = await build(createElement(ForwardConditionJoinAlias));
    const slide = result.report.slides[0];
    const motion = slide?.capabilities?.["motion"] as MotionObservation;
    const objectId = (testId: string) =>
      slide?.objects.find(({ testId: id }) => id === testId)?.id;
    const aliasScope = motion.scopes?.find(({ members }) =>
      members.some(
        (member) =>
          member.kind === "object" &&
          member.id === objectId("forward-alias-consumer"),
      ),
    );
    const delayedScope = motion.scopes?.find(({ members }) =>
      members.some(
        (member) =>
          member.kind === "object" &&
          member.id === objectId("forward-alias-delayed"),
      ),
    );
    const innerScope = motion.scopes?.find(
      ({ members }) =>
        members.some(
          (member) =>
            member.kind === "object" &&
            member.id === objectId("forward-alias-first"),
        ) &&
        members.some(
          (member) => member.kind === "scope" && member.id === delayedScope?.id,
        ),
    );
    const aliasEffect = motion.steps[0]?.effects.find(
      ({ target }) => target.testId === "forward-alias-consumer",
    );
    assert.ok(aliasScope && innerScope && delayedScope && aliasEffect);
    assert.deepEqual(aliasScope.startAfterAll, [
      { scopeId: innerScope.id, point: "end", offsetMs: 19 },
      { scopeId: delayedScope.id, point: "end", offsetMs: 11019 },
    ]);
    assert.deepEqual(
      [aliasScope.startMs, aliasScope.endMs, aliasScope.completion],
      [undefined, undefined, "unknown"],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.extensions, []);
    const fileSlide = file.slides[0];
    const fileEffect = fileSlide?.animations?.[0]?.effects.find(
      ({ targetId }) => targetId === objectId("forward-alias-consumer"),
    );
    const containers = fileSlide?.animationContainers ?? [];
    const childrenOf = (parentId: number | undefined) =>
      containers.find(({ timingId }) => timingId === parentId);
    const joins = containers.filter(
      ({ children, durationMs, endSync }) =>
        endSync === "all" &&
        durationMs === undefined &&
        children.length === 2 &&
        children.every(({ timingId, kind }) => {
          const gate = childrenOf(timingId);
          return (
            kind === "container" &&
            gate?.durationMs === 0 &&
            gate.children.length === 0
          );
        }),
    );
    const join = joins[0];
    assert.ok(fileEffect && join);
    assert.equal(joins.length, 1);
    assert.deepEqual(
      [fileEffect.startAfter?.timingId, fileEffect.startAfter?.offsetMs],
      [join.timingId, 19],
    );

    const slidePart = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
    assert.ok(slidePart);
    const xml = strFromU8(slidePart);
    assertTimingReferencesFollowDefinitions(xml);
    const timingNodeIndex = (timingId: number | undefined) =>
      xml.indexOf(`<p:cTn id="${timingId}"`);
    const joinIndex = timingNodeIndex(join.timingId);
    const aliasIndex = timingNodeIndex(fileEffect.timingId);
    const sourceIndexes = join.children.flatMap((gate) => {
      const sourceId = childrenOf(gate.timingId)?.startAfter?.timingId;
      return sourceId === undefined ? [] : [timingNodeIndex(sourceId)];
    });
    assert.ok(joinIndex >= 0);
    assert.ok(sourceIndexes.length === 2);
    assert.ok(sourceIndexes.every((index) => index >= 0 && index < joinIndex));
    assert.ok(joinIndex < aliasIndex);
    assert.match(
      startConditions(xml, fileEffect.timingId),
      new RegExp(
        `<p:cond evt="end" delay="19"><p:tn val="${join.timingId}"><\\/p:tn><\\/p:cond>`,
        "u",
      ),
    );
  });

  it("排序 alias 时同时保留普通作用域完成依赖", async () => {
    const variants: readonly RootEndChainOptions[] = [
      {
        rootOrder: ["alias", "afterAlias", "host", "follower"],
      },
      {
        rootOrder: ["host", "afterAlias", "alias", "follower"],
      },
      {
        rootOrder: [
          "afterAfterAlias",
          "alias",
          "afterAlias",
          "host",
          "follower",
          "targetGroup",
        ],
        includeAfterAfterAlias: true,
        includeTargetGroup: true,
        explicitStep: true,
      },
    ];
    for (const options of variants) {
      const result = await build(
        createElement(ForwardConditionJoinOrdinaryEndChain, options),
      );
      const slide = result.report.slides[0];
      const motion = slide?.capabilities?.["motion"] as MotionObservation;
      const objectId = (testId: string) =>
        slide?.objects.find(({ testId: id }) => id === testId)?.id;
      const scopeContaining = (testId: string) =>
        motion.scopes?.find(({ members }) =>
          members.some(
            (member) =>
              member.kind === "object" && member.id === objectId(testId),
          ),
        );
      const aliasScope = scopeContaining("chain-alias");
      const afterAliasScope = scopeContaining("chain-after-alias");
      assert.ok(aliasScope && afterAliasScope);
      assert.deepEqual(afterAliasScope.startAfter, {
        scopeId: aliasScope.id,
        point: "end",
        offsetMs: 0,
      });
      if (options.includeAfterAfterAlias === true) {
        const afterAfterAliasScope = scopeContaining("chain-after-after-alias");
        assert.ok(afterAfterAliasScope);
        assert.deepEqual(afterAfterAliasScope.startAfter, {
          scopeId: afterAliasScope.id,
          point: "end",
          offsetMs: 7,
        });
      }
      if (options.includeTargetGroup === true) {
        const targetGroupFollower = scopeContaining(
          "chain-target-group-follower",
        );
        assert.ok(targetGroupFollower);
        assert.deepEqual(targetGroupFollower.startAfter, {
          targetId: objectId("chain-target-group-source"),
          group: "effects",
          point: "end",
          offsetMs: 0,
        });
      }

      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      assert.deepEqual(file.extensions, []);
      assert.equal(
        file.slides[0]?.animations?.[0]?.trigger,
        options.explicitStep === true ? "click" : "auto",
      );
      const slidePart = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
      assert.ok(slidePart);
      const xml = strFromU8(slidePart);
      assertTimingReferencesFollowDefinitions(xml);
      if (options.includeTargetGroup === true) {
        const fileSlide = file.slides[0];
        const groupFollowerEffect = fileSlide?.animations?.[0]?.effects.find(
          ({ targetId }) =>
            targetId === objectId("chain-target-group-follower"),
        );
        const group = fileSlide?.animationContainers?.find(
          ({ timingId }) =>
            timingId === groupFollowerEffect?.startAfter?.timingId,
        );
        assert.ok(group && groupFollowerEffect);
        assert.equal(group.endSync, "all");
        assert.equal(group.children.length, 2);
        assert.deepEqual(
          [
            groupFollowerEffect.startAfter?.timingId,
            groupFollowerEffect.startAfter?.offsetMs,
          ],
          [group.timingId, 0],
        );
      }
    }
  });

  it("早宿主条件集合保留来源声明顺序、未知时长和点击步入口", async () => {
    const variants: readonly EarlyHostParallelJoinOptions[] = [
      { longFirst: false, sourceOrder: "first-late", trigger: "auto" },
      { longFirst: false, sourceOrder: "late-first", trigger: "auto" },
      { longFirst: false, sourceOrder: "first-late", trigger: "click" },
      { longFirst: false, sourceOrder: "late-first", trigger: "click" },
      { longFirst: true, sourceOrder: "first-late", trigger: "auto" },
      { longFirst: true, sourceOrder: "late-first", trigger: "auto" },
      { longFirst: true, sourceOrder: "first-late", trigger: "click" },
      { longFirst: true, sourceOrder: "late-first", trigger: "click" },
    ];
    for (const options of variants) {
      const result = await build(createElement(EarlyHostParallelJoin, options));
      const slide = result.report.slides[0];
      const motion = slide?.capabilities?.["motion"] as MotionObservation;
      const objectId = (testId: string) =>
        slide?.objects.find(({ testId: id }) => id === testId)?.id;
      const scopeContaining = (testId: string) =>
        motion.scopes?.find(({ members }) =>
          members.some(
            (member) =>
              member.kind === "object" && member.id === objectId(testId),
          ),
        );
      const firstScope = scopeContaining("early-host-first");
      const lateScope = scopeContaining("early-host-late");
      const consumerScope = scopeContaining("early-host-consumer");
      const consumerParentScope = motion.scopes?.find(({ members }) =>
        members.some(
          (member) =>
            member.kind === "scope" && member.id === consumerScope?.id,
        ),
      );
      const hostScope = motion.scopes?.find(
        ({ members }) =>
          members.some(
            (member) => member.kind === "scope" && member.id === firstScope?.id,
          ) &&
          members.some(
            (member) =>
              member.kind === "scope" && member.id === consumerParentScope?.id,
          ),
      );
      const followerScope = scopeContaining("early-host-follower");
      assert.ok(
        firstScope &&
          lateScope &&
          consumerScope &&
          consumerParentScope &&
          hostScope &&
          followerScope,
      );
      assert.deepEqual(consumerScope.startAfterAll, [
        { scopeId: firstScope.id, point: "end", offsetMs: 1000 },
        { scopeId: lateScope.id, point: "end", offsetMs: 12000 },
      ]);
      assert.deepEqual(consumerScope.completion, "unknown");
      assert.deepEqual(consumerParentScope.startAfter, {
        scopeId: firstScope.id,
        point: "end",
        offsetMs: 0,
      });
      assert.deepEqual(followerScope.startAfter, {
        scopeId: hostScope.id,
        point: "end",
        offsetMs: 0,
      });

      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      assert.deepEqual(file.extensions, []);
      const fileSlide = file.slides[0];
      assert.equal(
        fileSlide?.animations?.[0]?.trigger,
        options.trigger ?? "auto",
      );
      const effects = fileSlide?.animations?.[0]?.effects ?? [];
      assert.equal(effects.length, 4);
      const effect = (testId: string) =>
        effects.find(({ targetId }) => targetId === objectId(testId));
      const firstEffect = effect("early-host-first");
      const lateEffect = effect("early-host-late");
      const consumerEffect = effect("early-host-consumer");
      const followerEffect = effect("early-host-follower");
      assert.ok(firstEffect && lateEffect && consumerEffect && followerEffect);
      for (const source of [firstEffect, lateEffect]) {
        assert.equal(source.by, "letter");
        assert.equal(source.completion, "unknown");
      }

      const containers = fileSlide?.animationContainers ?? [];
      const consumerParent = containers.find(({ children }) =>
        children.some(
          ({ kind, timingId }) =>
            kind === "effect" && timingId === consumerEffect.timingId,
        ),
      );
      assert.ok(consumerEffect.startAfter && consumerParent);
      const join = containers.find(
        ({ timingId }) => timingId === consumerEffect.startAfter?.timingId,
      );
      const firstSourceContainer = containers.find(({ children }) =>
        children.some(
          ({ kind, timingId }) =>
            kind === "effect" && timingId === firstEffect.timingId,
        ),
      );
      const host = containers.find(
        ({ timingId }) => timingId === join?.parentTimingId,
      );
      assert.ok(join && host && firstSourceContainer);
      assert.equal(join.endSync, "all");
      assert.equal(join.durationMs, undefined);
      assert.equal(join.children.length, 2);
      assert.equal(consumerParent.parentTimingId, host.timingId);
      assert.equal(
        consumerParent.startAfter?.timingId,
        firstSourceContainer.timingId,
      );
      assert.equal(consumerParent.startAfter?.point, "end");
      assert.equal(consumerParent.startAfter?.offsetMs, 0);
      assert.equal(followerEffect.startAfter?.timingId, host.timingId);
      assert.equal(followerEffect.startAfter?.point, "end");
      assert.ok(
        host.children.some(
          ({ kind, timingId }) =>
            kind === "container" && timingId === join.timingId,
        ),
      );

      const slidePart = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
      assert.ok(slidePart);
      const xml = strFromU8(slidePart);
      assertTimingReferencesFollowDefinitions(xml);
      assert.equal(
        firstMainSequenceEntryNodeType(xml, 0),
        options.trigger === "click" ? "clickEffect" : "afterEffect",
      );
    }
  });

  it("跨并行分支放置条件集合并按来源关系排序两种声明顺序", async () => {
    const variants: readonly CrossBranchJoinOptions[] = [
      { kind: "cross-forward", trigger: "auto" },
      { kind: "cross-reverse", trigger: "auto" },
      { kind: "nonsequence", trigger: "auto" },
      { kind: "cross-forward", trigger: "click" },
      { kind: "cross-reverse", trigger: "click" },
      { kind: "nonsequence", trigger: "click" },
    ];
    for (const options of variants) {
      const result = await build(
        createElement(CrossBranchConditionJoin, options),
      );
      const slide = result.report.slides[0];
      const motion = slide?.capabilities?.["motion"] as MotionObservation;
      const objectId = (testId: string) =>
        slide?.objects.find(({ testId: id }) => id === testId)?.id;
      const scopeContaining = (testId: string) =>
        motion.scopes?.find(({ members }) =>
          members.some(
            (member) =>
              member.kind === "object" && member.id === objectId(testId),
          ),
        );
      const firstScope = scopeContaining("cross-first");
      const lateScope = scopeContaining("cross-late");
      const consumerScope = scopeContaining("cross-consumer");
      const consumerParent = motion.scopes?.find(({ members }) =>
        members.some(
          (member) =>
            member.kind === "scope" && member.id === consumerScope?.id,
        ),
      );
      assert.ok(firstScope && lateScope && consumerScope && consumerParent);
      const conditionSources =
        options.kind === "nonsequence"
          ? [
              {
                scopeId: firstScope.id,
                point: "end" as const,
                offsetMs: 1000,
              },
              {
                scopeId: lateScope.id,
                point: "end" as const,
                offsetMs: 12000,
              },
            ]
          : [
              {
                scopeId: lateScope.id,
                point: "end" as const,
                offsetMs: 0,
              },
              {
                scopeId: firstScope.id,
                point: "end" as const,
                offsetMs: 11000,
              },
            ];
      assert.deepEqual(consumerScope.startAfterAll, conditionSources);
      assert.equal(consumerScope.completion, "unknown");

      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      assert.deepEqual(file.extensions, []);
      const fileSlide = file.slides[0];
      assert.equal(fileSlide?.animations?.[0]?.trigger, options.trigger);
      const effects = fileSlide?.animations?.[0]?.effects ?? [];
      assert.equal(effects.length, 4);
      const effect = (testId: string) =>
        effects.find(({ targetId }) => targetId === objectId(testId));
      const firstEffect = effect("cross-first");
      const lateEffect = effect("cross-late");
      const consumerEffect = effect("cross-consumer");
      const followerEffect = effect("cross-follower");
      assert.ok(firstEffect && lateEffect && consumerEffect && followerEffect);
      for (const source of [firstEffect, lateEffect]) {
        assert.equal(source.by, "letter");
        assert.equal(source.completion, "unknown");
      }

      const containers = fileSlide?.animationContainers ?? [];
      const consumerParentContainer = containers.find(({ children }) =>
        children.some(
          ({ kind, timingId }) =>
            kind === "effect" && timingId === consumerEffect.timingId,
        ),
      );
      assert.ok(consumerEffect.startAfter && consumerParentContainer);
      const join = containers.find(
        ({ timingId }) => timingId === consumerEffect.startAfter?.timingId,
      );
      assert.ok(join);
      const joinHost = containers.find(
        ({ timingId }) => timingId === join.parentTimingId,
      );
      const outerHost = containers.find(
        ({ timingId }) => timingId === consumerParentContainer.parentTimingId,
      );
      assert.ok(joinHost && outerHost);
      assert.equal(join.endSync, "all");
      assert.equal(join.children.length, 2);
      assert.equal(join.durationMs, undefined);
      if (options.kind === "nonsequence") {
        assert.equal(consumerParentContainer.parentTimingId, joinHost.timingId);
        assert.equal(consumerParentContainer.startAfter?.offsetMs, 0);
      } else {
        assert.equal(joinHost.timingId, consumerParentContainer.timingId);
      }
      assert.equal(followerEffect.startAfter?.timingId, outerHost.timingId);

      const firstContainer = containers.find(({ children }) =>
        children.some(
          ({ kind, timingId }) =>
            kind === "effect" && timingId === firstEffect.timingId,
        ),
      );
      const lateContainer = containers.find(({ children }) =>
        children.some(
          ({ kind, timingId }) =>
            kind === "effect" && timingId === lateEffect.timingId,
        ),
      );
      assert.ok(firstContainer && lateContainer);
      const slidePart = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
      assert.ok(slidePart);
      const xml = strFromU8(slidePart);
      assertTimingReferencesFollowDefinitions(xml);
      const definitionAt = (timingId: number) =>
        xml.indexOf(`<p:cTn id="${timingId}"`);
      assert.ok(
        definitionAt(firstContainer.timingId) < definitionAt(join.timingId),
      );
      assert.ok(
        definitionAt(lateContainer.timingId) < definitionAt(join.timingId),
      );
      assert.equal(
        firstMainSequenceEntryNodeType(xml, 0),
        options.trigger === "click" ? "clickEffect" : "afterEffect",
      );
    }
  });

  it("按最大候选逐项证明嵌套条件集合严格晚于其父起点", async () => {
    const variants: readonly Omit<NestedMaxConditionJoinOptions, "control">[] =
      [
        { reverse: false, trigger: "auto" },
        { reverse: true, trigger: "auto" },
        { reverse: false, trigger: "click" },
        { reverse: true, trigger: "click" },
      ];

    for (const options of variants) {
      for (const control of [false, true]) {
        const result = await build(
          createElement(NestedMaxConditionJoin, { ...options, control }),
        );
        const slide = result.report.slides[0];
        const motion = slide?.capabilities?.["motion"] as MotionObservation;
        const objectId = (testId: string) =>
          slide?.objects.find(({ testId: id }) => id === testId)?.id;
        const scopeContaining = (testId: string) =>
          motion.scopes?.find(({ members }) =>
            members.some(
              (member) =>
                member.kind === "object" && member.id === objectId(testId),
            ),
          );
        const firstScope = scopeContaining("nested-max-first");
        const delayedScope = scopeContaining("nested-max-delayed");
        const consumerScope = scopeContaining("nested-max-consumer");
        const sourceScope = motion.scopes?.find(
          ({ members }) =>
            members.some(
              (member) =>
                member.kind === "scope" && member.id === firstScope?.id,
            ) &&
            members.some(
              (member) =>
                member.kind === "scope" && member.id === delayedScope?.id,
            ),
        );
        const parentScope = motion.scopes?.find(({ members }) =>
          members.some(
            (member) =>
              member.kind === "scope" && member.id === consumerScope?.id,
          ),
        );
        assert.ok(
          firstScope &&
            delayedScope &&
            consumerScope &&
            sourceScope &&
            parentScope,
        );
        assert.deepEqual(parentScope.startAfterAll, [
          { scopeId: sourceScope.id, point: "end", offsetMs: 0 },
          { scopeId: delayedScope.id, point: "end", offsetMs: 11000 },
        ]);
        assert.deepEqual(consumerScope.startAfterAll, [
          { scopeId: sourceScope.id, point: "end", offsetMs: 1 },
          {
            scopeId: delayedScope.id,
            point: "end",
            offsetMs: control ? 11001 : 12001,
          },
        ]);

        const file = await validatePptx(result.pptx);
        assert.deepEqual(file.errors, []);
        assert.deepEqual(file.extensions, []);
        const fileSlide = file.slides[0];
        assert.equal(fileSlide?.animations?.[0]?.trigger, options.trigger);
        const effects = fileSlide?.animations?.[0]?.effects ?? [];
        assert.equal(effects.length, 3);
        const effect = (testId: string) =>
          effects.find(({ targetId }) => targetId === objectId(testId));
        const firstEffect = effect("nested-max-first");
        const delayedEffect = effect("nested-max-delayed");
        const consumerEffect = effect("nested-max-consumer");
        assert.ok(firstEffect && delayedEffect && consumerEffect);
        for (const sourceEffect of [firstEffect, delayedEffect]) {
          assert.equal(sourceEffect.by, "letter");
          assert.equal(sourceEffect.completion, "unknown");
        }
        const containers = fileSlide?.animationContainers ?? [];
        const containerById = (timingId: number | undefined) =>
          containers.find(({ timingId: id }) => id === timingId);
        const effectContainer = (timingId: number) =>
          containers.find(({ children }) =>
            children.some(
              ({ kind, timingId: childId }) =>
                kind === "effect" && childId === timingId,
            ),
          );
        const firstContainer = effectContainer(firstEffect.timingId);
        const delayedContainer = effectContainer(delayedEffect.timingId);
        assert.ok(
          consumerEffect.startAfter && firstContainer && delayedContainer,
        );
        const join = containerById(consumerEffect.startAfter.timingId);
        assert.ok(join);
        assert.equal(join.endSync, "all");
        assert.equal(join.durationMs, undefined);
        assert.equal(join.children.length, 2);
        if (control) {
          assert.equal(consumerEffect.startAfter.offsetMs, 1);
        } else {
          assert.equal(consumerEffect.startAfter.offsetMs, 0);
          const gates = join.children.map(({ timingId }) =>
            containerById(timingId),
          );
          const ancestors = (container: (typeof containers)[number]) => {
            const result = [container];
            let parent = containerById(container.parentTimingId);
            while (parent !== undefined) {
              result.push(parent);
              parent = containerById(parent.parentTimingId);
            }
            return result;
          };
          const delayedAncestorIds = ancestors(delayedContainer).map(
            ({ timingId }) => timingId,
          );
          const sourceContainer = ancestors(firstContainer).find(
            ({ timingId }) => delayedAncestorIds.includes(timingId),
          );
          assert.ok(sourceContainer);
          assert.deepEqual(
            gates.map((gate) => [
              gate?.startAfter?.timingId,
              gate?.startAfter?.offsetMs,
            ]),
            [
              [sourceContainer.timingId, 1],
              [delayedContainer.timingId, 12001],
            ],
          );
        }

        const slidePart = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
        assert.ok(slidePart);
        const xml = strFromU8(slidePart);
        assertTimingReferencesFollowDefinitions(xml);
        assert.equal(
          firstMainSequenceEntryNodeType(xml, 0),
          options.trigger === "click" ? "clickEffect" : "afterEffect",
        );
      }
    }
  });

  it("只有部分父起点候选有严格下界时仍拒绝条件集合", async () => {
    await assert.rejects(
      build(
        createElement(NestedMaxConditionJoin, {
          reverse: false,
          trigger: "auto",
          partialCoverage: true,
        }),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "unknown-motion-start");
        assert.match(
          error.diagnostics[0]?.message ?? "",
          /不能证明条件集合完成晚于消费成员的父作用域激活/u,
        );
        return true;
      },
    );
  });

  it("保留交叉顺序分支的双成员整体完成及外部消费者", async () => {
    for (const longFirst of [false, true]) {
      for (const branchOrder of ["p-first", "q-first"] as const) {
        for (const consumerOrder of ["p-first", "q-first"] as const) {
          for (const trigger of ["auto", "click"] as const) {
            const result = await build(
              createElement(DfsEndCollectors, {
                longFirst,
                branchOrder,
                consumerOrder,
                trigger,
              }),
            );
            const slide = result.report.slides[0];
            const motion = slide?.capabilities?.["motion"] as MotionObservation;
            const objectId = (testId: string) =>
              slide?.objects.find(({ testId: id }) => id === testId)?.id;
            const scopeContaining = (testId: string) =>
              motion.scopes?.find(({ members }) =>
                members.some(
                  (member) =>
                    member.kind === "object" && member.id === objectId(testId),
                ),
              );
            const firstScope = scopeContaining("dfs-first");
            const secondScope = scopeContaining("dfs-second");
            const dScope = scopeContaining("dfs-d");
            const eScope = scopeContaining("dfs-e");
            const pScope = motion.scopes?.find(
              ({ members }) =>
                members.some(
                  (member) =>
                    member.kind === "scope" && member.id === firstScope?.id,
                ) &&
                members.some(
                  (member) =>
                    member.kind === "scope" && member.id === dScope?.id,
                ),
            );
            const qScope = motion.scopes?.find(
              ({ members }) =>
                members.some(
                  (member) =>
                    member.kind === "scope" && member.id === secondScope?.id,
                ) &&
                members.some(
                  (member) =>
                    member.kind === "scope" && member.id === eScope?.id,
                ),
            );
            assert.ok(
              firstScope && secondScope && dScope && eScope && pScope && qScope,
            );
            assert.equal(pScope.completion, "unknown");
            assert.equal(qScope.completion, "unknown");
            assert.deepEqual(pScope.members, [
              { kind: "scope", id: firstScope.id },
              { kind: "scope", id: dScope.id },
            ]);
            assert.deepEqual(qScope.members, [
              { kind: "scope", id: secondScope.id },
              { kind: "scope", id: eScope.id },
            ]);
            const afterPScope = scopeContaining("dfs-after-p");
            const afterQScope = scopeContaining("dfs-after-q");
            const hostFollowerScope = scopeContaining("dfs-host-follower");
            assert.ok(afterPScope && afterQScope && hostFollowerScope);
            assert.deepEqual(afterPScope.startAfter, {
              scopeId: pScope.id,
              point: "end",
              offsetMs: 0,
            });
            assert.deepEqual(afterQScope.startAfter, {
              scopeId: qScope.id,
              point: "end",
              offsetMs: 0,
            });
            const hostScope = motion.scopes?.find(
              ({ id }) => id === pScope.parentId,
            );
            assert.ok(hostScope);
            assert.deepEqual(
              hostScope.members,
              branchOrder === "p-first"
                ? [
                    { kind: "scope", id: pScope.id },
                    { kind: "scope", id: qScope.id },
                  ]
                : [
                    { kind: "scope", id: qScope.id },
                    { kind: "scope", id: pScope.id },
                  ],
            );
            assert.deepEqual(hostFollowerScope.startAfter, {
              scopeId: hostScope.id,
              point: "end",
              offsetMs: 0,
            });

            const file = await validatePptx(result.pptx);
            assert.deepEqual(file.errors, []);
            assert.deepEqual(file.extensions, []);
            const fileSlide = file.slides[0];
            assert.equal(fileSlide?.animations?.[0]?.trigger, trigger);
            const effects = fileSlide?.animations?.[0]?.effects ?? [];
            assert.equal(effects.length, 7);
            assert.equal(
              new Set(effects.map(({ targetId }) => targetId)).size,
              7,
            );
            const effect = (testId: string) =>
              effects.find(({ targetId }) => targetId === objectId(testId));
            const firstEffect = effect("dfs-first");
            const secondEffect = effect("dfs-second");
            const dEffect = effect("dfs-d");
            const eEffect = effect("dfs-e");
            const afterPEffect = effect("dfs-after-p");
            const afterQEffect = effect("dfs-after-q");
            const hostFollowerEffect = effect("dfs-host-follower");
            assert.ok(
              firstEffect &&
                secondEffect &&
                dEffect &&
                eEffect &&
                afterPEffect &&
                afterQEffect &&
                hostFollowerEffect,
            );
            const containers = fileSlide?.animationContainers ?? [];
            const containerById = (timingId: number | undefined) =>
              containers.find(({ timingId: id }) => id === timingId);
            const effectContainer = (timingId: number) =>
              containers.find(({ children }) =>
                children.some(
                  ({ kind, timingId: childId }) =>
                    kind === "effect" && childId === timingId,
                ),
              );
            const firstContainer = effectContainer(firstEffect.timingId);
            const secondContainer = effectContainer(secondEffect.timingId);
            const dContainer = effectContainer(dEffect.timingId);
            const eContainer = effectContainer(eEffect.timingId);
            assert.ok(
              firstContainer && secondContainer && dContainer && eContainer,
            );
            const conditionGates = (
              consumer: NonNullable<typeof dContainer>,
            ) => {
              const join = containerById(consumer.startAfter?.timingId);
              assert.ok(join);
              assert.equal(join.endSync, "all");
              assert.equal(join.durationMs, undefined);
              assert.equal(join.children.length, 2);
              return join.children.map(
                ({ timingId }) => containerById(timingId)?.startAfter,
              );
            };
            const dGates = conditionGates(dContainer);
            const eGates = conditionGates(eContainer);
            assert.deepEqual(
              dGates
                .map((gate) => [gate?.timingId, gate?.offsetMs])
                .sort(
                  ([leftId], [rightId]) => Number(leftId) - Number(rightId),
                ),
              [
                [firstContainer.timingId, 0],
                [secondContainer.timingId, 11000],
              ].sort(([leftId], [rightId]) => Number(leftId) - Number(rightId)),
            );
            assert.deepEqual(
              eGates
                .map((gate) => [gate?.timingId, gate?.offsetMs])
                .sort(
                  ([leftId], [rightId]) => Number(leftId) - Number(rightId),
                ),
              [
                [secondContainer.timingId, 0],
                [firstContainer.timingId, 11000],
              ].sort(([leftId], [rightId]) => Number(leftId) - Number(rightId)),
            );
            const pCollector = containerById(afterPEffect.startAfter?.timingId);
            const qCollector = containerById(afterQEffect.startAfter?.timingId);
            const hostContainer = containerById(
              hostFollowerEffect.startAfter?.timingId,
            );
            assert.ok(pCollector && qCollector && hostContainer);
            for (const collector of [pCollector, qCollector]) {
              assert.equal(collector.kind, "parallel");
              assert.equal(collector.endSync, "all");
              assert.equal(collector.durationMs, undefined);
              assert.equal(collector.children.length, 2);
              assert.ok(
                collector.children.every(({ kind }) => kind === "container"),
              );
            }
            const collectorSources = (collector: typeof pCollector) =>
              collector.children.map(
                ({ timingId }) => containerById(timingId)?.startAfter?.timingId,
              );
            assert.deepEqual(
              new Set(collectorSources(pCollector)),
              new Set([firstContainer.timingId, dContainer.timingId]),
            );
            assert.deepEqual(
              new Set(collectorSources(qCollector)),
              new Set([secondContainer.timingId, eContainer.timingId]),
            );
            assert.equal(pCollector.parentTimingId, hostContainer.timingId);
            assert.equal(qCollector.parentTimingId, hostContainer.timingId);
            assert.equal(hostContainer.endSync, "all");
            assert.ok(
              !hostContainer.children.some(({ timingId }) =>
                [
                  afterPEffect.timingId,
                  afterQEffect.timingId,
                  hostFollowerEffect.timingId,
                ].includes(timingId),
              ),
            );

            const slidePart = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
            assert.ok(slidePart);
            const xml = strFromU8(slidePart);
            assertTimingReferencesFollowDefinitions(xml);
            assert.equal(
              firstMainSequenceEntryNodeType(xml, 0),
              trigger === "click" ? "clickEffect" : "afterEffect",
            );
          }
        }
      }
    }
  });

  it("逐词真实单次行为在中英文、空白、标点和空文本上保留正下界", async () => {
    const texts = ["中文词语", "e\u0301 👩🏽‍💻", "   \t", "，。!?", ""];
    for (const [index, text] of texts.entries()) {
      const result = await build(
        createElement(DfsEndCollectors, {
          wordSources: true,
          firstText: text,
          secondText: text,
          branchOrder: index % 2 === 0 ? "p-first" : "q-first",
          consumerOrder: index % 2 === 0 ? "q-first" : "p-first",
          trigger: index % 2 === 0 ? "auto" : "click",
        }),
      );
      const slide = result.report.slides[0];
      const motion = slide?.capabilities?.["motion"] as MotionObservation;
      const first = motion.steps[0]?.effects.find(
        ({ target }) => target.testId === "dfs-first",
      );
      const second = motion.steps[0]?.effects.find(
        ({ target }) => target.testId === "dfs-second",
      );
      assert.ok(first && second);
      for (const source of [first, second]) {
        assert.deepEqual(
          [
            source.by,
            source.intervalMs,
            source.durationMs,
            source.startMs,
            source.endMs,
            source.completion,
          ],
          ["word", 200, 600, 0, undefined, "unknown"],
        );
      }
      const sourceScope = motion.scopes?.find(({ members }) =>
        members.some(
          (member) =>
            member.kind === "object" &&
            (member.id === first.target.id || member.id === second.target.id),
        ),
      );
      assert.equal(sourceScope?.completion, "unknown");
      assert.equal(sourceScope?.endMs, undefined);

      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      assert.deepEqual(file.extensions, []);
      const fileSlide = file.slides[0];
      const effects = fileSlide?.animations?.[0]?.effects ?? [];
      for (const testId of ["dfs-first", "dfs-second"]) {
        const targetId = slide?.objects.find(
          ({ testId: id }) => id === testId,
        )?.id;
        const wordEffect = effects.find(({ targetId: id }) => id === targetId);
        assert.ok(wordEffect);
        assert.deepEqual(
          [
            wordEffect.by,
            wordEffect.intervalMs,
            wordEffect.durationMs,
            wordEffect.completion,
          ],
          ["word", 200, 600, "unknown"],
        );
      }
      const slidePart = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
      assert.ok(slidePart);
      const xml = strFromU8(slidePart);
      assert.match(
        xml,
        /<p:iterate type="wd"><p:tmAbs val="200"><\/p:tmAbs><\/p:iterate>/u,
      );
      assertTimingReferencesFollowDefinitions(xml);
    }
  });

  it("真实嵌套成员完成可收集，但零下界成员仍不能作为早宿主监听源", async () => {
    await assert.rejects(
      build(
        createElement(DfsEndCollectors, {
          emptyFirstLetter: true,
          trigger: "auto",
        }),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "unknown-motion-start");
        assert.match(
          error.diagnostics[0]?.message ?? "",
          /不能证明真实来源结束晚于候选宿主激活/u,
        );
        return true;
      },
    );
  });

  it("晚起点别名保留逻辑起点，完成收集器仍从早宿主注册", async () => {
    const result = await build(
      createElement(DfsEndCollectors, {
        pDelayMs: 2500,
        pStartAlias: true,
      }),
    );
    const slide = result.report.slides[0];
    const motion = slide?.capabilities?.["motion"] as MotionObservation;
    const scopeContaining = (testId: string) =>
      motion.scopes?.find(({ members }) =>
        members.some(
          (member) =>
            member.kind === "object" &&
            member.id ===
              slide?.objects.find(({ testId: id }) => id === testId)?.id,
        ),
      );
    const firstScope = scopeContaining("dfs-first");
    const dScope = scopeContaining("dfs-d");
    const pScope = motion.scopes?.find(
      ({ members }) =>
        members.some(
          (member) => member.kind === "scope" && member.id === firstScope?.id,
        ) &&
        members.some(
          (member) => member.kind === "scope" && member.id === dScope?.id,
        ),
    );
    const aliasScope = scopeContaining("dfs-p-start-alias");
    assert.ok(pScope && aliasScope);
    assert.equal(pScope.startMs, 2500);
    assert.equal(aliasScope.startMs, 2500);
    assert.equal(pScope.completion, "unknown");

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.extensions, []);
    const fileSlide = file.slides[0];
    const effects = fileSlide?.animations?.[0]?.effects ?? [];
    const afterPEffect = effects.find(
      ({ targetId }) =>
        targetId ===
        slide?.objects.find(({ testId }) => testId === "dfs-after-p")?.id,
    );
    const aliasEffect = effects.find(
      ({ targetId }) =>
        targetId ===
        slide?.objects.find(({ testId }) => testId === "dfs-p-start-alias")?.id,
    );
    assert.ok(afterPEffect && aliasEffect);
    const containers = fileSlide?.animationContainers ?? [];
    const pCollector = containers.find(
      ({ timingId }) => timingId === afterPEffect.startAfter?.timingId,
    );
    assert.ok(pCollector);
    assert.equal(pCollector.endSync, "all");
    assert.equal(pCollector.startAfter, undefined);
    const slidePart = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
    assert.ok(slidePart);
    const xml = strFromU8(slidePart);
    assertTimingReferencesFollowDefinitions(xml);
    assert.match(
      xml,
      new RegExp(
        `<p:cTn id="${aliasEffect.timingId}"[^>]*>[\\s\\S]*?<p:cond delay="2500"`,
        "u",
      ),
    );
    assert.match(
      xml,
      new RegExp(
        `<p:cTn id="${pCollector.timingId}"[^>]*><p:stCondLst><p:cond delay="0"`,
        "u",
      ),
    );
  });

  it("同父顺序成员间保留不可比较的未知完成约束和真实整体完成", async () => {
    const variants: readonly IncomparableJoinOptions[] = [
      { firstText: "AB", trigger: "auto" },
      { firstText: "ABCDEFGHIJKLMNOPQRSTUVWXYZ123456", trigger: "auto" },
      { firstText: "AB", trigger: "click" },
      { firstText: "ABCDEFGHIJKLMNOPQRSTUVWXYZ123456", trigger: "click" },
      {
        firstText: "AB",
        trigger: "explicit",
        unknownParentStart: true,
      },
      {
        firstText: "ABCDEFGHIJKLMNOPQRSTUVWXYZ123456",
        trigger: "explicit",
        unknownParentStart: true,
      },
      { firstText: "AB", trigger: "auto", parentStartGate: true },
      { firstText: "AB", trigger: "click", parentStartGate: true },
      {
        firstText: "AB",
        trigger: "auto",
        parentStartGate: true,
        effectDelayMs: 37,
      },
      {
        firstText: "AB",
        trigger: "click",
        parentStartGate: true,
        scopeDelayMs: 41,
        effectDelayMs: 37,
      },
      { firstText: "AB", trigger: "auto", effectDelayMs: 37 },
      {
        firstText: "AB",
        trigger: "click",
        scopeDelayMs: 41,
        effectDelayMs: 37,
      },
    ];

    for (const options of variants) {
      const result = await build(
        createElement(IncomparableUnknownJoin, options),
      );
      const slide = result.report.slides[0];
      const motion = slide?.capabilities?.["motion"] as MotionObservation;
      const objectId = (testId: string) =>
        slide?.objects.find(({ testId: id }) => id === testId)?.id;
      const scopeContaining = (testId: string) =>
        motion.scopes?.find(({ members }) =>
          members.some(
            (member) =>
              member.kind === "object" && member.id === objectId(testId),
          ),
        );
      const delayedScope = scopeContaining("join-late");
      const innerScope = motion.scopes?.find(
        ({ members }) =>
          members.some(
            (member) =>
              member.kind === "object" && member.id === objectId("join-first"),
          ) &&
          members.some(
            (member) =>
              member.kind === "scope" && member.id === delayedScope?.id,
          ),
      );
      const afterJoinScope = scopeContaining("join-after");
      const chainScope = motion.scopes?.find(
        ({ members }) =>
          members.some(
            (member) => member.kind === "scope" && member.id === innerScope?.id,
          ) &&
          members.some(
            (member) =>
              member.kind === "scope" && member.id === afterJoinScope?.id,
          ),
      );
      const selfStartedScope = scopeContaining("join-independent");
      const outerScope = motion.scopes?.find(
        ({ members }) =>
          members.some(
            (member) => member.kind === "scope" && member.id === chainScope?.id,
          ) &&
          members.some(
            (member) =>
              member.kind === "scope" && member.id === selfStartedScope?.id,
          ),
      );
      const conditionSources = [
        {
          scopeId: innerScope?.id,
          point: "end" as const,
          offsetMs: options.scopeDelayMs ?? 0,
        },
        options.parentStartGate === true
          ? {
              scopeId: chainScope?.id,
              point: "start" as const,
              offsetMs: 5000 + (options.scopeDelayMs ?? 0),
            }
          : {
              scopeId: delayedScope?.id,
              point: "end" as const,
              offsetMs: 11000 + (options.scopeDelayMs ?? 0),
            },
      ];
      const effectSources = conditionSources.map((dependency) => ({
        ...dependency,
        offsetMs: dependency.offsetMs + (options.effectDelayMs ?? 0),
      }));
      const firstObservation = motion.steps[0]?.effects.find(
        ({ target }) => target.testId === "join-first",
      );
      const lateObservation = motion.steps[0]?.effects.find(
        ({ target }) => target.testId === "join-late",
      );
      const afterJoinObservation = motion.steps[0]?.effects.find(
        ({ target }) => target.testId === "join-after",
      );
      const followerObservation = motion.steps[0]?.effects.find(
        ({ target }) => target.testId === "join-follower",
      );
      assert.deepEqual(
        [firstObservation?.completion, firstObservation?.endMs],
        ["unknown", undefined],
      );
      assert.deepEqual(
        [lateObservation?.completion, lateObservation?.endMs],
        ["unknown", undefined],
      );
      assert.deepEqual(
        [
          afterJoinScope?.startMs,
          afterJoinScope?.endMs,
          afterJoinScope?.completion,
        ],
        [undefined, undefined, "unknown"],
      );
      assert.deepEqual(afterJoinScope?.startAfterAll, conditionSources);
      assert.equal(afterJoinScope?.startAfter, undefined);
      assert.deepEqual(afterJoinObservation?.startAfterAll, effectSources);
      assert.equal(afterJoinObservation?.startAfter, undefined);
      assert.equal(afterJoinObservation?.completion, "unknown");
      assert.deepEqual(
        [
          innerScope?.completion,
          chainScope?.completion,
          outerScope?.completion,
        ],
        ["unknown", "unknown", "unknown"],
      );
      assert.equal(
        followerObservation?.startAfter &&
          "scopeId" in followerObservation.startAfter &&
          followerObservation.startAfter.scopeId,
        outerScope?.id,
      );
      if (options.unknownParentStart === true) {
        const parentSource = scopeContaining("join-parent-source");
        assert.equal(
          outerScope?.startAfter &&
            "scopeId" in outerScope.startAfter &&
            outerScope.startAfter.scopeId,
          parentSource?.id,
        );
      }
      assert.equal(
        motion.steps[0]?.trigger,
        options.trigger === "auto" ? "auto" : "click",
      );
      assert.equal(
        motion.steps[0]?.number,
        options.trigger === "explicit" ? 2 : undefined,
      );

      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      assert.deepEqual(file.extensions, []);
      const fileSlide = file.slides[0];
      const fileEffects = fileSlide?.animations?.[0]?.effects ?? [];
      const fileEffect = (testId: string) =>
        fileEffects.find(({ targetId }) => targetId === objectId(testId));
      const firstEffect = fileEffect("join-first");
      const lateEffect = fileEffect("join-late");
      const afterJoinEffect = fileEffect("join-after");
      const independentEffect = fileEffect("join-independent");
      const followerEffect = fileEffect("join-follower");
      const containers = fileSlide?.animationContainers ?? [];
      const childrenOf = (parentId: number | undefined) =>
        containers.find(({ timingId }) => timingId === parentId);
      assert.ok(
        firstEffect &&
          lateEffect &&
          afterJoinEffect &&
          independentEffect &&
          followerEffect,
      );
      assert.equal(firstEffect.by, "letter");
      assert.equal(firstEffect.completion, "unknown");
      assert.equal(lateEffect.by, "letter");
      assert.equal(lateEffect.completion, "unknown");
      const fileInner = containers.find(
        ({ children, endSync }) =>
          endSync === "all" &&
          children.some(
            ({ kind, timingId }) =>
              kind === "effect" && timingId === firstEffect.timingId,
          ) &&
          children.some(
            ({ kind, timingId }) =>
              kind === "container" ||
              (kind === "effect" && timingId === lateEffect.timingId),
          ),
      );
      const fileDelayed = containers.find(
        ({ children, endSync }) =>
          endSync === "all" &&
          children.some(
            ({ kind, timingId }) =>
              kind === "effect" && timingId === lateEffect.timingId,
          ),
      );
      assert.ok(fileInner);
      if (options.parentStartGate !== true) assert.ok(fileDelayed);
      assert.equal(fileInner.durationMs, undefined);
      const join = containers.find(({ children, endSync, durationMs }) => {
        if (
          endSync !== "all" ||
          durationMs !== undefined ||
          children.length !== 2
        ) {
          return false;
        }
        return children.every(({ kind, timingId }) => {
          const gate = childrenOf(timingId);
          return (
            kind === "container" &&
            gate?.durationMs === 0 &&
            gate.children.length === 0
          );
        });
      });
      assert.ok(join);
      const gates = join.children.map(({ timingId }) => childrenOf(timingId));
      const secondGateSource =
        options.parentStartGate === true ? undefined : fileDelayed?.timingId;
      assert.deepEqual(
        gates.map((gate) => [
          gate?.startAfter?.timingId,
          gate?.startAfter?.offsetMs,
        ]),
        [
          [fileInner.timingId, options.scopeDelayMs ?? 0],
          [
            secondGateSource,
            options.parentStartGate === true
              ? undefined
              : 11000 + (options.scopeDelayMs ?? 0),
          ],
        ],
      );
      assert.equal(afterJoinEffect.startAfter?.timingId, join.timingId);
      assert.equal(
        afterJoinEffect.startAfter?.offsetMs,
        options.effectDelayMs ?? 0,
      );
      const fileChain = childrenOf(join.parentTimingId);
      assert.ok(fileChain);
      assert.ok(
        fileChain.children.some(
          ({ kind, timingId }) =>
            kind === "container" && timingId === fileInner.timingId,
        ),
      );
      const fileOuter = childrenOf(followerEffect.startAfter?.timingId);
      assert.equal(fileOuter?.endSync, "all");
      assert.equal(fileOuter?.durationMs, undefined);
      assert.equal(followerEffect.startAfter?.timingId, fileOuter?.timingId);
      assert.ok(
        fileOuter?.children.some(
          ({ kind, timingId }) =>
            kind === "container" && timingId === fileChain.timingId,
        ),
      );
      assert.ok(
        fileOuter?.children.some(
          ({ kind, timingId }) =>
            kind === "effect" && timingId === independentEffect.timingId,
        ),
      );

      const archive = unzipSync(result.pptx);
      const slidePart = archive["ppt/slides/slide1.xml"];
      assert.ok(slidePart);
      const xml = strFromU8(slidePart);
      const timingNodeIndex = (timingId: number) =>
        xml.indexOf(`<p:cTn id="${timingId}"`);
      assert.ok(timingNodeIndex(fileInner.timingId) >= 0);
      assert.ok(
        timingNodeIndex(join.timingId) > timingNodeIndex(fileInner.timingId),
      );
      assert.ok(
        timingNodeIndex(afterJoinEffect.timingId) >
          timingNodeIndex(join.timingId),
      );
      assert.match(
        startConditions(xml, join.timingId),
        /<p:cond delay="0"><\/p:cond>/u,
      );
      assert.match(
        xml.slice(
          timingNodeIndex(join.timingId),
          timingNodeIndex(join.timingId) + 500,
        ),
        /<p:endSync evt="end" delay="0"><p:rtn val="all"><\/p:rtn><\/p:endSync>/u,
      );
      for (const [index, gate] of gates.entries()) {
        assert.ok(gate);
        const conditions = startConditions(xml, gate.timingId);
        const offsetMs =
          index === 0
            ? (options.scopeDelayMs ?? 0)
            : options.parentStartGate === true
              ? 5000 + (options.scopeDelayMs ?? 0)
              : 11000 + (options.scopeDelayMs ?? 0);
        if (index === 1 && options.parentStartGate === true) {
          assert.match(
            conditions,
            new RegExp(`<p:cond delay="${offsetMs}"><\\/p:cond>`, "u"),
          );
        } else {
          assert.match(
            conditions,
            new RegExp(
              `<p:cond evt="end" delay="${offsetMs}"><p:tn val="${gate.startAfter?.timingId}"><\\/p:tn><\\/p:cond>`,
              "u",
            ),
          );
        }
        assert.match(
          xml.slice(
            timingNodeIndex(gate.timingId),
            timingNodeIndex(gate.timingId) + 180,
          ),
          /dur="0"/u,
        );
      }
      assert.match(
        startConditions(xml, afterJoinEffect.timingId),
        new RegExp(
          `<p:cond evt="end" delay="${options.effectDelayMs ?? 0}"><p:tn val="${join.timingId}"><\\/p:tn><\\/p:cond>`,
          "u",
        ),
      );
      assert.equal(
        fileOuter?.children.some(({ timingId }) => timingId === join.timingId),
        false,
      );
    }
  });

  it("嵌套未知完成的自动步与点击步保留真实入口角色和完整依赖", async () => {
    const result = await build(createElement(UnknownCompletionEntryDeck));
    const slide = result.report.slides[0];
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);

    const animations = file.slides[0]?.animations ?? [];
    assert.deepEqual(
      animations.map(({ trigger }) => trigger),
      ["auto", "click"],
    );
    const slidePart = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
    assert.ok(slidePart);
    const xml = strFromU8(slidePart);
    assert.deepEqual(
      [
        firstMainSequenceEntryNodeType(xml, 0),
        firstMainSequenceEntryNodeType(xml, 1),
      ],
      ["afterEffect", "clickEffect"],
    );

    const objectId = (testId: string) =>
      slide?.objects.find(({ testId: id }) => id === testId)?.id;
    const fileEffects = file.slides[0]?.animations?.flatMap(
      ({ effects }) => effects,
    );
    const fileEffect = (testId: string) => {
      const targetId = objectId(testId);
      return fileEffects?.find((effect) => effect.targetId === targetId);
    };
    const containers = file.slides[0]?.animationContainers ?? [];
    for (const [sourceNames, followerName, offsetMs] of [
      [["auto-unknown-long", "auto-unknown-short"], "auto-after-unknown", 25],
      [
        ["click-unknown-long", "click-unknown-short"],
        "click-after-unknown",
        30,
      ],
    ] as const) {
      const sourceEffects = sourceNames.map((name) => fileEffect(name));
      const follower = fileEffect(followerName);
      assert.ok(
        sourceEffects.every((effect) => effect?.completion === "unknown"),
      );
      assert.ok(follower);
      const source = containers.find(
        ({ children, endSync }) =>
          endSync === "all" &&
          sourceEffects.every((effect) =>
            children.some(
              ({ kind, timingId }) =>
                kind === "effect" && timingId === effect?.timingId,
            ),
          ),
      );
      assert.ok(source);
      assert.equal(source.durationMs, undefined);
      assert.deepEqual(
        [follower.startAfter?.timingId, follower.startAfter?.offsetMs],
        [source.timingId, offsetMs],
      );
      assert.match(
        startConditions(xml, follower.timingId),
        new RegExp(
          `<p:cond evt="end" delay="${offsetMs}"><p:tn val="${source.timingId}"><\\/p:tn><\\/p:cond>`,
          "u",
        ),
      );
    }
  });

  it("未知外部起点后的顺序成员可依赖实际前驱，且后继等待完整作用域", async () => {
    const result = await build(createElement(UnknownStartSequence, {}));
    const slide = result.report.slides[0];
    const motion = slide?.capabilities?.["motion"] as MotionObservation;
    const objectId = (testId: string) =>
      slide?.objects.find((object) => object.testId === testId)?.id;
    const source = motion.steps[0]?.effects.find(
      ({ target }) => target.testId === "unknown-source",
    );
    const first = motion.steps[0]?.effects.find(
      ({ target }) => target.testId === "sequence-first",
    );
    const second = motion.steps[0]?.effects.find(
      ({ target }) => target.testId === "sequence-second",
    );
    const follower = motion.steps[0]?.effects.find(
      ({ target }) => target.testId === "sequence-follower",
    );
    assert.deepEqual(
      [source?.completion, source?.endMs, first?.startMs, second?.startMs],
      ["unknown", undefined, undefined, undefined],
    );
    assert.equal(
      first?.startAfter &&
        "scopeId" in first.startAfter &&
        first.startAfter.scopeId,
      "scope-1",
    );
    assert.equal(first?.startAfter?.offsetMs, 25);
    assert.deepEqual(
      second?.startAfter && "targetId" in second.startAfter
        ? [second.startAfter.targetId, second.startAfter.point]
        : undefined,
      [objectId("sequence-first"), "end"],
    );
    assert.equal(
      follower?.startAfter &&
        "scopeId" in follower.startAfter &&
        follower.startAfter.scopeId,
      "scope-2",
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const fileEffects = file.slides[0]?.animations?.[0]?.effects ?? [];
    const fileEffect = (testId: string) =>
      fileEffects.find(({ targetId }) => targetId === objectId(testId));
    const fileFirst = fileEffect("sequence-first");
    const fileSecond = fileEffect("sequence-second");
    const fileFollower = fileEffect("sequence-follower");
    assert.equal(fileFirst?.startAfter?.source, "container");
    assert.equal(fileFirst?.startAfter?.offsetMs, 25);
    assert.equal(fileSecond?.startAfter?.timingId, fileFirst?.timingId);
    assert.equal(fileSecond?.startAfter?.offsetMs, 0);
    const sequence = file.slides[0]?.animationContainers?.find(
      ({ children, endSync }) =>
        endSync === "all" &&
        children.some(
          ({ kind, timingId }) =>
            kind === "effect" && timingId === fileFirst?.timingId,
        ) &&
        children.some(
          ({ kind, timingId }) =>
            kind === "effect" && timingId === fileSecond?.timingId,
        ),
    );
    assert.ok(sequence);
    assert.equal(sequence.durationMs, undefined);
    assert.equal(fileFollower?.startAfter?.timingId, sequence.timingId);
  });

  it("顺序前驱完成覆盖冗余自定起点，保留未知时长与真实整体完成", async () => {
    const variants: readonly SequenceStartCausalityOptions[] = [
      {
        sourceUnknown: true,
        firstUnknown: true,
        offsetMs: 0,
      },
      {
        sourceUnknown: true,
        firstUnknown: true,
        offsetMs: 17,
      },
      {
        sourceUnknown: true,
        firstUnknown: false,
        offsetMs: 17,
      },
      {
        sourceUnknown: true,
        firstUnknown: false,
        offsetMs: 136,
      },
      {
        sourceUnknown: true,
        firstUnknown: false,
        offsetMs: 137,
      },
      {
        sourceUnknown: true,
        firstUnknown: false,
        offsetMs: 200,
      },
      {
        sourceUnknown: true,
        firstUnknown: false,
        offsetMs: 200,
        automaticClickStep: true,
      },
      {
        sourceUnknown: true,
        firstUnknown: false,
        offsetMs: 200,
        explicitStep: true,
      },
      {
        sourceUnknown: true,
        firstUnknown: false,
        offsetMs: 200,
        wrapFirstInScope: true,
      },
      {
        sourceUnknown: false,
        firstUnknown: true,
        offsetMs: 17,
        explicitStep: true,
      },
      {
        sourceUnknown: false,
        firstUnknown: false,
        offsetMs: 17,
      },
    ];

    for (const options of variants) {
      const result = await build(
        createElement(SequenceStartCausality, options),
      );
      const slide = result.report.slides[0];
      const motion = slide?.capabilities?.["motion"] as MotionObservation;
      const objectId = (testId: string) =>
        slide?.objects.find(({ testId: id }) => id === testId)?.id;
      const source = motion.steps[0]?.effects.find(
        ({ target }) => target.testId === "causality-source",
      );
      const first = motion.steps[0]?.effects.find(
        ({ target }) => target.testId === "causality-first",
      );
      const second = motion.steps[0]?.effects.find(
        ({ target }) => target.testId === "causality-second",
      );
      const follower = motion.steps[0]?.effects.find(
        ({ target }) => target.testId === "causality-follower",
      );

      if (options.sourceUnknown) {
        assert.deepEqual(
          [source?.completion, source?.endMs],
          ["unknown", undefined],
        );
      }
      if (options.firstUnknown) {
        assert.deepEqual(
          [first?.completion, first?.endMs],
          ["unknown", undefined],
        );
      }
      assert.ok(second);
      assert.ok(follower);
      if (options.sourceUnknown || options.firstUnknown) {
        if (
          options.sourceUnknown &&
          !options.firstUnknown &&
          options.offsetMs >= 200
        ) {
          assert.deepEqual(
            second.startAfter && "scopeId" in second.startAfter
              ? [second.startAfter.scopeId, second.startAfter.offsetMs]
              : undefined,
            ["scope-1", 225],
          );
        } else {
          assert.deepEqual(
            second.startAfter && "targetId" in second.startAfter
              ? [second.startAfter.targetId, second.startAfter.point]
              : undefined,
            [objectId("causality-first"), "end"],
          );
        }
      } else {
        assert.equal(second.startMs, 299);
      }

      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      const effects = file.slides[0]?.animations?.[0]?.effects ?? [];
      const fileEffect = (testId: string) =>
        effects.find(({ targetId }) => targetId === objectId(testId));
      const fileSource = fileEffect("causality-source");
      const fileFirst = fileEffect("causality-first");
      const fileSecond = fileEffect("causality-second");
      const fileFollower = fileEffect("causality-follower");
      const containers = file.slides[0]?.animationContainers ?? [];
      assert.ok(fileSource && fileFirst && fileSecond && fileFollower);
      const sourceContainer = options.sourceUnknown
        ? containers.find(
            ({ children, endSync }) =>
              endSync === "all" &&
              children.some(
                ({ kind, timingId }) =>
                  kind === "effect" && timingId === fileSource.timingId,
              ),
          )
        : undefined;
      const follow =
        options.sourceUnknown || options.firstUnknown
          ? containers.find(
              ({ timingId }) => timingId === fileFollower.startAfter?.timingId,
            )
          : undefined;
      if (options.firstUnknown) {
        assert.equal(fileFirst.by, "letter");
        assert.equal(fileFirst.completion, "unknown");
        assert.equal(fileSecond.startAfter?.timingId, fileFirst.timingId);
        assert.equal(fileSecond.startAfter?.offsetMs, 0);
      }
      if (options.sourceUnknown) {
        assert.ok(sourceContainer);
        assert.equal(fileFirst.startAfter?.source, "container");
        assert.equal(fileFirst.startAfter?.offsetMs, 25);
      }
      if (options.sourceUnknown || options.firstUnknown) {
        assert.ok(follow);
        assert.equal(follow.endSync, "all");
      }
      if (
        options.sourceUnknown &&
        !options.firstUnknown &&
        options.offsetMs >= 200
      ) {
        assert.ok(sourceContainer);
        assert.equal(fileSecond.startAfter?.source, "container");
        assert.equal(fileSecond.startAfter?.timingId, sourceContainer.timingId);
        assert.equal(fileSecond.startAfter?.offsetMs, 25);
        assert.equal(fileSecond.delayMs, 225);
        const slideXml = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
        assert.ok(slideXml);
        assert.match(
          startConditions(strFromU8(slideXml), fileSecond.timingId),
          /<p:cond delay="200"><\/p:cond>/u,
        );
      } else if (options.sourceUnknown || options.firstUnknown) {
        assert.equal(fileSecond.startAfter?.timingId, fileFirst.timingId);
        assert.equal(fileSecond.startAfter?.offsetMs, 0);
      }
      if (options.sourceUnknown || options.firstUnknown) {
        assert.ok(follow);
        assert.equal(fileFollower.startAfter?.timingId, follow.timingId);
        assert.equal(fileFollower.startAfter?.offsetMs, 11);
      } else {
        assert.equal(fileFollower.startAfter, undefined);
        assert.equal(follower.startMs, 447);
      }
      assert.equal(
        file.slides[0]?.animations?.[0]?.trigger,
        options.explicitStep === true || options.automaticClickStep === true
          ? "click"
          : "auto",
      );
    }
  });

  it("平行作用域中的完成引用不受兄弟声明顺序影响", async () => {
    for (const sourceFirst of [true, false]) {
      const result = await build(
        createElement(ForwardSibling, { sourceFirst }),
      );
      const slide = result.report.slides[0];
      const objectId = (testId: string) =>
        slide?.objects.find(({ testId: id }) => id === testId)?.id;
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      const effects = file.slides[0]?.animations?.[0]?.effects ?? [];
      const slidePart = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
      assert.ok(slidePart);
      assertTimingReferencesFollowDefinitions(strFromU8(slidePart));
      const sourceEffect = effects.find(
        ({ targetId }) => targetId === objectId("forward-source"),
      );
      const waiterEffect = effects.find(
        ({ targetId }) => targetId === objectId("forward-waiter"),
      );
      assert.ok(sourceEffect && waiterEffect);
      const containers = file.slides[0]?.animationContainers ?? [];
      const source = containers.find(
        ({ children, endSync }) =>
          endSync === "all" &&
          children.some(
            ({ kind, timingId }) =>
              kind === "effect" && timingId === sourceEffect.timingId,
          ),
      );
      assert.ok(source);
      assert.equal(waiterEffect.startAfter?.timingId, source.timingId);
      assert.equal(waiterEffect.startAfter?.offsetMs, 0);
    }
  });

  it("未知起点后的并行、交错、嵌套顺序及同一点击步保持可写", async () => {
    const variants: readonly UnknownStartSequenceOptions[] = [
      { mode: "parallel" as const },
      { mode: "stagger" as const },
      { nested: true },
      { explicitStep: true },
    ];
    for (const options of variants) {
      const result = await build(createElement(UnknownStartSequence, options));
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      assert.equal(
        file.slides[0]?.animations?.[0]?.trigger,
        options.explicitStep === true ? "click" : "auto",
      );
    }
  });

  it("顺序前项以真实自定起点子作用域整体完成后启动", async () => {
    const result = await build(createElement(SelfStartedNestedPredecessor));
    const slide = result.report.slides[0];
    const motion = slide?.capabilities?.["motion"] as MotionObservation;
    const objectId = (testId: string) =>
      slide?.objects.find((object) => object.testId === testId)?.id;
    const predecessorScope = motion.scopes?.find(
      ({ members }) =>
        members.some(
          (member) =>
            member.kind === "object" &&
            member.id === objectId("predecessor-long"),
        ) &&
        members.some(
          (member) =>
            member.kind === "object" &&
            member.id === objectId("predecessor-short"),
        ),
    );
    const first = motion.steps[0]?.effects.find(
      ({ target }) => target.testId === "after-predecessor",
    );
    const follower = motion.steps[0]?.effects.find(
      ({ target }) => target.testId === "after-self-started-sequence",
    );
    assert.deepEqual(
      [predecessorScope?.completion, predecessorScope?.endMs],
      ["unknown", undefined],
    );
    assert.deepEqual(
      predecessorScope?.startAfter && "scopeId" in predecessorScope.startAfter
        ? [
            predecessorScope.startAfter.scopeId,
            predecessorScope.startAfter.offsetMs,
          ]
        : undefined,
      ["scope-1", 65],
    );
    assert.equal(
      first?.startAfter &&
        "scopeId" in first.startAfter &&
        first.startAfter.scopeId,
      predecessorScope?.id,
    );
    assert.equal(
      follower?.startAfter &&
        "scopeId" in follower.startAfter &&
        follower.startAfter.scopeId,
      "scope-2",
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const effects = file.slides[0]?.animations?.[0]?.effects ?? [];
    const fileEffect = (testId: string) =>
      effects.find(({ targetId }) => targetId === objectId(testId));
    const fileFirst = fileEffect("after-predecessor");
    const fileFollower = fileEffect("after-self-started-sequence");
    const containers = file.slides[0]?.animationContainers ?? [];
    const predecessor = containers.find(
      ({ children, endSync }) =>
        endSync === "all" &&
        children.length === 2 &&
        children.every(({ kind }) => kind === "effect") &&
        children.some(
          ({ timingId }) =>
            timingId === fileEffect("predecessor-long")?.timingId,
        ) &&
        children.some(
          ({ timingId }) =>
            timingId === fileEffect("predecessor-short")?.timingId,
        ),
    );
    assert.ok(predecessor);
    const sequence = containers.find(
      ({ children, endSync }) =>
        endSync === "all" &&
        children.some(
          ({ kind, timingId }) =>
            kind === "container" && timingId === predecessor.timingId,
        ) &&
        children.some(
          ({ kind, timingId }) =>
            kind === "effect" && timingId === fileFirst?.timingId,
        ),
    );
    assert.ok(sequence);
    assert.equal(fileFirst?.startAfter?.timingId, predecessor.timingId);
    assert.equal(fileFirst?.startAfter?.offsetMs, 0);
    assert.equal(fileFollower?.startAfter?.timingId, sequence.timingId);
  });

  it("独立读取器拒绝相互等待及包含在自身完成条件中的环", async () => {
    const result = await build(createElement(ReaderCycleDeck));
    const original = await validatePptx(result.pptx);
    assert.deepEqual(original.errors, []);
    const scopes = (original.slides[0]?.animationContainers ?? []).filter(
      ({ endSync }) => endSync === "all",
    );
    assert.equal(scopes.length, 2);
    const [first, second] = scopes;
    assert.ok(first && second);
    const firstEffect = first.children.find(({ kind }) => kind === "effect");
    assert.ok(firstEffect);
    const archive = unzipSync(result.pptx);
    const slidePart = archive["ppt/slides/slide1.xml"];
    assert.ok(slidePart);
    const xml = strFromU8(slidePart);

    const mutualCycle = startAfterCondition(
      startAfterCondition(xml, first.timingId, second.timingId),
      second.timingId,
      first.timingId,
    );
    const parentCycle = startAfterCondition(
      xml,
      firstEffect.timingId,
      first.timingId,
    );
    for (const brokenXml of [mutualCycle, parentCycle]) {
      const brokenFile = await validatePptx(
        zipSync({
          ...archive,
          "ppt/slides/slide1.xml": strToU8(brokenXml),
        }),
      );
      assert.ok(
        brokenFile.errors.some(
          ({ code }) => code === "invalid-animation-end-dependency",
        ),
      );
    }
  });

  it("独立读取器把父启动前置传给实际子节点，并接受合法前向引用", async () => {
    const result = await build(createElement(UnknownStartSequence, {}));
    const original = await validatePptx(result.pptx);
    assert.deepEqual(original.errors, []);
    const slide = result.report.slides[0];
    const objectId = (testId: string) =>
      slide?.objects.find(({ testId: id }) => id === testId)?.id;
    const effects = original.slides[0]?.animations?.[0]?.effects ?? [];
    const sourceEffect = effects.find(
      ({ targetId }) => targetId === objectId("unknown-source"),
    );
    const firstEffect = effects.find(
      ({ targetId }) => targetId === objectId("sequence-first"),
    );
    assert.ok(sourceEffect && firstEffect);
    const containers = original.slides[0]?.animationContainers ?? [];
    const source = containers.find(
      ({ children, endSync }) =>
        endSync === "all" &&
        children.some(
          ({ kind, timingId }) =>
            kind === "effect" && timingId === sourceEffect.timingId,
        ),
    );
    const follow = containers.find(
      ({ children, endSync }) =>
        endSync === "all" &&
        children.some(
          ({ kind, timingId }) =>
            kind === "effect" && timingId === firstEffect.timingId,
        ),
    );
    assert.ok(source && follow);
    const archive = unzipSync(result.pptx);
    const slidePart = archive["ppt/slides/slide1.xml"];
    assert.ok(slidePart);
    const xml = strFromU8(slidePart);

    const inheritedCycle = startAfterCondition(
      xml,
      source.timingId,
      firstEffect.timingId,
    );
    const cycleRead = await validatePptx(
      zipSync({
        ...archive,
        "ppt/slides/slide1.xml": strToU8(inheritedCycle),
      }),
    );
    assert.ok(
      cycleRead.errors.some(
        ({ code }) => code === "invalid-animation-end-dependency",
      ),
    );

    const forwardReference = startAtDelay(inheritedCycle, follow.timingId, 0);
    const forwardRead = await validatePptx(
      zipSync({
        ...archive,
        "ppt/slides/slide1.xml": strToU8(forwardReference),
      }),
    );
    assert.deepEqual(forwardRead.errors, []);
  });

  it("空子作用域保留顺序边界，但不占交错位置", async () => {
    const sequenceResult = await build(createElement(EmptySequenceMember));
    const sequenceMotion = sequenceResult.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(
      sequenceMotion.steps[0]?.effects.map(({ target, startMs, endMs }) => [
        target.testId,
        startMs,
        endMs,
      ]),
      [
        ["before-empty", 0, 200],
        ["after-empty", 200, 300],
        ["after-empty-completion", 210, 310],
      ],
    );
    assert.deepEqual(
      sequenceMotion.scopes?.find(({ id }) => id === "scope-2"),
      {
        id: "scope-2",
        parentId: "scope-1",
        mode: "parallel",
        delayMs: 0,
        startMs: 200,
        endMs: 200,
        completion: "known",
        members: [],
      },
    );
    assert.deepEqual((await validatePptx(sequenceResult.pptx)).errors, []);

    const staggerResult = await build(createElement(EmptyStaggerMember));
    const staggerMotion = staggerResult.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(
      staggerMotion.steps[0]?.effects.map(({ target, startMs }) => [
        target.testId,
        startMs,
      ]),
      [
        ["stagger-before-empty", 0],
        ["stagger-after-empty", 100],
      ],
    );
    assert.deepEqual((await validatePptx(staggerResult.pptx)).errors, []);
  });

  it("非活动延迟子作用域不拉长父完成，且保留自身完成边界", async () => {
    for (const sourceMode of ["empty", "static", "disabled"] as const) {
      const result = await build(
        createElement(DelayedInactiveSequenceMember, { sourceMode }),
      );
      const slide = result.report.slides[0];
      const motion = slide?.capabilities?.["motion"] as MotionObservation;
      assert.deepEqual(
        motion.steps[0]?.effects.map(({ target, startMs, endMs }) => [
          target.testId,
          startMs,
          endMs,
        ]),
        [
          ["first", 0, 100],
          ["second", 100, 200],
          ["after-sequence", 200, 300],
          ["after-inactive", 1125, 1225],
        ],
      );

      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      assert.deepEqual(
        file.slides[0]?.animations?.[0]?.effects.map(
          ({ targetId, delayMs }) => [
            slide?.objects.find(({ id }) => id === targetId)?.testId,
            delayMs,
          ],
        ),
        [
          ["first", 0],
          ["second", 100],
          ["after-sequence", 200],
          ["after-inactive", 1125],
        ],
      );
    }
  });

  it("自行指定起点的并行成员仍计入父作用域整体完成", async () => {
    const result = await build(createElement(SelfStartedMember));
    const slide = result.report.slides[0];
    const motion = slide?.capabilities?.["motion"] as MotionObservation;
    assert.deepEqual(
      motion.steps[0]?.effects.map(({ target, startMs, endMs }) => [
        target.testId,
        startMs,
        endMs,
      ]),
      [
        ["leader", 0, 100],
        ["self-started", 120, 170],
        ["after-group", 180, 230],
      ],
    );
    assert.equal(motion.scopes?.[0]?.endMs, 170);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
  });

  it("Box 默认入场只补给没有自带入场的对象，静态和禁用对象不占交错位置", async () => {
    const result = await build(createElement(BoxAndDisabledMembers));
    const slide = result.report.slides[0];
    const motion = slide?.capabilities?.["motion"] as MotionObservation;
    assert.deepEqual(
      motion.steps[0]?.effects.map(({ target, startMs, delayMs }) => [
        target.testId,
        startMs,
        delayMs,
      ]),
      [
        ["own-entry", 0, 0],
        ["stagger-first", 0, 0],
        ["stagger-second", 100, 0],
        ["box-default", 800, 800],
      ],
    );
    const stagger = motion.scopes?.find(({ mode }) => mode === "stagger");
    assert.deepEqual(
      stagger?.members.map((member) =>
        member.kind === "object"
          ? slide?.objects.find(({ id }) => id === member.id)?.testId
          : member.kind,
      ),
      ["stagger-first", "stagger-second"],
    );
    assert.equal(
      slide?.objects.some(({ testId }) => testId === "disabled-descendant"),
      true,
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.animations?.[0]?.effects.map(({ targetId, delayMs }) => [
        slide?.objects.find(({ id }) => id === targetId)?.testId,
        delayMs,
      ]),
      [
        ["own-entry", 0],
        ["stagger-first", 0],
        ["stagger-second", 100],
        ["box-default", 800],
      ],
    );
  });

  it("可见 Box 背景继承默认入场，但不覆盖子项自有入场或禁用状态", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { slideId: "visible-box-motion" },
          createElement(
            Box,
            {
              testId: "visible-background",
              enter: fade({ duration: 200, delay: 800 }),
              style: { width: 120, height: 80, backgroundColor: "accent1" },
            },
            createElement(Shape, {
              testId: "inherited-child",
              x: 0,
              y: 0,
              width: 20,
              height: 20,
            }),
            createElement(Shape, {
              testId: "own-entry-child",
              x: 24,
              y: 0,
              width: 20,
              height: 20,
              enter: fade({ duration: 100 }),
            }),
          ),
          createElement(
            Box,
            {
              testId: "disabled-background",
              motion: false,
              enter: fade(),
              style: { width: 100, height: 60, backgroundColor: "accent2" },
            },
            createElement(Shape, {
              testId: "disabled-child",
              x: 0,
              y: 0,
              width: 20,
              height: 20,
              enter: fade(),
            }),
          ),
        ),
      ),
    );
    const slide = result.report.slides[0];
    const motion = slide?.capabilities?.["motion"] as MotionObservation;
    const file = await validatePptx(result.pptx);

    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      motion.steps[0]?.effects.map(
        ({ target, startMs, delayMs, durationMs }) => [
          target.testId,
          startMs,
          delayMs,
          durationMs,
        ],
      ),
      [
        ["own-entry-child", 0, 0, 100],
        ["visible-background", 800, 800, 200],
        ["inherited-child", 800, 800, 200],
      ],
    );
    assert.deepEqual(
      file.slides[0]?.animations?.[0]?.effects.map(({ targetId, delayMs }) => [
        slide?.objects.find(({ id }) => id === targetId)?.testId,
        delayMs,
      ]),
      [
        ["own-entry-child", 0],
        ["visible-background", 800],
        ["inherited-child", 800],
      ],
    );
  });

  it("同名组件实例互相隔离，普通函数、Fragment 和原生 Group 不自动成作用域", async () => {
    const reused = await build(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { slideId: "reused" },
          createElement(ReusedNamedScope),
          createElement(ReusedNamedScope),
        ),
      ),
    );
    const reusedMotion = reused.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(
      reusedMotion.steps[0]?.effects.map(({ target, startMs }) => [
        target.testId,
        startMs,
      ]),
      [
        ["source", 0],
        ["source", 0],
        ["follower", 110],
        ["follower", 110],
      ],
    );
    assert.equal(reusedMotion.scopes?.length, 4);

    const ordinary = await build(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { slideId: "ordinary" },
          createElement(OrdinaryComposition),
        ),
      ),
    );
    const ordinaryMotion = ordinary.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.equal(ordinaryMotion.scopes, undefined);
    assert.deepEqual(
      ordinaryMotion.steps[0]?.effects.map(({ target, startMs }) => [
        target.testId,
        startMs,
      ]),
      [
        ["function-first", 0],
        ["group-first", 0],
        ["group-second", 0],
      ],
    );
  });
});
