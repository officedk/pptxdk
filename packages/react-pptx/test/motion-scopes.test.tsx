import assert from "node:assert/strict";

import {
  after,
  Box,
  build,
  BuildError,
  Deck,
  fade,
  Group,
  MotionScope,
  Shape,
  Slide,
  Step,
  Text,
  useMotionScope,
  type MotionObservation,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createElement, Fragment } from "react";
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
}: {
  readonly parentStartGate: boolean;
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
  return createElement(
    Deck,
    { width: 320, height: 190 },
    createElement(
      Slide,
      { slideId: "condition-join-start-alias" },
      createElement(
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
      ),
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
          ? { scopeId: chainScope.id, point: "start" as const, offsetMs: 5000 }
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
