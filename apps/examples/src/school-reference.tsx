import { readFileSync } from "node:fs";

import {
  after,
  Box,
  Deck,
  defineTheme,
  fade,
  floatUp,
  font,
  Group,
  MotionScope,
  path,
  preset,
  Shape,
  Slide,
  Step,
  Text,
  useMotionScope,
  wipe,
  zoom,
  type MotionMoment,
  type MotionScopeHandle,
} from "@react-pptx/react-pptx";
import type { ReactNode } from "react";

const asset = (name: string) =>
  readFileSync(new URL(`./assets/fonts/${name}`, import.meta.url));

const montserrat = font("Montserrat", [
  { data: asset("Montserrat-Regular.ttf") },
  { data: asset("Montserrat-Bold.ttf"), weight: 700 },
  { data: asset("Montserrat-ExtraBold.ttf"), weight: 800 },
]);

// 原稿使用 Tw Cen MT；用有明确许可的 Jost 提供度量，交付仍保留原字体名。
const twCenMt = font(
  "Tw Cen MT",
  [
    { data: asset("Jost-400-Book.ttf") },
    { data: asset("Jost-700-Bold.ttf"), weight: 700 },
  ],
  { substitute: true, embeddable: false },
);

const schoolTheme = defineTheme({
  name: "PowerPoint School",
  colors: {
    lavender: "#9188E5",
    deep: "#837AD9",
    sun: "#FFC000",
    paper: "#F2F2F2",
    silver: "#D9D9D9",
    gray: "#808080",
  },
  fonts: { major: montserrat, minor: montserrat },
});

const brief = "Here You can Add Some Brief Text That will Explain Your Tittle";
const serviceBody = `${brief} This is Just a Demo Text So You Can Replace These Texts`;

type SchoolDot = {
  readonly x: number;
  readonly y: number;
  readonly size: number;
  readonly color: "deep" | "sun";
  readonly half?: "left" | "right" | "top" | "bottom";
};

const halfDisc = {
  left: path("M1 0 A1 1 0 0 0 1 2 Z", { viewBox: [1, 2] }),
  right: path("M0 0 A1 1 0 0 1 0 2 Z", { viewBox: [1, 2] }),
  top: path("M0 1 A1 1 0 0 1 2 1 Z", { viewBox: [2, 1] }),
  bottom: path("M0 0 A1 1 0 0 0 2 0 Z", { viewBox: [2, 1] }),
} as const;

const serviceDots: readonly SchoolDot[] = [
  { x: 252.7, y: 145.1, size: 17, color: "deep" },
  { x: 108.3, y: 70.7, size: 28, color: "sun" },
  { x: -13.9, y: 651.4, size: 28, color: "sun", half: "right" },
  { x: -13.9, y: 195.7, size: 28, color: "deep", half: "right" },
  { x: 1139, y: 350.7, size: 17, color: "deep" },
  { x: 1149.5, y: 70.7, size: 17, color: "deep" },
  { x: 1176.6, y: 623, size: 28, color: "sun" },
  { x: 1015.8, y: 175.3, size: 28, color: "sun" },
  { x: 1265.8, y: 575.4, size: 28, color: "deep", half: "left" },
  { x: 1266, y: 119.7, size: 28, color: "sun", half: "left" },
  { x: 647.5, y: 706, size: 28, color: "deep", half: "top" },
];

const agendaDots: readonly SchoolDot[] = [
  { x: 68.9, y: 32.1, size: 28, color: "sun" },
  { x: 240.4, y: 563.6, size: 17, color: "deep" },
  { x: -13.8, y: 651.4, size: 28, color: "sun", half: "right" },
  { x: -13.8, y: 195.7, size: 28, color: "deep", half: "right" },
  { x: 495.5, y: 66.7, size: 17, color: "deep" },
  { x: 455.8, y: 584.3, size: 28, color: "sun" },
  { x: 625.7, y: 18.3, size: 28, color: "deep", half: "left" },
  { x: 625.7, y: 587.6, size: 28, color: "deep", half: "left" },
  { x: 890, y: 612.6, size: 17, color: "deep" },
  { x: 695.1, y: 706, size: 28, color: "sun", half: "top" },
  { x: 626.2, y: 263.8, size: 28, color: "deep", half: "right" },
  { x: 1135.5, y: 66.7, size: 17, color: "deep" },
  { x: 1108, y: 652.9, size: 28, color: "sun" },
  { x: 1265.7, y: 18.3, size: 28, color: "deep", half: "left" },
  { x: 1265.7, y: 587.6, size: 28, color: "deep", half: "left" },
];

function Decorations({ dots }: { readonly dots: readonly SchoolDot[] }) {
  return dots.map((dot, index) => {
    const half = dot.half;
    const box =
      half === undefined
        ? { left: dot.x, top: dot.y, width: dot.size, height: dot.size }
        : {
            left: half === "right" ? dot.x + dot.size / 2 : dot.x,
            top: half === "bottom" ? dot.y + dot.size / 2 : dot.y,
            width:
              half === "left" || half === "right" ? dot.size / 2 : dot.size,
            height:
              half === "top" || half === "bottom" ? dot.size / 2 : dot.size,
          };
    return (
      <Shape
        key={index}
        geometry={half === undefined ? "ellipse" : halfDisc[half]}
        style={{ position: "absolute", ...box, backgroundColor: dot.color }}
      />
    );
  });
}

function PagePill({
  testId,
  children,
  width = 280,
}: {
  readonly testId: string;
  readonly children: string;
  readonly width?: number;
}) {
  return (
    <Shape
      testId={testId}
      geometry="roundRect"
      style={{
        width,
        height: 48,
        borderRadius: 24,
        backgroundColor: "deep",
      }}
      textStyle={{
        color: "lt1",
        fontFamily: twCenMt,
        fontSize: "20pt",
        fontWeight: 700,
        align: "center",
        verticalAlign: "middle",
      }}
    >
      {children}
    </Shape>
  );
}

function SchoolHeader({
  title,
  lead,
}: {
  readonly title: string;
  readonly lead: string;
}) {
  return (
    <Box
      testId="school-services-header"
      style={{
        display: "flex",
        flexDirection: "column",
        position: "absolute",
        left: 0,
        right: 0,
        top: 13.8,
        alignItems: "center",
        fontFamily: twCenMt,
        color: "paper",
        align: "center",
      }}
    >
      <Text
        testId="school-services-title"
        style={{ fontSize: "48pt", fontWeight: 700 }}
      >
        {title}
      </Text>
      <Text
        testId="school-services-lead"
        style={{ width: 648, fontSize: "16pt" }}
      >
        {lead}
      </Text>
    </Box>
  );
}

function MotionMember({ children }: { readonly children: ReactNode }) {
  const scope = useMotionScope();
  return <MotionScope scope={scope}>{children}</MotionScope>;
}

const serviceTileLayers = [
  { size: 157, radius: 53.4, color: "paper" },
  { size: 136, radius: 45.2, color: "lavender" },
  { size: 117, radius: 39.8, color: "deep" },
  { size: 95, radius: 30.8, color: "lavender" },
  { size: 76, radius: 24.6, color: "deep" },
] as const;

function CheckIcon({ testId }: { readonly testId: string }) {
  const unit = 1;
  const bar = 3 * unit;
  return (
    <Group
      testId={testId}
      enter={zoom()}
      style={{
        position: "absolute",
        left: 52.5,
        top: 52.5,
        width: 52,
        height: 52,
      }}
    >
      <Shape
        geometry={preset("donut", { adj: 4704 })}
        style={{ position: "absolute", inset: 0, backgroundColor: "lt1" }}
      />
      <Group
        style={{
          position: "absolute",
          left: 19 * unit,
          top: 8 * unit,
          width: 12 * unit,
          height: 31 * unit,
          rotate: 45,
        }}
      >
        <Shape
          geometry="roundRect"
          style={{
            position: "absolute",
            left: 4 * unit,
            top: 24 * unit,
            width: bar,
            height: 12 * unit,
            borderRadius: bar / 2,
            rotate: 90,
            backgroundColor: "lt1",
          }}
        />
        <Shape
          geometry="roundRect"
          style={{
            position: "absolute",
            left: 9 * unit,
            top: 0,
            width: bar,
            height: 30 * unit,
            borderRadius: bar / 2,
            backgroundColor: "lt1",
          }}
        />
      </Group>
    </Group>
  );
}

function ServiceCard({
  id,
  title,
  body,
}: {
  readonly id: string;
  readonly title: string;
  readonly body: string;
}) {
  const tileScope = useMotionScope();
  const captionScope = useMotionScope();
  return (
    <Box
      testId={`school-service-${id}`}
      style={{
        display: "flex",
        flex: 1,
        flexDirection: "column",
        alignItems: "center",
        color: "paper",
        align: "center",
      }}
    >
      <MotionScope scope={tileScope} mode="stagger" each={250}>
        <Box
          testId={`school-service-${id}-icon-tile`}
          style={{
            position: "relative",
            display: "flex",
            flexDirection: "column",
            width: 157,
            height: 157,
            alignItems: "center",
            justifyContent: "center",
          }}
        >
          {serviceTileLayers.map((layer, index) => (
            <MotionMember key={layer.size}>
              <Shape
                testId={`school-service-${id}-tile-${index + 1}`}
                geometry="roundRect"
                enter={zoom()}
                style={{
                  position: "absolute",
                  left: (157 - layer.size) / 2,
                  top: (157 - layer.size) / 2,
                  width: layer.size,
                  height: layer.size,
                  borderRadius: layer.radius,
                  backgroundColor: layer.color,
                }}
              />
            </MotionMember>
          ))}
          <MotionMember>
            <CheckIcon testId={`school-service-${id}-check`} />
          </MotionMember>
        </Box>
      </MotionScope>
      <MotionScope scope={captionScope} startAt={after(tileScope.end, 0)}>
        <Group
          testId={`school-service-${id}-caption`}
          enter={floatUp()}
          style={{
            display: "flex",
            flexDirection: "column",
            alignSelf: "stretch",
            alignItems: "center",
            marginLeft: 20,
            marginRight: 20,
          }}
        >
          <Text
            testId={`school-service-${id}-title`}
            style={{ marginTop: 21.8, fontSize: "28pt", fontWeight: 700 }}
          >
            {title}
          </Text>
          <Text
            testId={`school-service-${id}-body`}
            style={{ marginTop: -2.8, alignSelf: "stretch" }}
          >
            {body}
          </Text>
        </Group>
      </MotionScope>
    </Box>
  );
}

function ServicesPage() {
  const services = ["first", "second", "third"].map((id) => ({
    id,
    title: "Your Text",
    body: brief,
  }));
  return (
    <Slide slideId="school-services" background="lavender">
      <SchoolHeader title="Our Services" lead={serviceBody} />
      <Box
        testId="school-services-row"
        style={{
          display: "flex",
          position: "absolute",
          left: 61,
          top: 280,
          width: 1155,
        }}
      >
        {services.map((service) => (
          <Step key={service.id}>
            <ServiceCard {...service} />
          </Step>
        ))}
      </Box>
      <Decorations dots={serviceDots} />
    </Slide>
  );
}

const taskList = [
  { id: "first", day: 7 },
  { id: "second", day: 17 },
  { id: "third", day: 19 },
] as const;

function CalendarHighlight({
  testId,
  startAt,
  step,
}: {
  readonly testId: string;
  readonly startAt: MotionMoment;
  readonly step: number;
}) {
  const scope = useMotionScope();
  return (
    <Step n={step}>
      <MotionScope scope={scope} startAt={startAt}>
        <Shape
          testId={testId}
          geometry="ellipse"
          enter={fade()}
          style={{
            position: "absolute",
            left: 16.95,
            top: -4.7,
            width: 35,
            height: 35,
            backgroundColor: "deep",
          }}
        />
      </MotionScope>
    </Step>
  );
}

type CalendarReveal = {
  readonly step: number;
  readonly startAt: MotionMoment;
};

function Calendar({
  revealForDay,
}: {
  readonly revealForDay: (day: number) => CalendarReveal | undefined;
}) {
  const weekdays = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  return (
    <Box
      testId="school-calendar"
      style={{
        display: "flex",
        flexDirection: "column",
        position: "absolute",
        left: 0,
        top: 0,
        width: 640,
        height: 720,
        backgroundColor: "lavender",
        alignItems: "center",
        paddingTop: 91,
        align: "center",
      }}
    >
      <PagePill testId="school-calendar-title">Agenda Slide</PagePill>
      <Text
        testId="school-calendar-month"
        style={{
          marginTop: 75.8,
          fontSize: "24pt",
          fontWeight: 700,
          color: "lt1",
        }}
      >
        October
      </Text>
      <Box
        style={{
          marginTop: 27.6,
          display: "grid",
          gridTemplateColumns: "repeat(7, 68.9px)",
          gridTemplateRows: "44px repeat(5, 42.6px)",
          fontSize: "16pt",
          fontWeight: 700,
          color: "lt1",
        }}
      >
        {weekdays.map((day) => (
          <Text key={day} testId={`school-weekday-${day.toLowerCase()}`}>
            {day}
          </Text>
        ))}
        {Array.from({ length: 31 }, (_, index) => {
          const day = index + 1;
          const task = taskList.find((item) => item.day === day);
          const reveal = task === undefined ? undefined : revealForDay(day);
          return (
            <Box
              key={day}
              testId={`school-day-${day}`}
              style={{
                position: "relative",
                ...(day === 1 && { gridColumn: 5 }),
              }}
            >
              {task !== undefined && reveal !== undefined && (
                <CalendarHighlight
                  testId={`school-day-${day}-highlight`}
                  startAt={reveal.startAt}
                  step={reveal.step}
                />
              )}
              <Text testId={`school-day-${day}-label`}>{String(day)}</Text>
            </Box>
          );
        })}
      </Box>
    </Box>
  );
}

function TimelineTask({
  id,
  day,
  index,
  eventScope,
}: {
  readonly id: string;
  readonly day: number;
  readonly index: number;
  readonly eventScope: MotionScopeHandle;
}) {
  const taskScope = useMotionScope();
  const markerScope = useMotionScope();
  const revealStart = eventScope.cue("revealStart");
  return (
    <Box
      testId={`school-task-${id}`}
      style={{ position: "relative", display: "flex", height: 145.5 }}
    >
      <Step n={index + 1}>
        <MotionScope scope={eventScope}>
          <MotionScope scope={taskScope} mode="sequence">
            {index > 0 && (
              <Shape
                testId={`school-task-${id}-connector`}
                geometry="rect"
                enter={wipe({ from: "top", duration: 250 })}
                style={{
                  position: "absolute",
                  left: 12,
                  top: 39 - 145.5,
                  width: 3,
                  height: 118.5,
                  backgroundColor: "deep",
                }}
              />
            )}
            <MotionScope
              scope={markerScope}
              mode="stagger"
              each={250}
              publishStart={revealStart}
            >
              <Box style={{ position: "relative", width: 27, paddingTop: 12 }}>
                <Shape
                  testId={`school-task-${id}-ring`}
                  geometry="ellipse"
                  enter={zoom()}
                  style={{
                    width: 27,
                    height: 27,
                    borderColor: "deep",
                    borderWidth: 3,
                  }}
                />
                <Shape
                  testId={`school-task-${id}-dot`}
                  geometry="ellipse"
                  enter={zoom()}
                  style={{
                    position: "absolute",
                    left: 9,
                    top: 21,
                    width: 9,
                    height: 9,
                    backgroundColor: "deep",
                  }}
                />
              </Box>
            </MotionScope>
            <Group
              testId={`school-task-${id}-text`}
              enter={fade()}
              style={{ marginLeft: 21.6, paddingTop: 4.8 }}
            >
              <Text
                testId={`school-task-${id}-date`}
                style={{ fontSize: "24pt", fontWeight: 700, color: "deep" }}
              >
                October {day}, 2020
              </Text>
              <Text
                testId={`school-task-${id}-body`}
                style={{
                  marginTop: 0.6,
                  width: 418,
                  fontSize: "16pt",
                  color: "gray",
                }}
              >
                {brief}
              </Text>
            </Group>
          </MotionScope>
        </MotionScope>
      </Step>
    </Box>
  );
}

function Timeline({
  revealForTask,
}: {
  readonly revealForTask: (id: string) => MotionScopeHandle;
}) {
  return (
    <Box
      testId="school-timeline"
      style={{
        display: "flex",
        flexDirection: "column",
        position: "absolute",
        left: 640,
        top: 0,
        width: 640,
        height: 720,
        alignItems: "center",
        paddingTop: 91,
      }}
    >
      <PagePill testId="school-timeline-title">Our Tasks</PagePill>
      <Box style={{ position: "absolute", left: 84, top: 200 }}>
        {taskList.map((task, index) => (
          <TimelineTask
            key={task.id}
            id={task.id}
            day={task.day}
            index={index}
            eventScope={revealForTask(task.id)}
          />
        ))}
      </Box>
    </Box>
  );
}

function AgendaPage() {
  const taskScopes = [
    useMotionScope(),
    useMotionScope(),
    useMotionScope(),
  ] as const;
  const revealForTask = (id: string) => {
    const index = taskList.findIndex((task) => task.id === id);
    const scope = taskScopes[index];
    if (scope === undefined) throw new Error(`未知任务：${id}`);
    return scope;
  };
  const revealForDay = (day: number) => {
    const index = taskList.findIndex((task) => task.day === day);
    const scope = taskScopes[index];
    return scope === undefined
      ? undefined
      : { step: index + 1, startAt: scope.cue("revealStart") };
  };

  return (
    <Slide slideId="school-agenda" background="paper">
      <Calendar revealForDay={revealForDay} />
      <Timeline revealForTask={revealForTask} />
      <Decorations dots={agendaDots} />
    </Slide>
  );
}

/** 两页参考稿使用公开的 Deck/Slide/动效 API；字体度量文件由本 examples 包提供。 */
export function schoolReferenceDeck() {
  return (
    <Deck
      title="PowerPoint School 参考页"
      width={1280}
      height={720}
      theme={schoolTheme}
      style={{ fontFamily: "major", fontSize: "18pt", color: "paper" }}
    >
      <ServicesPage />
      <AgendaPage />
    </Deck>
  );
}

export { AgendaPage as SchoolAgendaPage, ServicesPage as SchoolServicesPage };
