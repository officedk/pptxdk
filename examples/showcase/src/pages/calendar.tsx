import {
  fillColor,
  floatUp,
  Group,
  MotionScope,
  Shape,
  Slide,
  Step,
  Text,
  useMotionScope,
} from "@officedk/pptxdk/react";

import { colors, motion } from "../theme.ts";

const events = [
  {
    day: 4,
    title: "Lorem ipsum",
    body: "Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
  },
  {
    day: 12,
    title: "Dolor sit",
    body: "Sed do eiusmod tempor incididunt ut labore et dolore.",
  },
  {
    day: 23,
    title: "Amet elit",
    body: "Ut enim ad minim veniam, quis nostrud exercitation.",
  },
];

function DateCell({
  day,
  emphasis = false,
}: {
  day: number;
  emphasis?: boolean;
}) {
  return (
    <Shape
      id={`date-${day}`}
      style={{
        position: "absolute",
        left: 80 + ((day - 1) % 7) * 88,
        top: 248 + Math.floor((day - 1) / 7) * 72,
        width: 76,
        height: 58,
        background: colors.surface,
        fontSize: 21,
        fontFamily: "medium",
        fontWeight: 500,
        color: colors.ink,
        align: "center",
        verticalAlign: "middle",
      }}
      {...(emphasis
        ? {
            emphasis: fillColor({
              from: colors.surface,
              to: colors.accentSoft,
              duration: motion.reveal,
            }),
          }
        : {})}
    >
      {String(day).padStart(2, "0")}
    </Shape>
  );
}

function CalendarEvent({
  event,
  index,
}: {
  event: (typeof events)[number];
  index: number;
}) {
  const date = useMotionScope();
  const note = useMotionScope();
  const y = 246 + index * 126;
  return (
    <Step n={index + 1}>
      <MotionScope scope={date}>
        <DateCell day={event.day} emphasis />
      </MotionScope>
      <MotionScope scope={note} startAt={date.start}>
        <Group
          id={`event-${event.day}`}
          enter={floatUp({ duration: motion.reveal })}
        >
          <Shape
            style={{
              position: "absolute",
              left: 790,
              top: y,
              width: 5,
              height: 88,
              background: colors.accent,
            }}
          />
          <Text
            style={{
              position: "absolute",
              left: 820,
              top: y - 4,
              width: 330,
              height: 38,
              fontSize: 28,
              lineHeight: 32,
              fontWeight: 700,
              color: colors.ink,
            }}
          >
            {String(event.day).padStart(2, "0")} / {event.title}
          </Text>
          <Text
            style={{
              position: "absolute",
              left: 820,
              top: y + 44,
              width: 350,
              height: 54,
              fontSize: 17,
              lineHeight: 23,
              color: colors.muted,
            }}
          >
            {event.body}
          </Text>
        </Group>
      </MotionScope>
    </Step>
  );
}

export function CalendarPage() {
  return (
    <Slide id="calendar" background={colors.paper}>
      <Shape
        style={{
          position: "absolute",
          left: 80,
          top: 58,
          width: 42,
          height: 5,
          background: colors.accent,
        }}
      />
      <Text
        style={{
          position: "absolute",
          left: 80,
          top: 90,
          width: 550,
          height: 76,
          fontSize: 58,
          lineHeight: 62,
          fontWeight: 700,
        }}
      >
        October
      </Text>
      <Text
        style={{
          position: "absolute",
          left: 790,
          top: 94,
          width: 400,
          height: 46,
          fontSize: 34,
          lineHeight: 38,
          fontFamily: "medium",
          fontWeight: 500,
          color: colors.accent,
        }}
      >
        Plans, in sequence.
      </Text>
      <Text
        style={{
          position: "absolute",
          left: 790,
          top: 150,
          width: 370,
          height: 54,
          fontSize: 17,
          lineHeight: 23,
          color: colors.muted,
        }}
      >
        Lorem ipsum dolor sit amet, consectetur adipiscing elit.
      </Text>
      {["M", "T", "W", "T", "F", "S", "S"].map((label, index) => (
        <Text
          key={`${label}-${index}`}
          style={{
            position: "absolute",
            left: 80 + index * 88,
            top: 204,
            width: 76,
            height: 30,
            fontSize: 15,
            fontWeight: 700,
            align: "center",
            color: colors.muted,
          }}
        >
          {label}
        </Text>
      ))}
      {Array.from({ length: 28 }, (_, index) => index + 1)
        .filter((day) => !events.some((event) => event.day === day))
        .map((day) => (
          <DateCell key={day} day={day} />
        ))}
      {events.map((event, index) => (
        <CalendarEvent key={event.day} event={event} index={index} />
      ))}
      <Text
        style={{
          position: "absolute",
          left: 790,
          top: 634,
          width: 240,
          height: 26,
          fontSize: 14,
          fontFamily: "medium",
          fontWeight: 500,
          color: colors.muted,
        }}
      >
        08 / CALENDAR
      </Text>
    </Slide>
  );
}
