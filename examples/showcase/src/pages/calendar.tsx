import {
  Group,
  MotionScope,
  Shape,
  Step,
  Text,
  fade,
  fillColor,
  useMotionScope,
} from "@react-pptx/react-pptx";

import { PageFrame } from "../templates/page-frame.tsx";
import { colors } from "../theme.ts";

const events = [
  { day: 4, title: "确定主题", body: "先明确要回答的问题", color: colors.teal },
  {
    day: 12,
    title: "组织内容",
    body: "再安排材料的先后关系",
    color: colors.coral,
  },
  {
    day: 23,
    title: "完成讲述",
    body: "最后检查整段表达的节奏",
    color: "#7369A4",
  },
];

function DateCell({
  day,
  emphasis = false,
}: {
  day: number;
  emphasis?: boolean;
}) {
  const event = events.find((item) => item.day === day);
  return (
    <Shape
      testId={`date-${day}`}
      x={64 + ((day - 1) % 7) * 92}
      y={240 + Math.floor((day - 1) / 7) * 78}
      width={80}
      height={64}
      fill={colors.pale}
      text={String(day)}
      textStyle={{ fontSize: 22, align: "center", verticalAlign: "middle" }}
      {...(emphasis
        ? {
            emphasis: fillColor({
              from: colors.pale,
              to: event!.color,
              duration: 450,
            }),
          }
        : {})}
    />
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
  return (
    <Step n={index + 1}>
      <MotionScope scope={date}>
        <DateCell day={event.day} emphasis />
      </MotionScope>
      <MotionScope scope={note} startAt={date.start}>
        <Group testId={`event-${event.day}`} enter={fade({ duration: 450 })}>
          <Shape
            x={790}
            y={245 + index * 115}
            width={6}
            height={80}
            fill={event.color}
          />
          <Text
            x={814}
            y={240 + index * 115}
            width={390}
            height={48}
            style={{ fontSize: 28 }}
          >
            {event.day} 日 · {event.title}
          </Text>
          <Text
            x={814}
            y={292 + index * 115}
            width={390}
            height={45}
            style={{ fontSize: 22, color: colors.muted }}
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
    <PageFrame
      slideId="calendar"
      title="同步讲述日期与事项"
      subtitle="日期告诉听众什么时候；对应事项解释那一天发生什么。"
    >
      {["一", "二", "三", "四", "五", "六", "日"].map((label, i) => (
        <Text
          key={label}
          x={64 + i * 92}
          y={190}
          width={80}
          height={40}
          style={{ fontSize: 20, align: "center", color: colors.muted }}
        >
          {label}
        </Text>
      ))}
      {Array.from({ length: 28 }, (_, i) => i + 1)
        .filter((day) => !events.some((event) => event.day === day))
        .map((day) => (
          <DateCell key={day} day={day} />
        ))}
      {events.map((event, index) => (
        <CalendarEvent key={event.day} event={event} index={index} />
      ))}
      <Text
        x={64}
        y={620}
        width={1100}
        height={45}
        style={{ fontSize: 18, color: colors.muted }}
      >
        教学日历：以周一开始的示意月份，日期仅用于说明讲述顺序。
      </Text>
    </PageFrame>
  );
}
