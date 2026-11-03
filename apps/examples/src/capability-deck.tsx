import { readFileSync } from "node:fs";

import {
  after,
  Box,
  Cell,
  Chart,
  createMorphScope,
  Deck,
  defineTheme,
  fade,
  font,
  Group,
  Image,
  MotionScope,
  P,
  repeat,
  Row,
  Shape,
  Slide,
  SlideNumber,
  Span,
  Step,
  Table,
  Text,
  transitions,
  useMotionScope,
  breathe,
} from "@react-pptx/react-pptx";

const asset = (name: string) =>
  readFileSync(new URL(`./assets/fonts/${name}`, import.meta.url));

const jost = font("Jost*", [
  { data: asset("Jost-400-Book.ttf") },
  { data: asset("Jost-700-Bold.ttf"), weight: 700 },
]);

const xiaowei = font("ZCOOL XiaoWei", [
  { data: asset("ZCOOLXiaoWei_400Regular.ttf") },
]);

const observatoryImage = readFileSync(
  new URL("./assets/capability-observatory.png", import.meta.url),
);

const capabilityTheme = defineTheme({
  name: "灵感实验室",
  colors: {
    accent1: "#147C80",
    accent2: "#E37955",
    accent3: "#E6B95B",
    accent4: "#6C7FA6",
    accent5: "#8FA89A",
    accent6: "#D5A18E",
    ink: "#182B3A",
    muted: "#61717D",
    paper: "#F6F4EE",
    surface: "#FFFFFF",
    border: "#D8DFDA",
    positive: "#367C62",
    focus: "#147C80",
  },
  fonts: { display: jost, body: xiaowei },
});

function PageHeader({
  section,
  title,
  subtitle,
}: {
  readonly section: string;
  readonly title: string;
  readonly subtitle: string;
}) {
  return (
    <>
      <Text
        x={60}
        y={28}
        width={500}
        height={20}
        style={{
          color: "accent2",
          fontFamily: jost,
          fontSize: 11,
          fontWeight: "bold",
          letterSpacing: 1,
          textTransform: "uppercase",
        }}
      >
        {section}
      </Text>
      <Text
        x={60}
        y={55}
        width={1120}
        height={48}
        style={{
          color: "ink",
          fontFamily: xiaowei,
          fontSize: 31,
        }}
      >
        {title}
      </Text>
      <Text
        x={62}
        y={108}
        width={1080}
        height={26}
        style={{ color: "muted", fontFamily: xiaowei, fontSize: 15 }}
      >
        {subtitle}
      </Text>
    </>
  );
}

function Footer({ slideId }: { readonly slideId: string }) {
  return (
    <Box
      testId={`${slideId}-footer`}
      style={{
        position: "absolute",
        left: 60,
        top: 672,
        width: 1160,
        height: 24,
        display: "flex",
        flexDirection: "row",
        justifyContent: "space-between",
        alignItems: "center",
      }}
    >
      <Text
        width={320}
        height={24}
        style={{ color: "muted", fontFamily: jost, fontSize: 10 }}
      >
        FIELD NOTES / 2026
      </Text>
      <Text
        testId={`${slideId}-home-link`}
        link={{ action: "home" }}
        width={72}
        height={24}
        style={{ color: "accent1", fontFamily: xiaowei, fontSize: 12 }}
      >
        返回首页
      </Text>
      <Text
        testId={`${slideId}-folio`}
        width={48}
        height={24}
        style={{
          align: "right",
          color: "muted",
          fontFamily: xiaowei,
          fontSize: 12,
        }}
      >
        <SlideNumber style={{ fontFamily: xiaowei, fontSize: 12 }} />
      </Text>
    </Box>
  );
}

function CapabilityCover() {
  return (
    <Slide
      slideId="cover"
      background="paper"
      transition={transitions.fade({ duration: 650 })}
    >
      <Shape
        testId="cover-sun-disc"
        geometry="ellipse"
        x={846}
        y={59}
        width={330}
        height={330}
        style={{ backgroundColor: "accent3" }}
      />
      <Image
        testId="cover-observatory-image"
        src={observatoryImage}
        alt="夕阳下的原创天文台山景插画"
        x={704}
        y={138}
        width={510}
        height={340}
        fit="cover"
      />
      <Text
        x={68}
        y={106}
        width={570}
        height={104}
        style={{
          color: "ink",
          fontFamily: xiaowei,
          fontSize: 39,
        }}
      >
        把想法，写进演示文稿
      </Text>
      <Text
        x={70}
        y={224}
        width={560}
        height={58}
        style={{ color: "accent1", fontFamily: jost, fontSize: 19 }}
      >
        One source · twelve pages · native PowerPoint
      </Text>
      <Text
        x={70}
        y={307}
        width={530}
        height={58}
        style={{ color: "muted", fontFamily: xiaowei, fontSize: 18 }}
      >
        用公开组件组织对象、文字、数据与播放关系。
      </Text>
      <Shape
        testId="cover-browse-button"
        link={{ slideId: "objects" }}
        geometry="roundRect"
        x={70}
        y={420}
        width={210}
        height={58}
        style={{ backgroundColor: "accent1", borderRadius: 16 }}
        textStyle={{
          align: "center",
          color: "surface",
          fontFamily: xiaowei,
          fontSize: 16,
          verticalAlign: "middle",
        }}
      >
        开始浏览
      </Shape>
      <Footer slideId="cover" />
    </Slide>
  );
}

function NativeObjectsPage() {
  return (
    <Slide
      slideId="objects"
      background="paper"
      transition={transitions.strips({ direction: "rightDown", duration: 720 })}
    >
      <PageHeader
        section="01 / Native objects"
        title="对象都能接着编辑"
        subtitle="图片、文本框、形状、原生组合和数据对象各有自己的结构。"
      />
      <Box
        testId="native-object-layout"
        style={{
          position: "absolute",
          left: 60,
          top: 183,
          width: 1160,
          height: 418,
          display: "flex",
          flexDirection: "row",
          gap: 22,
        }}
      >
        <Box
          testId="native-shape-card"
          style={{
            width: 286,
            height: 388,
            padding: 20,
            backgroundColor: "surface",
            borderColor: "border",
            borderWidth: 1,
            borderRadius: 18,
            display: "flex",
            flexDirection: "column",
            gap: 16,
          }}
        >
          <Shape
            testId="native-text-shape"
            geometry="roundRect"
            width={244}
            height={112}
            style={{ backgroundColor: "accent1", borderRadius: 20 }}
            textStyle={{
              align: "center",
              color: "surface",
              fontFamily: xiaowei,
              fontSize: 22,
              verticalAlign: "middle",
            }}
          >
            原生文字形状
          </Shape>
          <Text
            width={242}
            height={100}
            style={{ color: "muted", fontFamily: xiaowei, fontSize: 15 }}
          >
            形状自身可以包含可编辑文字；独立文本框也保留段落和字形样式。
          </Text>
        </Box>
        <Box
          testId="native-image-card"
          style={{
            width: 398,
            height: 388,
            padding: 18,
            backgroundColor: "surface",
            borderColor: "border",
            borderWidth: 1,
            borderRadius: 18,
            display: "flex",
            flexDirection: "column",
            gap: 14,
          }}
        >
          <Image
            testId="native-object-image"
            src={observatoryImage}
            alt="原创天文台风景"
            width={360}
            height={240}
            fit="cover"
          />
          <Text
            width={360}
            height={68}
            style={{ color: "ink", fontFamily: xiaowei, fontSize: 16 }}
          >
            这是一张独立图片对象；标题与说明仍然是原生文字。
          </Text>
        </Box>
        <Box
          testId="native-group-card"
          style={{
            width: 420,
            height: 388,
            padding: 18,
            backgroundColor: "surface",
            borderColor: "border",
            borderWidth: 1,
            borderRadius: 18,
            display: "flex",
            flexDirection: "column",
            gap: 12,
          }}
        >
          <Group
            testId="native-group"
            style={{
              display: "flex",
              flexDirection: "row",
              alignItems: "center",
              gap: 20,
              width: 376,
              height: 190,
            }}
          >
            <Shape
              testId="native-group-tile"
              geometry="roundRect"
              width={174}
              height={154}
              style={{ backgroundColor: "accent2", borderRadius: 18 }}
            />
            <Shape
              testId="native-group-dot"
              geometry="ellipse"
              width={92}
              height={92}
              style={{ backgroundColor: "accent3" }}
            />
          </Group>
          <Text
            width={376}
            height={76}
            style={{ color: "muted", fontFamily: xiaowei, fontSize: 15 }}
          >
            组合保留成员关系，形状与文字可分别选择和修改。
          </Text>
        </Box>
      </Box>
      <Footer slideId="objects" />
    </Slide>
  );
}

function TablePage() {
  return (
    <Slide
      slideId="table"
      background="paper"
      transition={transitions.push({ from: "right", duration: 680 })}
    >
      <PageHeader
        section="02 / Native table"
        title="表格仍是一张表格"
        subtitle="行、列与合并关系写进原生表格；文字和单元格边界可以继续调整。"
      />
      <Table
        testId="capability-merged-table"
        x={62}
        y={202}
        width={1156}
        columns={[264, 264, 264, 364]}
        style={{
          fontFamily: xiaowei,
          fontSize: 16,
          color: "ink",
        }}
      >
        <Row
          style={{ height: 54, backgroundColor: "accent1", color: "surface" }}
        >
          <Cell colSpan={2}>阶段目标</Cell>
          <Cell>负责人</Cell>
          <Cell>交付与进展</Cell>
        </Row>
        <Row style={{ height: 64, backgroundColor: "surface" }}>
          <Cell colSpan={2}>需求验证</Cell>
          <Cell rowSpan={2}>林青</Cell>
          <Cell>访谈与问题定义</Cell>
        </Row>
        <Row style={{ height: 64, backgroundColor: "#EEF3EF" }}>
          <Cell colSpan={2}>交互原型</Cell>
          <Cell>关键路径通过</Cell>
        </Row>
        <Row style={{ height: 64, backgroundColor: "surface" }}>
          <Cell colSpan={2}>试点发布</Cell>
          <Cell>复盘与下一步</Cell>
          <Cell>第 3 周 / 进行中</Cell>
        </Row>
      </Table>
      <Shape
        testId="table-note-accent"
        geometry="roundRect"
        x={62}
        y={524}
        width={1156}
        height={66}
        style={{ backgroundColor: "#E4EFEC", borderRadius: 14 }}
      />
      <Text
        x={84}
        y={541}
        width={1110}
        height={34}
        style={{ color: "accent1", fontFamily: xiaowei, fontSize: 17 }}
      >
        首行横向合并，负责人纵向合并：文件中保留真实 gridSpan / rowSpan 关系。
      </Text>
      <Footer slideId="table" />
    </Slide>
  );
}

const monthlyValues = {
  categories: ["一月", "二月", "三月", "四月"],
  series: [
    { name: "实际营收", values: [8.2, 11.4, 10.7, 13.1], color: "accent1" },
  ],
} as const;

const chartTextStyle = { fontFamily: xiaowei, fontSize: 10 } as const;
const chartAxis = {
  min: 0,
  max: 15,
  majorUnit: 5,
  numberFormat: '0.0"万元"',
  gridlines: { color: "border", width: 1 },
  style: chartTextStyle,
} as const;

function ChartCard({
  testId,
  title,
  x,
}: {
  readonly testId: string;
  readonly title: string;
  readonly x: number;
}) {
  return (
    <>
      <Shape
        testId={`${testId}-card`}
        geometry="roundRect"
        x={x}
        y={184}
        width={364}
        height={442}
        style={{
          backgroundColor: "surface",
          borderColor: "border",
          borderWidth: 1,
          borderRadius: 18,
        }}
      />
      <Text
        x={x + 20}
        y={198}
        width={324}
        height={38}
        style={{ color: "ink", fontFamily: xiaowei, fontSize: 18 }}
      >
        {title}
      </Text>
    </>
  );
}

function CartesianChartsPage() {
  const x = [58, 458, 858] as const;
  return (
    <Slide
      slideId="charts-cartesian"
      background="paper"
      transition={transitions.flip({
        direction: "right",
        duration: 740,
        fallback: transitions.fade(),
      })}
    >
      <PageHeader
        section="03 / Native charts"
        title="六类图表，直接接数据"
        subtitle="柱形、条形、折线、面积、饼图与圆环图均保留可编辑数据系列。"
      />
      <ChartCard testId="column-chart" title="01 / 柱形 / Column" x={x[0]} />
      <ChartCard testId="bar-chart" title="02 / 条形 / Bar" x={x[1]} />
      <ChartCard testId="line-chart" title="03 / 折线 / Line" x={x[2]} />
      <Chart
        testId="native-column-chart"
        kind="column"
        x={x[0] + 10}
        y={244}
        width={344}
        height={340}
        data={monthlyValues}
        legend={false}
        style={chartTextStyle}
        valueAxis={chartAxis}
        categoryAxis={{ style: chartTextStyle }}
        dataLabels={{
          value: true,
          numberFormat: '0.0"万元"',
          style: chartTextStyle,
        }}
      />
      <Chart
        testId="native-bar-chart"
        kind="bar"
        x={x[1] + 10}
        y={244}
        width={344}
        height={340}
        data={monthlyValues}
        legend={false}
        style={chartTextStyle}
        valueAxis={chartAxis}
        categoryAxis={{ style: chartTextStyle }}
      />
      <Chart
        testId="native-line-chart"
        kind="line"
        x={x[2] + 10}
        y={244}
        width={344}
        height={340}
        data={{
          categories: ["一季度", "二季度", "三季度", "四季度"],
          series: [
            { name: "本年", values: [8.2, 11.4, 10.7, 13.1], color: "accent1" },
            { name: "目标", values: [9, 10, 11, 12], color: "accent2" },
          ],
        }}
        legend="bottom"
        style={chartTextStyle}
        valueAxis={chartAxis}
        categoryAxis={{ style: chartTextStyle }}
      />
      <Footer slideId="charts-cartesian" />
    </Slide>
  );
}

function CompositionChartsPage() {
  const x = [58, 458, 858] as const;
  return (
    <Slide
      slideId="charts-composition"
      background="paper"
      transition={transitions.morph({
        by: "word",
        duration: 820,
        fallback: transitions.push({ from: "bottom" }),
      })}
    >
      <PageHeader
        section="04 / Native charts"
        title="选择适合问题的图形"
        subtitle="面积强调趋势，饼图比较占比，圆环图留出中心信息空间。"
      />
      <ChartCard testId="area-chart" title="04 / 面积 / Area" x={x[0]} />
      <ChartCard testId="pie-chart" title="05 / 饼图 / Pie" x={x[1]} />
      <ChartCard
        testId="doughnut-chart"
        title="06 / 圆环 / Doughnut"
        x={x[2]}
      />
      <Chart
        testId="native-area-chart"
        kind="area"
        x={x[0] + 10}
        y={244}
        width={344}
        height={340}
        data={{
          categories: ["一月", "二月", "三月", "四月"],
          series: [
            {
              name: "服务收入",
              values: [5.2, 6.4, 7.1, 9.3],
              color: "accent1",
            },
            {
              name: "产品收入",
              values: [3.0, 4.1, 3.6, 4.8],
              color: "accent3",
            },
          ],
        }}
        legend="bottom"
        style={chartTextStyle}
        valueAxis={chartAxis}
        categoryAxis={{ style: chartTextStyle }}
      />
      <Chart
        testId="native-pie-chart"
        kind="pie"
        x={x[1] + 10}
        y={244}
        width={344}
        height={340}
        data={{
          categories: ["研发", "设计", "运营", "服务"],
          series: [{ name: "团队投入", values: [38, 24, 21, 17] }],
        }}
        legend="bottom"
        style={chartTextStyle}
        dataLabels={{
          percent: true,
          category: true,
          position: "bestFit",
          style: chartTextStyle,
        }}
      />
      <Chart
        testId="native-doughnut-chart"
        kind="doughnut"
        holeSize={62}
        firstSliceAngle={285}
        x={x[2] + 10}
        y={244}
        width={344}
        height={340}
        data={{
          categories: ["已完成", "处理中", "待开始"],
          series: [{ name: "迭代进度", values: [64, 23, 13] }],
        }}
        legend="bottom"
        style={chartTextStyle}
        dataLabels={{ percent: true, style: chartTextStyle }}
      />
      <Footer slideId="charts-composition" />
    </Slide>
  );
}

function TypographyPage() {
  return (
    <Slide
      slideId="typography"
      background="paper"
      transition={transitions.fade({ duration: 520 })}
    >
      <PageHeader
        section="05 / Text & typography"
        title="文字保留段落、列表与样式"
        subtitle="中文、English、项目符号、自动编号和真实的大写转换都来自同一份源码。"
      />
      <Text
        testId="native-list-text"
        x={72}
        y={192}
        width={532}
        height={338}
        style={{ color: "ink", fontFamily: xiaowei, fontSize: 20 }}
      >
        <P
          style={{ list: { type: "bullet", character: "•" }, spacingAfter: 15 }}
        >
          先讲清楚问题，再选择表达方式
        </P>
        <P
          style={{ list: { type: "bullet", character: "•" }, spacingAfter: 15 }}
        >
          文字与图表共享一致的层级
        </P>
        <P style={{ list: { type: "number", startAt: 1 }, spacingAfter: 15 }}>
          将结论放在读者容易找到的位置
        </P>
        <P style={{ list: { type: "number", startAt: 2 } }}>
          给后续修改留下充足空间
        </P>
      </Text>
      <Shape
        testId="typography-sample-card"
        geometry="roundRect"
        x={666}
        y={190}
        width={540}
        height={350}
        style={{ backgroundColor: "#E8EFEB", borderRadius: 22 }}
      />
      <Text
        x={704}
        y={232}
        width={468}
        height={42}
        style={{ color: "accent1", fontFamily: jost, fontSize: 14 }}
      >
        <Span style={{ fontWeight: "bold" }}>DESIGN SYSTEM</Span>
        <Span style={{ fontFamily: xiaowei }}> / 中文排版</Span>
      </Text>
      <Text
        testId="transformed-uppercase"
        x={704}
        y={296}
        width={468}
        height={62}
        style={{
          color: "ink",
          fontFamily: jost,
          fontSize: 31,
          fontWeight: "bold",
          textTransform: "uppercase",
        }}
      >
        native typography
      </Text>
      <Text
        x={706}
        y={384}
        width={460}
        height={90}
        style={{ color: "muted", fontFamily: xiaowei, fontSize: 18 }}
      >
        字族、字重和字色按文字片段设置。源内容保留小写，输出文字体现大写转换。
      </Text>
      <Footer slideId="typography" />
    </Slide>
  );
}

function LayoutThemePage() {
  return (
    <Slide
      slideId="layout-theme"
      background="paper"
      transition={transitions.push({ from: "bottom", duration: 650 })}
    >
      <PageHeader
        section="06 / Layout & theme"
        title="用布局组织页面，用局部主题表达语境"
        subtitle="Box 负责对齐和间距；容器范围内可以覆盖已定义的颜色令牌。"
      />
      <Box
        testId="capability-layout-row"
        style={{
          position: "absolute",
          left: 64,
          top: 196,
          width: 1152,
          height: 366,
          display: "flex",
          flexDirection: "row",
          gap: 24,
        }}
      >
        <Box
          testId="layout-card-one"
          style={{
            flex: 1,
            padding: 22,
            backgroundColor: "surface",
            borderColor: "border",
            borderWidth: 1,
            borderRadius: 18,
            display: "flex",
            flexDirection: "column",
            gap: 14,
          }}
        >
          <Shape
            geometry="roundRect"
            width={88}
            height={10}
            style={{ backgroundColor: "accent1", borderRadius: 5 }}
          />
          <Text
            width={300}
            height={34}
            style={{ color: "ink", fontFamily: xiaowei, fontSize: 19 }}
          >
            01 / 网格对齐
          </Text>
          <Text
            width={300}
            height={110}
            style={{ color: "muted", fontFamily: xiaowei, fontSize: 15 }}
          >
            三列内容共享同一容器。间距、内边距和宽度由 Box 布局计算。
          </Text>
        </Box>
        <Box
          testId="layout-card-two"
          theme={{ colors: { focus: "#C45D3F" } }}
          style={{
            flex: 1,
            padding: 22,
            backgroundColor: "surface",
            borderColor: "border",
            borderWidth: 1,
            borderRadius: 18,
            display: "flex",
            flexDirection: "column",
            gap: 14,
          }}
        >
          <Shape
            testId="local-theme-mark"
            geometry="roundRect"
            width={88}
            height={10}
            style={{ backgroundColor: "focus", borderRadius: 5 }}
          />
          <Text
            width={300}
            height={34}
            style={{ color: "focus", fontFamily: xiaowei, fontSize: 19 }}
          >
            02 / 局部强调
          </Text>
          <Text
            width={300}
            height={110}
            style={{ color: "muted", fontFamily: xiaowei, fontSize: 15 }}
          >
            只在本卡片中替换强调色；其他页面继续沿用文稿主题。
          </Text>
        </Box>
        <Box
          testId="layout-card-three"
          style={{
            flex: 1,
            padding: 22,
            backgroundColor: "surface",
            borderColor: "border",
            borderWidth: 1,
            borderRadius: 18,
            display: "flex",
            flexDirection: "column",
            gap: 14,
          }}
        >
          <Shape
            geometry="roundRect"
            width={88}
            height={10}
            style={{ backgroundColor: "accent3", borderRadius: 5 }}
          />
          <Text
            width={300}
            height={34}
            style={{ color: "ink", fontFamily: xiaowei, fontSize: 19 }}
          >
            03 / 内容分组
          </Text>
          <Text
            width={300}
            height={110}
            style={{ color: "muted", fontFamily: xiaowei, fontSize: 15 }}
          >
            灵活组合文稿主题令牌与组件局部样式，保持全稿的视觉秩序。
          </Text>
        </Box>
      </Box>
      <Footer slideId="layout-theme" />
    </Slide>
  );
}

function TransitionPage() {
  const samples = [
    { title: "FADE", description: "柔和切换", color: "accent1" },
    { title: "PUSH", description: "方向推进", color: "accent2" },
    { title: "FLIP", description: "扩展翻页", color: "accent4" },
  ] as const;
  return (
    <Slide
      slideId="transitions"
      background="paper"
      transition={transitions.flip({
        direction: "left",
        duration: 690,
        fallback: transitions.strips({ direction: "leftUp" }),
      })}
    >
      <PageHeader
        section="07 / Slide transitions"
        title="基础与扩展转场都可声明"
        subtitle="每页选择转场参数；扩展转场同时提供明确的基础格式回退。"
      />
      {samples.map((sample, index) => (
        <Box
          key={sample.title}
          testId={`transition-sample-${index + 1}`}
          style={{
            position: "absolute",
            left: 74 + index * 390,
            top: 220,
            width: 352,
            height: 274,
            padding: 28,
            backgroundColor: "surface",
            borderColor: "border",
            borderWidth: 1,
            borderRadius: 22,
            display: "flex",
            flexDirection: "column",
            justifyContent: "space-between",
          }}
        >
          <Shape
            geometry="roundRect"
            width={80}
            height={8}
            style={{ backgroundColor: sample.color, borderRadius: 4 }}
          />
          <Text
            width={288}
            height={52}
            style={{ color: sample.color, fontFamily: jost, fontSize: 31 }}
          >
            {sample.title}
          </Text>
          <Text
            width={288}
            height={46}
            style={{ color: "muted", fontFamily: xiaowei, fontSize: 18 }}
          >
            {sample.description}
          </Text>
        </Box>
      ))}
      <Text
        x={76}
        y={534}
        width={1110}
        height={44}
        style={{ color: "ink", fontFamily: xiaowei, fontSize: 17 }}
      >
        淡入、推进与翻转，分别服务于内容的不同节奏。
      </Text>
      <Footer slideId="transitions" />
    </Slide>
  );
}

const morphScope = createMorphScope();
const morphTitle = morphScope.handle("title");
const morphPanel = morphScope.handle("panel");
const morphMarker = morphScope.handle("marker");

function MorphStartPage() {
  return (
    <Slide
      slideId="morph-start"
      background="paper"
      transition={transitions.fade({ duration: 420 })}
    >
      <PageHeader
        section="08 / Object continuity"
        title="显式标记跨页连续对象"
        subtitle="两个页面共享同一组稳定句柄；下一页改变位置、尺寸与源码顺序。"
      />
      <Shape
        testId="morph-start-title"
        morphKey={morphTitle}
        geometry="roundRect"
        x={88}
        y={198}
        width={466}
        height={142}
        style={{ backgroundColor: "accent1", borderRadius: 24 }}
        textStyle={{
          color: "surface",
          fontFamily: xiaowei,
          fontSize: 27,
          align: "center",
          verticalAlign: "middle",
        }}
      >
        从一个问题开始
      </Shape>
      <Shape
        testId="morph-start-panel"
        morphKey={morphPanel}
        geometry="roundRect"
        x={642}
        y={198}
        width={514}
        height={270}
        style={{ backgroundColor: "#E4ECE8", borderRadius: 24 }}
      />
      <Shape
        testId="morph-start-marker"
        morphKey={morphMarker}
        geometry="ellipse"
        x={842}
        y={270}
        width={104}
        height={104}
        style={{ backgroundColor: "accent3" }}
      />
      <Text
        x={672}
        y={505}
        width={452}
        height={50}
        style={{ color: "muted", fontFamily: xiaowei, fontSize: 17 }}
      >
        继续前进，查看相同对象如何重新排列。
      </Text>
      <Footer slideId="morph-start" />
    </Slide>
  );
}

function MorphEndPage() {
  return (
    <Slide
      slideId="morph-end"
      background="paper"
      transition={transitions.morph({
        by: "object",
        duration: 900,
        fallback: transitions.fade({ duration: 520 }),
      })}
    >
      <PageHeader
        section="09 / Object continuity"
        title="身份跟着句柄，不跟着排列顺序"
        subtitle="源码顺序已调整，匹配仍由作者给出的 Morph 句柄决定。"
      />
      <Shape
        testId="morph-end-marker"
        morphKey={morphMarker}
        geometry="ellipse"
        x={1012}
        y={220}
        width={142}
        height={142}
        style={{ backgroundColor: "accent2" }}
      />
      <Shape
        testId="morph-end-panel"
        morphKey={morphPanel}
        geometry="roundRect"
        x={124}
        y={206}
        width={630}
        height={282}
        style={{ backgroundColor: "#E4ECE8", borderRadius: 24 }}
      />
      <Shape
        testId="morph-end-title"
        morphKey={morphTitle}
        geometry="roundRect"
        x={228}
        y={274}
        width={470}
        height={142}
        style={{ backgroundColor: "accent1", borderRadius: 24 }}
        textStyle={{
          color: "surface",
          fontFamily: xiaowei,
          fontSize: 27,
          align: "center",
          verticalAlign: "middle",
        }}
      >
        再次查看同一对象
      </Shape>
      <Text
        x={804}
        y={426}
        width={358}
        height={52}
        style={{ color: "muted", fontFamily: xiaowei, fontSize: 17 }}
      >
        三个对象均使用跨页共享的连续性标识。
      </Text>
      <Footer slideId="morph-end" />
    </Slide>
  );
}

function MotionCompositionPage() {
  const finiteBody = useMotionScope();
  return (
    <Slide
      slideId="motion-composition"
      background="paper"
      transition={transitions.push({ from: "right", duration: 640 })}
    >
      <PageHeader
        section="10 / Motion composition"
        title="有限正文与持续呼吸各自编排"
        subtitle="正文分两次点击进入；组合中的形状成员独立持续呼吸。"
      />
      <Shape
        testId="motion-panel"
        geometry="roundRect"
        x={72}
        y={194}
        width={652}
        height={360}
        style={{ backgroundColor: "surface", borderRadius: 22 }}
      />
      <MotionScope scope={finiteBody} mode="parallel">
        <Step n={1}>
          <Text
            testId="finite-body-step-one"
            enter={fade({ duration: 360 })}
            x={108}
            y={236}
            width={566}
            height={64}
            style={{ color: "ink", fontFamily: xiaowei, fontSize: 23 }}
          >
            第一次点击：先呈现行动方向。
          </Text>
        </Step>
        <Step n={2}>
          <Text
            testId="finite-body-step-two"
            enter={fade({ duration: 360 })}
            x={108}
            y={344}
            width={566}
            height={64}
            style={{ color: "accent1", fontFamily: xiaowei, fontSize: 23 }}
          >
            第二次点击：补充下一步与结果。
          </Text>
        </Step>
      </MotionScope>
      <Group testId="breathing-native-group">
        <Shape
          testId="breathing-member-one"
          geometry="roundRect"
          x={816}
          y={236}
          width={164}
          height={136}
          emphasis={repeat(
            breathe({ period: 1800, from: "#147C80", to: "#E6B95B" }),
            "forever",
          )}
          style={{ backgroundColor: "accent1", borderRadius: 24 }}
        />
        <Shape
          testId="breathing-member-two"
          geometry="ellipse"
          x={1010}
          y={275}
          width={146}
          height={146}
          emphasis={repeat(
            breathe({
              period: 1800,
              delay: 240,
              from: "#E37955",
              to: "#E6B95B",
            }),
            "forever",
          )}
          style={{ backgroundColor: "accent2" }}
        />
      </Group>
      <Text
        x={796}
        y={464}
        width={390}
        height={56}
        style={{ color: "muted", fontFamily: xiaowei, fontSize: 16 }}
      >
        呼吸效果属于 Group 中的两个独立形状成员。
      </Text>
      <Text
        testId="motion-page-link"
        link={{ slideId: "unknown-completion" }}
        x={72}
        y={584}
        width={326}
        height={38}
        style={{ color: "accent1", fontFamily: xiaowei, fontSize: 17 }}
      >
        查看逐词与逐段编排
      </Text>
      <Footer slideId="motion-composition" />
    </Slide>
  );
}

function UnknownCompletionPage() {
  const allReveals = useMotionScope();
  const afterAllReveals = useMotionScope();
  return (
    <Slide
      slideId="unknown-completion"
      background="paper"
      transition={transitions.fade({ duration: 560 })}
    >
      <PageHeader
        section="11 / Fine-grained motion"
        title="细粒度文字以原生单元进入"
        subtitle="逐词与逐段采用原生编排；总完成时刻未知时，整体作用域等待全部成员。"
      />
      <Shape
        testId="unknown-motion-card"
        geometry="roundRect"
        x={76}
        y={194}
        width={1128}
        height={350}
        style={{ backgroundColor: "surface", borderRadius: 22 }}
      />
      <MotionScope scope={allReveals} mode="parallel">
        <Text
          testId="word-by-word-reveal"
          enter={fade({ by: "word", duration: 320, intervalMs: 120 })}
          x={116}
          y={238}
          width={1034}
          height={56}
          style={{ color: "ink", fontFamily: xiaowei, fontSize: 22 }}
        >
          每个词在自己的节奏中出现，整段总时长由词数决定。
        </Text>
        <Text
          testId="paragraph-by-paragraph-reveal"
          enter={fade({ by: "paragraph", duration: 340 })}
          x={116}
          y={324}
          width={1034}
          height={142}
          style={{ color: "accent1", fontFamily: xiaowei, fontSize: 20 }}
        >
          <P style={{ spacingAfter: 14 }}>第一段先交代目标与上下文。</P>
          <P>第二段补充条件、行动和需要关注的结果。</P>
        </Text>
      </MotionScope>
      <MotionScope scope={afterAllReveals} startAt={after(allReveals.end, 0)}>
        <Text
          testId="wait-for-all-completion"
          enter={fade({ duration: 340 })}
          x={116}
          y={484}
          width={1034}
          height={40}
          style={{ color: "muted", fontFamily: xiaowei, fontSize: 17 }}
        >
          下一项从全部文字完成后开始。
        </Text>
      </MotionScope>
      <Footer slideId="unknown-completion" />
    </Slide>
  );
}

/** 公开 API 能力册；同一组件同时用于默认、完整与子集字体构建。 */
export function CapabilityDeck() {
  return (
    <Deck
      title="灵感实验室 · 原创能力册"
      width={1280}
      height={720}
      theme={capabilityTheme}
      style={{ color: "ink", fontFamily: "body", fontSize: 16 }}
    >
      <CapabilityCover />
      <NativeObjectsPage />
      <TablePage />
      <CartesianChartsPage />
      <CompositionChartsPage />
      <TypographyPage />
      <LayoutThemePage />
      <TransitionPage />
      <MorphStartPage />
      <MorphEndPage />
      <MotionCompositionPage />
      <UnknownCompletionPage />
    </Deck>
  );
}

export const capabilityDeck = <CapabilityDeck />;
