import { createMorphScope, Deck, Section } from "@officedk/pptxdk/react";
import { Suspense, useMemo } from "react";

import { fetchGitHubProjects } from "./github-projects.ts";
import { AuxLossFreeMtpPage } from "./pages/aux-loss-free-mtp.tsx";
import { DecodePipelinePage } from "./pages/decode-pipeline.tsx";
import { DualPipePage } from "./pages/dualpipe.tsx";
import { Fp8KernelsPage } from "./pages/fp8-kernels.tsx";
import { FrontierBenchmarksPage } from "./pages/frontier-benchmarks.tsx";
import { GitHubProjectsPage } from "./pages/github-projects.tsx";
import { MlaMoeEfficiencyPage } from "./pages/mla-moe-efficiency.tsx";
import { MpftDeepEpPage } from "./pages/mpft-deepep.tsx";
import { OpenInfraSynthesisPage } from "./pages/open-infra-synthesis.tsx";
import { OpeningPage } from "./pages/opening.tsx";
import { PdDisaggregationPage } from "./pages/pd-disaggregation.tsx";
import { ProductionMetricsPage } from "./pages/production-metrics.tsx";
import { R1RlDistillPage } from "./pages/r1-rl-distill.tsx";
import { ThreeFsPage } from "./pages/three-fs.tsx";
import { TrainingCostPage } from "./pages/training-cost.tsx";
import { colors, theme } from "./theme.ts";

/**
 * 15-slide technical architecture presentation deck for DeepSeek-V3 / R1
 * and the Open-Source Week infrastructure stack.
 */
export function DeepSeekSystemsDeck({
  fetchProject,
}: {
  readonly fetchProject?: typeof fetch;
}) {
  // Keep the request stable while the page retries inside Suspense.
  const projects = useMemo(
    () => fetchGitHubProjects(fetchProject),
    [fetchProject],
  );
  const pdMorphScope = useMemo(() => createMorphScope(), []);

  return (
    <Deck
      width={1280}
      height={720}
      title="稀疏激活与通信隐藏：DeepSeek-V3 / R1 训练与推理全栈架构剖析"
      theme={theme}
      style={{
        fontFamily: "body",
        fontSize: 14,
        color: colors.ink,
      }}
    >
      <Section name="01 · 模型架构基石：算力成本拆解、MLA 与细粒度 MoE">
        <OpeningPage />
        <TrainingCostPage />
        <MlaMoeEfficiencyPage />
        <AuxLossFreeMtpPage />
      </Section>

      <Section name="02 · 训练与基础设施：FP8、DualPipe、MPFT/DeepEP 与 3FS">
        <Fp8KernelsPage />
        <DualPipePage />
        <MpftDeepEpPage />
        <ThreeFsPage />
      </Section>

      <Section name="03 · 生产推理系统：PD 分离、5 阶段解码流水与 24 小时实测">
        <PdDisaggregationPage scope={pdMorphScope} />
        <DecodePipelinePage scope={pdMorphScope} />
        <ProductionMetricsPage />
      </Section>

      <Section name="04 · 强化学习推理涌现与全栈硬件协同总结">
        <R1RlDistillPage />
        <FrontierBenchmarksPage />
        <Suspense fallback={null}>
          <GitHubProjectsPage snapshot={projects} />
        </Suspense>
        <OpenInfraSynthesisPage />
      </Section>
    </Deck>
  );
}
