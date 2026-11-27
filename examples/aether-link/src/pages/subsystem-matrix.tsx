import { Box, transitions } from "@officedk/pptxdk/react";
import { useMemo } from "react";

import { EditorialSlide, TakeawayStrip } from "../components/chrome.tsx";
import { createSubsystemMatrixTable } from "../components/diagrams.ts";
import { PALETTE, editorialTableLayout } from "../theme.ts";

/** Slide 8: Subsystem qualification matrix bound via native table-placeholder. */
export function SubsystemMatrixSlide() {
  const matrix = useMemo(() => createSubsystemMatrixTable(), []);

  return (
    <EditorialSlide
      id="subsystem-matrix"
      layout={editorialTableLayout}
      kicker="03 · FLIGHT CAMPAIGN & ROADMAP  |  VERIFICATION MATRIX"
      title="Subsystem Qualification & Flight Telemetry Compliance"
      transition={transitions.fade({ duration: 450 })}
      extraPlaceholders={{ matrix }}
      notes="This native PowerPoint table is bound into a table-placeholder slot and exercises vertical rowSpan grouping plus horizontal colSpan summary cells."
    >
      <Box
        style={{
          position: "absolute",
          left: 56,
          top: 604,
          width: 1168,
        }}
      >
        <TakeawayStrip
          tag="AUDIT RESULT"
          accentColor={PALETTE.cobalt}
          text="Zero open Class-A waivers; cislunar LRT-90 environmental thermal-vacuum qualification completed October 2026."
        />
      </Box>
    </EditorialSlide>
  );
}
