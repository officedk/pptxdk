import { chartMechanismVerifications } from "./chart/mechanisms.ts";
import { groupMechanismVerifications } from "./group/mechanisms.ts";
import { imageMechanismVerifications } from "./image/mechanisms.ts";
import { tableMechanismVerifications } from "./table/mechanisms.ts";
import {
  textBoxRoundtripScenario,
  textMechanismVerifications,
} from "./text/mechanisms.ts";

export { textBoxRoundtripScenario };

/**
 * Aggregated mechanism verifications for canvas/content capabilities
 * (Group, Text, Image, Table, Chart).
 */
export const contentMechanismVerifications = {
  ...groupMechanismVerifications,
  ...textMechanismVerifications,
  ...imageMechanismVerifications,
  ...tableMechanismVerifications,
  ...chartMechanismVerifications,
};
