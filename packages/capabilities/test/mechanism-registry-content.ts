import { chartMechanismVerifications } from "../src/chart/test/mechanisms.ts";
import { groupMechanismVerifications } from "../src/group/test/mechanisms.ts";
import { imageMechanismVerifications } from "../src/image/test/mechanisms.ts";
import { tableMechanismVerifications } from "../src/table/test/mechanisms.ts";
import {
  textBoxRoundtripScenario,
  textMechanismVerifications,
} from "../src/text/test/mechanisms.ts";

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
