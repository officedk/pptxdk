import { footerPlaceholderMechanismVerifications } from "../src/masters/test/footer-mechanisms.ts";
import { masterOwnershipVerifications } from "../src/masters/test/master-layout-mechanisms.ts";
import { masterSharedCanvasMechanismVerifications } from "../src/masters/test/shared-canvas-mechanisms.ts";
import { themeMechanismVerifications } from "../src/masters/test/theme-mechanisms.ts";
import { motionMechanismVerifications } from "../src/motion/test/mechanisms.ts";
import { notesMechanismVerifications } from "../src/notes/test/mechanisms.ts";
import { placeholderMechanismVerifications } from "../src/placeholder/test/mechanisms.ts";
import { slideBackgroundMechanismVerifications } from "../src/slide-background/test/mechanisms.ts";
import { transitionMechanismVerifications } from "../src/transition/test/mechanisms.ts";

/**
 * Aggregated mechanism verifications for masters, placeholders, backgrounds,
 * notes/sections, transitions, and motion capabilities.
 */
export const presentationMechanismVerifications = {
  ...masterOwnershipVerifications,
  ...themeMechanismVerifications,
  ...masterSharedCanvasMechanismVerifications,
  ...footerPlaceholderMechanismVerifications,
  ...slideBackgroundMechanismVerifications,
  ...placeholderMechanismVerifications,
  ...notesMechanismVerifications,
  ...transitionMechanismVerifications,
  ...motionMechanismVerifications,
};
