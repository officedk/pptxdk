import lintConfiguration from "@react-pptx/tooling-config/oxlint";
import { defineConfig } from "oxlint";

export default defineConfig({
  options: { typeAware: true },
  extends: [lintConfiguration],
});
