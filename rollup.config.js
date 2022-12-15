// Lot of examples taken from https://github.com/atti187/esmodules
import { babel } from "@rollup/plugin-babel";
import { nodeResolve } from "@rollup/plugin-node-resolve";
import cleanup from "rollup-plugin-cleanup";

const extensions = [".ts", ".js"];

function preventTreeShakingPlugin() {
  return {
    name: "no-treeshaking",
    resolveId(id, importer) {
      if (!importer) {
        // No need for tree-shaking; entry point files shouldn't have anything
        // exported when working with Apps Script
        return { id, moduleSideEffects: "no-treeshake" };
      }
      return null;
    },
  };
}

export default {
  input: "./src/index.ts",
  output: {
    dir: "./dist",
    format: "cjs",
  },
  plugins: [
    cleanup({ extensions: [".js", ".ts"] }),
    preventTreeShakingPlugin(),
    nodeResolve({ extensions }),
    babel({ extensions, babelHelpers: "runtime" }),
  ],
};
