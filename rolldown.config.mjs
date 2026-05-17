import { defineConfig } from "rolldown";

export default defineConfig([
  {
    input: "src/index.ts",
    external: ["exceljs", "fs", "path"],
    output: {
      format: "esm",
      dir: "dist/esm",
      entryFileNames: "[name].js",
      sourcemap: true,
    },
  },
  {
    input: "src/index.ts",
    external: ["exceljs", "fs", "path"],
    output: {
      format: "cjs",
      dir: "dist/cjs",
      entryFileNames: "[name].js",
      sourcemap: true,
    },
  },
]);
