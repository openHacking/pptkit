import path from "node:path";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";

export default defineConfig(({ mode }) => {
  if (mode === "pages") {
    return {
      root: path.resolve(__dirname, "playground"),
      base: "./",
      plugins: [react()],
      build: {
        outDir: path.resolve(__dirname, "dist/playground"),
        emptyOutDir: true,
      },
    };
  }

  return {
    plugins: [react(), tailwindcss()],
    resolve: {
      alias: {
        "@": path.resolve(__dirname, "src/client"),
      },
    },
    server: {
      port: 5173,
      proxy: {
        "/api": {
          target: "http://127.0.0.1:3210",
          changeOrigin: true,
        },
      },
    },
    build: {
      outDir: "dist/client",
    },
  };
});
