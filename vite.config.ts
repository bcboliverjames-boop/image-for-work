import { defineConfig } from "vite";

// GitHub Pages 部署时的基础路径
// 格式: /仓库名/
const base = process.env.GITHUB_ACTIONS ? "/image-for-work/" : "/";

export default defineConfig({
  base,
  server: {
    host: "::",
    port: 3002,
    https: true,
    headers: {
      // 禁用缓存，确保每次都加载最新代码
      "Cache-Control": "no-cache, no-store, must-revalidate",
      "Pragma": "no-cache",
      "Expires": "0"
    }
  },
  build: {
    rollupOptions: {
      input: {
        taskpane: "taskpane.html",
        commands: "commands.html"
      }
    }
  },
  test: {
    environment: "jsdom",
    globals: true,
  }
});
