import { defineConfig } from "vite";
import { resolve } from "path";

export default defineConfig(() => {
  // For GitHub Pages, set BASE_PATH to '/<repo>/' in the GitHub Action.
  const base = process.env.BASE_PATH ?? "/";

  return {
    base,
    build: {
      rollupOptions: {
        input: {
          taskpane: resolve(__dirname, "taskpane.html"),
          commands: resolve(__dirname, "commands.html")
        }
      }
    }
  };
});
