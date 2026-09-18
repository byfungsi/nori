import { defineConfig } from "vitepress";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";
const base = process.env.NORI_DOCS_BASE ?? "/";
const repository = fileURLToPath(new URL("../../", import.meta.url));
export default defineConfig({
  title: "Nori",
  description:
    "A platform-independent spreadsheet runtime. One public package, composable layers.",
  base,
  cleanUrls: base === "/",
  vite: {
    build: { target: "es2022" },
    plugins: [
      {
        name: "nori-agent-docs",
        handleHotUpdate({ file }) {
          const relative = file.slice(repository.length);
          if (
            /^docs\/[^/]+\.md$/.test(relative) ||
            relative.startsWith("docs/examples/")
          ) {
            execFileSync(process.execPath, ["scripts/generate-docs.mjs"], {
              cwd: repository,
            });
          }
        },
      },
    ],
  },
  srcExclude: ["public/**", "examples/**"],
  head: [["meta", { name: "theme-color", content: "#244b39" }]],
  themeConfig: {
    siteTitle: "Nori / docs",
    socialLinks: [{ icon: "github", link: "https://github.com/byfungsi/nori" }],
    search: { provider: "local" },
    nav: [
      { text: "Playground", link: "/playground" },
      { text: "Guide", link: "/getting-started" },
      { text: "API", link: "/api" },
      { text: "For agents", link: "/agents" },
    ],
    sidebar: [
      {
        text: "Start here",
        items: [
          { text: "Introduction", link: "/" },
          { text: "Live playground", link: "/playground" },
          { text: "Quick start", link: "/getting-started" },
          { text: "React integration", link: "/react" },
          { text: "XLSX import", link: "/importing" },
        ],
      },
      {
        text: "Use the library",
        items: [
          { text: "Headless recipes", link: "/recipes" },
          { text: "Editor interactions", link: "/interactions" },
          { text: "Chat preview", link: "/preview" },
          { text: "API reference", link: "/api" },
        ],
      },
      {
        text: "Understand the boundaries",
        items: [
          { text: "Architecture", link: "/architecture" },
          { text: "Excel support", link: "/support" },
          { text: "Troubleshooting", link: "/troubleshooting" },
          { text: "Roadmap", link: "/roadmap" },
        ],
      },
      {
        text: "Build with agents",
        items: [
          { text: "Agent integration guide", link: "/agents" },
          { text: "Contributing & verification", link: "/contributing" },
          { text: "Verification record", link: "/verification" },
        ],
      },
    ],
    outline: [2, 3],
    footer: {
      message:
        "Nori 0.1.0 · Initial milestone · Explicit compatibility boundaries.",
    },
  },
});
