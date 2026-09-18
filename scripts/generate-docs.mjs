import {
  readFileSync,
  writeFileSync,
  mkdirSync,
  readdirSync,
  cpSync,
  rmSync,
} from "node:fs";
import { resolve, basename } from "node:path";
import { createHash } from "node:crypto";
import { docsBase, sitePath } from "./docs-site.mjs";
const root = resolve("docs"),
  output = resolve(root, "public");
mkdirSync(output, { recursive: true });
for (const directory of ["markdown", "types"]) {
  rmSync(resolve(output, directory), { recursive: true, force: true });
  mkdirSync(resolve(output, directory), { recursive: true });
}
const version = JSON.parse(
  readFileSync("packages/nori/package.json", "utf8"),
).version;
const pages = readdirSync(root)
  .filter((name) => name.endsWith(".md"))
  .sort()
  .map((name) => {
    const source = readFileSync(resolve(root, name), "utf8");
    const title = source.match(/^# (.+)$/m)?.[1] ?? "Nori documentation";
    const expanded = source
      .replace(/^---\n[\s\S]*?\n---\n/, "# Nori documentation\n")
      .replace(/^<<< (.+)$/gm, (_match, relative) => {
        const path = resolve(root, relative.trim());
        if (!path.startsWith(root + "/examples/"))
          throw Error("Unexpected documentation include");
        return (
          "```" +
          (path.endsWith(".tsx") ? "tsx" : "ts") +
          "\n" +
          readFileSync(path, "utf8").trimEnd() +
          "\n```"
        );
      })
      .replace(/\]\(\/(?!\/)([^)]*)\)/g, (_match, path) => {
        const [page, anchor] = path.split("#");
        const route =
          docsBase !== "/" && page && !page.includes(".") && !page.includes("/")
            ? page + ".html"
            : page;
        return `](${sitePath(route)}${anchor === undefined ? "" : "#" + anchor})`;
      });
    writeFileSync(resolve(output, "markdown", name), expanded);
    return {
      title,
      route: sitePath(
        name === "index.md"
          ? ""
          : basename(name, ".md") + (docsBase === "/" ? "" : ".html"),
      ),
      markdown: sitePath("markdown/" + name),
      source: "docs/" + name,
      sha256: createHash("sha256").update(expanded).digest("hex"),
      content: expanded,
    };
  });
for (const name of readdirSync("packages/nori/dist").filter((name) =>
  name.endsWith(".d.ts"),
))
  cpSync(resolve("packages/nori/dist", name), resolve(output, "types", name));
writeFileSync(
  resolve(output, "docs-manifest.json"),
  JSON.stringify(
    {
      package: "@byfungsi/nori",
      version,
      base: docsBase,
      pages: pages.map(({ content, ...page }) => page),
    },
    null,
    2,
  ) + "\n",
);
writeFileSync(
  resolve(output, "llms.txt"),
  `# Nori ${version}\n\n> Platform-independent spreadsheet/XLSX library with a React DOM adapter. Initial milestone; not full Excel parity. Public package not yet published by this repository.\n\nRead ${sitePath("markdown/agents.md")} first, then ${sitePath("markdown/getting-started.md")}, ${sitePath("markdown/support.md")} and the relevant API guide. Resolve these root-relative paths against this documentation host.\n\nFull context: ${sitePath("llms-full.txt")}\nMachine manifest: ${sitePath("docs-manifest.json")}\nPublic declarations: ${sitePath("types/index.d.ts")} (relative declaration dependencies are served alongside it).\n\n## Documentation\n\n` +
    pages.map((page) => `- [${page.title}](${page.markdown})`).join("\n") +
    "\n",
);
writeFileSync(
  resolve(output, "llms-full.txt"),
  `# Nori ${version}: complete documentation\n\nGenerated from repository Markdown. Refer to ${sitePath("docs-manifest.json")} for per-page sources and hashes.\n\n` +
    pages
      .map(
        (page) =>
          `---\nSource: ${page.source}\nPage: ${page.route}\nMarkdown: ${page.markdown}\n\n${page.content}`,
      )
      .join("\n\n"),
);
console.log(
  `Generated ${pages.length} Markdown pages, agent index, full context and public declarations.`,
);
