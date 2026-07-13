import { cp, mkdir, readFile, readdir, rm, writeFile } from "node:fs/promises";
import path from "node:path";

const root = process.cwd();
const versionDir = path.join(root, "public", "stock_dashboard_versions");
const versionPages = (await readdir(versionDir))
  .filter((file) => /^stock_holdings_percent_dashboard_.*\.html$/.test(file))
  .sort()
  .map((file) => `/stock_dashboard_versions/${file}`);
const pages = ["/index.html", "/stock_holdings_percent_dashboard.html", ...versionPages];

const entries = [];
for (const page of pages) {
  const filePath = path.join(root, "public", page);
  entries.push([page, await readFile(filePath, "utf8")]);
}

const source = await readFile(path.join(root, "worker", "index.js"), "utf8");
const bundled = source.replace(
  "const PAGES = new Map();",
  `const PAGES = new Map(${JSON.stringify(entries)});`,
);

await rm(path.join(root, "dist"), { recursive: true, force: true });
await mkdir(path.join(root, "dist", "server"), { recursive: true });
await mkdir(path.join(root, "dist", "_appgen_meta"), { recursive: true });
await writeFile(path.join(root, "dist", "server", "index.js"), bundled);
await cp(
  path.join(root, "public"),
  path.join(root, "dist", "server", "public"),
  { recursive: true },
);
await cp(
  path.join(root, ".openai", "hosting.json"),
  path.join(root, "dist", "_appgen_meta", "appgarden.json"),
);
