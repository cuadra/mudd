import { cp, mkdir, rm } from "node:fs/promises";
import { resolve } from "node:path";

const root = process.cwd();
const src = resolve(root, "src");
const dist = resolve(root, "dist");

await rm(dist, { recursive: true, force: true });
await mkdir(resolve(dist, "lib"), { recursive: true });

await cp(resolve(src, "background.js"), resolve(dist, "background.js"));
await cp(resolve(src, "content.js"), resolve(dist, "content.js"));
await cp(resolve(src, "doc-builder.js"), resolve(dist, "doc-builder.js"));
await cp(resolve(src, "lib", "docx-lite.js"), resolve(dist, "lib", "docx-lite.js"));
