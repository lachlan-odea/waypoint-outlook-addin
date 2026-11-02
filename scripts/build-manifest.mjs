// Reads manifest.xml, replaces {{BASE_URL}} with the given base, writes manifest.<mode>.xml.
// Usage:
//   node scripts/build-manifest.mjs dev
//   node scripts/build-manifest.mjs prod
//
// Configure prod base in .env.manifest (or env var BASE_URL).

import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = resolve(__dirname, "..");

const mode = process.argv[2] ?? "dev";

let baseUrl = process.env.BASE_URL;
if (!baseUrl) {
  if (mode === "dev") {
    baseUrl = "https://localhost:3000/waypoint-outlook-addin";
  } else {
    const envPath = resolve(root, ".env.manifest");
    if (existsSync(envPath)) {
      const env = readFileSync(envPath, "utf8");
      const match = env.match(/^BASE_URL=(.*)$/m);
      if (match) baseUrl = match[1].trim();
    }
  }
}
if (!baseUrl) {
  console.error(
    "BASE_URL not set. Provide via env var (BASE_URL=...) or .env.manifest file."
  );
  process.exit(1);
}
// Strip trailing slash for consistent join behaviour.
baseUrl = baseUrl.replace(/\/+$/, "");

const src = readFileSync(resolve(root, "manifest.xml"), "utf8");
const out = src.replaceAll("{{BASE_URL}}", baseUrl);
const outPath = resolve(root, `manifest.${mode}.xml`);
writeFileSync(outPath, out);
console.log(`Wrote ${outPath} (BASE_URL=${baseUrl})`);
