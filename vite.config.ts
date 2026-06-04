import { defineConfig, type ServerOptions } from "vite";
import { resolve } from "node:path";
import { existsSync, readFileSync } from "node:fs";
import { homedir } from "node:os";

// GitHub Pages serves the site under /<repo-name>/, so set base for that.
// Override with VITE_BASE=/some-path/ at build time if you host it elsewhere.
const base = process.env.VITE_BASE ?? "/waypoint-outlook-addin/";

// Read certs installed by `office-addin-dev-certs install`. The util writes them
// to <homedir>/.office-addin-dev-certs/ by default. If they're not present we
// fall back to no HTTPS so the user gets a clear error to run `npm run certs:install`.
function devHttps(): ServerOptions["https"] | undefined {
  const certDir = resolve(homedir(), ".office-addin-dev-certs");
  const key = resolve(certDir, "localhost.key");
  const cert = resolve(certDir, "localhost.crt");
  if (existsSync(key) && existsSync(cert)) {
    return { key: readFileSync(key), cert: readFileSync(cert) };
  }
  console.warn(
    "[waypoint] No office-addin-dev-certs found.\n" +
      "Run `npm run certs:install` once so Outlook can trust https://localhost:3000."
  );
  return undefined;
}

export default defineConfig({
  base,
  server: {
    https: devHttps(),
    host: "localhost",
    port: 3000,
    strictPort: true,
  },
  build: {
    outDir: "dist",
    sourcemap: true,
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, "index.html"),
        commands: resolve(__dirname, "commands.html"),
      },
    },
  },
});
