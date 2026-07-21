/*
 * Builds a store-ready package in dist/:
 *   node package-for-store.js --url https://your-domain.example --email you@example.com [--provider "Your Name"]
 *
 * - dist/web/       -> static files to upload to your HTTPS host (host them at the given URL)
 * - dist/manifest.xml -> production manifest to upload to Partner Center
 */
const fs = require("fs");
const path = require("path");

function arg(name, fallback) {
  const i = process.argv.indexOf("--" + name);
  return i >= 0 && process.argv[i + 1] ? process.argv[i + 1] : fallback;
}

const baseUrl = (arg("url", "") || "").replace(/\/+$/, "");
const email = arg("email", "");
const provider = arg("provider", "JGold");

if (!baseUrl || !/^https:\/\//.test(baseUrl)) {
  console.error("ERROR: pass your production HTTPS hosting URL, e.g.:");
  console.error('  node package-for-store.js --url https://your-domain.example --email you@example.com');
  process.exit(1);
}
if (!email) {
  console.error("ERROR: pass a support email with --email you@example.com");
  process.exit(1);
}

const ROOT = __dirname;
const DIST = path.join(ROOT, "dist");
const WEB = path.join(DIST, "web");

fs.rmSync(DIST, { recursive: true, force: true });
fs.mkdirSync(path.join(WEB, "src", "taskpane"), { recursive: true });
fs.mkdirSync(path.join(WEB, "assets"), { recursive: true });

// --- Copy web assets, substituting the support email placeholder. ---
const copies = [
  ["src/taskpane/taskpane.html", "src/taskpane/taskpane.html"],
  ["src/taskpane/taskpane.css", "src/taskpane/taskpane.css"],
  ["src/taskpane/taskpane.js", "src/taskpane/taskpane.js"],
  ["src/support.html", "support.html"],
  ["src/privacy.html", "privacy.html"],
  ["assets/icon-16.png", "assets/icon-16.png"],
  ["assets/icon-32.png", "assets/icon-32.png"],
  ["assets/icon-80.png", "assets/icon-80.png"]
];

for (const [from, to] of copies) {
  const src = path.join(ROOT, from);
  const dest = path.join(WEB, to);
  if (/\.(html|css|js)$/.test(from)) {
    let text = fs.readFileSync(src, "utf8");
    text = text.split("SUPPORT_EMAIL_PLACEHOLDER").join(email);
    fs.writeFileSync(dest, text);
  } else {
    fs.copyFileSync(src, dest);
  }
}

// --- Production manifest: swap localhost for the production URL. ---
let manifest = fs.readFileSync(path.join(ROOT, "manifest.xml"), "utf8");
manifest = manifest.split("https://localhost:3000").join(baseUrl);
manifest = manifest.replace(
  /<SupportUrl DefaultValue="[^"]*"\/>/,
  `<SupportUrl DefaultValue="${baseUrl}/support.html"/>`
);
manifest = manifest.replace(
  /<ProviderName>[^<]*<\/ProviderName>/,
  `<ProviderName>${provider}</ProviderName>`
);
// The dev manifest serves the pane from /src/taskpane/; keep that path in dist/web too.
fs.writeFileSync(path.join(DIST, "manifest.xml"), manifest);

console.log("Store package written to dist/");
console.log("  1. Upload the contents of dist/web/ to " + baseUrl + " (paths must match exactly).");
console.log("  2. Verify " + baseUrl + "/src/taskpane/taskpane.html loads in a browser.");
console.log("  3. Validate:  npx office-addin-manifest validate dist/manifest.xml -p");
console.log("  4. Upload dist/manifest.xml in Partner Center.");
