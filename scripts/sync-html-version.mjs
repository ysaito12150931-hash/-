import { readFileSync, writeFileSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const versionJs = readFileSync(join(root, "js", "version.js"), "utf8");
const match = versionJs.match(/APP_VERSION\s*=\s*"([^"]+)"/);
if (!match) {
  console.error("js/version.js に APP_VERSION が見つかりません。");
  process.exit(1);
}

const version = match[1];
const htmlPath = join(root, "index.html");
const html = readFileSync(htmlPath, "utf8");
const updated = html.replace(/\?v=[\d.]+/g, `?v=${version}`);

if (updated === html) {
  console.log(`index.html は既に v=${version} です。`);
} else {
  writeFileSync(htmlPath, updated);
  console.log(`index.html を v=${version} に更新しました。`);
}
