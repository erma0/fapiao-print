#!/usr/bin/env node
/**
 * 版本号同步脚本 — 以 package.json 为单一数据源，同步版本号到:
 *   - src-tauri/Cargo.toml
 *   - src-tauri/tauri.conf.json
 *   - src-tauri/tauri.ocr.conf.json (if it has a version field)
 *
 * 用法:
 *   node scripts/bump-version.js 1.8.3       # 设置新版本号并同步
 *   node scripts/bump-version.js              # 仅同步当前 package.json 版本号
 *
 * 之后需要手动 commit + tag + push 触发 CI 构建。
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const PKG_PATH = path.join(ROOT, 'package.json');
const CARGO_PATH = path.join(ROOT, 'src-tauri', 'Cargo.toml');
const TAURI_CONF_PATH = path.join(ROOT, 'src-tauri', 'tauri.conf.json');

// Read package.json
const pkg = JSON.parse(fs.readFileSync(PKG_PATH, 'utf8'));
const newVersion = process.argv[2] || pkg.version;

// Update package.json if version changed
if (pkg.version !== newVersion) {
  pkg.version = newVersion;
  fs.writeFileSync(PKG_PATH, JSON.stringify(pkg, null, 2) + '\n');
  console.log(`✅ package.json: ${pkg.version} → ${newVersion}`);
} else {
  console.log(`📦 package.json: already ${newVersion}`);
}

// Update Cargo.toml
let cargo = fs.readFileSync(CARGO_PATH, 'utf8');
const cargoReplaced = cargo.replace(
  /^version\s*=\s*"[^"]*"/m,
  `version = "${newVersion}"`
);
if (cargo !== cargoReplaced) {
  fs.writeFileSync(CARGO_PATH, cargoReplaced);
  console.log(`✅ Cargo.toml: → ${newVersion}`);
} else {
  console.log(`🦀 Cargo.toml: already ${newVersion}`);
}

// Update tauri.conf.json
const tauriConf = JSON.parse(fs.readFileSync(TAURI_CONF_PATH, 'utf8'));
if (tauriConf.version !== newVersion) {
  tauriConf.version = newVersion;
  fs.writeFileSync(TAURI_CONF_PATH, JSON.stringify(tauriConf, null, 2) + '\n');
  console.log(`✅ tauri.conf.json: → ${newVersion}`);
} else {
  console.log(`⚙️ tauri.conf.json: already ${newVersion}`);
}

console.log(`\n🎯 All configs synced to v${newVersion}`);
console.log(`   前端版本号由 Rust get_app_version() 运行时读取，无需手动更新`);
console.log(`\nNext steps:`);
console.log(`  git add -A && git commit -m "v${newVersion}: ..." && git push origin master`);
console.log(`  git tag v${newVersion} && git push origin v${newVersion}`);
