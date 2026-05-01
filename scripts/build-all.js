#!/usr/bin/env node
/**
 * 一键全量构建脚本 — 产出 4 个产物:
 *   1. 轻量版安装包 (NSIS, ~3.6MB)
 *   2. 轻量版绿色便携 (exe zip, ~14MB)
 *   3. OCR 版安装包 (NSIS, ~24MB)
 *   4. OCR 版绿色便携 (exe + models/ zip, ~22MB)
 *
 * 用法: node scripts/build-all.js
 *        npm run build:all
 */

const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

// ─── 配置 ───────────────────────────────────────────
const ROOT = path.resolve(__dirname, '..');
const pkg = require(path.join(ROOT, 'package.json'));
const VERSION = pkg.version;
const PRODUCT_NAME = '发票打印工具';
const EXE_NAME = 'fapiao-print.exe';           // Cargo 编译出的二进制名
const ARCH = 'x64';

const TARGET_RELEASE = path.join(ROOT, 'src-tauri', 'target', 'release');
const BUNDLE_NSIS = path.join(TARGET_RELEASE, 'bundle', 'nsis');
const MODELS_SRC = path.join(ROOT, 'src-tauri', 'models');
const DIST = path.join(ROOT, 'dist');

// ─── 工具函数 ───────────────────────────────────────
function run(cmd, opts = {}) {
  console.log(`\n\x1b[36m▶ ${cmd}\x1b[0m`);
  execSync(cmd, { stdio: 'inherit', cwd: ROOT, ...opts });
}

function sizeMB(filepath) {
  const stat = fs.statSync(filepath);
  return (stat.size / 1024 / 1024).toFixed(1);
}

function copyFile(src, dest) {
  fs.mkdirSync(path.dirname(dest), { recursive: true });
  fs.copyFileSync(src, dest);
  console.log(`  ✓ ${path.basename(dest)} (${sizeMB(dest)} MB)`);
}

/**
 * 用 PowerShell Compress-Archive 创建 zip
 * @param {string} zipPath  输出 zip 路径
 * @param {string[]} items  要打包的文件/目录路径
 * @param {string} workDir  解压后的根目录名（zip 内的一级目录）
 */
function createZip(zipPath, items, workDir) {
  // 先删掉已存在的 zip（Compress-Archive 不允许覆盖）
  if (fs.existsSync(zipPath)) fs.unlinkSync(zipPath);

  // 在临时目录中组装好结构，再整目录打包
  const tmpDir = path.join(DIST, `_zip_tmp_${Date.now()}`);
  const contentDir = path.join(tmpDir, workDir);
  fs.mkdirSync(contentDir, { recursive: true });

  for (const item of items) {
    const basename = path.basename(item);
    const dest = path.join(contentDir, basename);
    if (fs.statSync(item).isDirectory()) {
      copyDirRecursive(item, dest);
    } else {
      fs.copyFileSync(item, dest);
    }
  }

  // PowerShell Compress-Archive
  const psCmd = `Compress-Archive -Path '${contentDir}' -DestinationPath '${zipPath}' -Force`;
  run(`powershell -NoProfile -Command "${psCmd}"`, { cwd: ROOT });

  // 清理临时目录
  fs.rmSync(tmpDir, { recursive: true, force: true });
  console.log(`  ✓ ${path.basename(zipPath)} (${sizeMB(zipPath)} MB)`);
}

function copyDirRecursive(src, dest) {
  fs.mkdirSync(dest, { recursive: true });
  for (const entry of fs.readdirSync(src, { withFileTypes: true })) {
    const srcPath = path.join(src, entry.name);
    const destPath = path.join(dest, entry.name);
    if (entry.isDirectory()) {
      copyDirRecursive(srcPath, destPath);
    } else {
      fs.copyFileSync(srcPath, destPath);
    }
  }
}

// ─── 主流程 ─────────────────────────────────────────
async function main() {
  const t0 = Date.now();
  console.log(`\n${'═'.repeat(60)}`);
  console.log(`  发票打印工具 v${VERSION} — 全量构建`);
  console.log(`${'═'.repeat(60)}`);

  // 清理 staging
  fs.rmSync(DIST, { recursive: true, force: true });
  fs.mkdirSync(DIST, { recursive: true });

  // ─── Step 1: 轻量版编译 ─────────────────────────
  console.log(`\n${'─'.repeat(60)}`);
  console.log('  [1/4] 编译轻量版 (无 OCR)...');
  console.log(`${'─'.repeat(60)}`);
  run('npx tauri build');

  // 保存轻量版产物（OCR 编译会覆盖）
  const lwInstaller = path.join(BUNDLE_NSIS, `${PRODUCT_NAME}_${VERSION}_${ARCH}-setup.exe`);
  const lwExe = path.join(TARGET_RELEASE, EXE_NAME);
  const lwStaging = path.join(DIST, 'lightweight');
  fs.mkdirSync(lwStaging, { recursive: true });

  console.log('\n📦 保存轻量版产物...');
  copyFile(lwInstaller, path.join(lwStaging, `${PRODUCT_NAME}_${VERSION}_${ARCH}-setup.exe`));
  copyFile(lwExe, path.join(lwStaging, `${PRODUCT_NAME}.exe`));

  // ─── Step 2: OCR 版编译 ─────────────────────────
  console.log(`\n${'─'.repeat(60)}`);
  console.log('  [2/4] 编译 OCR 版 (含 PP-OCRv5)...');
  console.log(`${'─'.repeat(60)}`);
  run('npx tauri build --features ocr --config src-tauri/tauri.ocr.conf.json');

  // 保存 OCR 版产物
  const ocrInstaller = path.join(BUNDLE_NSIS, `${PRODUCT_NAME}_${VERSION}_${ARCH}-setup.exe`);
  const ocrExe = path.join(TARGET_RELEASE, EXE_NAME);
  const ocrStaging = path.join(DIST, 'ocr');
  fs.mkdirSync(ocrStaging, { recursive: true });

  console.log('\n📦 保存 OCR 版产物...');
  copyFile(ocrInstaller, path.join(ocrStaging, `${PRODUCT_NAME}_${VERSION}_${ARCH}-setup.exe`));
  copyFile(ocrExe, path.join(ocrStaging, `${PRODUCT_NAME}.exe`));

  // ─── Step 3: 绿色便携版打包 ─────────────────────
  console.log(`\n${'─'.repeat(60)}`);
  console.log('  [3/4] 打包绿色便携版...');
  console.log(`${'─'.repeat(60)}`);

  // 轻量版绿色便携: 仅 exe
  const lwZipName = `${PRODUCT_NAME}_${VERSION}_${ARCH}_绿色版.zip`;
  const lwZipPath = path.join(DIST, lwZipName);
  console.log('\n📦 轻量版绿色便携 (仅 exe)...');
  createZip(lwZipPath, [path.join(lwStaging, `${PRODUCT_NAME}.exe`)], `${PRODUCT_NAME}_${VERSION}_绿色版`);

  // OCR 版绿色便携: exe + models/
  const ocrZipName = `${PRODUCT_NAME}_${VERSION}_${ARCH}_OCR绿色版.zip`;
  const ocrZipPath = path.join(DIST, ocrZipName);
  console.log('\n📦 OCR 版绿色便携 (exe + models/)...');
  createZip(ocrZipPath, [
    path.join(ocrStaging, `${PRODUCT_NAME}.exe`),
    MODELS_SRC,
  ], `${PRODUCT_NAME}_${VERSION}_OCR绿色版`);

  // ─── Step 4: 汇总 ────────────────────────────────
  console.log(`\n${'─'.repeat(60)}`);
  console.log('  [4/4] 产物汇总');
  console.log(`${'─'.repeat(60)}\n`);

  const artifacts = [
    { name: '轻量版安装包', path: path.join(lwStaging, `${PRODUCT_NAME}_${VERSION}_${ARCH}-setup.exe`) },
    { name: '轻量版绿色便携', path: lwZipPath },
    { name: 'OCR 版安装包', path: path.join(ocrStaging, `${PRODUCT_NAME}_${VERSION}_${ARCH}-setup.exe`) },
    { name: 'OCR 版绿色便携', path: ocrZipPath },
  ];

  for (const a of artifacts) {
    if (fs.existsSync(a.path)) {
      console.log(`  ✅ ${a.name}: ${sizeMB(a.path)} MB`);
      console.log(`     ${a.path}`);
    } else {
      console.log(`  ❌ ${a.name}: 未找到!`);
    }
  }

  const elapsed = ((Date.now() - t0) / 1000 / 60).toFixed(1);
  console.log(`\n⏱  总耗时: ${elapsed} 分钟`);
  console.log(`\n产物目录: ${DIST}\n`);
}

main().catch(err => {
  console.error('\n❌ 构建失败:', err.message);
  process.exit(1);
});
