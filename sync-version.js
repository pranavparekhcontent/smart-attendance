// ============================================================
//  sync-version.js — Auto-sync version across ALL config files
//  
//  Usage:
//    node sync-version.js          → reads highest version, syncs everywhere
//    node sync-version.js 1.0.5   → sets 1.0.5 everywhere
//    node sync-version.js bump     → auto-bumps patch (1.0.3 → 1.0.4)
//
//  Files synced:
//    1. version.json                    (root)
//    2. appstart/config.js              (APP_VERSION)
//    3. js/config.js                    (APP_VERSION) — if exists
// ============================================================

const fs = require('fs');
const path = require('path');

// ── File paths ─────────────────────────────────────────────
const ROOT = __dirname;
const FILES = {
  versionJson:   path.join(ROOT, 'version.json'),
  appstartConfig: path.join(ROOT, 'appstart', 'config.js'),
  jsConfig:      path.join(ROOT, 'js', 'config.js'),
};

// ── Helpers ────────────────────────────────────────────────
function readFile(filepath) {
  try { return fs.readFileSync(filepath, 'utf-8'); } catch { return null; }
}

function writeFile(filepath, content) {
  fs.writeFileSync(filepath, content, 'utf-8');
}

/** Extract version string from file content */
function extractVersion(content, type) {
  if (!content) return null;
  let match;
  if (type === 'json') {
    match = content.match(/"version"\s*:\s*"([^"]+)"/);
  } else {
    // JS file — matches both: APP_VERSION: "1.0.3" and APP_VERSION = '1.0.3'
    match = content.match(/APP_VERSION\s*[:=]\s*['"]([^'"]+)['"]/);
  }
  return match ? match[1] : null;
}

/** Compare two semver strings → returns 1, -1, or 0 */
function compareSemver(a, b) {
  const pa = a.split('.').map(Number);
  const pb = b.split('.').map(Number);
  for (let i = 0; i < 3; i++) {
    const va = pa[i] || 0;
    const vb = pb[i] || 0;
    if (va > vb) return 1;
    if (va < vb) return -1;
  }
  return 0;
}

/** Bump patch version: 1.0.3 → 1.0.4 */
function bumpPatch(version) {
  const parts = version.split('.').map(Number);
  parts[2] = (parts[2] || 0) + 1;
  return parts.join('.');
}

/** Write version into version.json */
function updateVersionJson(content, newVersion) {
  return content.replace(
    /"version"\s*:\s*"[^"]+"/,
    `"version": "${newVersion}"`
  );
}

/** Write version into a JS config file */
function updateJsConfig(content, newVersion) {
  // Handles both formats:
  //   APP_VERSION: "1.0.3",    (object property)
  //   APP_VERSION = '1.0.3';   (export const)
  return content.replace(
    /(APP_VERSION\s*[:=]\s*)(['"])([^'"]+)\2/,
    `$1$2${newVersion}$2`
  );
}

// ── Main ───────────────────────────────────────────────────
function main() {
  const arg = process.argv[2]; // optional: version string or "bump"

  console.log('');
  console.log('  ╔═══════════════════════════════════╗');
  console.log('  ║   🔄 VERSION SYNC TOOL            ║');
  console.log('  ╚═══════════════════════════════════╝');
  console.log('');

  // 1. Read current versions from all files
  const contents = {};
  const versions = {};

  // version.json
  contents.versionJson = readFile(FILES.versionJson);
  if (contents.versionJson) {
    versions.versionJson = extractVersion(contents.versionJson, 'json');
    console.log(`  📄 version.json         → ${versions.versionJson || '❌ not found'}`);
  } else {
    console.log(`  📄 version.json         → ❌ file missing`);
  }

  // appstart/config.js
  contents.appstartConfig = readFile(FILES.appstartConfig);
  if (contents.appstartConfig) {
    versions.appstartConfig = extractVersion(contents.appstartConfig, 'js');
    console.log(`  📄 appstart/config.js   → ${versions.appstartConfig || '❌ not found'}`);
  } else {
    console.log(`  📄 appstart/config.js   → ❌ file missing`);
  }

  // js/config.js (optional — may not exist in template)
  contents.jsConfig = readFile(FILES.jsConfig);
  if (contents.jsConfig) {
    versions.jsConfig = extractVersion(contents.jsConfig, 'js');
    console.log(`  📄 js/config.js         → ${versions.jsConfig || '❌ no APP_VERSION found'}`);
  } else {
    console.log(`  📄 js/config.js         → (not present, skipping)`);
  }

  console.log('');

  // 2. Determine target version
  let targetVersion;
  const allVersions = Object.values(versions).filter(Boolean);

  if (!allVersions.length) {
    console.log('  ❌ No versions found in any file!');
    process.exit(1);
  }

  if (arg === 'bump') {
    // Auto-bump: find highest, bump patch
    const highest = allVersions.sort(compareSemver).pop();
    targetVersion = bumpPatch(highest);
    console.log(`  🚀 Bumping: ${highest} → ${targetVersion}`);
  } else if (arg && /^\d+\.\d+\.\d+$/.test(arg)) {
    // Explicit version provided
    targetVersion = arg;
    console.log(`  🎯 Setting explicit version: ${targetVersion}`);
  } else if (arg) {
    console.log(`  ❌ Invalid version format: "${arg}"`);
    console.log(`     Use: node sync-version.js 1.0.5`);
    console.log(`     Or:  node sync-version.js bump`);
    process.exit(1);
  } else {
    // No argument: use highest found version
    targetVersion = allVersions.sort(compareSemver).pop();
    console.log(`  🔍 Highest version found: ${targetVersion}`);
  }

  console.log('');

  // 3. Check if anything needs updating
  const allMatch = allVersions.every(v => v === targetVersion);
  if (allMatch && !arg) {
    console.log(`  ✅ All files already at v${targetVersion}. Nothing to do.`);
    console.log('');
    return;
  }

  // 4. Write to all files
  let updated = 0;

  if (contents.versionJson) {
    const newContent = updateVersionJson(contents.versionJson, targetVersion);
    if (newContent !== contents.versionJson) {
      writeFile(FILES.versionJson, newContent);
      console.log(`  ✅ version.json         → v${targetVersion}`);
      updated++;
    } else {
      console.log(`  ── version.json         → already v${targetVersion}`);
    }
  }

  if (contents.appstartConfig) {
    const newContent = updateJsConfig(contents.appstartConfig, targetVersion);
    if (newContent !== contents.appstartConfig) {
      writeFile(FILES.appstartConfig, newContent);
      console.log(`  ✅ appstart/config.js   → v${targetVersion}`);
      updated++;
    } else {
      console.log(`  ── appstart/config.js   → already v${targetVersion}`);
    }
  }

  if (contents.jsConfig && versions.jsConfig !== undefined) {
    const newContent = updateJsConfig(contents.jsConfig, targetVersion);
    if (newContent !== contents.jsConfig) {
      writeFile(FILES.jsConfig, newContent);
      console.log(`  ✅ js/config.js         → v${targetVersion}`);
      updated++;
    } else {
      console.log(`  ── js/config.js         → already v${targetVersion}`);
    }
  }

  console.log('');
  console.log(`  🎉 Done! ${updated} file(s) updated to v${targetVersion}`);
  console.log('');
}

main();
