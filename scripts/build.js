#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");

const ROOT = path.resolve(__dirname, "..");
const SRC_UI = path.join(ROOT, "frontend", "public");
const SRC_BACKEND = path.join(ROOT, "backend", "src");
const DIST_DIR = path.join(ROOT, "dist");
const DIST_UI = path.join(DIST_DIR, "public");
const DIST_BACKEND = path.join(DIST_DIR, "backend", "src");

async function exists(p) {
  try {
    await fs.promises.access(p, fs.constants.F_OK);
    return true;
  } catch {
    return false;
  }
}

async function rmDir(p) {
  await fs.promises.rm(p, { recursive: true, force: true });
}

async function ensureDir(p) {
  await fs.promises.mkdir(p, { recursive: true });
}

async function copyDir(src, dest) {
  // Node 18+ supports fs.cp
  await fs.promises.cp(src, dest, { recursive: true });
}

function checkNodeSyntax(absFilePath) {
  // node --check parses the file and exits non-zero on syntax errors.
  const r = spawnSync(process.execPath, ["--check", absFilePath], {
    stdio: "pipe",
    env: process.env,
  });
  if (r.status === 0) return;
  const stdout = (r.stdout || "").toString().trim();
  const stderr = (r.stderr || "").toString().trim();
  const msg = [stdout, stderr].filter(Boolean).join("\n");
  throw new Error(`Syntax check failed for ${absFilePath}\n${msg}`);
}

async function main() {
  if (!(await exists(SRC_UI))) {
    throw new Error(`UI source folder not found: ${SRC_UI}`);
  }
  if (!(await exists(SRC_BACKEND))) {
    throw new Error(`Backend source folder not found: ${SRC_BACKEND}`);
  }

  await rmDir(DIST_DIR);
  await ensureDir(DIST_DIR);

  // "Build" backend: validate JS syntax and copy sources to dist (no bundling/compilation needed).
  const backendEntries = await fs.promises.readdir(SRC_BACKEND, { withFileTypes: true });
  const backendFiles = backendEntries
    .filter((d) => d.isFile() && d.name.endsWith(".js"))
    .map((d) => path.join(SRC_BACKEND, d.name));

  if (!backendFiles.length) {
    throw new Error(`No backend .js files found in: ${SRC_BACKEND}`);
  }

  for (const f of backendFiles) checkNodeSyntax(f);
  await copyDir(SRC_BACKEND, DIST_BACKEND);

  await copyDir(SRC_UI, DIST_UI);

  const files = await fs.promises.readdir(DIST_UI).catch(() => []);
  console.log("[build] OK");
  console.log(
    `[build] Backend checked+copied: ${backendFiles.length} files -> ${path.relative(
      ROOT,
      DIST_BACKEND,
    )}`,
  );
  console.log(`[build] Copied UI: ${path.relative(ROOT, SRC_UI)} -> ${path.relative(ROOT, DIST_UI)}`);
  console.log(`[build] Files: ${files.length}`);
}

main().catch((err) => {
  console.error("[build] FAILED:", err?.message || err);
  process.exitCode = 1;
});


