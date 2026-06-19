#!/usr/bin/env node
/**
 * scripts/totp.js
 *
 * CLI helper — generates TOTP codes from TOTP_SECRET.
 * Reads .env.local automatically (same source of truth as the Next.js app).
 *
 * Usage:
 *   npm run totp              → current 6-digit code + time remaining
 *   npm run totp:watch        → live countdown, refreshes every second
 *   npm run totp:range        → table: prev 1 min → next 3 min at 30s intervals
 *   npm run totp -- --offset -30  → code for a specific time offset
 */
"use strict";

const fs   = require("fs");
const path = require("path");

// ── Load .env.local — always takes priority over shell env ───────────────────
// Shell env may be stale from a previous session; .env.local is the source of
// truth. We always override so changing the file takes effect immediately.
const envPath = path.resolve(__dirname, "../.env.local");
if (fs.existsSync(envPath)) {
  const lines = fs.readFileSync(envPath, "utf8").split("\n");
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const eqIdx = trimmed.indexOf("=");
    if (eqIdx < 0) continue;
    const key   = trimmed.slice(0, eqIdx).trim();
    const value = trimmed.slice(eqIdx + 1).trim().replace(/^["']|["']$/g, "");
    if (process.env[key] && process.env[key] !== value) {
      console.warn(`⚠️  Shell had a stale ${key} — overriding with .env.local value.`);
    }
    process.env[key] = value; // .env.local always wins
  }
} else {
  console.warn("⚠️  .env.local not found — using shell environment variables.");
}

// ── Load generation helper ────────────────────────────────────────────────────
const { generateTotp } = require("../totp-offsets");

// ── Helpers ───────────────────────────────────────────────────────────────────
function secondsRemaining() {
  return 30 - (Math.floor(Date.now() / 1000) % 30);
}

function getSecret() {
  const secret = process.env.TOTP_SECRET;
  if (!secret) {
    console.error("❌  TOTP_SECRET is not set. Add it to .env.local:\n   TOTP_SECRET=your_base32_secret");
    process.exit(1);
  }
  return secret;
}

function formatTime(offsetSeconds) {
  const ts  = new Date(Date.now() + offsetSeconds * 1000);
  const hh  = String(ts.getHours()).padStart(2, "0");
  const mm  = String(ts.getMinutes()).padStart(2, "0");
  const ss  = String(ts.getSeconds()).padStart(2, "0");
  return `${hh}:${mm}:${ss}`;
}

function formatOffset(secs) {
  if (secs === 0)  return "now    ";
  const sign = secs > 0 ? "+" : "-";
  const abs  = Math.abs(secs);
  // Use whole-minute labels only when evenly divisible, else show seconds
  if (abs % 60 === 0) return `${sign}${abs / 60}min  `;
  return `${sign}${abs}s   `;
}

// ── Parse args ────────────────────────────────────────────────────────────────
const args       = process.argv.slice(2);
const watchMode  = args.includes("--watch") || args.includes("-w");
const rangeMode  = args.includes("--range") || args.includes("-r");
const tightMode  = args.includes("--tight") || args.includes("-t");
const offsetIdx  = args.indexOf("--offset");
const offsetSecs = offsetIdx >= 0 ? parseInt(args[offsetIdx + 1] ?? "0", 10) : 0;

// ── Range mode: prev 1 min → next 3 min at 30s intervals ─────────────────────
//
//   Window    Offset    Time          Code
//   ───────   ───────   ─────────     ──────
//   prev 2    -60s      hh:mm:ss      xxxxxx
//   prev 1    -30s      hh:mm:ss      xxxxxx
// ► current     0s      hh:mm:ss      xxxxxx   ← active window
//   next 1    +30s      hh:mm:ss      xxxxxx
//   next 2    +60s      hh:mm:ss      xxxxxx
//   next 3    +90s      hh:mm:ss      xxxxxx
//   next 4   +120s      hh:mm:ss      xxxxxx
//   next 5   +150s      hh:mm:ss      xxxxxx
//   next 6   +180s      hh:mm:ss      xxxxxx
//
function printRange() {
  const secret  = getSecret();
  const windows = [-60, -30, 0, 30, 60, 90, 120, 150, 180];
  const remaining = secondsRemaining();
  const bar       = "█".repeat(Math.round(remaining / 30 * 20)).padEnd(20, "░");

  console.log("");
  console.log("  \x1b[1mTOTP Window Table\x1b[0m" +
    `  (current window expires in \x1b[33m${remaining}s\x1b[0m  [${bar}])\n`);
  console.log(
    "  " +
    "\x1b[2mWindow   Offset    Clock       Code\x1b[0m"
  );
  console.log("  " + "─".repeat(44));

  for (const offset of windows) {
    const code      = generateTotp(secret, offset);
    const isCurrent = offset === 0;
    const prefix    = isCurrent ? "► " : "  ";
    const label     = offset < 0  ? `prev ${Math.abs(offset / 30)}`
                    : offset === 0 ? "current"
                    : `next ${offset / 30}`;
    const offsetStr = formatOffset(offset).padEnd(8);
    const timeStr   = formatTime(offset);
    const codeStr   = isCurrent
      ? `\x1b[1;32m${code}\x1b[0m`   // bright green for current
      : offset < 0
        ? `\x1b[2m${code}\x1b[0m`    // dim for past
        : `\x1b[36m${code}\x1b[0m`;  // cyan for future

    console.log(
      `  ${prefix}\x1b[1m${label.padEnd(8)}\x1b[0m  ${offsetStr}  ${timeStr}  ${codeStr}`
    );
  }
  console.log("");
}

// ── Single code mode ──────────────────────────────────────────────────────────
function printCode() {
  const secret    = getSecret();
  const code      = generateTotp(secret, offsetSecs);
  const remaining = secondsRemaining();
  const bar       = "█".repeat(Math.round(remaining / 30 * 20)).padEnd(20, "░");

  if (watchMode) process.stdout.write("\x1B[2K\r"); // clear line in-place
  process.stdout.write(
    `  🔑  TOTP: \x1b[1;32m${code}\x1b[0m   ` +
    `⏳ ${String(remaining).padStart(2)}s  [${bar}]` +
    (watchMode ? "" : "\n")
  );
}

// ── Tight mode: every second from -20s to +20s ───────────────────────────────
//
// Useful for checking exactly when your code will flip to the next window.
// Rows sharing the same TOTP code are grouped visually; a bright separator
// marks every 30-second boundary (where the code actually changes).
//
function printTight() {
  const secret    = getSecret();
  const remaining = secondsRemaining();
  const bar       = "█".repeat(Math.round(remaining / 30 * 20)).padEnd(20, "░");

  console.log("");
  console.log("  \x1b[1mTOTP Second-by-Second (±20s)\x1b[0m" +
    `  (expires in \x1b[33m${remaining}s\x1b[0m  [${bar}])\n`);
  console.log("  \x1b[2mOffset   Clock       Code      Window\x1b[0m");
  console.log("  " + "─".repeat(46));

  let prevCode = null;
  let prevWindow = null;

  for (let offset = -20; offset <= 20; offset++) {
    const code      = generateTotp(secret, offset);
    const ts        = new Date(Date.now() + offset * 1000);
    const hh        = String(ts.getHours()).padStart(2, "0");
    const mm        = String(ts.getMinutes()).padStart(2, "0");
    const ss        = String(ts.getSeconds()).padStart(2, "0");
    const clockStr  = `${hh}:${mm}:${ss}`;
    const offsetStr = (offset >= 0 ? "+" : "") + String(offset).padStart(3);
    const isNow     = offset === 0;

    // Compute which 30s window this second falls in
    const epochSec  = Math.floor((Date.now() + offset * 1000) / 1000);
    const window30  = Math.floor(epochSec / 30);

    // Insert a boundary separator when the TOTP window changes
    if (prevCode !== null && code !== prevCode) {
      console.log(
        "  " + "┄".repeat(14) +
        " \x1b[1;33m▲ code changes here\x1b[0m " +
        "┄".repeat(9)
      );
    }

    const isNewWindow = prevWindow !== null && window30 !== prevWindow;
    const codeColor   = isNow    ? "\x1b[1;32m"  // bright green = current second
                      : offset < 0 ? "\x1b[2m"   // dim = past
                      : "\x1b[36m";              // cyan = future
    const arrow       = isNow ? " ◄ now" : "";

    console.log(
      `  ${offsetStr}s   ${clockStr}   ${codeColor}${code}\x1b[0m${arrow}`
    );

    prevCode   = code;
    prevWindow = window30;
  }

  console.log("");
}

// ── Run ───────────────────────────────────────────────────────────────────────
// Print key preview once at startup (before entering any mode)
const _secret = process.env.TOTP_SECRET || '';
if (_secret) {
  console.log(`  🔑  Key: \x1b[2m${_secret.slice(0, 6)}...\x1b[0m (length ${_secret.length})`);
}

if (tightMode) {
  printTight();
} else if (rangeMode) {
  printRange();
} else if (watchMode) {
  console.log("👁  Watching TOTP codes (Ctrl+C to stop)...\n");
  printCode();
  const interval = setInterval(printCode, 1000);
  process.on("SIGINT", () => {
    clearInterval(interval);
    process.stdout.write("\n");
    process.exit(0);
  });
} else {
  printCode();
}
