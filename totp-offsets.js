"use strict";

/**
 * totp-offsets.js
 *
 * Generates a 6-digit TOTP from a base32 secret using otplib.
 * Supports an optional time offset (in seconds) for testing adjacent windows.
 *
 * Works with the otplib version installed in this project (which exposes
 * generateSync at the top level rather than via authenticator.clone()).
 */
const { generateSync } = require("otplib");

function normalizeSecret(secret) {
  // Remove spaces/newlines that are commonly copied along with base32 secrets.
  return secret.replace(/\s+/g, "").trim();
}

function generateTotp(secret, timeOffsetSeconds = 0) {
  if (!secret || !secret.trim()) {
    throw new Error("TOTP secret is empty. Please set TOTP_SECRET in .env.local.");
  }

  return generateSync({
    secret:    normalizeSecret(secret),
    algorithm: "sha1",
    digits:    6,
    step:      30,
    // epoch must be in SECONDS (not ms) for the step division to work correctly
    epoch:     Math.floor(Date.now() / 1000) + timeOffsetSeconds,
  });
}

module.exports = {
  generateTotp,
};