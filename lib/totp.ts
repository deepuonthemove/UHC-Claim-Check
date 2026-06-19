/**
 * lib/totp.ts
 *
 * Thin wrapper around totp-offsets.js so the server automation and the
 * `npm run totp` CLI command share exactly the same generation logic.
 *
 * Secret is read from TOTP_SECRET in .env.local (loaded by Next.js automatically).
 */
// eslint-disable-next-line @typescript-eslint/no-require-imports
const { generateTotp } = require('../totp-offsets');

/** Generate a 6-digit TOTP from TOTP_SECRET (throws if secret is missing). */
export function generateTOTP(timeOffsetSeconds = 0): string {
  const secret = process.env.TOTP_SECRET;
  if (!secret) {
    throw new Error('TOTP_SECRET is not set in environment variables.');
  }
  return generateTotp(secret, timeOffsetSeconds) as string;
}

/** Returns seconds remaining until the current TOTP window expires (1–30). */
export function totpSecondsRemaining(): number {
  const epoch = Math.floor(Date.now() / 1000);
  return 30 - (epoch % 30);
}
