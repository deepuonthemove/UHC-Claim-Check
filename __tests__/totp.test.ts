/**
 * __tests__/totp.test.ts
 *
 * Tests for lib/totp.ts:
 * - Generates a valid 6-digit TOTP code
 * - Throws when TOTP_SECRET is not set
 * - Returns correct seconds remaining
 *
 * .env.local is always preferred over shell env so stale sessions don't
 * cause confusion. The first 6 chars of the active key are printed at
 * startup so you can verify which key is in use.
 */
import { generateTOTP, totpSecondsRemaining } from '@/lib/totp';
import * as fs   from 'fs';
import * as path from 'path';

// ── Always read from .env.local first (same priority as scripts/totp.js) ──────
function readSecretFromEnvLocal(): string {
  try {
    const envFile = fs.readFileSync(path.resolve(__dirname, '../.env.local'), 'utf8');
    const match   = envFile.match(/^TOTP_SECRET=(.+)$/m);
    if (match) return match[1].trim();
  } catch { /* .env.local may not exist in CI */ }
  return '';
}

// .env.local wins over stale shell env
const envLocalSecret = readSecretFromEnvLocal();
const REAL_SECRET    = envLocalSecret || process.env.TOTP_SECRET || '';

// ── Print which key is active so mismatches are instantly visible ─────────────
const keyPreview = REAL_SECRET
  ? `${REAL_SECRET.slice(0, 6)}... (length ${REAL_SECRET.length})`
  : '(NOT SET — tests will fail)';
const source = envLocalSecret ? '.env.local' : 'shell env';
console.log(`\n  🔑  TOTP key [${source}]: ${keyPreview}\n`);

describe('generateTOTP', () => {
  beforeEach(() => {
    process.env.TOTP_SECRET = REAL_SECRET;
  });

  afterEach(() => {
    delete process.env.TOTP_SECRET;
  });

  it('generates a 6-digit numeric code', () => {
    const code = generateTOTP();
    expect(code).toMatch(/^\d{6}$/);
  });

  it('generates consistent codes within the same 30-second window', () => {
    const code1 = generateTOTP();
    const code2 = generateTOTP();
    // Both calls are within the same test run — must be the same window
    expect(code1).toBe(code2);
  });

  it('generates a different code for the next window (+30s offset)', () => {
    const code0    = generateTOTP(0);
    const codePlus = generateTOTP(30);
    // Adjacent 30s windows must produce different codes
    expect(codePlus).toMatch(/^\d{6}$/);
    expect(codePlus).not.toBe(code0);
  });

  it('throws when TOTP_SECRET is not set', () => {
    delete process.env.TOTP_SECRET;
    expect(() => generateTOTP()).toThrow('TOTP_SECRET is not set');
  });
});

describe('totpSecondsRemaining', () => {
  it('returns a number between 1 and 30', () => {
    const remaining = totpSecondsRemaining();
    expect(remaining).toBeGreaterThanOrEqual(1);
    expect(remaining).toBeLessThanOrEqual(30);
  });
});
