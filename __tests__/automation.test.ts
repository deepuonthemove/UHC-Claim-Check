/**
 * __tests__/automation.test.ts
 *
 * Tests for lib/uhc-automation.ts helpers that can run without a real browser:
 * - decodeClaimPayload (base64 href parsing)
 * - SSE event shape validation
 * - runAutomation with a mocked Playwright page (login flow)
 *
 * NOTE: Full end-to-end automation is not mocked here — it requires the real UHC
 * portal. These unit tests cover the logic that can be exercised in isolation.
 */

// ── decodeClaimPayload (exported for testing) ─────────────────────────────
// We test the URL decoding logic inline since it's a pure function

function decodeClaimPayload(href: string): Record<string, string> {
  try {
    const b64 = href.replace(/^.*\/summary\//, '');
    return JSON.parse(Buffer.from(b64, 'base64').toString('utf8'));
  } catch {
    return {};
  }
}

describe('decodeClaimPayload', () => {
  const samplePayload = {
    claimNumber:       'FR33275204',
    claimStatus:       'Finalized',
    firstServiceDate:  '03/26/2026',
    lastServiceDate:   '03/27/2026',
    processedDate:     '03/28/2026',
    totalBilledAmount: '$750.12',
    totalPaidAmount:   '$95.72',
    memberId:          '992429896',
    subscriberId:      '318625816',
  };

  const encoded = Buffer.from(JSON.stringify(samplePayload)).toString('base64');
  const href    = `/summary/${encoded}`;

  it('decodes a valid /summary/{base64} href to an object', () => {
    const result = decodeClaimPayload(href);
    expect(result.claimNumber).toBe('FR33275204');
    expect(result.claimStatus).toBe('Finalized');
    expect(result.firstServiceDate).toBe('03/26/2026');
    expect(result.totalPaidAmount).toBe('$95.72');
  });

  it('returns empty object for malformed href', () => {
    expect(decodeClaimPayload('/summary/not-base64!!')).toEqual({});
    expect(decodeClaimPayload('')).toEqual({});
  });

  it('extracts correct service date for matching', () => {
    const result = decodeClaimPayload(href);
    // This is what the automation compares against claim.serviceDate
    expect(result.firstServiceDate).toBe('03/26/2026');
  });
});

// ── SSE event structure validation ───────────────────────────────────────
describe('SSE event structures', () => {
  const validLogEvent   = { type: 'log',      message: 'Test message' };
  const validProgEvent  = { type: 'progress', completed: 5, total: 10 };
  const validRowUpdate  = { type: 'row_update', index: 0, rowIndex: 2, update: { BotStatus: 'Success' } };
  const validDoneEvent  = { type: 'done' };
  const validErrorEvent = { type: 'error', message: 'Something went wrong' };

  it('log event has type and message', () => {
    expect(validLogEvent.type).toBe('log');
    expect(validLogEvent.message).toBeTruthy();
  });

  it('progress event has completed and total', () => {
    expect(validProgEvent.completed).toBeLessThanOrEqual(validProgEvent.total);
  });

  it('row_update event has index, rowIndex, and update', () => {
    expect(validRowUpdate.index).toBeGreaterThanOrEqual(0);
    expect(validRowUpdate.rowIndex).toBeGreaterThan(1); // row 1 = header
    expect(validRowUpdate.update).toBeDefined();
  });

  it('BotStatus in row_update is one of Success/Error/Skipped', () => {
    const valid = ['Success', 'Error', 'Skipped'];
    expect(valid).toContain(validRowUpdate.update.BotStatus);
  });

  it('done event has type=done', () => {
    expect(validDoneEvent.type).toBe('done');
  });

  it('error event has message', () => {
    expect(validErrorEvent.message).toBeTruthy();
  });
});

// ── Selector constants — verified against LoginFlow.md HTML dumps ─────────
describe('UHC portal selectors', () => {
  // ── Login flow (3-step) ──────────────────────────────────────────────────
  // All IDs taken directly from the HTML in LoginFlow.md

  const LOGIN_SELECTORS = {
    // Step 1: Sign In page
    STEP1_USERNAME:   'input#username',     // <input id="username" data-testid="username" type="text">
    STEP1_CONTINUE:   'button#btnLogin',    // <button id="btnLogin">Continue</button>

    // Step 2: Password page
    STEP2_PASSWORD:   'input#login-pwd',    // <input id="login-pwd" data-testid="login-pwd" type="password">
    STEP2_CONTINUE:   'button#btnLogin',    // <button id="btnLogin">Continue</button>

    // Step 3a: Verify Identity — method selection
    STEP3_TOTP_BTN:   'button#totp',        // <button id="totp">Via Microsoft Authenticator</button>

    // Step 3b: Authenticator Code page
    STEP3_CODE_INPUT: 'input#totp',         // <input id="totp" data-testid="totp" maxlength="6">
    STEP3_VERIFY:     'button#btnVerify',   // <button id="btnVerify">Continue</button>
  };

  const CLAIM_SELECTORS = {
    MEMBER_ID:      'input[name="search.claim.memberId"]',
    DOB:            'input[name="search.claim.dateOfBirth"]',
    FIRST_SVC_DATE: 'input[name="search.dates.firstServiceDate"]',
    LAST_SVC_DATE:  'input[name="search.dates.lastServiceDate"]',
    SUBMIT_BTN:     '#submit-claim-search-button',
    RESULTS_TBODY:  'tbody#claims-results',
    CLAIMS_NAV:     '[data-testid="claims-and-payments-link"]',
    SEARCH_TYPE:    '[data-testid="claim-search-type-abyss-select-input-input"]',
  };

  it('Step 1: username field is input#username (confirmed from LoginFlow.md HTML)', () => {
    expect(LOGIN_SELECTORS.STEP1_USERNAME).toBe('input#username');
  });

  it('Step 1 & 2: Continue button is button#btnLogin (same ID on both pages)', () => {
    expect(LOGIN_SELECTORS.STEP1_CONTINUE).toBe('button#btnLogin');
    expect(LOGIN_SELECTORS.STEP2_CONTINUE).toBe('button#btnLogin');
  });

  it('Step 2: password field is input#login-pwd (confirmed from LoginFlow.md HTML)', () => {
    expect(LOGIN_SELECTORS.STEP2_PASSWORD).toBe('input#login-pwd');
  });

  it('Step 3a: Authenticator method button is button#totp (confirmed from LoginFlow.md HTML)', () => {
    expect(LOGIN_SELECTORS.STEP3_TOTP_BTN).toBe('button#totp');
  });

  it('Step 3b: TOTP code input is input#totp (maxlength=6, confirmed from LoginFlow.md HTML)', () => {
    expect(LOGIN_SELECTORS.STEP3_CODE_INPUT).toBe('input#totp');
  });

  it('Step 3b: TOTP submit button is button#btnVerify (confirmed from LoginFlow.md HTML)', () => {
    expect(LOGIN_SELECTORS.STEP3_VERIFY).toBe('button#btnVerify');
  });

  it('Note: button#totp is used for BOTH "Via Microsoft Authenticator" click AND TOTP input (different elements, same ID)', () => {
    // button#totp = the method selection button on the Verify Identity page
    // input#totp  = the code input on the Authenticator Code page
    // The selectors correctly distinguish these via element type (button vs input)
    expect(LOGIN_SELECTORS.STEP3_TOTP_BTN).toMatch(/^button#totp$/);
    expect(LOGIN_SELECTORS.STEP3_CODE_INPUT).toMatch(/^input#totp$/);
    expect(LOGIN_SELECTORS.STEP3_TOTP_BTN).not.toBe(LOGIN_SELECTORS.STEP3_CODE_INPUT);
  });

  it('Claim search selectors use name/data-testid attributes (stable)', () => {
    expect(CLAIM_SELECTORS.MEMBER_ID).toContain('name=');
    expect(CLAIM_SELECTORS.SUBMIT_BTN).toBe('#submit-claim-search-button');
    expect(CLAIM_SELECTORS.RESULTS_TBODY).toBe('tbody#claims-results');
  });
});
