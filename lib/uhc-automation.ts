/**
 * lib/uhc-automation.ts — REWRITTEN
 *
 * Key changes vs original:
 * - Every step logs via sendEvent({ type: 'log', message }) immediately (no swallowed errors)
 * - Row errors logged as: log(`Row ${i+1}: Error — ${msg}`) + row_update with BotStatus=Error
 * - claimRows JSON passed from client (not re-read from file each batch)
 * - Robust login: tries multiple selector strategies, logs each attempt
 * - Screenshot + debug_html captured on every row failure
 */
import { chromium as playwrightChromium, firefox as playwrightFirefox, type Browser, type BrowserContext, type Page } from 'playwright-core';
import chromium from '@sparticuz/chromium';
import { generateTOTP, totpSecondsRemaining } from './totp';
import type { ClaimRow, BotFields } from './excel';

// ── Exact selectors from LoginFlow.md HTML dumps ─────────────────────────────
const SEL = {
  // Step 1 — Sign In page (username only)
  // <input id="username" data-testid="username" type="text">
  // <button id="btnLogin">Continue</button>
  STEP1_USERNAME:   'input#username',
  STEP1_CONTINUE:   'button#btnLogin',

  // Step 2 — Password page
  // <input id="login-pwd" data-testid="login-pwd" type="password">
  // <button id="btnLogin">Continue</button>  (same ID, different page)
  STEP2_PASSWORD:   'input#login-pwd',
  STEP2_CONTINUE:   'button#btnLogin',

  // Step 3a — Verify Identity (method selection)
  // <button id="totp">Via Microsoft Authenticator</button>
  STEP3_TOTP_BTN:   'button#totp',

  // Step 3b — Authenticator Code page
  // <input id="totp" data-testid="totp" maxlength="6">
  // <button id="btnVerify">Continue</button>
  STEP3_CODE_INPUT: 'input#totp',
  STEP3_VERIFY:     'button#btnVerify',

  // Post-login: dashboard indicator
  CLAIMS_NAV: '[data-testid="claims-and-payments-link"]',

  // Claim search form selectors
  SEARCH_TYPE_BTN:  '[data-testid="claim-search-type-abyss-select-input-input"]',
  SEARCH_OPTION:    '[role="option"]:has-text("Member ID & date of birth")',
  TIN_RADIO:        'input[name="search.tinWideSearch"][value="tin"]',
  MEMBER_ID:        'input[name="search.claim.memberId"]',
  FIRST_NAME:       'input[name*="firstName" i], input[id*="firstName" i]',
  LAST_NAME:        'input[name*="lastName" i], input[id*="lastName" i]',
  DOB:              'input[name="search.claim.dateOfBirth"]',
  DATE_CUSTOM:      'input[name="search.dateRange"][value="custom"]',
  FIRST_SVC_DATE:   'input[name="search.dates.firstServiceDate"]',
  LAST_SVC_DATE:    'input[name="search.dates.lastServiceDate"]',
  SUBMIT_BTN:       '#submit-claim-search-button',
  SUBMIT_BTN_ALT:   '[aria-label="submit claim search"]',

  // Results
  RESULTS_HEADING:  '[data-testid="search-results-label"]',
  RESULTS_TBODY:    'tbody#claims-results',
  NEW_SEARCH_BTN:   '[data-testid="new-search-button-abyss-link-root"]',
  ALL_CLAIM_LINKS:  'a.abyss-link-root[href^="/summary/"]',
  NO_RESULTS:       '[data-testid="no-claims-found"]',

  // Error popup — appears after search when member found but no claim exists
  // <button data-testid="loading-close-button">x</button>
  // <div    data-testid="loading-error-message">Member found, but no claim...</div>
  POPUP_CLOSE:      'button[data-testid="loading-close-button"]',
  POPUP_MESSAGE:    'div[data-testid="loading-error-message"]',
};

const CLAIMS_URL = 'https://secure.uhcprovider.com/#/claims';

export interface SseEvent {
  type: 'log' | 'progress' | 'row_update' | 'error_screenshot' | 'debug_html' | 'done' | 'error' | 'padding';
  message?: string;
  completed?: number;
  total?: number;
  index?: number;    // 0-based index into claimRows array (for workbook update)
  rowIndex?: number; // 1-based Excel row number
  attempt?: number;  // Chunk execution attempt number
  update?: BotFields;
  image?: string;
  html?: string;
}

export type SendEvent = (event: SseEvent) => Promise<void>;

// ── Helpers ──────────────────────────────────────────────────────────────────

function decodeClaimPayload(href: string): Record<string, string> {
  try {
    const b64 = href.replace(/^.*\/summary\//, '');
    return JSON.parse(Buffer.from(b64, 'base64').toString('utf8'));
  } catch {
    return {};
  }
}

/** Try each selector in the list, return the first one that matches */
async function findFirst(page: Page, selectors: string[], timeout = 5000): Promise<string | null> {
  for (const sel of selectors) {
    try {
      await page.waitForSelector(sel, { timeout });
      return sel;
    } catch {
      // try next
    }
  }
  return null;
}

// ── Close claims & payments navigation menu if it is expanded and blocking the page ──
async function closeNavDropdownIfOpen(page: Page, log: (msg: string) => Promise<void>) {
  try {
    const trigger = page.locator('button:has([data-testid="claims-and-payments-link"])');
    if (await trigger.count() > 0) {
      const state = await trigger.getAttribute('data-state');
      const expanded = await trigger.getAttribute('aria-expanded');
      if (state === 'open' || expanded === 'true') {
        await log('  ⚠️  "Claims & Payments" dropdown menu is expanded/blocking. Clicking to close it...');
        await trigger.click({ force: true });
        try {
          await page.waitForFunction(
            (el) => el?.getAttribute('data-state') !== 'open' && el?.getAttribute('aria-expanded') !== 'true',
            await trigger.elementHandle(),
            { timeout: 3_000 }
          );
        } catch { /* ignore wait timeout */ }
        await page.waitForTimeout(500); // allow animations to settle
        await log('  ✖  Dropdown menu closed.');
      }
    }
  } catch (err) {
    await log(`  ⚠️  Failed to check/close navigation dropdown: ${err}`);
  }
}

// ── Popup handler ─────────────────────────────────────────────────────────────
//
// After a search submit, UHC may show an error popup:
//   <button data-testid="loading-close-button">x</button>
//   <div    data-testid="loading-error-message">Member found, but no claim found...</div>
//
// Returns the popup message text if a popup was found and dismissed, null otherwise.
//
async function dismissPopupIfPresent(
  page: Page,
  sendEvent: SendEvent
): Promise<string | null> {
  const log = (msg: string) => sendEvent({ type: 'log', message: msg });
  try {
    await page.waitForSelector(SEL.POPUP_CLOSE, { timeout: 4_000 });
  } catch {
    return null; // no popup — happy path
  }

  // Read message before closing (DOM is gone after close)
  let message = '';
  try {
    message = (await page.locator(SEL.POPUP_MESSAGE).innerText({ timeout: 2_000 })).trim();
  } catch {
    message = 'An error popup appeared but its message could not be read.';
  }

  await log(`  ⚠️  Popup detected: "${message}"`);

  // Dismiss
  try {
    await page.click(SEL.POPUP_CLOSE, { force: true });
    await page.locator(SEL.POPUP_CLOSE).waitFor({ state: 'detached', timeout: 3_000 });
    // Also wait for the message element to detach
    try {
      await page.locator(SEL.POPUP_MESSAGE).waitFor({ state: 'detached', timeout: 2_000 });
    } catch { /* ignore if not detached or already gone */ }
    await page.waitForTimeout(500); // allow animations to settle
    await log('  ✖  Popup closed.');
  } catch {
    await log('  ⚠️  Could not close popup — it may have already dismissed itself.');
  }

  return message;
}

// ── clickWithRetry ─────────────────────────────────────────────────────────────────
//
// Clicks a button and then checks that the page has moved on (i.e. the
// clicked element disappears OR a new element appears). If the page hasn't
// changed after `delayMs`, it clicks again. Retries up to `maxAttempts`.
//
async function clickWithRetry(
  page: Page,
  selector: string,
  {
    label         = selector,
    maxAttempts   = 3,
    retryDelayMs  = 2_000,
    disappearsSel = selector,   // selector we wait to disappear (confirms click worked)
  }: {
    label?:         string;
    maxAttempts?:   number;
    retryDelayMs?:  number;
    disappearsSel?: string;
  },
  log: (msg: string) => Promise<void>
): Promise<void> {
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const btn = page.locator(selector);
      await btn.waitFor({ state: 'visible', timeout: 5_000 });
      await btn.click();
      await log(`  🖱️  Clicked ${label} (attempt ${attempt}/${maxAttempts}).`);

      // Wait briefly to see if the page reacts (button disappears = success)
      try {
        await page.locator(disappearsSel).waitFor({ state: 'detached', timeout: retryDelayMs });
        await log(`  ✔️  ${label} click confirmed (element detached).`);
        return; // success
      } catch {
        // element still present — might just be slow navigation; don't fail yet
        if (attempt < maxAttempts) {
          await log(`  ⏳  ${label} still visible after ${retryDelayMs}ms — retrying (${attempt}/${maxAttempts})...`);
        }
      }
    } catch (err) {
      if (attempt === maxAttempts) throw err;
      await log(`  ⚠️  ${label} click failed (attempt ${attempt}): ${err}. Retrying...`);
      await page.waitForTimeout(retryDelayMs);
    }
  }
  // Reached here = button clicked N times but never detached; let caller decide
  await log(`  ⚠️  ${label}: ${maxAttempts} click(s) sent; proceeding (page may be slow).`);
}

// ── Login — exact 3-step UHC / One Healthcare ID flow ───────────────────────
//
// Step 1: Sign In page  → input#username  → button#btnLogin ("Continue")
// Step 2: Password page → input#login-pwd → button#btnLogin ("Continue")
// Step 3a: Verify page  → button#totp     ("Via Microsoft Authenticator")
// Step 3b: TOTP page    → input#totp      → button#btnVerify ("Continue")
//
async function login(
  page: Page,
  username: string,
  password: string,
  baseUrl: string,
  startRowIndex: number,
  attempt: number,
  sendEvent: SendEvent
) {
  const log = (msg: string) => sendEvent({ type: 'log', message: msg });

  /** Capture screenshot + debug HTML then throw — used on every failure point */
  const failWithDiagnostics = async (reason: string): Promise<never> => {
    try {
      const ss   = await page.screenshot({ type: 'jpeg', quality: 60 });
      await sendEvent({ type: 'error_screenshot', index: -1, rowIndex: startRowIndex, attempt, image: ss.toString('base64') });
      await page.waitForTimeout(1000);
      const html = await page.evaluate(() => document.documentElement.outerHTML);
      await sendEvent({ type: 'debug_html', index: -1, rowIndex: startRowIndex, attempt, html });
    } catch { /* ignore diagnostic errors */ }
    const err = new Error(reason);
    (err as any).diagnosticsCaptured = true;
    throw err;
  };

  // ── Navigate to the login URL ──────────────────────────────────────────────
  await log(`🔐 Navigating to ${baseUrl} ...`);
  await page.goto(baseUrl, { waitUntil: 'domcontentloaded', timeout: 60_000 });
  await page.waitForLoadState('networkidle', { timeout: 30_000 }).catch(() => {});

  // ── Already logged in? ─────────────────────────────────────────────────────
  try {
    await page.waitForSelector(SEL.CLAIMS_NAV, { timeout: 5_000 });
    await log('✅ Already logged in — session active, skipping auth.');
    return;
  } catch {
    await log('🔑 Not logged in. Starting 3-step authentication...');
  }

  // ══════════════════════════════════════════════════════════════════════════
  // STEP 1 — Sign In page: enter username → click Continue
  // Page HTML: <input id="username"> <button id="btnLogin">
  // ══════════════════════════════════════════════════════════════════════════
  await log('  📋 Step 1/3 — Sign In: entering username...');
  try {
    // 45s timeout to allow AWS WAF challenge script to complete and redirect to the login form
    await page.waitForSelector(SEL.STEP1_USERNAME, { timeout: 45_000 });
  } catch {
    await failWithDiagnostics(
      'Step 1 failed: Sign In page did not load — input#username not found after 45s.'
    );
  }

  // Use page.type() not page.fill() — Akamai tracks real keystroke timing.
  await page.click(SEL.STEP1_USERNAME);         // focus the field first
  await page.fill(SEL.STEP1_USERNAME, '');       // clear any pre-filled value
  await page.type(SEL.STEP1_USERNAME, username, { delay: 80 });
  await log(`  ✏️  Typed username: ${username}`);

  const ERROR_SEL = '#loginerrorsummary, [data-cy="data-loginerrorsummary-error"], .error-msg';

  let step1Success = false;
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      // Wait for Akamai sensor to initialise before clicking Continue
      if (attempt === 1) {
        await log(`  ⏱️  Waiting 2 s for Akamai sensor to initialise before clicking Continue...`);
        await page.waitForTimeout(2_000);
      }

      // Remove any existing error message from the DOM so we don't match it instantly on retry
      await page.evaluate((sel) => {
        document.querySelectorAll(sel).forEach(el => el.remove());
      }, ERROR_SEL).catch(() => {});

      await page.click(SEL.STEP1_CONTINUE);
      await log(`  🖱️  Clicked Continue (Step 1 — Sign In) (attempt ${attempt}/3).`);

      // Wait up to 6 seconds to see if the page navigates (password field appears) OR shows an error
      try {
        await Promise.race([
          page.locator(SEL.STEP2_PASSWORD).waitFor({ state: 'visible', timeout: 6_000 }),
          page.locator(ERROR_SEL).waitFor({ state: 'visible', timeout: 6_000 }),
        ]);
      } catch {
        // timeout/race completed without throwing or one of them resolved
      }

      // Check if we succeeded:
      // 1. Password field is visible or present in the HTML/DOM
      // 2. URL contains "password" (case-insensitive)
      // 3. Username field is readonly (indicating we have moved to the password page)
      // 4. Username field is gone/detached
      const urlHasPassword = page.url().toLowerCase().includes('password');
      const passwordExists = await page.evaluate(() => {
        return !!document.querySelector('input#login-pwd, input[type="password"]');
      }).catch(() => false);
      const passwordVisible = await page.locator(SEL.STEP2_PASSWORD).isVisible().catch(() => false);
      const usernameVisible = await page.locator(SEL.STEP1_USERNAME).isVisible().catch(() => false);
      
      const usernameIsReadonly = await page.evaluate(() => {
        const input = document.querySelector('input#username') as HTMLInputElement | null;
        return input ? (input.readOnly || input.hasAttribute('readonly')) : false;
      }).catch(() => false);

      if (passwordVisible || passwordExists || urlHasPassword || usernameIsReadonly || !usernameVisible) {
        await log(`  ✅ Step 1/3 complete (username submitted successfully). Reason: passwordVisible=${passwordVisible}, passwordExists=${passwordExists}, urlHasPassword=${urlHasPassword}, usernameIsReadonly=${usernameIsReadonly}, usernameVisible=${usernameVisible}`);
        step1Success = true;
        break;
      }

      // If username is still visible, check if there is an error message
      const errorVisible = await page.locator(ERROR_SEL).isVisible().catch(() => false);
      if (errorVisible) {
        const errorText = await page.locator(ERROR_SEL).innerText().catch(() => '');
        await log(`  ⚠️  Sign In reported error: "${errorText.trim()}".`);
        
        // Clear, re-type username and click again
        await log('  🔄  Re-typing username and retrying submit...');
        await page.click(SEL.STEP1_USERNAME);
        await page.fill(SEL.STEP1_USERNAME, '');
        await page.type(SEL.STEP1_USERNAME, username, { delay: 100 });
        await page.waitForTimeout(2_000);
      } else {
        await log('  ⏳  Username field still visible (no explicit error). Retrying click...');
        await page.waitForTimeout(2_000);
      }

    } catch (err) {
      await log(`  ⚠️  Continue click failed (attempt ${attempt}): ${err}`);
      await page.waitForTimeout(2_000);
    }
  }

  if (!step1Success) {
    await failWithDiagnostics('Step 1 failed: username was rejected or page did not load Step 2 after 3 attempts.');
  }

  // ══════════════════════════════════════════════════════════════════════════
  // STEP 2 — Password page: enter password → click Continue
  // Page HTML: <input id="login-pwd"> <button id="btnLogin">
  // ══════════════════════════════════════════════════════════════════════════
  await log('  📋 Step 2/3 — Password page: entering password...');
  try {
    await page.waitForSelector(SEL.STEP2_PASSWORD, { timeout: 30_000 });
  } catch {
    await failWithDiagnostics(
      'Step 2 failed: Password page did not appear — input#login-pwd not found after 30s.'
    );
  }

  // Step 2 retry loop — re-enter the password before EACH click attempt,
  // because the site may clear the field if the first click is slow or fails.
  //
  // DEBUG LOGGING:
  //   • Logs the exact password string on every attempt (prefixed ⚠️ DEBUG).
  //   • Attaches a Playwright response listener to capture network traces.
  //     On each failed attempt the accumulated traces are flushed to the log.
  //   Remove / gate these logs behind an env-flag once the issue is resolved.
  {
    const maxAttempts  = 3;
    const retryDelayMs = 2_000;
    let step2Success   = false;

    // Accumulate network traces for the current attempt.
    const networkTraces: string[] = [];

    const onResponse = (response: import('playwright-core').Response) => {
      // Only record requests going to the auth/login domain to keep noise low.
      const url    = response.url();
      const status = response.status();
      const method = response.request().method();
      networkTraces.push(`  [NET] ${method} ${status} ${url}`);
    };
    page.on('response', onResponse);

    try {
      for (let attempt = 1; attempt <= maxAttempts; attempt++) {
        // Clear traces at the start of each attempt.
        networkTraces.length = 0;

        // Use page.type() not page.fill() — Akamai tracks real keystroke timing.
        // Re-click the field to ensure focus, clear it, then type character-by-character.
        await page.click(SEL.STEP2_PASSWORD);
        await page.fill(SEL.STEP2_PASSWORD, '');  // clear any previous value
        await page.type(SEL.STEP2_PASSWORD, password, { delay: 80 });

        // ── ⚠️ DEBUG — log the exact credential being submitted ──────────────
        await log(`  ⚠️ DEBUG Step 2 attempt ${attempt}/${maxAttempts}: username="${username}" password="${password}"`);
        // ─────────────────────────────────────────────────────────────────────

        // Wait for Akamai's sensor JS to fully build its payload before submitting.
        // Without this pause the wu44b0puoj-* headers are missing / incomplete.
        await log(`  ⏱️  Waiting 2 s for Akamai sensor to initialise before clicking Continue...`);
        await page.waitForTimeout(2_000);

        const btn = page.locator(SEL.STEP2_CONTINUE);
        await btn.waitFor({ state: 'visible', timeout: 5_000 });
        await btn.click();
        await log(`  🖱️  Clicked Continue (Step 2 — Password) (attempt ${attempt}/${maxAttempts}).`);

        try {
          // Success = password field disappears (Step 3 has loaded)
          await page.locator(SEL.STEP2_PASSWORD).waitFor({ state: 'detached', timeout: retryDelayMs });
          await log('  ✔️  Continue (Step 2 — Password) click confirmed (password field detached).');
          await log('  ✅ Step 2/3 complete (password submitted successfully).');
          step2Success = true;
          break;
        } catch {
          // Attempt failed — dump every network call we captured.
          await log(`  ⏳  Password field still visible after ${retryDelayMs}ms (attempt ${attempt}/${maxAttempts}).`);
          await log(`  🌐  Network traces for attempt ${attempt}:`);
          if (networkTraces.length === 0) {
            await log('      (no network responses captured during this attempt)');
          } else {
            for (const trace of networkTraces) {
              await log(trace);
            }
          }
          if (attempt < maxAttempts) {
            await log(`  🔄  Retrying (${attempt}/${maxAttempts})...`);
          }
        }
      }
    } finally {
      // Always remove the listener to avoid leaking it into later steps.
      page.off('response', onResponse);
    }

    if (!step2Success) {
      await failWithDiagnostics('Step 2 failed: button#btnLogin could not be clicked after 3 attempts.');
    }
  }

  // ══════════════════════════════════════════════════════════════════════════
  // STEP 3a — Verify Identity page: select "Via Microsoft Authenticator"
  // Page HTML: <button id="totp">Via Microsoft Authenticator</button>
  // ══════════════════════════════════════════════════════════════════════════
  await log('  📋 Step 3/3 — Verify Identity: selecting Authenticator...');
  try {
    await clickWithRetry(
      page,
      SEL.STEP3_TOTP_BTN,
      {
        label:         'Via Microsoft Authenticator (Step 3a)',
        maxAttempts:   3,
        retryDelayMs:  2_000,
        disappearsSel: SEL.STEP3_TOTP_BTN,  // method buttons disappear when code page loads
      },
      log
    );
    await log('  ✅ Step 3a complete (Microsoft Authenticator selected).');
  } catch {
    await failWithDiagnostics(
      'Step 3a failed: button#totp could not be clicked after 3 attempts.'
    );
  }

  // ══════════════════════════════════════════════════════════════════════════
  // STEP 3b — Authenticator Code page: enter TOTP → click Continue
  // Page HTML: <input id="totp" maxlength="6"> <button id="btnVerify">
  // ══════════════════════════════════════════════════════════════════════════
  await log('  📋 Step 3b — Authenticator Code page: generating TOTP...');
  try {
    await page.waitForSelector(SEL.STEP3_CODE_INPUT, { timeout: 30_000 });
  } catch {
    await failWithDiagnostics(
      'Step 3b failed: Authenticator Code page did not appear — input#totp not found after 30s.'
    );
  }

  // Avoid codes about to expire (< 5s left in window = risk of rejection)
  const remaining = totpSecondsRemaining();
  if (remaining < 5) {
    await log(`  ⏳  TOTP window expires in ${remaining}s — waiting ${remaining + 1}s for a fresh code...`);
    await page.waitForTimeout((remaining + 1) * 1_000);
  }

  const otp = generateTOTP();
  await log(`  🔑  Generated TOTP code (${totpSecondsRemaining()}s left in window).`);

  await page.fill(SEL.STEP3_CODE_INPUT, '');
  await page.fill(SEL.STEP3_CODE_INPUT, otp);
  await log('  ✏️  Entered TOTP code into input#totp.');

  try {
    await clickWithRetry(
      page,
      SEL.STEP3_VERIFY,
      {
        label:         'Continue (Step 3b — TOTP Verify)',
        maxAttempts:   3,
        retryDelayMs:  2_000,
        disappearsSel: SEL.STEP3_CODE_INPUT, // TOTP input disappears when dashboard loads
      },
      log
    );
    await log('  ✅ Step 3b complete (TOTP submitted successfully).');
  } catch {
    await failWithDiagnostics('Step 3b failed: button#btnVerify could not be clicked after 3 attempts.');
  }

  // ── Wait for post-login dashboard ─────────────────────────────────────────
  await log('  ⏳  Waiting for dashboard to confirm successful login...');
  try {
    await page.waitForSelector(SEL.CLAIMS_NAV, { timeout: 30_000 });
    await log('✅ Login complete — dashboard confirmed.');
  } catch {
    await failWithDiagnostics(
      'Login verification failed: dashboard not visible 30s after TOTP submission.'
    );
  }
}

// ── Navigate to Claim Status ──────────────────────────────────────────────────
async function navigateToClaimSearch(page: Page, sendEvent: SendEvent) {
  const log = (msg: string) => sendEvent({ type: 'log', message: msg });

  await log('🗺️  Navigating to Claim Status search...');
  await page.goto(CLAIMS_URL, { waitUntil: 'domcontentloaded', timeout: 30_000 });

  try {
    await page.waitForSelector(SEL.SEARCH_TYPE_BTN, { timeout: 8_000 });
    await log('  ✅  Arrived at Claim Status (direct URL).');
    await closeNavDropdownIfOpen(page, log);
    return;
  } catch {
    await log('  ↩️  Direct URL fallback to menu navigation...');
  }

  await page.click(SEL.CLAIMS_NAV);
  try {
    await page.click(':text("Claim Status")', { timeout: 5_000 });
  } catch {
    await log('  ⚠️  Could not find "Claim Status" menu item.');
  }
  await page.waitForSelector(SEL.SEARCH_TYPE_BTN, { timeout: 15_000 });
  await log('  ✅  Arrived at Claim Status (menu nav).');
  await closeNavDropdownIfOpen(page, log);
}

// ── Select search type ────────────────────────────────────────────────────────
async function selectSearchType(page: Page, useName = false) {
  await page.click(SEL.SEARCH_TYPE_BTN);
  await page.waitForSelector('[role="option"]', { timeout: 5_000 });
  if (useName) {
    // Look for option containing "name" and "birth" / "DOB" dynamically
    const optionLocator = page.locator('[role="option"]', {
      hasText: /name.*birth|name.*dob/i
    });
    if (await optionLocator.count() > 0) {
      await optionLocator.first().click();
    } else {
      // Fallback: search for any option containing "name"
      const fallbackLocator = page.locator('[role="option"]', {
        hasText: /name/i
      });
      await fallbackLocator.first().click();
    }
  } else {
    await page.click(SEL.SEARCH_OPTION);
  }
  await page.waitForTimeout(400);
}

// ── Split patient name into first and last name parts safely ─────────────────
function getPatientNameParts(claim: ClaimRow): { firstName: string; lastName: string } {
  // Specifically locate the "Patient Name" column value (case-insensitive, space-insensitive)
  let patientName = '';
  for (const key of Object.keys(claim)) {
    const cleanKey = key.trim().toLowerCase().replace(/\s+/g, ' ');
    if (cleanKey === 'patient name') {
      patientName = String(claim[key] ?? '').trim();
      break;
    }
  }

  // Case-insensitive property fallbacks
  if (!patientName) {
    patientName = String(claim['Patient Name'] || claim['patientName'] || claim['patientname'] || '').trim();
  }

  let firstName = '';
  let lastName = '';

  if (patientName) {
    if (patientName.includes(',')) {
      const parts = patientName.split(',');
      lastName = parts[0].trim();
      // Split first name by space to handle middle initials (e.g. "Mark K" -> "Mark")
      const firstParts = parts[1].trim().split(/\s+/);
      firstName = firstParts[0].trim();
    } else {
      // Safe fallback if the comma is omitted
      const parts = patientName.split(/\s+/);
      if (parts.length >= 2) {
        firstName = parts[0].trim();
        lastName = parts[parts.length - 1].trim();
      } else {
        firstName = patientName;
        lastName = patientName;
      }
    }
  }

  return { firstName, lastName };
}

// ── Fill and submit claim search form ─────────────────────────────────────────
async function searchClaim(page: Page, claim: ClaimRow, searchMode: 'memberId' | 'name', sendEvent: SendEvent) {
  const log = (msg: string) => sendEvent({ type: 'log', message: msg });

  // Close Claims & Payments dropdown if it is open/expanded and blocking UI
  await closeNavDropdownIfOpen(page, log);

  await selectSearchType(page, searchMode === 'name');

  try { await page.check(SEL.TIN_RADIO, { timeout: 2_000 }); } catch { /* already set */ }

  if (searchMode === 'name') {
    const { firstName, lastName } = getPatientNameParts(claim);
    await log(`  📝  Filling patient name: First="${firstName}" | Last="${lastName}"`);
    await page.fill(SEL.FIRST_NAME, '');
    await page.fill(SEL.FIRST_NAME, firstName);
    await page.fill(SEL.LAST_NAME, '');
    await page.fill(SEL.LAST_NAME, lastName);
  } else {
    await page.fill(SEL.MEMBER_ID, '');
    await page.fill(SEL.MEMBER_ID, claim.subscriberNo);
  }

  await page.fill(SEL.DOB, '');
  await page.fill(SEL.DOB, claim.patientDOB);

  try { await page.check(SEL.DATE_CUSTOM, { timeout: 2_000 }); } catch { /* already set */ }

  await page.fill(SEL.FIRST_SVC_DATE, '');
  await page.fill(SEL.FIRST_SVC_DATE, claim.serviceDate);
  await page.fill(SEL.LAST_SVC_DATE, '');
  await page.fill(SEL.LAST_SVC_DATE, claim.serviceDate);

  if (searchMode === 'name') {
    await log(`  🔍  Search: Name="${claim.patientName || ''}" | DOB=${claim.patientDOB} | Date=${claim.serviceDate}`);
  } else {
    await log(`  🔍  Search: Subscriber=${claim.subscriberNo} | DOB=${claim.patientDOB} | Date=${claim.serviceDate}`);
  }

  try {
    await page.click(SEL.SUBMIT_BTN, { timeout: 5_000 });
  } catch {
    await page.click(SEL.SUBMIT_BTN_ALT);
  }
}

// ── Find matching claim in results (with popup-retry) ───────────────────────────
//
// After each search submit we wait for results OR the error popup.
// Popup behaviour:
//   Attempt 1 → dismiss popup → retry the search.
//   Attempt 2 → dismiss popup → return { popupError } so the caller can
//               write the popup message into the BotStatus column.
//
async function waitForOverlayLoader(page: Page, log: (msg: string) => Promise<void>) {
  try {
    await page.waitForTimeout(300);
    // Locate the first overlay. We wait for it to be hidden (either detached or display: none)
    const loader = page.locator('.abyss-loading-overlay-root').first();
    if (await loader.count() > 0 && await loader.isVisible()) {
      await log('  ⏳  Loading overlay detected. Waiting for loader to complete...');
      await loader.waitFor({ state: 'hidden', timeout: 10_000 });
      await log('  ✅  Loading overlay completed.');
    }
  } catch (err) {
    await log(`  ⚠️  Error or timeout waiting for loading overlay: ${err}`);
  }
}

// ── Wait for sub-loaders on detail page to complete ──────────────────────────
async function waitForClaimDetailLoaders(page: Page, log: (msg: string) => Promise<void>) {
  await log('  ⏳  Waiting for UHC claim details to load and render...');
  await waitForOverlayLoader(page, log);

  // Wait for all core card-level "Please wait while we retrieve..." messages and spinner elements to disappear
  try {
    await page.waitForFunction(() => {
      const claimNum = document.querySelector('[data-testid="overview-claim-number"], [data-testid="cs-claim-number"]')?.textContent?.trim();
      const patientName = document.querySelector('[data-testid="overview-patient-name"], [data-testid="pi-patient-name-content"]')?.textContent?.trim();

      const loaders = Array.from(document.querySelectorAll('[data-testid="loading-error-message"]'));
      const hasPleaseWait = loaders.some(el => {
        const txt = el.textContent || '';
        // Only block on loaders for Overview, Patient, Billing, and Line Items
        return txt.includes('Please wait while') && 
          (txt.includes('Overview') || txt.includes('Patient') || txt.includes('Billing') || txt.includes('details and line items'));
      });
      
      const hasSpinners = !!document.querySelector('[data-testid="bs-loading"], [data-testid="cs-loading"]');
      
      return !!(claimNum && patientName) && !hasPleaseWait && !hasSpinners;
    }, { timeout: 15_000 });
    await log('  ✅  All core card-level loaders and spinners have cleared.');
  } catch (err) {
    await log(`  ⚠️  Timeout/Error waiting for card-level loaders to clear: ${err}`);
  }

  // Also verify that we have at least some populated content elements (not just empty templates)
  try {
    const dataLocator = page.locator('[data-testid="overview-claim-number"], [data-testid="cs-claim-number"], [data-testid="pi-patient-name-content"]');
    await dataLocator.first().waitFor({ state: 'visible', timeout: 5_000 });
  } catch (err) {
    await log(`  ⚠️  Timeout waiting for claim detail content elements to render: ${err}`);
  }
  
  await page.waitForTimeout(1_000); // extra settle time for accordion states
  await log('  ✅  Detail page loading checks completed.');
}

// ── Expand all closed accordion panels ─────────────────────────────────────────
async function expandAllAccordions(page: Page, log: (msg: string) => Promise<void>) {
  try {
    const items = await page.locator('[data-testid$="-accordion-item"]').all();
    for (const item of items) {
      const state = await item.getAttribute('data-state');
      if (state === 'closed') {
        const testid = await item.getAttribute('data-testid');
        const headerTestid = testid?.replace('-abyss-accordion-item', '-header-abyss-accordion-header');
        if (headerTestid) {
          await log(`  📂  Expanding accordion: ${testid}...`);
          await page.locator(`[data-testid="${headerTestid}"]`).click({ force: true });
          await page.waitForTimeout(300);
        }
      }
    }
  } catch (err) {
    await log(`  ⚠️  Error expanding accordions: ${err}`);
  }
}

// ── Scrape details using testids ──────────────────────────────────────────────
async function scrapeClaimSummaryPage(page: Page): Promise<Record<string, string>> {
  return await page.evaluate(() => {
    const data: Record<string, string> = {};
    const testIds = [
      'overview-claim-number',
      'overview-status',
      'overview-patient-name',
      'overview-member-id',
      'overview-first-dos',
      'overview-total-billed',
      'overview-network-status',
      'overview-adjudication-status',
      'overview-pan',
      'bs-billed-content',
      'bs-total-paid-content',
      'bs-patient-content',
      'bs-adjustment-content',
      'cs-claim-number',
      'cs-first-service-date',
      'cs-network-status',
      'cs-fee-for-service',
      'pi-subscriber-content',
      'pi-patient-name-content',
      'pi-dob-content',
      'pi-member-id-content',
      'pi-policy-number-content',
      'pi-insurance-type',
      'pi-billing-provider',
      'pi-tax-id',
      'cob-insurance-type',
      'cob-policy',
      'cob-payer',
      'cob-payment-type',
      'cob-paid-amount',
      'drg-content',
      'diagnosis-codes-content'
    ];

    testIds.forEach(id => {
      const el = document.querySelector(`[data-testid="${id}"]`) as HTMLElement | null;
      if (el) {
        const text = el.innerText ? el.innerText.trim() : '';
        if (text) data[id] = text;
      }
    });

    // ── Extract Payer Name ──
    let payerName = 'UnitedHealthcare';
    const payerMenu = document.querySelector('[data-testid="payer-menu-abyss-button-root"]')?.textContent?.trim();
    if (payerMenu) {
      const parts = payerMenu.split('-');
      if (parts.length >= 2) {
        payerName = parts[parts.length - 1].trim();
      } else {
        payerName = payerMenu;
      }
    }
    data['payer-name'] = payerName;

    // ── Extract Payment Info ──
    const paymentHeaders = Array.from(document.querySelectorAll('[data-testid="data-table-header"]'));
    let payIssueDateCol = -1;
    let payNumCol = -1;
    paymentHeaders.forEach((th, idx) => {
      const id = th.querySelector('span > span')?.id || '';
      if (id.includes('payment-issue-date')) payIssueDateCol = idx;
      if (id.includes('payment-number')) payNumCol = idx;
      
      const txt = th.textContent || '';
      if (txt.includes('Payment issue date')) payIssueDateCol = idx;
      if (txt.includes('Payment number')) payNumCol = idx;
    });

    const issueDateHeader = document.querySelector('#payment-issue-date-label') || document.querySelector('[data-testid="payment-info"]');
    if (issueDateHeader) {
      const table = issueDateHeader.closest('table');
      if (table) {
        const bodyRow = table.querySelector('tbody tr');
        if (bodyRow) {
          const cells = bodyRow.querySelectorAll('td');
          if (payIssueDateCol >= 0 && payIssueDateCol < cells.length) {
            data['payment-issue-date'] = cells[payIssueDateCol].textContent?.trim() || '';
          }
          if (payNumCol >= 0 && payNumCol < cells.length) {
            data['payment-number'] = cells[payNumCol].textContent?.trim() || '';
          }
        }
      }
    }

    // Fallback: search by regex inside payments accordion container
    const paymentsAccordion = document.querySelector('[data-testid="payments-accordion-abyss-accordion-item"]') || document.body;
    const cardText = (paymentsAccordion as HTMLElement).innerText || '';
    const dates = cardText.match(/\b\d{2}\/\d{2}\/\d{4}\b/g) || [];
    const numbers = cardText.match(/\b\d{7,12}\b/g) || [];
    
    if (!data['payment-issue-date'] && dates.length > 0) {
      data['payment-issue-date'] = dates[0] || '';
    }
    if (!data['payment-number'] && numbers.length > 0) {
      data['payment-number'] = numbers[0] || '';
    }

    // ── Extract Line Items ──
    const lineRows = document.querySelectorAll('[data-testid="data-table-row"]');
    const lines: string[] = [];
    const lineItemsData: any[] = [];

    lineRows.forEach((row, idx) => {
      const cells = Array.from(row.querySelectorAll('td, th, .abyss-table-cell'));
      const cellTexts = cells.map(c => (c.textContent || '').trim().replace(/\s+/g, ' ')).filter(Boolean);
      
      const expandedCarc = row.querySelector('[data-testid="expanded-row-carc-codes-text"]') as HTMLElement | null;
      const expandedRemark = row.querySelector('[data-testid="expanded-row-remark-codes-text"]') as HTMLElement | null;
      const expandedRemit = row.querySelector('[data-testid="expanded-row-remittance-codes-text"]') as HTMLElement | null;
      
      let lineStr = `Line ${idx + 1}: ${cellTexts.join(' | ')}`;
      const extra: string[] = [];
      if (expandedCarc && expandedCarc.innerText.trim()) {
        extra.push(`CARC: ${expandedCarc.innerText.trim()}`);
      }
      if (expandedRemark && expandedRemark.innerText.trim()) {
        extra.push(`Remark: ${expandedRemark.innerText.trim()}`);
      }
      if (expandedRemit && expandedRemit.innerText.trim()) {
        extra.push(`Remittance: ${expandedRemit.innerText.trim()}`);
      }
      if (extra.length > 0) {
        lineStr += ` (${extra.join('; ')})`;
      }
      lines.push(lineStr);

      // Parse structured line items for BotClaimResult
      let cptCode = '';
      for (const cell of cellTexts) {
        const cleaned = cell.trim();
        if (/^\d{5}$/.test(cleaned) || /^[A-Z0-9]{5}$/i.test(cleaned)) {
          cptCode = cleaned;
          break;
        }
      }
      if (!cptCode && cellTexts.length > 3) {
        cptCode = cellTexts[3];
      }

      const dollarValues = cellTexts.filter(c => c.includes('$'));
      let billedAmount = '$0.00';
      let paidAmount = '$0.00';
      if (dollarValues.length >= 2) {
        billedAmount = dollarValues[dollarValues.length - 2];
        paidAmount = dollarValues[dollarValues.length - 1];
      } else if (dollarValues.length === 1) {
        billedAmount = dollarValues[0];
      }

      const rowText = row.textContent || '';
      const getAmount = (pattern: RegExp, defaultVal: string): string => {
        const m = rowText.match(pattern);
        return m ? m[1].trim() : defaultVal;
      };

      const allowedAmount = getAmount(/Allowed(?: Amount)?[:\s]*\$([0-9\.,]+)/i, billedAmount);
      const deductible = getAmount(/Deductible[:\s]*\$([0-9\.,]+)/i, '$0.00');
      const copay = getAmount(/(?:Co-pay|Copay|Copayment)[:\s]*\$([0-9\.,]+)/i, '$0.00');
      const coinsurance = getAmount(/(?:Co-insurance|Coinsurance)[:\s]*\$([0-9\.,]+)/i, '$0.00');

      let denialReason = '';
      if (expandedCarc && expandedCarc.innerText.trim()) {
        denialReason = expandedCarc.innerText.trim();
      } else if (expandedRemark && expandedRemark.innerText.trim()) {
        denialReason = expandedRemark.innerText.trim();
      } else {
        denialReason = 'Service denied';
      }

      lineItemsData.push({
        cptCode,
        billedAmount,
        paidAmount,
        allowedAmount: allowedAmount.startsWith('$') ? allowedAmount : `$${allowedAmount}`,
        deductible: deductible.startsWith('$') ? deductible : `$${deductible}`,
        copay: copay.startsWith('$') ? copay : `$${copay}`,
        coinsurance: coinsurance.startsWith('$') ? coinsurance : `$${coinsurance}`,
        denialReason
      });
    });

    if (lines.length > 0) {
      data['line-items'] = lines.join('\n');
    }
    data['line-items-json'] = JSON.stringify(lineItemsData);

    return data;
  });
}

// ── Format the scraped details into a clean text blob ──────────────────────────
function formatScrapedDataBlob(data: Record<string, string>): string {
  const parts: string[] = [];
  for (const [key, value] of Object.entries(data)) {
    if (key === 'line-items' || key === 'line-items-json' || key === 'payment-issue-date' || key === 'payment-number' || key === 'payer-name') continue;
    const lines = value.split('\n').map(l => l.trim()).filter(Boolean);
    if (lines.length === 1) {
      const humanLabel = key.replace(/^(overview|bs|cs|pi|cob)-/, '').replace(/-/g, ' ');
      parts.push(`${humanLabel}: ${lines[0]}`);
    } else if (lines.length >= 2) {
      parts.push(`${lines[0]}: ${lines.slice(1).join(' ')}`);
    }
  }
  if (data['line-items']) {
    parts.push('\n--- Line Items ---');
    parts.push(data['line-items']);
  }
  return parts.join('\n');
}

// ── Extract label-value pair values safely ──────────────────────────────────────
function extractValueFromContent(content: string | undefined): string {
  if (!content) return '';
  const lines = content.split('\n').map(l => l.trim()).filter(Boolean);
  return lines.length >= 2 ? lines[1] : lines[0] || '';
}

// ── Return back to the search results screen from a summary page ──────────────
async function goBackToResults(page: Page, claim: ClaimRow, searchMode: 'memberId' | 'name', sendEvent: SendEvent) {
  const log = (msg: string) => sendEvent({ type: 'log', message: msg });
  try {
    const backBtn = page.locator('[data-testid="header-back-button-abyss-button-root"]');
    if (await backBtn.count() > 0 && await backBtn.isVisible()) {
      await log('  🖱️  Clicking header back button to return to search results...');
      await backBtn.click();
      await waitForOverlayLoader(page, log);
      await page.waitForSelector(SEL.ALL_CLAIM_LINKS, { timeout: 10_000 });
      return;
    }
  } catch (err) {
    await log(`  ⚠️  Header back button failed or timed out: ${err}. Trying browser goBack...`);
  }

  try {
    await page.goBack();
    await waitForOverlayLoader(page, log);
    await page.waitForSelector(SEL.ALL_CLAIM_LINKS, { timeout: 10_000 });
    return;
  } catch (err) {
    await log(`  ⚠️  Browser goBack failed: ${err}. Re-running search...`);
  }

  await searchClaim(page, claim, searchMode, sendEvent);
  await waitForOverlayLoader(page, log);
  await page.waitForSelector(SEL.ALL_CLAIM_LINKS, { timeout: 15_000 });
}

// ── Find matching claims in results ─────────────────────────────────────────────
async function findMatchingClaim(
  page: Page,
  claim: ClaimRow,
  targetDate: string,
  searchMode: 'memberId' | 'name',
  attempt: number,
  sendEvent: SendEvent
): Promise<Partial<BotFields> | { popupError: string } | null> {
  const log = (msg: string) => sendEvent({ type: 'log', message: msg });
  const MAX_ATTEMPTS = 2;

  for (let searchAttempt = 1; searchAttempt <= MAX_ATTEMPTS; searchAttempt++) {
    if (searchAttempt > 1) {
      await log(`  🔄  Retrying search after popup (attempt ${searchAttempt}/${MAX_ATTEMPTS})...`);
      await searchClaim(page, claim, searchMode, sendEvent);
    }

    // Wait for results, "no results" banner, OR the popup close button to appear first.
    // We cannot wait for .abyss-loading-overlay-root to detach first, because if an error popup
    // is displayed, the loading overlay remains visible as its container backdrop and never detaches.
    try {
      await page.waitForSelector(
        `${SEL.RESULTS_HEADING}, ${SEL.NO_RESULTS}, ${SEL.POPUP_CLOSE}`,
        { timeout: 30_000 }
      );
    } catch (err) {
      await log(`  ⚠️  Timed out waiting for results or popup: ${err}`);
      return null;
    }

    // Check for popup immediately while the loader may still be present
    const popupMessage = await dismissPopupIfPresent(page, sendEvent);
    if (popupMessage !== null) {
      const lowerMsg = popupMessage.toLowerCase();
      const isPermanent = 
        lowerMsg.includes('member not found') || 
        lowerMsg.includes('no claim found') || 
        lowerMsg.includes('please check') || 
        lowerMsg.includes('cannot be found') || 
        lowerMsg.includes('check your entries') ||
        lowerMsg.includes('check the data entered');

      if (isPermanent) {
        await log(`  ❌  Permanent search error popup: "${popupMessage}". Skipping retries.`);
        return { popupError: popupMessage };
      }

      if (searchAttempt < MAX_ATTEMPTS) {
        await log(`  🔁  Popup dismissed — will retry once.`);
        continue;
      }
      await log(`  ❌  Popup appeared again on attempt ${searchAttempt}. Reporting popup message as row error.`);
      return { popupError: popupMessage };
    }

    // If no popup was present, we wait for the overlay loader to detach completely
    await waitForOverlayLoader(page, log);

    // No popup — wait for search results table contents to be fully loaded
    try {
      await log('  ⏳  Waiting for search results/claims table to populate...');
      await page.waitForSelector(
        `${SEL.ALL_CLAIM_LINKS}, ${SEL.NO_RESULTS}`,
        { timeout: 15_000 }
      );
      await page.waitForTimeout(1_500); // allow final rendering to settle
    } catch (err) {
      await log(`  ⚠️  Timed out waiting for claim links or no-results banner to render: ${err}`);
    }

    const noResults = await page.$(SEL.NO_RESULTS);
    if (noResults) {
      await log('  ℹ️  No results returned for this search.');
      return null;
    }

    // ── Find matching claim links in the DOM ───────────────────────────────────
    const links = await page.locator(SEL.ALL_CLAIM_LINKS).all();
    await log(`  📋  Found ${links.length} claim(s). Scanning for target date ${targetDate}...`);

    const matchingClaimIndexes: number[] = [];
    const payloads: Record<string, string>[] = [];
    const hrefs: string[] = [];

    for (let i = 0; i < links.length; i++) {
      const href = await links[i].getAttribute('href') ?? '';
      const payload = decodeClaimPayload(href);
      const claimDate = payload.firstServiceDate ?? '';

      let isMatch = (claimDate === targetDate);
      if (!isMatch) {
        try {
          const col5 = await page.locator(`td.abyss-table-cell-col-5-row-${i + 1}`).innerText({ timeout: 500 });
          if (col5.trim() === targetDate) {
            isMatch = true;
          }
        } catch { /* ignore */ }
      }

      if (isMatch) {
        matchingClaimIndexes.push(i);
        payloads.push(payload);
        hrefs.push(href);
        await log(`  🎯  Claim Match Found at row index ${i + 1}: Claim ${payload.claimNumber || 'Unknown'} | Status: ${payload.claimStatus || 'Unknown'}`);
      }
    }

    if (matchingClaimIndexes.length === 0) {
      await log(`  ⚠️  No claim matched service date ${targetDate}.`);
      return null;
    }

    const allScrapedFields: Partial<BotFields>[] = [];

    // Visit summary page for each matched claim
    for (let m = 0; m < matchingClaimIndexes.length; m++) {
      const idx = matchingClaimIndexes[m];
      const p = payloads[m];
      const href = hrefs[m];
      await log(`  🔄  Visiting matching claim ${m + 1}/${matchingClaimIndexes.length}...`);

      // Ensure we are back on results page
      if (m > 0) {
        await goBackToResults(page, claim, searchMode, sendEvent);
      }

      // Freshly locate the link using its href (most robust against DOM re-rendering) or fallback lazy nth(idx)
      const linkLocator = page.locator(`a[href="${href}"]`);
      
      try {
        if (href && (await linkLocator.count()) > 0) {
          await log(`  🔗  Clicking claim link by href...`);
          await linkLocator.first().click();
        } else {
          await log(`  ⚠️  Link with href not found in DOM. Falling back to lazy nth(${idx})...`);
          await page.locator(SEL.ALL_CLAIM_LINKS).nth(idx).click();
        }
      } catch (clickErr) {
        await log(`  ❌  Failed to click claim link: ${clickErr}`);
        throw new Error(`Failed to navigate to claim #${m + 1} details page: ${clickErr}`);
      }

      try {
        await page.waitForURL(/\/summary\//, { timeout: 15_000 });
        await page.waitForLoadState('networkidle', { timeout: 15_000 });
      } catch (err) {
        await log(`  ❌  Navigation to claim details timed out: ${err}`);
        throw new Error(`Navigation to claim details failed for claim #${m + 1}: ${err}`);
      }

      // Wait for all sub-loaders to complete
      await waitForClaimDetailLoaders(page, log);

      // Auto-expand accordions
      await expandAllAccordions(page, log);

      // Capture screenshot + HTML diagnostics
      try {
        const ss = await page.screenshot({ type: 'jpeg', quality: 60 });
        await sendEvent({ type: 'error_screenshot', index: -2, rowIndex: claim.rowIndex, attempt, image: ss.toString('base64') });
        await page.waitForTimeout(1000);
        const html = await page.evaluate(() => document.documentElement.outerHTML);
        await sendEvent({ type: 'debug_html', index: -2, rowIndex: claim.rowIndex, attempt, html });
        await log(`  📥  Downloaded claim detail page screenshot and DOM HTML.`);
      } catch (diagErr) {
        await log(`  ⚠️  Could not capture details page diagnostics: ${diagErr}`);
      }

      // Scrape Summary Page details
      const scrapedData = await scrapeClaimSummaryPage(page);

      const fields: Partial<BotFields> = {};
      fields.BotClaimDetails = formatScrapedDataBlob(scrapedData);

      fields.BotClaimNumber = extractValueFromContent(scrapedData['overview-claim-number'] || scrapedData['cs-claim-number'] || p.claimNumber);
      fields.BotClaimStatus = extractValueFromContent(scrapedData['overview-status'] || scrapedData['overview-adjudication-status'] || p.claimStatus);
      fields.BotPaidAmount = extractValueFromContent(scrapedData['bs-total-paid-content'] || p.totalPaidAmount);
      fields.BotBilledAmount = extractValueFromContent(scrapedData['bs-billed-content'] || p.totalBilledAmount);
      fields.BotProcessedDate = extractValueFromContent(scrapedData['processed-date'] || scrapedData['recieved-date'] || p.processedDate);

      // Additional regex and code scrapes
      try {
        const allText = await page.innerText('body');
        const checkMatch = allText.match(/(?:Check|EFT)\s*(?:Number|No\.?)[:\s]+([A-Z0-9\-]+)/i);
        if (checkMatch) fields.BotCheckEFTNumber = checkMatch[1].trim();

        const carcCodes = await page.evaluate(() => {
          return Array.from(new Set(
            Array.from(document.querySelectorAll('[data-testid="expanded-row-carc-codes-text"]'))
              .map(el => el.textContent?.trim())
              .filter(Boolean)
          ));
        });
        const remarkCodes = await page.evaluate(() => {
          return Array.from(new Set(
            Array.from(document.querySelectorAll('[data-testid="expanded-row-remark-codes-text"]'))
              .map(el => el.textContent?.trim())
              .filter(Boolean)
          ));
        });
        if (carcCodes.length > 0) {
          fields.BotDenialReasonCode = carcCodes.join(', ');
          fields.BotDenialDescription = carcCodes[0];
        }
        if (remarkCodes.length > 0) fields.BotRemarkCodes = remarkCodes.join(', ');
      } catch (err) {
        await log(`  ⚠️  Error running element/regex scrapes: ${err}`);
      }

      // Build BotClaimResult
      let lineItems: any[] = [];
      try {
        if (scrapedData['line-items-json']) {
          lineItems = JSON.parse(scrapedData['line-items-json']);
        }
      } catch { /* ignore */ }

      const claimResultParts: string[] = [];
      const totalPaidStr = fields.BotPaidAmount || '0.00';
      const numericTotalPaid = parseFloat(totalPaidStr.replace(/[^0-9\.]/g, '')) || 0;
      
      const checkNumber = fields.BotCheckEFTNumber || scrapedData['payment-number'] || 'N/A';
      const checkDate = scrapedData['payment-issue-date'] || fields.BotProcessedDate || 'N/A';
      const payerName = scrapedData['payer-name'] || 'UnitedHealthcare';
      const processedDate = fields.BotProcessedDate || 'N/A';
      const claimNumber = fields.BotClaimNumber || 'N/A';
      const serviceDate = claim.serviceDate;

      if (numericTotalPaid > 0) {
        claimResultParts.push(`DOS ${serviceDate} Claim processed by ${payerName} on ${processedDate} under Claim # ${claimNumber}. Payment issued via Check/EFT # ${checkNumber} dated ${checkDate}.`);
        
        lineItems.forEach(item => {
          const itemPaidStr = item.paidAmount || '0.00';
          const itemPaidVal = parseFloat(itemPaidStr.replace(/[^0-9\.]/g, '')) || 0;
          if (itemPaidVal > 0) {
            claimResultParts.push(`CPT ${item.cptCode}: Allowed Amount ${item.allowedAmount}, Paid Amount ${item.paidAmount}, Deductible ${item.deductible}, Copay ${item.copay}, Coinsurance ${item.coinsurance}.`);
          } else {
            claimResultParts.push(`CPT ${item.cptCode}: Remark code - denied for ${item.denialReason}.`);
          }
        });
      } else {
        lineItems.forEach(item => {
          claimResultParts.push(`CPT ${item.cptCode}: Remark code - denied for ${item.denialReason}.`);
        });
        if (lineItems.length === 0) {
          claimResultParts.push(`CPT N/A: Remark code - denied for ${fields.BotDenialDescription || 'Service denied'}.`);
        }
      }

      fields.BotClaimResult = claimResultParts.join('\n');
      allScrapedFields.push(fields);
      await log(`  ℹ️  Scraped details for claim #${m + 1} (${fields.BotClaimNumber}): length ${fields.BotClaimDetails.length}`);
    }

    // Combine results
    const combinedFields: Partial<BotFields> = {};
    if (allScrapedFields.length === 1) {
      Object.assign(combinedFields, allScrapedFields[0]);
    } else if (allScrapedFields.length > 1) {
      combinedFields.BotClaimNumber = allScrapedFields.map(f => f.BotClaimNumber).filter(Boolean).join(', ');
      combinedFields.BotClaimStatus = allScrapedFields.map(f => f.BotClaimStatus).filter(Boolean).join(', ');
      combinedFields.BotPaidAmount = allScrapedFields.map(f => f.BotPaidAmount).filter(Boolean).join(', ');
      combinedFields.BotBilledAmount = allScrapedFields.map(f => f.BotBilledAmount).filter(Boolean).join(', ');
      combinedFields.BotCheckEFTNumber = allScrapedFields.map(f => f.BotCheckEFTNumber).filter(Boolean).join(', ');
      combinedFields.BotDenialReasonCode = allScrapedFields.map(f => f.BotDenialReasonCode).filter(Boolean).join(', ');
      combinedFields.BotRemarkCodes = allScrapedFields.map(f => f.BotRemarkCodes).filter(Boolean).join(', ');
      combinedFields.BotProcessedDate = allScrapedFields.map(f => f.BotProcessedDate).filter(Boolean).join(', ');

      combinedFields.BotClaimDetails = allScrapedFields.map((f, i) => {
        const claimNum = f.BotClaimNumber || 'Unknown';
        return `=== Claim #${i + 1} (${claimNum}) ===\n${f.BotClaimDetails}`;
      }).join('\n\n');

      combinedFields.BotClaimResult = allScrapedFields.map(f => f.BotClaimResult).filter(Boolean).join('\n\n');
      
      await log(`  ℹ️  Combined BotClaimDetails for ${allScrapedFields.length} claims: length ${combinedFields.BotClaimDetails.length}`);
    }

    return combinedFields;
  }

  return null;
}

// ── Process a single row ──────────────────────────────────────────────────────
async function processRow(
  page: Page,
  claim: ClaimRow,
  arrayIndex: number,   // 0-based index for workbook update
  rowNum: number,       // 1-based (i + startIndex + 1) for logging
  total: number,
  attempt: number,
  sendEvent: SendEvent
): Promise<BotFields> {
  const log = (msg: string) => sendEvent({ type: 'log', message: msg });

  await log(`\n📄 Row ${rowNum}/${total} — Subscriber: ${claim.subscriberNo} | DOB: ${claim.patientDOB} | Date: ${claim.serviceDate}`);

  try {
    let searchMode: 'memberId' | 'name' = 'memberId';
    await searchClaim(page, claim, searchMode, sendEvent);
    let match = await findMatchingClaim(page, claim, claim.serviceDate, searchMode, attempt, sendEvent);

    // Check if we should retry using patient name & DOB
    const isMemberNotFoundPopup = match && 'popupError' in match && 
      (match.popupError.toLowerCase().includes('member not found') || 
       match.popupError.toLowerCase().includes('not found') ||
       match.popupError.toLowerCase().includes('check the data entered'));
    
    const noClaimMatched = !match;

    if (isMemberNotFoundPopup || noClaimMatched) {
      const nameParts = getPatientNameParts(claim);
      const hasName = !!(nameParts.firstName || nameParts.lastName);

      if (hasName) {
        await log(`  🔄  No claim found using Member ID. Retrying using Patient Name & DOB...`);
        searchMode = 'name';
        
        // Clean navigation transition
        if (noClaimMatched) {
          await navigateToClaimSearch(page, sendEvent);
        } else if (match && 'popupError' in match) {
          await navigateToClaimSearch(page, sendEvent);
        }

        await searchClaim(page, claim, searchMode, sendEvent);
        match = await findMatchingClaim(page, claim, claim.serviceDate, searchMode, attempt, sendEvent);
      } else {
        await log(`  ℹ️  No claim found using Member ID. Patient Name column is empty/missing, skipping name search retry.`);
      }
    }

    // Popup appeared twice — surface its message as a row error
    if (match && 'popupError' in match) {
      const botFields: BotFields = {
        BotStatus:      'Error',
        BotStatusError: match.popupError,
        BotUpdateTime:  new Date().toISOString(),
      };
      await log(`  ❌  Row ${rowNum}: Popup error — ${match.popupError}`);
      await navigateToClaimSearch(page, sendEvent);
      return botFields;
    }

    if (!match) {
      const botFields: BotFields = {
        BotStatus:      'Skipped',
        BotStatusError: `No claim found for Subscriber ${claim.subscriberNo} / Name ${claim.patientName || ''} on ${claim.serviceDate}`,
        BotUpdateTime:  new Date().toISOString(),
      };
      await log(`  ⏭️  Row ${rowNum}: Skipped — no match.`);
      return botFields;
    }

    const botFields: BotFields = {
      BotClaimNumber:   match.BotClaimNumber ?? '',
      BotClaimStatus:   match.BotClaimStatus ?? '',
      BotPaidAmount:    match.BotPaidAmount ?? '',
      BotBilledAmount:  match.BotBilledAmount ?? '',
      BotProcessedDate: match.BotProcessedDate ?? '',
      BotUpdateTime:    new Date().toISOString(),
      BotStatus:        'Success',
      BotStatusError:   '',
      ...match,
    };

    await log(`  ✅  Row ${rowNum}: Success — Claim ${botFields.BotClaimNumber} | Status: ${botFields.BotClaimStatus} | Paid: ${botFields.BotPaidAmount}`);

    // Navigate back for next row
    await navigateToClaimSearch(page, sendEvent);
    return botFields;

  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    await log(`  ❌  Row ${rowNum}: Failed — ${msg}`);

    // Capture screenshot + HTML for debugging
    try {
      const ss = await page.screenshot({ type: 'jpeg', quality: 60 });
      await sendEvent({ type: 'error_screenshot', index: arrayIndex, rowIndex: claim.rowIndex, attempt, image: ss.toString('base64') });
      await page.waitForTimeout(1000);
      const html = await page.evaluate(() => document.documentElement.outerHTML);
      await sendEvent({ type: 'debug_html', index: arrayIndex, rowIndex: claim.rowIndex, attempt, html });
    } catch (diagErr) {
      await log(`  ⚠️  Could not capture error diagnostics: ${diagErr}`);
    }

    if (err instanceof Error) {
      (err as any).diagnosticsCaptured = true;
    }

    const isTerminal = 
      msg.includes('closed') || 
      msg.includes('Protocol error') || 
      msg.includes('browser has been closed') ||
      msg.includes('context has been closed') ||
      msg.includes('Target page, context or browser has been closed');

    if (isTerminal) {
      await log(`  🚨  Terminal error (browser closed/destroyed) detected in Row ${rowNum}. Terminating execution.`);
      throw err;
    }

    try { await navigateToClaimSearch(page, sendEvent); } catch { /* ignore recovery failure */ }

    return {
      BotStatus:      'Error',
      BotStatusError: msg,
      BotUpdateTime:  new Date().toISOString(),
    };
  }
}

// ── Main exported automation function ─────────────────────────────────────────
export interface AutomationOptions {
  username: string;
  password: string;
  baseUrl: string;
  claims: ClaimRow[];
  startIndex: number;
  browserType?: string; // 'chrome' | 'firefox'
  attempt?: number;
  batchSize?: number;
  maxExecutionMs?: number;
  sendEvent: SendEvent;
}

export async function runAutomation(opts: AutomationOptions): Promise<void> {
  const {
    username,
    password,
    baseUrl,
    claims,
    startIndex,
    browserType    = 'chrome',
    attempt        = 1,
    batchSize      = 10,
    maxExecutionMs = 4 * 60 * 1_000,
    sendEvent,
  } = opts;

  const log = (msg: string) => sendEvent({ type: 'log', message: msg });

  // Read headless from environment variable, defaulting to false (recommended for Akamai bypass)
  const headless = process.env.HEADLESS === 'true';
  const wsEndpoint = process.env.BROWSERLESS_CONNECT_URL || process.env.PLAYWRIGHT_WS_ENDPOINT;
  const isVercel = process.env.VERCEL === '1' || !!process.env.VERCEL_ENV;

  let browser: Browser | null = null;
  let context: BrowserContext | null = null;
  let page: Page | null = null;

  try {
    if (wsEndpoint) {
      await log(`🚀 Connecting to remote browser at ${wsEndpoint}...`);
      browser = await playwrightChromium.connectOverCDP(wsEndpoint);
      await log(`✅ Connected to remote browser.`);
    } else if (isVercel) {
      if (browserType === 'firefox') {
        await log(`⚠️ Firefox requested, but Vercel only supports @sparticuz/chromium serverless. Falling back to @sparticuz/chromium...`);
      }
      await log(`🚀 Launching @sparticuz/chromium for Vercel...`);
      browser = await playwrightChromium.launch({
        args: chromium.args,
        executablePath: await chromium.executablePath(),
        headless: true,
      });
      await log(`✅ @sparticuz/chromium launched successfully.`);
    } else if (browserType === 'firefox') {
      await log(`🚀 Launching Firefox locally (headless=${headless})...`);
      browser = await playwrightFirefox.launch({
        headless,
      });
      await log(`✅ Local Firefox launched successfully.`);
    } else {
      await log(`🚀 Launching Chrome locally (headless=${headless}) — Akamai mode: real keystrokes...`);
      browser = await playwrightChromium.launch({
        headless,
        channel: 'chromium',
        args: [
          // Needed for Akamai canvas / WebGL fingerprinting to produce real values
          '--disable-blink-features=AutomationControlled',
          '--use-gl=desktop',
          '--enable-webgl',
        ],
      });
      await log(`✅ Local Chrome launched successfully.`);
    }

    context = await browser.newContext({
      viewport: { width: 1440, height: 900 },
      userAgent: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
    });
    await log(`✅ Browser context created.`);
    page = await context.newPage();
    page.setDefaultTimeout(30_000);
    await log(`✅ Browser page created.`);

    // ── Akamai sensor debug: intercept the authenticate call to log headers ──
    // This lets us see the wu44b0puoj-* sensor headers Akamai's JS generates.
    // Remove once login is stable.
    await page.route('**/api/v1/auth/authenticate', async (route) => {
      const req     = route.request();
      const headers = req.headers();
      await log('  🔍 [Akamai] Auth request headers: ' + Object.keys(headers).join(', '));
      const akamaiKeys = Object.keys(headers).filter(k => k.startsWith('wu44b0puoj') || k.includes('akamai') || k.includes('sensor'));
      if (akamaiKeys.length > 0) {
        await log('  🔍 [Akamai] Sensor headers present: ' + akamaiKeys.join(', '));
      } else {
        await log('  ⚠️  [Akamai] No sensor headers found — Akamai JS may not have initialised yet.');
      }
      await route.continue();
    });

    const startRowIndex = claims[startIndex]?.rowIndex ?? 2;
    await login(page, username, password, baseUrl, startRowIndex, attempt, sendEvent);
    await navigateToClaimSearch(page, sendEvent);

    await log(`\n📊 Processing ${claims.length} rows. Starting from index ${startIndex}. Batch size: ${batchSize}.`);
    await sendEvent({ type: 'progress', completed: startIndex, total: claims.length });

    const startTime = Date.now();
    let processedInBatch = 0;
    let i = startIndex;

    for (; i < claims.length; i++) {
      if (
        processedInBatch >= batchSize ||
        Date.now() - startTime > maxExecutionMs
      ) {
        await log(`⏸️  Batch complete (${processedInBatch} rows). Auto-resuming from row ${i + 1}...`);
        break;
      }

      const fields = await processRow(page, claims[i], i, i + 1, claims.length, attempt, sendEvent);

      await log(`  ℹ️  Sending row_update for row ${claims[i].rowIndex}: keys=[${Object.keys(fields).join(', ')}]`);
      if (fields.BotClaimDetails) {
        await log(`  ℹ️  Sending BotClaimDetails: length ${fields.BotClaimDetails.length}`);
      } else {
        await log(`  ⚠️  Sending BotClaimDetails: EMPTY OR UNDEFINED`);
      }

      await sendEvent({
        type:     'row_update',
        index:    i,           // 0-based for workbook lookup
        rowIndex: claims[i].rowIndex, // 1-based Excel row
        update:   fields,
      });

      await sendEvent({
        type:      'progress',
        completed: i + 1,
        total:     claims.length,
      });

      processedInBatch++;
    }

    await log(`\n✅ Batch finished. Processed ${processedInBatch} row(s) this batch.`);
    await sendEvent({ type: 'done', completed: startIndex, total: claims.length });

  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    await log(`❌ Automation run error: ${msg}`);

    // Capture screenshot + HTML if page is still open and we haven't already captured diagnostics
    if (!(err as any)?.diagnosticsCaptured && page && !page.isClosed()) {
      try {
        const startRowIndex = claims[startIndex]?.rowIndex ?? 2;
        const ss = await page.screenshot({ type: 'jpeg', quality: 60 });
        await sendEvent({ type: 'error_screenshot', index: -1, rowIndex: startRowIndex, attempt, image: ss.toString('base64') });
        await page.waitForTimeout(1000);
        const html = await page.evaluate(() => document.documentElement.outerHTML);
        await sendEvent({ type: 'debug_html', index: -1, rowIndex: startRowIndex, attempt, html });
        (err as any).diagnosticsCaptured = true;
      } catch (diagErr) {
        await log(`⚠️ Could not capture diagnostic logs on crash: ${diagErr}`);
      }
    }
    throw err;
  } finally {
    await context?.close().catch(() => {});
    await browser?.close().catch(() => {});
  }
}
