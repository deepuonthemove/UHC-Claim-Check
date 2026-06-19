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
import { chromium as playwright, type Browser, type BrowserContext, type Page } from 'playwright-core';
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

// ── SSE Event types ──────────────────────────────────────────────────────────
export interface SseEvent {
  type: 'log' | 'progress' | 'row_update' | 'error_screenshot' | 'debug_html' | 'done' | 'error' | 'padding';
  message?: string;
  completed?: number;
  total?: number;
  index?: number;    // 0-based index into claimRows array (for workbook update)
  rowIndex?: number; // 1-based Excel row number
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
    await page.click(SEL.POPUP_CLOSE);
    await page.locator(SEL.POPUP_CLOSE).waitFor({ state: 'detached', timeout: 3_000 });
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
  sendEvent: SendEvent
) {
  const log = (msg: string) => sendEvent({ type: 'log', message: msg });

  /** Capture screenshot + debug HTML then throw — used on every failure point */
  const failWithDiagnostics = async (reason: string): Promise<never> => {
    try {
      const ss   = await page.screenshot({ type: 'jpeg', quality: 60 });
      await sendEvent({ type: 'error_screenshot', index: -1, image: ss.toString('base64') });
      const html = await page.evaluate(() => document.documentElement.outerHTML);
      await sendEvent({ type: 'debug_html', index: -1, html });
    } catch { /* ignore diagnostic errors */ }
    throw new Error(reason);
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
    await page.waitForSelector(SEL.STEP1_USERNAME, { timeout: 15_000 });
  } catch {
    await failWithDiagnostics(
      'Step 1 failed: Sign In page did not load — input#username not found after 15s.'
    );
  }

  // Use page.type() not page.fill() — Akamai tracks real keystroke timing.
  await page.click(SEL.STEP1_USERNAME);         // focus the field first
  await page.fill(SEL.STEP1_USERNAME, '');       // clear any pre-filled value
  await page.type(SEL.STEP1_USERNAME, username, { delay: 80 });
  await log(`  ✏️  Typed username: ${username}`);

  try {
    await clickWithRetry(
      page,
      SEL.STEP1_CONTINUE,
      {
        label:         'Continue (Step 1 — Sign In)',
        maxAttempts:   3,
        retryDelayMs:  2_000,
        disappearsSel: SEL.STEP1_USERNAME,  // username field disappears when Step 2 loads
      },
      log
    );
  } catch {
    await failWithDiagnostics('Step 1 failed: button#btnLogin could not be clicked after 3 attempts.');
  }

  // ══════════════════════════════════════════════════════════════════════════
  // STEP 2 — Password page: enter password → click Continue
  // Page HTML: <input id="login-pwd"> <button id="btnLogin">
  // ══════════════════════════════════════════════════════════════════════════
  await log('  📋 Step 2/3 — Password page: entering password...');
  try {
    await page.waitForSelector(SEL.STEP2_PASSWORD, { timeout: 20_000 });
  } catch {
    await failWithDiagnostics(
      'Step 2 failed: Password page did not appear — input#login-pwd not found after 20s.'
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
    await page.waitForSelector(SEL.STEP3_CODE_INPUT, { timeout: 15_000 });
  } catch {
    await failWithDiagnostics(
      'Step 3b failed: Authenticator Code page did not appear — input#totp not found after 15s.'
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
}

// ── Select search type ────────────────────────────────────────────────────────
async function selectSearchType(page: Page) {
  await page.click(SEL.SEARCH_TYPE_BTN);
  await page.waitForSelector(SEL.SEARCH_OPTION, { timeout: 5_000 });
  await page.click(SEL.SEARCH_OPTION);
  await page.waitForTimeout(400);
}

// ── Fill and submit claim search form ─────────────────────────────────────────
async function searchClaim(page: Page, claim: ClaimRow, sendEvent: SendEvent) {
  const log = (msg: string) => sendEvent({ type: 'log', message: msg });

  await selectSearchType(page);

  try { await page.check(SEL.TIN_RADIO, { timeout: 2_000 }); } catch { /* already set */ }

  await page.fill(SEL.MEMBER_ID, '');
  await page.fill(SEL.MEMBER_ID, claim.subscriberNo);

  await page.fill(SEL.DOB, '');
  await page.fill(SEL.DOB, claim.patientDOB);

  try { await page.check(SEL.DATE_CUSTOM, { timeout: 2_000 }); } catch { /* already set */ }

  await page.fill(SEL.FIRST_SVC_DATE, '');
  await page.fill(SEL.FIRST_SVC_DATE, claim.serviceDate);
  await page.fill(SEL.LAST_SVC_DATE, '');
  await page.fill(SEL.LAST_SVC_DATE, claim.serviceDate);

  await log(`  🔍  Search: Subscriber=${claim.subscriberNo} | DOB=${claim.patientDOB} | Date=${claim.serviceDate}`);

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
async function findMatchingClaim(
  page: Page,
  claim: ClaimRow,
  targetDate: string,
  sendEvent: SendEvent
): Promise<{ href: string; payload: Record<string, string> } | { popupError: string } | null> {
  const log = (msg: string) => sendEvent({ type: 'log', message: msg });
  const MAX_ATTEMPTS = 2;

  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    if (attempt > 1) {
      await log(`  🔄  Retrying search after popup (attempt ${attempt}/${MAX_ATTEMPTS})...`);
      await searchClaim(page, claim, sendEvent);
    }

    // Wait for results, "no results" banner, OR the popup close button
    try {
      await page.waitForSelector(
        `${SEL.RESULTS_HEADING}, ${SEL.NO_RESULTS}, ${SEL.POPUP_CLOSE}`,
        { timeout: 30_000 }
      );
    } catch {
      await log('  ⚠️  Timed out waiting for results or popup.');
      return null;
    }

    // Check for popup
    const popupMessage = await dismissPopupIfPresent(page, sendEvent);
    if (popupMessage !== null) {
      if (attempt < MAX_ATTEMPTS) {
        await log(`  🔁  Popup dismissed — will retry once.`);
        continue;
      }
      await log(`  ❌  Popup appeared again on attempt ${attempt}. Reporting popup message as row error.`);
      return { popupError: popupMessage };
    }

    // No popup — process results normally
    const noResults = await page.$(SEL.NO_RESULTS);
    if (noResults) {
      await log('  ℹ️  No results returned for this search.');
      return null;
    }

    const links = await page.locator(SEL.ALL_CLAIM_LINKS).all();
    await log(`  📋  Found ${links.length} claim(s). Looking for date ${targetDate}...`);

    for (let i = 0; i < links.length; i++) {
      const href    = await links[i].getAttribute('href') ?? '';
      const payload = decodeClaimPayload(href);
      const claimDate = payload.firstServiceDate ?? '';

      if (claimDate === targetDate) {
        await log(`  ✅  Match found: Claim ${payload.claimNumber} | Status: ${payload.claimStatus} | Date: ${claimDate}`);
        return { href, payload };
      }

      try {
        const col5 = await page.locator(`td.abyss-table-cell-col-5-row-${i + 1}`).innerText({ timeout: 1_000 });
        if (col5.trim() === targetDate) {
          await log(`  ✅  Match found via table col5: row ${i + 1}`);
          return { href, payload };
        }
      } catch { /* column may not exist */ }
    }

    await log(`  ⚠️  No claim matched service date ${targetDate}.`);
    return null;
  }

  return null;
}

// ── Scrape detail page for denial/remark codes ────────────────────────────────
async function scrapeDetailPage(
  page: Page,
  href: string,
  sendEvent: SendEvent
): Promise<Partial<BotFields>> {
  const log = (msg: string) => sendEvent({ type: 'log', message: msg });
  const baseUrl = process.env.UHC_URL ?? 'https://secure.uhcprovider.com';
  const url = `${baseUrl}${href}`;

  await log(`  🔗  Loading claim detail page...`);
  await page.goto(url, { waitUntil: 'networkidle', timeout: 30_000 });

  const fields: Partial<BotFields> = {};
  try {
    const cells = await page.locator('[role="cell"]').allInnerTexts();
    fields.BotClaimDetails = cells.filter(t => t.trim()).join(' | ').substring(0, 5000);

    const allText = await page.innerText('body');
    const denialMatch = allText.match(/(?:Denial|Reason)\s*Code[:\s]+([A-Z0-9\-,\s]{1,50})/i);
    if (denialMatch) fields.BotDenialReasonCode = denialMatch[1].trim();

    const remarkMatch = allText.match(/Remark\s*Code[:\s]+([A-Z0-9\-,\s]{1,50})/i);
    if (remarkMatch) fields.BotRemarkCodes = remarkMatch[1].trim();

    const checkMatch = allText.match(/(?:Check|EFT)\s*(?:Number|No\.?)[:\s]+([A-Z0-9\-]+)/i);
    if (checkMatch) fields.BotCheckEFTNumber = checkMatch[1].trim();

    await log(`  📄  Detail page scraped. BotClaimDetails length: ${fields.BotClaimDetails?.length ?? 0} chars.`);
  } catch (err) {
    await log(`  ⚠️  Detail page scrape error: ${err}`);
  }
  return fields;
}

// ── Process a single row ──────────────────────────────────────────────────────
async function processRow(
  page: Page,
  claim: ClaimRow,
  arrayIndex: number,   // 0-based index for workbook update
  rowNum: number,       // 1-based (i + startIndex + 1) for logging
  total: number,
  sendEvent: SendEvent
): Promise<BotFields> {
  const log = (msg: string) => sendEvent({ type: 'log', message: msg });

  await log(`\n📄 Row ${rowNum}/${total} — Subscriber: ${claim.subscriberNo} | DOB: ${claim.patientDOB} | Date: ${claim.serviceDate}`);

  try {
    await searchClaim(page, claim, sendEvent);
    const match = await findMatchingClaim(page, claim, claim.serviceDate, sendEvent);

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
        BotStatusError: `No claim found for Subscriber ${claim.subscriberNo} on ${claim.serviceDate}`,
        BotUpdateTime:  new Date().toISOString(),
      };
      await log(`  ⏭️  Row ${rowNum}: Skipped — no match.`);
      return botFields;
    }


    const p = match.payload;
    const detailFields = await scrapeDetailPage(page, match.href, sendEvent);

    const botFields: BotFields = {
      BotClaimNumber:   p.claimNumber ?? '',
      BotClaimStatus:   p.claimStatus ?? '',
      BotPaidAmount:    p.totalPaidAmount ?? '',
      BotBilledAmount:  p.totalBilledAmount ?? '',
      BotProcessedDate: p.processedDate ?? '',
      BotUpdateTime:    new Date().toISOString(),
      BotStatus:        'Success',
      BotStatusError:   '',
      ...detailFields,
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
      await sendEvent({ type: 'error_screenshot', index: arrayIndex, image: ss.toString('base64') });
      const html = await page.evaluate(() => document.documentElement.outerHTML);
      await sendEvent({ type: 'debug_html', index: arrayIndex, html });
    } catch (diagErr) {
      await log(`  ⚠️  Could not capture error diagnostics: ${diagErr}`);
    }

    // Try to recover navigation
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

  try {
    if (wsEndpoint) {
      await log(`🚀 Connecting to remote browser at ${wsEndpoint}...`);
      browser = await playwright.connectOverCDP(wsEndpoint);
    } else if (isVercel) {
      await log(`🚀 Launching @sparticuz/chromium for Vercel...`);
      browser = await playwright.launch({
        args: chromium.args,
        executablePath: await chromium.executablePath(),
        headless: true,
      });
    } else {
      await log(`🚀 Launching Chrome locally (headless=${headless}) — Akamai mode: real keystrokes...`);
      browser = await playwright.launch({
        headless,
        channel: 'chromium',
        args: [
          // Needed for Akamai canvas / WebGL fingerprinting to produce real values
          '--disable-blink-features=AutomationControlled',
          '--use-gl=desktop',
          '--enable-webgl',
        ],
      });
    }

    context = await browser.newContext({
      viewport: { width: 1440, height: 900 },
      userAgent: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
    });
    const page = await context.newPage();
    page.setDefaultTimeout(30_000);

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

    await login(page, username, password, baseUrl, sendEvent);
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

      const fields = await processRow(page, claims[i], i, i + 1, claims.length, sendEvent);

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

  } finally {
    await context?.close().catch(() => {});
    await browser?.close().catch(() => {});
  }
}
