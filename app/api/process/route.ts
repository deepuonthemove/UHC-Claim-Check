/**
 * app/api/process/route.ts
 *
 * POST endpoint — streams progress via SSE (Server-Sent Events).
 * Matches reference repo pattern: claimRows sent as JSON string from client,
 * not re-parsed from Excel on every batch.
 *
 * FormData fields:
 *   loginExcel  : File  — UHC-Website Details.xlsx
 *   claimRows   : string — JSON.stringify(ClaimRow[])  (first batch + all resumes)
 *   startIndex  : string — '0', '10', '20', ...
 */
import { NextRequest } from 'next/server';
import { readLoginExcel } from '@/lib/excel';
import { runAutomation, type SseEvent } from '@/lib/uhc-automation';

export const runtime    = 'nodejs';
export const dynamic    = 'force-dynamic';
export const maxDuration = 300; // 5 minutes

export async function POST(req: NextRequest) {
  try {
    console.log("📥 Received API process request. Parsing request data...");
    
    // Parse request before establishing the stream so synchronous errors return proper HTTP 500
    const formData  = await req.formData();
    const startIndex = parseInt((formData.get('startIndex') as string) || '0', 10);
    const browserType = (formData.get('browserType') as string) || 'chrome';

    const claimRowsJson = formData.get('claimRows') as string | null;
    if (!claimRowsJson) {
      console.error("❌ Missing claimRows in request");
      return new Response(JSON.stringify({ error: 'Missing claimRows JSON in request.' }), {
        status: 400,
        headers: { 'Content-Type': 'application/json' },
      });
    }
    const claims = JSON.parse(claimRowsJson);

    const loginExcelFile = formData.get('loginExcel') as File | null;
    if (!loginExcelFile) {
      console.error("❌ Missing loginExcel file in request");
      return new Response(JSON.stringify({ error: 'Missing loginExcel file in request.' }), {
        status: 400,
        headers: { 'Content-Type': 'application/json' },
      });
    }

    const credentials = await readLoginExcel(await loginExcelFile.arrayBuffer());
    console.log(`🔑 Credentials parsed for user: ${credentials.username}`);

    const encoder = new TextEncoder();
    const stream = new ReadableStream({
      async start(controller) {

        // Keep-alive ping every second — prevents proxy/Vercel buffering
        const keepAliveInterval = setInterval(() => {
          try { controller.enqueue(encoder.encode(': ping\n\n')); }
          catch { clearInterval(keepAliveInterval); }
        }, 1_000);

        const sendEvent = async (event: SseEvent) => {
          try {
            controller.enqueue(encoder.encode(`data: ${JSON.stringify(event)}\n\n`));
            // Yield to event loop so Node.js actually flushes the socket
            await new Promise(resolve => setTimeout(resolve, 50));
          } catch { /* stream already closed */ }
        };

        const log = (message: string) => sendEvent({ type: 'log', message });

        try {
          // 1. Pad stream start to bypass Cloudflare / Vercel buffering (8 KB)
          await sendEvent({ type: 'padding', message: 'x'.repeat(8192) });

          await log(startIndex > 0
            ? `🔄 Resuming from row ${startIndex + 1} of ${claims.length}...`
            : `📊 Received ${claims.length} claim row(s) to process.`
          );
          await sendEvent({ type: 'progress', completed: startIndex, total: claims.length });
          await log(`🔑 Login: ${credentials.username} → ${credentials.url}`);

          // 2. Run the automation batch
          await runAutomation({
            username:   credentials.username,
            password:   credentials.password,
            baseUrl:    credentials.url,
            claims,
            startIndex,
            browserType,
            batchSize:      10,
            maxExecutionMs: 4 * 60 * 1_000,
            sendEvent,
          });

        } catch (err) {
          const msg = err instanceof Error ? err.message : String(err);
          console.error("❌ Automation execution error:", err);
          await log(`❌ Global error: ${msg}`);
          await sendEvent({ type: 'error', message: msg });
        } finally {
          clearInterval(keepAliveInterval);
          // 'done' is sent by runAutomation's finally block; close the stream here
          controller.close();
        }
      },
    });

    return new Response(stream, {
      headers: {
        'Content-Type':      'text/event-stream; charset=utf-8',
        'Cache-Control':     'no-cache, no-transform',
        'Connection':        'keep-alive',
        'X-Accel-Buffering': 'no',
        'Content-Encoding':  'none',
      },
    });

  } catch (err) {
    console.error("❌ Top-level API process error:", err);
    return new Response(JSON.stringify({ error: err instanceof Error ? err.message : String(err) }), {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
    });
  }
}
