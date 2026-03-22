# commands.js — Working Annotated Reference

Production-proven event handler for this Outlook add-in.
Confirmed working on New Outlook Mac, New Outlook Windows, Outlook on the web.

## Critical compatibility rules (do not break these)

1. NO `Office.onReady()` — not supported in event-based activation runtimes
2. NO `async/await` — use Promise chains and callbacks
3. Use `var` not `const`/`let` — safer for older JS engines
4. Always call `event.completed()` in BOTH success and error paths
5. `Office.actions.associate()` at the bottom — string must match manifest `FunctionName` exactly
6. `SIGNATURE_URL` must use extensionless URL — Cloudflare Pretty URLs redirect `.html` files (308)

## Working file

```js
// commands.js
// Outlook loads this silently — no UI, no task pane.
//
// HOW IT WORKS:
//   1. User opens a new compose window
//   2. Outlook fires OnNewMessageCompose
//   3. This script fetches signature HTML from a remote URL
//   4. It calls setSignatureAsync() to inject it into the compose window
//   5. event.completed() signals Outlook the handler is done

// ============================================================
// CONFIG
// ============================================================

// Extensionless URL — Cloudflare Pages redirects .html to extensionless (308).
// Microsoft runtimes do not follow redirects — use /signature not /signature.html
var SIGNATURE_URL = "https://your-domain.pages.dev/signature";

// Shown in Outlook's signature picker — cosmetic only
var SIGNATURE_NAME = "Company Signature";

// ============================================================
// EVENT HANDLER
// ============================================================

function onNewMessageCompose(event) {
  fetchSignature(SIGNATURE_URL)
    .then(function(html) {
      return setSignature(html);
    })
    .then(function() {
      event.completed();
    })
    .catch(function(err) {
      // Fail silently — never block the compose window
      console.error("Auto Signature: failed to inject signature.", err);
      // event.completed() MUST be called even on error or Outlook hangs
      event.completed();
    });
}

// ============================================================
// HELPERS
// ============================================================

/**
 * Fetches the remote HTML signature.
 * The hosting server must serve CORS headers (Access-Control-Allow-Origin: *)
 * Cloudflare Pages _headers file handles this.
 */
function fetchSignature(url) {
  return fetch(url).then(function(response) {
    if (!response.ok) {
      throw new Error("HTTP " + response.status + " fetching signature");
    }
    return response.text();
  });
}

/**
 * Injects the signature HTML into the compose item.
 *
 * Primary: setSignatureAsync (Mailbox 1.10+)
 *   - Respects user's existing native Outlook signature settings
 *   - Will not overwrite if user has a native signature configured
 *
 * Fallback: body.setAsync
 *   - Works on older Outlook builds
 *   - Appends to existing body rather than replacing
 */
function setSignature(html) {
  return new Promise(function(resolve, reject) {
    var item = Office.context.mailbox.item;

    if (item.body.setSignatureAsync) {
      item.body.setSignatureAsync(
        html,
        {
          coercionType: Office.CoercionType.Html,
          asyncContext: SIGNATURE_NAME,
        },
        function(result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            // setSignatureAsync failed — try body.setAsync fallback
            setBodyFallback(item, html, resolve, reject);
          }
        }
      );
    } else {
      // Older Outlook builds — setSignatureAsync not available
      setBodyFallback(item, html, resolve, reject);
    }
  });
}

/**
 * Fallback: read existing body, append signature, write back.
 * Less ideal but works on older Outlook builds.
 */
function setBodyFallback(item, signatureHtml, resolve, reject) {
  item.body.getAsync(Office.CoercionType.Html, function(result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      reject(new Error("Could not read email body for fallback."));
      return;
    }

    var existingBody = result.value || "";
    var newBody = existingBody + signatureHtml;

    item.body.setAsync(
      newBody,
      { coercionType: Office.CoercionType.Html },
      function(setResult) {
        if (setResult.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error("body.setAsync failed: " + setResult.error.message));
        }
      }
    );
  });
}

// ============================================================
// REGISTER
// Must be at the bottom of the file.
// String must exactly match FunctionName in manifest <LaunchEvent>.
// ============================================================

Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
```

## Per-user signature resolution

When moving from a single static signature to per-user signatures, resolve
`SIGNATURE_URL` dynamically (e.g. from a small lookup keyed on the user's email)
rather than hardcoding it. See `architecture/SKILL.md` for patterns.
