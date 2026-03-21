// commands.js
// Hosted remotely at https://your-domain.pages.dev/commands.js
// Outlook loads this silently — no UI, no task pane.
//
// HOW IT WORKS:
//   1. User opens a new compose window
//   2. Outlook fires OnNewMessageCompose
//   3. This script fetches your signature HTML from a remote URL
//   4. It calls setSignatureAsync() to inject it
//   5. event.completed() signals Outlook the handler is done
//
// REQUIREMENTS:
//   - Outlook for Mac with a Microsoft 365 account
//   - Outlook version that supports event-based add-ins (2023+)
//   - The signature URL must be HTTPS and CORS-accessible
//
// COMPATIBILITY:
//   - Office.onReady() is intentionally omitted — legacy Outlook Mac's
//     JS runtime does not support it in event-based activation mode.
//   - async/await is intentionally avoided for the same reason.
//     All async work uses Promise chains and callbacks instead.

// ============================================================
// CONFIG — update these two values
// ============================================================

// URL of your flat HTML signature file.
// Cloudflare Pages serves signature.html at /signature (extensionless).
var SIGNATURE_URL = "https://your-domain.pages.dev/signature";

// Name shown in Outlook's signature picker (cosmetic only)
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
      // Fail silently in production — don't block compose
      console.error("Auto Signature: failed to inject signature.", err);
      event.completed();
    });
}

// ============================================================
// HELPERS
// ============================================================

/**
 * Fetches the remote HTML signature.
 * The server hosting signature must allow CORS from null/outlook origins.
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
 * Injects the signature HTML into the current compose item.
 *
 * setSignatureAsync is the "proper" signature API — it respects the user's
 * existing Outlook signature settings (won't overwrite if they have one set).
 *
 * Falls back to body.setAsync() if setSignatureAsync is not available,
 * which works on older Outlook versions but replaces the whole body.
 */
function setSignature(html) {
  return new Promise(function(resolve, reject) {
    var item = Office.context.mailbox.item;

    // Preferred: setSignatureAsync (Mailbox 1.10+)
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
            // Fallback: user may have a native signature configured,
            // or API is not supported — try body.setAsync instead
            setBodyFallback(item, html, resolve, reject);
          }
        }
      );
    } else {
      // Fallback for older Outlook builds
      setBodyFallback(item, html, resolve, reject);
    }
  });
}

/**
 * Fallback: append signature to the body instead of using the signature API.
 * Less "proper" but works on older Outlook versions.
 * Appends rather than replaces so existing body content is preserved.
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
// REGISTER — Outlook looks for this on the global scope
// ============================================================

Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
