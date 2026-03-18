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

Office.onReady();

// ============================================================
// CONFIG — update these two values
// ============================================================

// URL of your flat HTML signature file
const SIGNATURE_URL = “https://your-domain.pages.dev/signature.html”;

// Name shown in Outlook’s signature picker (cosmetic only)
const SIGNATURE_NAME = “Company Signature”;

// ============================================================
// EVENT HANDLER
// ============================================================

async function onNewMessageCompose(event) {
try {
const signatureHtml = await fetchSignature(SIGNATURE_URL);
await setSignature(signatureHtml);
} catch (err) {
// Fail silently in production — don’t block compose
console.error(“Auto Signature: failed to inject signature.”, err);
} finally {
// Always call completed() or Outlook will hang
event.completed();
}
}

// ============================================================
// HELPERS
// ============================================================

/**

- Fetches the remote HTML signature.
- The server hosting signature.html must allow CORS from null/outlook origins,
- or you can inline the HTML string here as a fallback.
  */
  async function fetchSignature(url) {
  const response = await fetch(url);
  if (!response.ok) {
  throw new Error(`HTTP ${response.status} fetching signature`);
  }
  return await response.text();
  }

/**

- Injects the signature HTML into the current compose item.
- 
- setSignatureAsync is the “proper” signature API — it respects the user’s
- existing Outlook signature settings (won’t overwrite if they have one set).
- 
- Falls back to body.setAsync() if setSignatureAsync isn’t available,
- which works on older Outlook versions but replaces the whole body.
  */
  function setSignature(html) {
  return new Promise((resolve, reject) => {
  const item = Office.context.mailbox.item;
  
  // Preferred: setSignatureAsync (Mailbox 1.10+)
  if (item.body.setSignatureAsync) {
  item.body.setSignatureAsync(
  html,
  {
  coercionType: Office.CoercionType.Html,
  asyncContext: SIGNATURE_NAME,
  },
  (result) => {
  if (result.status === Office.AsyncResultStatus.Succeeded) {
  resolve();
  } else {
  // Fallback: user may have a native signature configured,
  // or API isn’t supported — try body.setAsync instead
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

- Fallback: append signature to the body instead of using the signature API.
- Less “proper” but works on older Outlook versions.
- Appends rather than replaces so existing body content is preserved.
  */
  function setBodyFallback(item, signatureHtml, resolve, reject) {
  item.body.getAsync(Office.CoercionType.Html, (result) => {
  if (result.status !== Office.AsyncResultStatus.Succeeded) {
  reject(new Error(“Could not read email body for fallback.”));
  return;
  }
  
  const existingBody = result.value || “”;
  const newBody = existingBody + signatureHtml;
  
  item.body.setAsync(
  newBody,
  { coercionType: Office.CoercionType.Html },
  (setResult) => {
  if (setResult.status === Office.AsyncResultStatus.Succeeded) {
  resolve();
  } else {
  reject(new Error(“body.setAsync failed: “ + setResult.error.message));
  }
  }
  );
  });
  }

// ============================================================
// REGISTER — Outlook looks for this on the global scope
// ============================================================

Office.actions.associate(“onNewMessageCompose”, onNewMessageCompose);