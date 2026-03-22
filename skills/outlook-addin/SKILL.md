---
name: email-signatures-app-outlook-addin
description: >
  Implementation knowledge for this Outlook auto-signature add-in. Use this skill whenever
  working on the add-in manifest, commands.js event handler, signature HTML files,
  Cloudflare Pages hosting, M365 deployment, or any aspect of the Office.js
  event-based activation architecture. Also use when debugging add-in installation
  failures, resid errors, URL routing issues, or runtime compatibility problems
  across Outlook clients.
---

# Auto Signature — Outlook Add-in Implementation

This skill captures hard-won knowledge from building and debugging this Outlook
auto-signature add-in. It covers manifest structure, the JavaScript event handler,
static hosting setup, and all the gotchas encountered during development.

See `architecture/SKILL.md` for ideas on extending this beyond a single static file.

---

## Architecture Overview

The add-in uses **event-based activation** (`OnNewMessageCompose`) to silently inject
a signature when a user opens a new compose window. No task pane, no UI — purely
background.

```
Outlook fires OnNewMessageCompose
  → loads commands.js (desktop) or commands.html (web)
  → fetch(SIGNATURE_URL) — pulls HTML from Cloudflare Pages
  → item.body.setSignatureAsync() — injects into compose window
  → event.completed() — signals Outlook handler is done
```

**Static pre-rendering**: Signature HTML is pre-built and hosted statically on
Cloudflare Pages. The add-in fetches it at compose time. No server-side rendering
per email — per-request costs are negligible.

---

## File Structure (Cloudflare Pages repo)

```
/
├── manifest.xml          # Office add-in manifest
├── commands.html         # Runtime host page (loads commands.js via <script>)
├── commands.js           # Event handler — the core logic
├── signature.html        # The actual signature HTML served to Outlook
├── _headers              # Cloudflare Pages CORS headers config
└── assets/
    ├── icon-64.png       # Required by manifest (white cursive on black)
    └── icon-128.png      # High-res version
```

---

## Critical: Cloudflare Pages Pretty URLs

Cloudflare Pages has "Pretty URLs" enabled by default, which causes all `.html`
files to redirect (308) to extensionless URLs:

- `commands.html` → redirects to `/commands`
- `signature.html` → redirects to `/signature`

**Microsoft's deployment validator and the Outlook runtime do not follow redirects.**
All URLs in the manifest and in `commands.js` must use the extensionless form:

```
https://your-domain.pages.dev/commands      ✅
https://your-domain.pages.dev/commands.html ❌ (308 redirect — will fail)
https://your-domain.pages.dev/signature     ✅
https://your-domain.pages.dev/signature.html ❌ (308 redirect — will fail)
```

`.js` files are NOT affected by Pretty URLs — `commands.js` returns 200 as-is.

**Fix**: either disable Pretty URLs in Cloudflare Pages settings, or use extensionless
URLs everywhere (preferred).

---

## Manifest — Key Patterns

See `references/manifest-annotated.md` for the full working manifest with inline
annotations.

### Critical manifest rules

**1. `<Runtimes>` block is required** for event-based activation — without it you
get "resid tag is invalid":

```xml
<Runtimes>
  <Runtime resid="webViewRuntime">
    <Override type="javascript" resid="jsRuntime"/>
  </Runtime>
</Runtimes>
```

- `webViewRuntime` → points to `commands` (the HTML page) — used by Outlook on the web
- `jsRuntime` → points to `commands.js` — used by Outlook desktop via the Override

**2. `resid` IDs must be plain camelCase** — no dots. Dots in IDs cause "resid tag
is invalid" in some Outlook versions:

```xml
<!-- Good -->
<bt:Url id="webViewRuntime" .../>
<bt:Url id="jsRuntime" .../>

<!-- Bad — dots cause resid errors -->
<bt:Url id="Commands.Html" .../>
<bt:Url id="Commands.Js" .../>
```

**3. `<FunctionFile>` must NOT be present** for LaunchEvent-only add-ins. It is only
for task pane/ribbon command add-ins. Including it causes "resid tag is invalid".

**4. `<SourceLocation>` on the `<LaunchEvent>` extension point** should reference the
`webViewRuntime` resid (the HTML page), not the JS file directly.

**5. `<FormSettings>` is required boilerplate** even in event-only mode — include an
`ItemRead` form pointing to the extensionless commands URL.

### Manifest validation

Run locally before deploying:
```bash
npx office-addin-manifest validate manifest.xml
```

The CLI validator checks schema only. The M365 admin centre deployment validator
additionally fetches all URLs — a 404 or redirect on any URL will cause deployment to
fail with a generic "Deployment failed" error.

---

## commands.js — Key Patterns

See `references/commands-annotated.md` for the full working file.

### Critical compatibility rules

**1. Never use `Office.onReady()`** — not supported in event-based activation runtimes
on any platform. Remove it entirely.

**2. Never use `async/await`** — use Promise chains and callbacks instead. The desktop
JS runtime on some Outlook versions uses an older engine that may not support
async/await reliably.

**3. Use `var` not `const`/`let`** — safer for older JS engines.

**4. Always call `event.completed()`** in both success and error paths — if it's not
called, Outlook hangs indefinitely.

**5. Register with `Office.actions.associate()`** at the bottom of the file — this maps
the manifest's `FunctionName` to the actual JS function:

```js
Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
```

The string must exactly match the `FunctionName` attribute in the manifest's
`<LaunchEvent>` element.

### Signature injection API

Use `setSignatureAsync` (Mailbox 1.10+) as the primary method — it respects the user's
existing Outlook signature settings. Fall back to `body.setAsync` for older builds:

```js
if (item.body.setSignatureAsync) {
  // preferred
} else {
  // fallback: getAsync body then setAsync with appended signature
}
```

### CORS requirement

`signature.html` (served at extensionless `/signature`) must be served with CORS headers
that allow Outlook's origin. Cloudflare Pages `_headers` file:

```
/*
  Access-Control-Allow-Origin: *
```

---

## commands.html — Purpose

This file is the runtime host page for Outlook on the web. It has no visible UI — it
simply loads `office.js` and `commands.js`:

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
<script src="commands.js"></script>
```

`office.js` must be loaded before `commands.js`. The file must exist and return 200
(use the extensionless `/commands` URL in the manifest).

---

## M365 Deployment

### Centralised deployment via Integrated Apps

**Path**: M365 Admin Center → Settings → Integrated Apps → Upload custom apps →
Office Add-in → Upload manifest file

**Common failure: `FailedWriteToExchange`**

If deployment fails with this error (visible in browser DevTools → Network tab):
1. Try the legacy Add-ins page: `admin.microsoft.com/adminportal/home#/Settings/AddIns`
2. Wait and retry — it is often transient
3. Raise Microsoft support ticket with the correlation ID from the error JSON

**Common failure: URL-related errors**

Check all URLs in the manifest return 200 with no redirects before retrying:
```bash
curl -I https://your-domain.pages.dev/commands
curl -I https://your-domain.pages.dev/commands.js
curl -I https://your-domain.pages.dev/assets/icon-64.png
curl -I https://your-domain.pages.dev/assets/icon-128.png
```

### User role permissions (for sideloading)

Enable these roles in Exchange Admin Center → Roles → User roles →
Default Role Assignment Policy:
- My Custom Apps
- My Marketplace Apps
- My ReadWriteMailbox Apps

---

## Platform Support Matrix

| Platform                | Status          | Notes                                      |
|-------------------------|-----------------|--------------------------------------------|
| New Outlook for Mac     | ✅ Working      | Uses webViewRuntime (HTML page path)       |
| New Outlook for Windows | ✅ Working      | Uses jsRuntime (JS file path)              |
| Outlook on the web      | ✅ Working      | Same as new Outlook                        |
| Legacy Outlook for Mac  | ⛔ Not targeted | End of support Oct 2026; removed from Exchange Online |
| Outlook on iOS/Android  | Not tested      | Would need mobile LaunchEvent support      |

---

## Per-User Signature Resolution

This POC serves one static signature to all users. If you want a different
signature per person instead of a hardcoded URL, see `architecture/SKILL.md` for
patterns — the short version is: resolve `SIGNATURE_URL` dynamically (e.g. from a
small lookup keyed on the user's email from Office.js) rather than hardcoding it.

---

## References

- `references/manifest-annotated.md` — full working manifest with inline annotations
- `references/commands-annotated.md` — full working commands.js with annotations
- `references/troubleshooting.md` — error messages and fixes
