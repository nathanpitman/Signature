# Troubleshooting Reference

Common errors encountered during this add-in's development and their fixes.

---

## "The VersionOverrides 1.1 section of the manifest is invalid. The resid tag is invalid."

**When**: Sideloading in legacy Outlook desktop add-ins dialog.

**Causes (in order of likelihood)**:

1. **Missing `<Runtimes>` block** — this is required for event-based activation. Without it, Outlook cannot resolve resid references at runtime even if the manifest schema is valid.

2. **Dots in resid IDs** — e.g. `Commands.Html`, `Commands.Js`. Some Outlook versions fail to parse these. Use plain camelCase: `webViewRuntime`, `jsRuntime`.

3. **`<FunctionFile>` present in LaunchEvent-only add-in** — remove it entirely. It is only valid for task pane / ribbon command add-ins.

**Fix**: Ensure the `<Runtimes>` block exists inside `<Host xsi:type="MailHost">`, resid IDs are camelCase, and `<FunctionFile>` is absent.

---

## "Deployment failed" in M365 Admin Center Integrated Apps

**When**: Uploading manifest via Settings → Integrated Apps → Upload custom apps.

**Step 1**: Open browser DevTools → Network tab → retry deployment → find the POST to `fd/addins/api/apps/uploadCustomApp` → read the response body for the actual error code.

**`FailedWriteToExchange`**:
- Known Microsoft backend bug, recurring and regional
- Workaround 1: Try legacy Add-ins page: `admin.microsoft.com/adminportal/home#/Settings/AddIns`
- Workaround 2: Wait and retry — often transient
- Workaround 3: Raise Microsoft support ticket with correlation ID from the error JSON

**`Exception thrown when extracting the given manifest`** or URL-related errors:
- One or more URLs in the manifest are returning non-200 or redirecting
- Check all URLs with `curl -I`:
  ```bash
  curl -I https://your-domain.pages.dev/commands
  curl -I https://your-domain.pages.dev/commands.js
  curl -I https://your-domain.pages.dev/assets/icon-64.png
  curl -I https://your-domain.pages.dev/assets/icon-128.png
  ```
- All must return `HTTP/2 200` — any `308` redirect will cause deployment to fail

---

## Signature not injecting (add-in installed but silent)

**Check 1**: Is `SIGNATURE_URL` using the extensionless URL?
```js
// Wrong — 308 redirect
var SIGNATURE_URL = "https://your-domain.pages.dev/signature.html";
// Correct
var SIGNATURE_URL = "https://your-domain.pages.dev/signature";
```

**Check 2**: Does the signature URL return 200?
```bash
curl -I https://your-domain.pages.dev/signature
```

**Check 3**: Enable runtime logging on Mac:
```bash
defaults write com.microsoft.Outlook CEFRuntimeLoggingFile -string "runtime_logs.txt"
# Log location:
cat ~/Library/Containers/com.microsoft.Outlook/Data/runtime_logs.txt
```
Quit and reopen Outlook after setting this, then compose a new message.

**Check 4**: Is `event.completed()` being called? If not, Outlook silently times out.

---

## "Failed to create role assignment policy" / "Error executing cmdlet"

**When**: Trying to create a new policy in Exchange Admin Center → Roles → User Roles.

**Cause**: The account doesn't have the Role Management role in Exchange Online.

**Fix**: Edit the existing Default Role Assignment Policy rather than creating a new one. Ensure My Custom Apps, My Marketplace Apps, and My ReadWriteMailbox Apps are all checked.

---

## Email bounced: 550 5.7.1 / 554 5.7.1

**When**: Sending from `@yourorg.onmicrosoft.com` to certain recipients.

**Cause**: Some mail servers (notably Apple/iCloud) reject `.onmicrosoft.com` sender addresses as a spam signal.

**Fix**: Add a custom domain to M365 (Settings → Domains → Add domain), configure MX/SPF/DKIM records, and send from your proper domain address.

---

## Cloudflare Pages — 308 redirect on HTML files

**Cause**: Cloudflare Pages "Pretty URLs" feature strips `.html` extensions and redirects.

**Fix**: Use extensionless URLs everywhere in manifest and JS:
- `/commands` not `/commands.html`
- `/signature` not `/signature.html`

Or disable Pretty URLs in Cloudflare Pages project Settings → Pages configuration.

---

## Platform compatibility notes

**Legacy Outlook for Mac**: Do not target. End of support for Exchange Online is October 2026. Already removed from Outlook 16.102+.

**New Outlook for Mac / Windows / Web**: All use the same runtime architecture — webViewRuntime for web, jsRuntime for desktop. Both confirmed working with the patterns in this skill.

**Windows classic Outlook**: Uses V8-based JS runtime. Minimum version for event-based activation: build 2206 (15330). Modern M365 installs will be above this.
