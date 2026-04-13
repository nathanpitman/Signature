# Outlook Auto Signature Add-in POC

An event-based Office.js add-in that automatically injects a remotely-managed HTML signature whenever a user opens a new compose window in Outlook. No task pane, no user interaction — it just works silently in the background.

Signature content is served from Cloudflare Pages, meaning it can be updated without touching the add-in or re-sideloading the manifest.

-----

## How It Works

1. User opens a new compose window in Outlook
1. Outlook fires the `OnNewMessageCompose` event
1. `commands.js` fetches `signature.html` from Cloudflare Pages
1. The signature is injected via `setSignatureAsync()` (falls back to `body.setAsync()` on older builds)
1. No UI, no button, no task pane

-----

## Repo Structure

```
├── manifest.xml       # Add-in manifest — sideloaded locally into Outlook
├── commands.js        # Event handler — hosted on Cloudflare Pages
├── signature.html     # HTML signature fragment — hosted on Cloudflare Pages
├── _headers           # Cloudflare Pages CORS header config
└── README.md
```

-----

## Requirements

- Outlook for Mac (2023 or later)
- Microsoft 365 account signed into Outlook
- A [Cloudflare Pages](https://pages.cloudflare.com) account (free tier)
- A GitHub account

-----

## Deploy

### 1. Fork or clone this repo

```bash
git clone https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
cd YOUR_REPO_NAME
```

### 2. Customise the signature

Edit `signature.html` with your real name, title, contact details, and logo URL.

### 3. Deploy to Cloudflare Pages

1. Log in to [Cloudflare Dashboard](https://dash.cloudflare.com) → **Workers & Pages** → **Create**
1. Connect your GitHub account and select this repository
1. Set build configuration:
- **Build command**: *(leave empty)*
- **Build output directory**: `/`
1. Click **Save and Deploy**

Your files will be live at `https://YOUR_PROJECT_NAME.pages.dev`.

The `_headers` file automatically sets permissive CORS headers on `commands.js` and `signature.html` — no extra configuration needed.

### 4. Update the manifest

Open `manifest.xml` and replace every instance of `YOUR_DOMAIN` with your Cloudflare Pages URL:

```xml
<bt:Url id="Commands.Url" DefaultValue="https://YOUR_PROJECT_NAME.pages.dev/commands.js"/>
```

Also replace the `<Id>` field with a fresh GUID. On macOS:

```bash
uuidgen
```

Update `commands.js` too:

```js
const SIGNATURE_URL = "https://YOUR_PROJECT_NAME.pages.dev/signature.html";
```

Push the changes — Cloudflare Pages will redeploy automatically.

### 5. Sideload in Outlook for Mac

1. Open Outlook
1. **Insert → Add-ins → My Add-ins → ··· → Add a Custom Add-in → Add from File…**
1. Select your local `manifest.xml`
1. Open a new compose window — the signature will appear automatically

-----

## Updating the Signature

Edit `signature.html` and push to `main`. Cloudflare Pages redeploys in seconds. No manifest changes, no re-sideloading required.

-----

## Troubleshooting

**Signature doesn’t appear after sideloading**

- Open `commands.js` and `signature.html` URLs directly in a browser to confirm they’re live and returning the expected content
- Check that all `YOUR_DOMAIN` placeholders in `manifest.xml` have been replaced
- Confirm the `<Id>` in `manifest.xml` is a valid GUID

**Add-in loads but signature is blank**

- Open browser DevTools on the Cloudflare Pages URL and check for CORS errors
- Verify the `_headers` file is present in the repo root and has deployed correctly

**`setSignatureAsync` has no effect**

- Outlook respects the user’s native signature settings — if a default signature is already configured in Outlook preferences, the API will not overwrite it. Remove the native signature under **Outlook → Preferences → Signatures** and try again.

**Manifest rejected silently by Outlook**

- Validate the XML is well-formed (no unclosed tags, no stray placeholder text)
- Ensure the Cloudflare Pages URLs are HTTPS

-----

## Architecture Notes

This is a proof of concept. The sideloaded manifest approach works for development and internal testing without requiring Microsoft 365 admin access or AppSource publishing. For a production rollout, the manifest would be deployed via the Microsoft 365 admin centre, making the add-in available to all users in an organisation automatically.

-----

## License

MIT
