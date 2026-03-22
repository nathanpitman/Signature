---
name: extending-auto-signature
description: >
  Ideas and patterns for extending this Outlook auto-signature POC beyond a single
  static file — per-user signatures, a small admin UI, template rendering. Use this
  skill when planning how to scale the add-in from "one signature for everyone" to
  "a signature per user or team" without over-engineering it.
---

# Extending This POC — Ideas for Rolling Your Own

This repo ships the simplest possible version: one `signature.html` file, fetched by
every user. That's genuinely enough for a lot of small teams. This skill is for when
you want more — different signatures per person, per department, or per role — and
you're deciding how much machinery that actually needs.

See `outlook-addin/SKILL.md` for the add-in implementation this builds on top of.

---

## Start here: do you need more than static files?

Before adding a server, ask whether static files still get you there:

- **One signature for the whole org** → what this repo does out of the box. Done.
- **A handful of variants** (e.g. one per department) → still doable with static
  files: `signature-sales.html`, `signature-support.html`, etc., and a small lookup
  table in `commands.js` keyed off something available locally (see below).
- **Per-user fields** (name, title, phone) → this is where you likely want a tiny
  templating step, described below.

Reach for a real backend only once you need to *edit* signatures without pushing
code — a per-user admin UI, self-service field entry, or non-technical staff
updating content.

---

## Where per-user identity comes from

The add-in can read the current user's identity straight from Office.js, with no
sign-in flow of its own:

```javascript
const email = Office.context.mailbox.userProfile.emailAddress;
const name  = Office.context.mailbox.userProfile.displayName;
```

Both are available at runtime without extra permissions or admin consent. That's
usually enough to key a lookup — e.g. `GET /signature?email=jane@acme.com` — without
needing SSO, Graph API calls, or OAuth on the server side.

---

## A simple per-user pattern

```
Outlook fires OnNewMessageCompose
         │
         ▼
commands.js reads email from Office.js
         │
         ▼
fetch(`https://your-api/signature?email=${email}`)
         │
         ▼
inject the returned HTML via setSignatureAsync()
```

A minimal version of the server side of this is just:

1. A small key/value store or table: `email → template fields` (name, title, phone,
   department, …).
2. A template with placeholder tokens, e.g. `{{name}}`, `{{title}}`.
3. An endpoint that looks up the user's fields, renders the template, and returns
   plain HTML (no JSON wrapper needed — the add-in just injects what it gets back).

Render once and cache the result rather than re-rendering on every compose — a
signature rarely changes between one email and the next, and the add-in already
tolerates a bit of latency on first load.

---

## If fields are missing

Not every user record will be fully populated. A safe default: if a template
references a field that's empty for a given user, drop that line/element from the
rendered output rather than showing a blank placeholder token. Nobody wants
`{{phone}}` printed literally into their signature.

---

## A basic data shape

If you do stand up a small backend, this is roughly the minimum shape that supports
per-user templates cleanly:

```
User record
  email          (lookup key)
  display_name
  job_title
  department
  phone
  template_id    (which signature template applies to this user)

Template
  id
  html           (with {{token}} placeholders)
```

Nothing here requires multi-tenancy, billing, or auth beyond "does this email have a
record" — add those only if your actual use case needs them (e.g. you're building
this for multiple separate companies rather than your own org).

---

## Keeping it static where you can

Whatever you build, keep the *hot path* — the thing Outlook fetches on every compose
— as cheap as possible:

- Serve the rendered HTML from a CDN or static host, not from a server that
  re-renders on every request.
- Re-render only when the underlying template or user record actually changes, and
  cache the result until then.
- Fail soft: if your lookup service is down, better to serve last-known-good HTML
  (or nothing) than to block or error out the compose window.

---

## Related Skills

- `outlook-addin/SKILL.md` — Add-in implementation: manifest, commands.js,
  Cloudflare Pages hosting, and debugging
