# Waypoint — Outlook add-in

Reads the open Outlook message and creates a project on a Waypoint team board
in one click. Pairs with the [Waypoint React app](https://github.com/lachlan-odea/waypoint)
(version 3.0) and writes to the same Firebase project, so the project appears
on everyone's board immediately.

## What it does

- Adds a **Brief to Waypoint** button to the Outlook ribbon on read messages.
- The task pane pre-fills a project from the email:
  - **Title** ← subject
  - **Overview** ← body (links kept inline, truncated to 600 chars), with the
    detected links listed under it
  - **Client** ← sender name / email
  - **Brand** ← detects CargoWise, Landside, WiseTech Global / WTG, WiseTech
    Academy / WTA, e2open, Blume, BorderWise in the subject or body
  - **Commenced** ← today
  - **Due date** ← best-effort parse of "due by …", "deadline …", ISO dates,
    dd/mm/yyyy
  - **Priority** ← detects "urgent", "asap", "important", "low priority" etc.
  - **URL** ← first http(s) link in the body
  - **Assignees** ← you, when your mailbox email matches a Waypoint profile
- The rest of the form matches Waypoint 3.0's New project form:
  - **Status** — Planning, Active, On hold, Completed or Cancelled (new
    projects start Active, as in the app).
  - **Team** — the boards you can see in Waypoint (Design, Video, Marketing,
    Comms, Events, or whatever is in `/workspaces`), with the same warning the
    app gives when an assignee isn't a member of the chosen team.
  - **Content types** — multi-select chips from Settings → Types & tags in
    Waypoint (`/socialConfig/default`), with an "Other…" escape hatch. Written
    as `contentTypes[]`; the legacy single `contentType` is no longer written.
  - **Assignees** — any number of people, as chips.
  - **Request review** — pick from the designers flagged as reviewers.
    Waypoint's `notifyReviewerAssigned` Cloud Function emails each one a link
    to the project.
- **Add project** writes the project to Firestore flagged `source: "outlook"`,
  then shows a confirmation with the title, team, status and due date, an
  **Open in Waypoint** button that deep-links to the project
  (`?project=<id>`), and a way back to the form to add another project from
  the same email.
- Follows the Waypoint 3.0 design system: the Waypoint logo on the WiseTech
  Blue header band, DM Sans and DM Mono, the `--wtg-*` palette, one 10px
  radius, no shadows. Dark mode follows Outlook desktop's theme or, on the
  web, the browser.

## Architecture

```
Outlook (Win/Mac/Web) ─▶ Office.js ─▶ task pane (this add-in)
                                          │  Firebase Anonymous Auth
                                          ▼
                                 Firestore (wtg-waypoint)  ◀── Waypoint React app
                                  /designers   read
                                  /workspaces  read
                                  /socialConfig/default  read (content types)
                                  /projects    write one document
```

The add-in signs in anonymously on every load (the task pane has no stored
session), reads the team, workspace and Types & tags lists, and writes one
project document. Firestore rules require any signed-in user, so Anonymous
sign-in must stay enabled on the Firebase project. The Firebase web config in
`src/firebase.ts` is a public project identifier, not a secret.

Keep `src/types.ts` in step with `waypoint/src/types.ts` and
`waypoint/src/constants.ts` (`BRANDS`, `SEED_WORKSPACES`, `CONTENT_TYPES`,
`PROJECT_STATUSES`): a field written here lands on the same document the app
renders.

## Local development

Prereqs: Node 20+ (CI uses 22), Outlook desktop or Outlook on the web. No
per-user config: the add-in writes straight to the shared Firebase project
that backs the main Waypoint app.

```bash
npm install
npm run certs:install   # installs a trusted CA cert for localhost, valid for a year
npm run manifest:dev    # writes manifest.dev.xml pointing at https://localhost:3000/waypoint-outlook-addin
npm run dev             # starts the HTTPS dev server on :3000
```

To look at the pane without Outlook, open
`https://localhost:3000/waypoint-outlook-addin/` in a browser: it loads the team
lists from Firestore and shows the empty form with a preview-mode notice, since
there is no message to pre-fill from.

`npm run certs:install` runs Microsoft's `office-addin-dev-certs` tool, which
adds a CA cert to your Windows Trusted Root store. After this Outlook (and your
browser) trust `https://localhost:3000` without any warnings — required for
Outlook on the web to load the task pane during install.

To check the manifest before sideloading:

```bash
npm run manifest:validate
```

### Sideload the add-in

**Outlook on the web** (easiest for testing):
1. Open Outlook on the web, click any message.
2. Click **... → Get Add-ins → My add-ins → Add a custom add-in → Add from file…**
3. Pick `manifest.dev.xml` from this folder.

**Outlook desktop (Windows)**:
1. Go to **File → Manage Add-ins** (opens browser).
2. Same flow as above.

The **Waypoint** group with **Brief to Waypoint** appears on the ribbon when
viewing an email. Click it → the form pre-fills from the message → submit.

### Troubleshooting

**"Add-in installation failed"** in Outlook on the web — almost always one of:
- `npm run dev` isn't running. The server must be up while Outlook fetches the
  manifest URLs at install time.
- The dev cert isn't trusted yet, or has expired (the default is only 30 days;
  `npm run certs:install` issues a one-year cert). Run it, then
  restart the dev server. Confirm by opening `https://localhost:3000/waypoint-outlook-addin/icons/icon-64.png`
  in the same browser; it should load with no warning.
- The manifest has a schema issue. Run `npm run manifest:validate` and fix
  anything it flags.

**"Missing or insufficient permissions"** in the pane — Anonymous sign-in has
been disabled on the Firebase project, or the Firestore rules changed. Both
are managed in the Firebase console for `wtg-waypoint`.

## Deploying to GitHub Pages

Pushing to `main` runs `.github/workflows/deploy.yml`, which:
- Builds the site with `VITE_BASE=/waypoint-outlook-addin/`.
- Writes `manifest.prod.xml` with `BASE_URL=https://<owner>.github.io/<repo>`.
- Publishes everything to Pages.

After a deploy, the production manifest is at
`https://<owner>.github.io/waypoint-outlook-addin/manifest.xml`. Sideload it
in Outlook, or upload it to **Microsoft 365 admin → Integrated apps** for
centralised deployment across the tenant. Outlook caches the task pane; bump
`<Version>` in `manifest.xml` when the manifest itself changes.

## File guide

| File | Purpose |
|---|---|
| `manifest.xml` | Templated Office Add-in manifest. `{{BASE_URL}}` is replaced by `scripts/build-manifest.mjs`. |
| `index.html` | Task pane markup. Loads Office.js, the DM Sans / DM Mono fonts and `src/main.ts`. |
| `commands.html` | Required by the manifest's `FunctionFile`; currently a no-op. |
| `src/main.ts` | UI controller — pre-fills the form, renders the chip pickers, submits or hands off to Waypoint. |
| `src/email.ts` | Reads the current Outlook message and best-effort parses dates, priority and brand. |
| `src/firebase.ts` | Firebase app + anonymous sign-in. |
| `src/firestore.ts` | Reads designers, workspaces and Types & tags; writes the project. |
| `src/linkify.ts` | Mirror of the app's link detection, for the overview's link preview. |
| `src/types.ts` | Mirror of Waypoint's `Project`, `Designer`, `Workspace` types and constants. |
| `src/styles.css` | Waypoint 3.0 design tokens and the pane's styles. |
| `public/waypoint-horizontal-*.svg` | The Waypoint logo lockups, copied from `waypoint/public/logos/`. |
| `scripts/generate-icons.mjs` | Renders the Waypoint app icon to the manifest's PNG sizes (same geometry as `waypoint/scripts/make-icons.mjs`). Runs before `dev` and `build`. |
| `scripts/build-manifest.mjs` | Replaces `{{BASE_URL}}` in `manifest.xml` for dev / prod. |
| `.github/workflows/deploy.yml` | Builds + publishes to GitHub Pages. |

## Privacy / security note

The add-in holds no credentials: it authenticates to Firebase anonymously and
relies on the Firestore rules, which let any signed-in user read the team lists
and write projects. That's the same trust model as the main app for a small
internal team; tighten the rules (and this add-in's auth) before opening
Waypoint to anyone outside it.
