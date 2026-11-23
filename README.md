# Waypoint — Outlook add-in

Reads the open Outlook message and creates a project on a Waypoint team board
in one click. Pairs with the [Waypoint React app](https://github.com/lachlan-odea/waypoint)
(version 3.0) and writes to the same Firebase project, so the project appears
on everyone's board immediately.

## What it does

- Adds a **Brief to Waypoint** button to the Outlook ribbon on read messages.
- Asks you to sign in with your Waypoint work email and password (with a
  **Forgot password?** link that sends Waypoint's reset email). The session is
  remembered, so this is usually once per device. Only approved Waypoint
  members get the form; a signed-in account that hasn't been approved yet sees
  a short note with a link to Waypoint to request access or check its status.
  The form shows **Signed in as … · Sign out** at the top.
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
  - **Assignees** ← you (your Waypoint profile)
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
- **Add project** writes the project to Firestore flagged `source: "outlook"`
  and stamped `updatedBy: <your uid>`, as the app does on every save,
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
                                          │  Firebase Auth (email + password)
                                          ▼
                                 Firestore (wtg-waypoint)  ◀── Waypoint React app
                                  /designers   read (incl. your own profile)
                                  /workspaces  read
                                  /socialConfig/default  read (content types)
                                  /projects    write one document
```

The add-in signs in with the same email/password accounts as the Waypoint
app. Auth state is kept in browser local storage (falling back to IndexedDB,
then memory if the Office iframe blocks storage), so people stay signed in
between task-pane loads. Once signed in it reads `/designers/{uid}`: the
Firestore rules only let *approved members* (a non-anonymous user with that
profile) read or write, and always let a user read their own profile, so a
missing document means "not approved yet" rather than an error. Members then
get the team, workspace and Types & tags lists, and the add-in writes one
project document. Anonymous sign-in is no longer used; a leftover anonymous
session is signed out on load. The Firebase web API key is read from
`VITE_FIREBASE_API_KEY` at build time (see Local development); it identifies
the project and is not a secret, but it isn't kept in the repo.

Keep `src/types.ts` in step with `waypoint/src/types.ts` and
`waypoint/src/constants.ts` (`BRANDS`, `SEED_WORKSPACES`, `CONTENT_TYPES`,
`PROJECT_STATUSES`): a field written here lands on the same document the app
renders.

## Local development

Prereqs: Node 20+ (CI uses 22), Outlook desktop or Outlook on the web. The
add-in writes straight to the shared Firebase project that backs the main
Waypoint app, and needs that project's web API key in `.env.local`
(git-ignored; the same value Waypoint uses). Deploys take it from the
repository variable `VITE_FIREBASE_API_KEY` (GitHub → Settings → Secrets and
variables → Actions → Variables).

```bash
npm install
cp .env.example .env.local   # then fill in VITE_FIREBASE_API_KEY
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

**"Your Waypoint account hasn't been approved yet"** after signing in — the
account exists but has no `/designers/{uid}` profile. A Waypoint admin
needs to approve it in Waypoint; then reopen the pane.

**"Missing or insufficient permissions"** in the pane — the Firestore rules
changed, or the signed-in member's profile was removed. Rules are managed in
the Firebase console for `wtg-waypoint`.

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
| `src/main.ts` | UI controller — sign-in / not-approved / form views, pre-fills the form, renders the chip pickers, submits. |
| `src/email.ts` | Reads the current Outlook message and best-effort parses dates, priority and brand. |
| `src/firebase.ts` | Firebase app, email/password sign-in, sign-out, password reset and the approved-member check. |
| `src/firestore.ts` | Reads designers, workspaces and Types & tags; writes the project. |
| `src/linkify.ts` | Mirror of the app's link detection, for the overview's link preview. |
| `src/types.ts` | Mirror of Waypoint's `Project`, `Designer`, `Workspace` types and constants. |
| `src/styles.css` | Waypoint 3.0 design tokens and the pane's styles. |
| `public/waypoint-horizontal-*.svg` | The Waypoint logo lockups, copied from `waypoint/public/logos/`. |
| `scripts/generate-icons.mjs` | Renders the Waypoint app icon to the manifest's PNG sizes (same geometry as `waypoint/scripts/make-icons.mjs`). Runs before `dev` and `build`. |
| `scripts/build-manifest.mjs` | Replaces `{{BASE_URL}}` in `manifest.xml` for dev / prod. |
| `.github/workflows/deploy.yml` | Builds + publishes to GitHub Pages. |

## Privacy / security note

The add-in holds no credentials of its own: each person signs in with their
Waypoint account, and Firebase Auth keeps the session in the task pane's
browser storage. Access is enforced by the Firestore rules, which let only
approved members (a non-anonymous user with a `/designers/{uid}` profile) read
the team lists and write projects, the same trust model as the main app.
