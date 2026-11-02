import "./styles.css";
import { ensureSignedIn } from "./firebase";
import {
  addProject,
  loadDesigners,
  loadSocialConfig,
  loadWorkspaces,
} from "./firestore";
import { hasMailItem, readCurrentEmail, type EmailContext } from "./email";
import {
  BRANDS,
  CONTENT_TYPES,
  DEFAULT_WORKSPACE_ID,
  PROJECT_STATUSES,
  SEED_WORKSPACES,
  WAYPOINT_APP_URL,
  type Designer,
  type Priority,
  type Project,
  type ProjectStatus,
  type Workspace,
} from "./types";
import { extractLinks } from "./linkify";

// ── State ───────────────────────────────────────────────────────────────
let designers: Designer[] = [];
let workspaces: Workspace[] = [];
let contentTypeOptions: string[] = [...CONTENT_TYPES];
let email: EmailContext | null = null;
// The signed-in Outlook user as a Waypoint designer, if their mailbox email
// matches a designer's `email` (or, failing that, their name).
let me: Designer | null = null;

// Multi-value fields live here rather than in a <select multiple>, which is
// near-unusable with a mouse. The chips in the DOM are rendered from these.
let assigneeIds: string[] = [];
let reviewerIds: string[] = [];
let contentTypes: string[] = [];

Office.onReady(() => {
  applyOfficeTheme();
  bindUi();
  void boot();
});

// ── Boot ────────────────────────────────────────────────────────────────
async function boot() {
  showView("loading");
  try {
    await ensureSignedIn();
    // The Types & tags document is optional: if it isn't there (or can't be
    // read) the picker shows the built-in list, like the main app does.
    const [ds, ws, cfg, ctx] = await Promise.all([
      loadDesigners(),
      loadWorkspaces(),
      loadSocialConfig().catch(() => null),
      readCurrentEmail(),
    ]);
    designers = ds;
    workspaces = ws.length > 0 ? ws : SEED_WORKSPACES;
    contentTypeOptions =
      cfg?.contentTypes && cfg.contentTypes.length > 0
        ? cfg.contentTypes
        : [...CONTENT_TYPES];
    email = ctx;
    me = findMe(designers);
    fillForm(ctx);
    if (!hasMailItem()) {
      setBanner(
        "Preview mode: no Outlook message is open, so nothing was pre-filled. " +
          "Inside Outlook the title, overview, client, brand, dates and priority come from the email.",
      );
    }
    showView("form");
  } catch (err) {
    showFormError(formatError(err));
    showView("form");
  }
}

// Match the Outlook user to a designer. Designers carry their sign-in email
// since Waypoint 2, so that's the reliable key; the name heuristic is kept
// for profiles created before the field existed.
function findMe(ds: Designer[]): Designer | null {
  const myEmail = Office.context?.mailbox?.userProfile?.emailAddress?.toLowerCase();
  if (!myEmail) return null;
  const byEmail = ds.find((d) => d.email?.toLowerCase() === myEmail);
  if (byEmail) return byEmail;
  const handle = myEmail.split("@")[0].replace(/[._-]/g, " ");
  return ds.find((d) => d.name.toLowerCase() === handle) ?? null;
}

// Teams the current user would see in Waypoint: open teams, teams they're a
// member of, or everything for a super user (or when we can't tell who they
// are). Never returns an empty list.
function visibleWorkspaces(): Workspace[] {
  const user = me;
  if (!user || user.isSuperUser) return workspaces;
  const mine = workspaces.filter(
    (w) => !w.memberIds || w.memberIds.length === 0 || w.memberIds.includes(user.id),
  );
  return mine.length > 0 ? mine : workspaces;
}

// ── UI wiring ───────────────────────────────────────────────────────────
function bindUi() {
  el("form-reset").addEventListener("click", () => {
    if (email) fillForm(email);
  });
  el("form-submit").addEventListener("click", () => void onSubmit());
  el("f-overview").addEventListener("input", updateOverviewPreview);
  el("f-workspace").addEventListener("change", updateOutsiderHint);

  el<HTMLSelectElement>("f-assignee-add").addEventListener("change", (e) => {
    const sel = e.currentTarget as HTMLSelectElement;
    if (sel.value && !assigneeIds.includes(sel.value)) {
      assigneeIds = [...assigneeIds, sel.value];
      renderAssignees();
    }
    sel.value = "";
  });
  el<HTMLSelectElement>("f-reviewer-add").addEventListener("change", (e) => {
    const sel = e.currentTarget as HTMLSelectElement;
    if (sel.value && !reviewerIds.includes(sel.value)) {
      reviewerIds = [...reviewerIds, sel.value];
      renderReviewers();
    }
    sel.value = "";
  });

  const otherInput = el<HTMLInputElement>("f-type-other");
  otherInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      commitOtherType();
    } else if (e.key === "Escape") {
      otherInput.value = "";
      otherInput.classList.add("hidden");
    }
  });
  otherInput.addEventListener("blur", commitOtherType);
}

function fillForm(ctx: EmailContext) {
  populateBrands();
  populateStatuses();
  populateWorkspaces(visibleWorkspaces());

  setVal("f-title", ctx.subject || "Untitled brief");
  setVal(
    "f-overview",
    truncate(ctx.bodyText.replace(/​/g, "").trim(), 600) ||
      "Brief sourced from email.",
  );
  setVal("f-client", ctx.senderName || ctx.senderEmail);
  setVal("f-commenced", todayISO());
  setVal("f-due", ctx.detectedDueDate ?? "");
  setVal("f-priority", ctx.detectedPriority ?? "Normal");
  // A new project in Waypoint starts Active unless told otherwise.
  setVal("f-status", "active");
  setVal("f-brand", ctx.detectedBrand ?? "");
  setVal("f-brief", ctx.firstUrl ?? "");

  const visible = visibleWorkspaces();
  setVal(
    "f-workspace",
    visible.some((w) => w.id === DEFAULT_WORKSPACE_ID)
      ? DEFAULT_WORKSPACE_ID
      : (visible[0]?.id ?? DEFAULT_WORKSPACE_ID),
  );

  // Default assignee: the Outlook user, if they're on Waypoint.
  assigneeIds = me ? [me.id] : [];
  reviewerIds = [];
  contentTypes = [];
  renderAssignees();
  renderReviewers();
  renderContentTypes();

  updateOverviewPreview();
  updateOutsiderHint();
  setBanner(null);
  showFormError(null);
}

// ── Pickers ─────────────────────────────────────────────────────────────
function populateBrands() {
  const sel = el<HTMLSelectElement>("f-brand");
  sel.replaceChildren(option("", "Select a brand…"));
  BRANDS.forEach((b) => sel.appendChild(option(b, b)));
}

function populateStatuses() {
  const sel = el<HTMLSelectElement>("f-status");
  sel.replaceChildren();
  PROJECT_STATUSES.forEach((s) => sel.appendChild(option(s.value, s.label)));
}

function populateWorkspaces(ws: Workspace[]) {
  const sel = el<HTMLSelectElement>("f-workspace");
  sel.replaceChildren();
  ws.slice()
    .sort((a, b) => a.name.localeCompare(b.name))
    .forEach((w) => sel.appendChild(option(w.id, w.name)));
}

const byName = (a: Designer, b: Designer) => a.name.localeCompare(b.name);

// Selected people render as removable chips; a select beneath adds another.
function renderPeople(
  chipsId: string,
  addId: string,
  pool: Designer[],
  selected: string[],
  onRemove: (id: string) => void,
  emptyLabel: string,
) {
  const chips = el(chipsId);
  chips.replaceChildren();
  selected
    .map((id) => designers.find((d) => d.id === id))
    .filter((d): d is Designer => !!d)
    .forEach((d) => chips.appendChild(personChip(d, () => onRemove(d.id))));

  const add = el<HTMLSelectElement>(addId);
  add.replaceChildren(option("", selected.length === 0 ? emptyLabel : "Add another…"));
  pool
    .filter((d) => !selected.includes(d.id))
    .sort(byName)
    .forEach((d) => add.appendChild(option(d.id, d.name)));
  add.disabled = add.options.length === 1;
}

function renderAssignees() {
  renderPeople(
    "f-assignee-chips",
    "f-assignee-add",
    designers,
    assigneeIds,
    (id) => {
      assigneeIds = assigneeIds.filter((x) => x !== id);
      renderAssignees();
    },
    "Add assignee…",
  );
  updateOutsiderHint();
}

function renderReviewers() {
  const reviewers = designers.filter((d) => d.isReviewer);
  const hint = el("f-reviewer-hint");
  hint.textContent =
    reviewers.length === 0
      ? "No reviewers yet. Mark someone as a reviewer in Waypoint's Settings."
      : "Reviewers get an email with a link to the project.";
  renderPeople(
    "f-reviewer-chips",
    "f-reviewer-add",
    reviewers,
    reviewerIds,
    (id) => {
      reviewerIds = reviewerIds.filter((x) => x !== id);
      renderReviewers();
    },
    "Request a review…",
  );
}

function personChip(d: Designer, onRemove: () => void): HTMLElement {
  const chip = document.createElement("span");
  chip.className = "chip person";
  const avatar = document.createElement("span");
  avatar.className = "chip-avatar";
  avatar.style.background = d.color || "#5D6B82";
  avatar.textContent = d.initials || d.name.slice(0, 2).toUpperCase();
  const name = document.createElement("span");
  name.textContent = d.name;
  const remove = document.createElement("button");
  remove.type = "button";
  remove.className = "chip-remove";
  remove.setAttribute("aria-label", `Remove ${d.name}`);
  remove.appendChild(closeIcon());
  remove.addEventListener("click", onRemove);
  chip.append(avatar, name, remove);
  return chip;
}

// Content types are toggle chips from the team-managed list, plus any
// one-off typed into "Other…". Mirrors ContentTypeField in the main app.
function renderContentTypes() {
  const wrap = el("f-type-chips");
  wrap.replaceChildren();
  const lower = new Set(contentTypes.map((t) => t.toLowerCase()));
  contentTypeOptions.forEach((t) => {
    wrap.appendChild(
      toggleChip(t, lower.has(t.toLowerCase()), () => toggleContentType(t)),
    );
  });
  contentTypes
    .filter((t) => !contentTypeOptions.some((o) => o.toLowerCase() === t.toLowerCase()))
    .forEach((t) => wrap.appendChild(toggleChip(t, true, () => toggleContentType(t), true)));

  const other = document.createElement("button");
  other.type = "button";
  other.className = "chip add-other";
  other.textContent = "+ Other…";
  other.addEventListener("click", () => {
    const input = el<HTMLInputElement>("f-type-other");
    input.classList.remove("hidden");
    input.focus();
  });
  wrap.appendChild(other);
}

function toggleChip(
  label: string,
  active: boolean,
  onClick: () => void,
  custom = false,
): HTMLElement {
  const b = document.createElement("button");
  b.type = "button";
  b.className = `chip toggle${active ? " active" : ""}${custom ? " custom" : ""}`;
  b.setAttribute("aria-pressed", String(active));
  b.textContent = label;
  if (custom) b.title = `Remove ${label}`;
  b.addEventListener("click", onClick);
  return b;
}

function toggleContentType(t: string) {
  const key = t.toLowerCase();
  contentTypes = contentTypes.some((x) => x.toLowerCase() === key)
    ? contentTypes.filter((x) => x.toLowerCase() !== key)
    : [...contentTypes, t];
  renderContentTypes();
}

function commitOtherType() {
  const input = el<HTMLInputElement>("f-type-other");
  const value = input.value.trim();
  if (value && !contentTypes.some((x) => x.toLowerCase() === value.toLowerCase())) {
    contentTypes = [...contentTypes, value];
  }
  input.value = "";
  input.classList.add("hidden");
  renderContentTypes();
}

// Anyone can be assigned regardless of team, but only a team's members get a
// column on its board. Say so when an assignee isn't a member of the chosen
// team (a team with no explicit members is open to everyone).
function updateOutsiderHint() {
  const hint = el("f-workspace-hint");
  const ws = workspaces.find((w) => w.id === getVal("f-workspace"));
  const members = ws?.memberIds ?? [];
  const outsiders =
    members.length === 0
      ? []
      : assigneeIds
          .filter((id) => !members.includes(id))
          .map((id) => designers.find((d) => d.id === id)?.name)
          .filter((n): n is string => !!n);
  if (outsiders.length === 0 || !ws) {
    hint.classList.remove("warn");
    hint.textContent = "Whose board this lands on. Movable later from the project.";
    return;
  }
  hint.classList.add("warn");
  hint.textContent = `${outsiders.join(", ")} ${
    outsiders.length === 1 ? "isn't" : "aren't"
  } in ${ws.name}, so they won't get a column on that board. The project still appears in their My work.`;
}

function updateOverviewPreview() {
  const preview = el("f-overview-preview");
  const links = extractLinks(getVal("f-overview"));
  preview.replaceChildren();
  preview.classList.toggle("empty", links.length === 0);
  if (links.length === 0) return;
  const label = document.createElement("span");
  label.className = "overview-links-label";
  label.textContent = `${links.length} link${links.length === 1 ? "" : "s"} detected ·`;
  preview.appendChild(label);
  links.forEach((l, i) => {
    const a = document.createElement("a");
    a.href = l.href;
    a.target = "_blank";
    a.rel = "noreferrer";
    a.textContent = l.label;
    preview.appendChild(a);
    if (i < links.length - 1) preview.appendChild(document.createTextNode(" · "));
  });
}

// ── Submit ──────────────────────────────────────────────────────────────
function buildProject(): Project | null {
  const title = getVal("f-title").trim();
  if (!title) {
    showFormError("Title is required.");
    return null;
  }
  const commenced = getVal("f-commenced");
  const project: Project = {
    id: `p-${Date.now()}`,
    workspaceId: getVal("f-workspace") || DEFAULT_WORKSPACE_ID,
    title,
    overview: getVal("f-overview").trim(),
    client: getVal("f-client").trim(),
    brand: getVal("f-brand").trim(),
    contentTypes,
    briefUrl: getVal("f-brief").trim(),
    dueDate: getVal("f-due"),
    priority: getVal("f-priority") as Priority,
    status: getVal("f-status") as ProjectStatus,
    assigneeIds,
    milestones: [],
    comments: [],
    createdAt: commenced
      ? new Date(`${commenced}T00:00:00Z`).toISOString()
      : new Date().toISOString(),
    source: "outlook",
  };
  if (reviewerIds.length > 0) project.reviewerIds = reviewerIds;
  return project;
}

async function onSubmit() {
  showFormError(null);
  const project = buildProject();
  if (!project) return;
  setSubmitting(true);
  try {
    await addProject(project);
    setBanner(
      `Added "${project.title}" to Waypoint.`,
      `${WAYPOINT_APP_URL}?project=${encodeURIComponent(project.id)}`,
    );
  } catch (err) {
    showFormError(formatError(err));
  } finally {
    setSubmitting(false);
  }
}

function openExternal(url: string) {
  const ui = Office.context.ui as { openBrowserWindow?: (u: string) => void } | undefined;
  if (ui?.openBrowserWindow) {
    ui.openBrowserWindow(url);
  } else {
    window.open(url, "_blank", "noopener");
  }
}

// ── View helpers ────────────────────────────────────────────────────────
function setSubmitting(busy: boolean) {
  const btn = el<HTMLButtonElement>("form-submit");
  btn.disabled = busy;
  btn.textContent = busy ? "Adding…" : "Add project";
}

function showView(name: "loading" | "form") {
  ["loading", "form"].forEach((n) => {
    document.getElementById(`view-${n}`)?.classList.toggle("hidden", n !== name);
  });
}

function showFormError(msg: string | null) {
  const node = el("form-error");
  node.classList.toggle("hidden", !msg);
  node.textContent = msg ?? "";
}

function setBanner(msg: string | null, href?: string) {
  const node = el("form-banner");
  node.replaceChildren();
  node.classList.toggle("hidden", !msg);
  if (!msg) return;
  node.appendChild(document.createTextNode(msg));
  if (href) {
    node.appendChild(document.createTextNode(" "));
    const a = document.createElement("a");
    a.href = href;
    a.target = "_blank";
    a.rel = "noreferrer";
    a.textContent = "Open in Waypoint";
    a.addEventListener("click", (e) => {
      e.preventDefault();
      openExternal(href);
    });
    node.appendChild(a);
  }
}

// Outlook desktop reports its theme through officeTheme; Outlook on the web
// follows the browser, which the prefers-color-scheme rules in styles.css
// cover. Flip the dark class when the pane's host is dark.
function applyOfficeTheme() {
  const context = Office.context as
    | { officeTheme?: { bodyBackgroundColor?: string } }
    | undefined;
  const bg = context?.officeTheme?.bodyBackgroundColor;
  if (!bg) return;
  const m = bg.match(/^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i);
  if (!m) return;
  const lum =
    (0.2126 * parseInt(m[1], 16) + 0.7152 * parseInt(m[2], 16) + 0.0722 * parseInt(m[3], 16)) /
    255;
  document.documentElement.classList.toggle("dark-mode", lum < 0.5);
  document.documentElement.classList.toggle("light-mode", lum >= 0.5);
}

// The design system's close glyph (Icon "close": 24×24, 1.5px stroke).
function closeIcon(): SVGSVGElement {
  const ns = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(ns, "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("width", "12");
  svg.setAttribute("height", "12");
  svg.setAttribute("fill", "none");
  svg.setAttribute("stroke", "currentColor");
  svg.setAttribute("stroke-width", "1.5");
  svg.setAttribute("stroke-linecap", "round");
  svg.setAttribute("aria-hidden", "true");
  ["M6 6l12 12", "M18 6L6 18"].forEach((d) => {
    const p = document.createElementNS(ns, "path");
    p.setAttribute("d", d);
    svg.appendChild(p);
  });
  return svg;
}

function el<T extends HTMLElement = HTMLElement>(id: string): T {
  const node = document.getElementById(id);
  if (!node) throw new Error(`Missing element #${id}`);
  return node as T;
}
function getVal(id: string): string {
  return el<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>(id).value;
}
function setVal(id: string, value: string) {
  el<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>(id).value = value;
}
function option(value: string, label: string): HTMLOptionElement {
  const o = document.createElement("option");
  o.value = value;
  o.textContent = label;
  return o;
}
function truncate(s: string, max: number): string {
  return s.length > max ? `${s.slice(0, max - 1)}…` : s;
}
function todayISO(): string {
  return new Date().toISOString().slice(0, 10);
}
function formatError(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}
