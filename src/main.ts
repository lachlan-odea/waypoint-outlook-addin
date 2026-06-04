import "./styles.css";
import { ensureSignedIn } from "./firebase";
import { addProject, loadDesigners } from "./firestore";
import { readCurrentEmail, type EmailContext } from "./email";
import {
  BRANDS,
  CONTENT_TYPES,
  type Designer,
  type Priority,
  type Project,
} from "./types";
import { extractLinks } from "./linkify";

let designers: Designer[] = [];
let email: EmailContext | null = null;

Office.onReady(async () => {
  bindUi();
  await boot();
});

function bindUi() {
  $("#form-reset")!.addEventListener("click", () => {
    if (email) fillForm(email, designers);
  });
  $("#form-submit")!.addEventListener("click", onSubmit);
  $("#f-overview")!.addEventListener("input", updateOverviewPreview);
}

function updateOverviewPreview() {
  const text = getVal("f-overview");
  const preview = document.getElementById("f-overview-preview");
  if (!preview) return;
  const links = extractLinks(text);
  preview.replaceChildren();
  if (links.length === 0) {
    preview.classList.add("empty");
    return;
  }
  preview.classList.remove("empty");
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
    if (i < links.length - 1)
      preview.appendChild(document.createTextNode(" · "));
  });
}

async function boot() {
  showView("loading");
  try {
    await ensureSignedIn();
    [designers, email] = await Promise.all([loadDesigners(), readCurrentEmail()]);
    fillForm(email, designers);
    showView("form");
  } catch (err) {
    showFormError(formatError(err));
    showView("form");
  }
}

function fillForm(ctx: EmailContext, ds: Designer[]) {
  populateAssignees(ds);
  populateBrands();
  populateContentTypes();

  setVal("f-title", ctx.subject || "Untitled brief");
  setVal(
    "f-overview",
    truncate(ctx.bodyText.replace(/​/g, "").trim(), 600) ||
      "Brief sourced from email."
  );
  setVal("f-client", ctx.senderName || ctx.senderEmail);
  setVal("f-commenced", todayISO());
  setVal("f-due", ctx.detectedDueDate ?? "");
  setVal("f-priority", ctx.detectedPriority ?? "Normal");
  setVal("f-brand", ctx.detectedBrand ?? "");
  setVal("f-area", "");
  setVal("f-brief", ctx.firstUrl ?? "");

  // Default assignee: signed-in user from the workspace if their name matches
  // the local part of their email, otherwise leave unassigned.
  const myEmail = Office.context.mailbox.userProfile?.emailAddress?.toLowerCase();
  let meId = "";
  if (myEmail) {
    const handle = myEmail.split("@")[0].replace(/[._]/g, " ");
    const match = ds.find((d) => d.name.toLowerCase().includes(handle));
    if (match) meId = match.id;
  }
  setVal("f-assignee", meId);
  updateOverviewPreview();

  setBanner(null);
}

function populateAssignees(ds: Designer[]) {
  const sel = el<HTMLSelectElement>("f-assignee");
  sel.innerHTML = "";
  sel.appendChild(option("", "— Unassigned —"));
  ds.forEach((d) => sel.appendChild(option(d.id, d.name)));
}

function populateBrands() {
  const sel = el<HTMLSelectElement>("f-brand");
  sel.innerHTML = "";
  sel.appendChild(option("", "Select a brand…"));
  BRANDS.forEach((b) => sel.appendChild(option(b, b)));
}

function populateContentTypes() {
  const sel = el<HTMLSelectElement>("f-area");
  sel.innerHTML = "";
  sel.appendChild(option("", "— select —"));
  CONTENT_TYPES.forEach((p) => sel.appendChild(option(p, p)));
  sel.appendChild(option("__other__", "Other…"));
  sel.addEventListener("change", () => {
    if (sel.value !== "__other__") return;
    const custom = prompt("Content type:");
    if (custom) {
      const opt = option(custom, custom);
      sel.insertBefore(opt, sel.querySelector('option[value="__other__"]'));
      sel.value = custom;
    } else {
      sel.value = "";
    }
  });
}

async function onSubmit() {
  const title = getVal("f-title").trim();
  if (!title) {
    showFormError("Title is required.");
    return;
  }
  const assigneeId = getVal("f-assignee");
  const project: Project = {
    id: `p-${Date.now()}`,
    title,
    overview: getVal("f-overview").trim(),
    client: getVal("f-client").trim(),
    brand: getVal("f-brand").trim(),
    contentType: getVal("f-area") === "__other__" ? "" : getVal("f-area"),
    briefUrl: getVal("f-brief").trim(),
    dueDate: getVal("f-due"),
    priority: getVal("f-priority") as Priority,
    assigneeIds: assigneeId ? [assigneeId] : [],
    milestones: [],
    comments: [],
    createdAt: getVal("f-commenced")
      ? new Date(`${getVal("f-commenced")}T00:00:00Z`).toISOString()
      : new Date().toISOString(),
    source: "outlook",
  };

  setSubmitting(true);
  showFormError(null);
  try {
    await addProject(project);
    setBanner(`✓ Added "${project.title}" to Waypoint.`);
  } catch (err) {
    showFormError(formatError(err));
  } finally {
    setSubmitting(false);
  }
}

function setSubmitting(busy: boolean) {
  const btn = el<HTMLButtonElement>("form-submit");
  btn.disabled = busy;
  btn.textContent = busy ? "Adding…" : "Add project";
}

function showView(name: "loading" | "form") {
  ["loading", "form"].forEach((n) => {
    const node = document.getElementById(`view-${n}`);
    if (!node) return;
    node.classList.toggle("hidden", n !== name);
  });
}

function showFormError(msg: string | null) {
  const node = document.getElementById("form-error")!;
  if (!msg) {
    node.classList.add("hidden");
    node.textContent = "";
    return;
  }
  node.classList.remove("hidden");
  node.textContent = msg;
}

function setBanner(msg: string | null) {
  const node = document.getElementById("form-banner")!;
  if (!msg) {
    node.classList.add("hidden");
    node.textContent = "";
    return;
  }
  node.classList.remove("hidden");
  node.textContent = msg;
}

function $(sel: string): HTMLElement | null {
  return document.querySelector(sel);
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
  if (err instanceof Error) return err.message;
  return String(err);
}
