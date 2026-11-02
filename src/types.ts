// Mirrors Waypoint's document shapes (waypoint/src/types.ts), trimmed to the
// fields the add-in reads or writes. Keep in step with the main app: a field
// written here lands on the same /projects document the React app renders.

export type Priority = "Urgent" | "High" | "Normal" | "Low";

// Lifecycle order is Planning → Active → On hold → Completed, plus Cancelled.
// The stored value "paused" is displayed as "On hold" — render it through
// PROJECT_STATUSES, never print the raw value.
export type ProjectStatus =
  | "planning"
  | "active"
  | "paused"
  | "completed"
  | "cancelled";

export const PROJECT_STATUSES: { value: ProjectStatus; label: string }[] = [
  { value: "planning", label: "Planning" },
  { value: "active", label: "Active" },
  { value: "paused", label: "On hold" },
  { value: "completed", label: "Completed" },
  { value: "cancelled", label: "Cancelled" },
];

export type Designer = {
  id: string;
  name: string;
  initials: string;
  color: string;
  email?: string;
  hubId?: string;
  photoUrl?: string;
  // Super users see every team regardless of membership.
  isSuperUser?: boolean;
  // Reviewers appear in the "Request review" picker. Independent from
  // isSuperUser; toggled from Waypoint's Settings by a super user.
  isReviewer?: boolean;
};

export type Milestone = {
  id: string;
  label: string;
  done: boolean;
};

export type Comment = {
  id: string;
  author: string;
  text: string;
  createdAt: string;
  likes?: string[];
  parentId?: string | null;
};

export type Project = {
  id: string;
  workspaceId: string;
  title: string;
  overview: string;
  client: string;
  brand: string;
  // Multi-value since Waypoint 2.x. The legacy single `contentType` is read
  // by the main app for old docs but never written any more, so the add-in
  // writes only `contentTypes`.
  contentTypes: string[];
  briefUrl: string;
  dueDate: string;
  priority: Priority;
  assigneeIds: string[];
  // Designer ids of reviewers asked to look at this project. Adding one here
  // triggers Waypoint's notifyReviewerAssigned Cloud Function, which emails
  // them a link to the project.
  reviewerIds?: string[];
  status?: ProjectStatus;
  archived?: boolean;
  milestones: Milestone[];
  comments: Comment[];
  createdAt: string;
  source?: "manual" | "outlook" | "teams";
};

// A team. An empty / missing `memberIds` means the team is open to everyone;
// a populated list restricts who sees its board (plus super users).
export type Workspace = {
  id: string;
  name: string;
  memberIds?: string[];
  icon?: string;
};

// Team-managed lists from Waypoint's Settings → Types & tags, kept as the
// single Firestore document /socialConfig/default. Only `contentTypes` is
// used here; the rest belong to the Social calendar and campaigns.
export type SocialConfig = {
  contentTypes?: string[];
};

// Fallback list if Firestore returns no workspaces (shouldn't happen once
// the main Waypoint app has booted, since it seeds these on first load).
// Mirrors waypoint/src/constants.ts SEED_WORKSPACES — keep in sync.
export const SEED_WORKSPACES: Workspace[] = [
  { id: "design", name: "Design" },
  { id: "video", name: "Video" },
  { id: "marketing", name: "Marketing" },
  { id: "comms", name: "Comms" },
  { id: "events", name: "Events" },
];

export const DEFAULT_WORKSPACE_ID = SEED_WORKSPACES[0].id;

// Mirrors waypoint/src/constants.ts BRANDS — keep in sync.
export const BRANDS = [
  "CargoWise",
  "CargoWise Landside",
  "WiseTech Global",
  "WiseTech Academy",
  "e2open",
  "Blume",
  "BorderWise",
  "No brand",
] as const;

// The built-in content types. The live list comes from /socialConfig/default
// (Settings → Types & tags in Waypoint); this is what the picker shows when
// that document doesn't exist yet or can't be read. Mirrors
// waypoint/src/constants.ts CONTENT_TYPES.
export const CONTENT_TYPES = [
  "Web",
  "Landing page",
  "Social",
  "Email",
  "Print",
  "Event",
  "Mobile",
  "Video",
  "Presentation",
  "Brand",
  "Illustration",
  "Internal comms",
] as const;

// Where the deployed Waypoint app lives. Used for the "Open in Waypoint"
// link shown after a project is created (`?project=<id>`).
export const WAYPOINT_APP_URL = "https://lachlan-odea.github.io/waypoint/";
