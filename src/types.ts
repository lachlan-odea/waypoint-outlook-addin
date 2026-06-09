// Mirrors the PM tool's workspace shape (only the bits the add-in needs).

export type Priority = "Urgent" | "High" | "Normal" | "Low";

export type Designer = {
  id: string;
  name: string;
  initials: string;
  color: string;
  pin: string;
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
};

export type Project = {
  id: string;
  workspaceId: string;
  title: string;
  overview: string;
  client: string;
  brand: string;
  contentType: string;
  briefUrl: string;
  dueDate: string;
  priority: Priority;
  assigneeIds: string[];
  milestones: Milestone[];
  comments: Comment[];
  createdAt: string;
  source?: "manual" | "outlook";
};

export type Workspace = {
  id: string;
  name: string;
};

// Fallback list if Firestore returns no workspaces (shouldn't happen once
// the main Waypoint app has booted, since it seeds these on first load).
// Mirrors waypoint/src/constants.ts SEED_WORKSPACES — keep in sync.
export const SEED_WORKSPACES: Workspace[] = [
  { id: "design", name: "Design" },
  { id: "video", name: "Video" },
  { id: "marketing", name: "Marketing" },
];

export const DEFAULT_WORKSPACE_ID = SEED_WORKSPACES[0].id;

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
