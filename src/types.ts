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
  title: string;
  overview: string;
  owner: string;
  client: string;
  brand: string;
  contentType: string;
  briefUrl: string;
  dueDate: string;
  priority: Priority;
  assigneeId: string | null;
  milestones: Milestone[];
  comments: Comment[];
  createdAt: string;
  source?: "manual" | "outlook";
};

export type Workspace = {
  designers: Designer[];
  projects: Project[];
  currentDesignerId: string;
};

export type JsonBinConfig = {
  binId: string;
  apiKey: string;
  accessKey?: string;
};

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
