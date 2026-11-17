import { collection, doc, getDoc, getDocs, setDoc } from "firebase/firestore";
import { auth, db } from "./firebase";
import type { Designer, Project, SocialConfig, Workspace } from "./types";

// Same collections the main app reads and writes (waypoint/src/firestore.ts).
const designersCol = () => collection(db, "designers");
const projectsCol = () => collection(db, "projects");
const workspacesCol = () => collection(db, "workspaces");
const socialConfigDoc = () => doc(collection(db, "socialConfig"), "default");

export async function loadDesigners(): Promise<Designer[]> {
  const snap = await getDocs(designersCol());
  return snap.docs.map((d) => d.data() as Designer);
}

export async function loadWorkspaces(): Promise<Workspace[]> {
  const snap = await getDocs(workspacesCol());
  return snap.docs.map((d) => d.data() as Workspace);
}

// The team-managed lists behind Waypoint's pickers. Null when the document
// hasn't been created yet — the caller falls back to the built-in defaults,
// exactly as the main app does.
export async function loadSocialConfig(): Promise<SocialConfig | null> {
  const snap = await getDoc(socialConfigDoc());
  return snap.exists() ? (snap.data() as SocialConfig) : null;
}

// Stamps the signed-in user's uid as `updatedBy`, as the main app does on
// every project save, so the notification Cloud Functions leave out whoever
// made the change.
export async function addProject(project: Project): Promise<void> {
  const uid = auth.currentUser?.uid;
  await setDoc(doc(projectsCol(), project.id), uid ? { ...project, updatedBy: uid } : project);
}
