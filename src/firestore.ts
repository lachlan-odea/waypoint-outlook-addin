import { collection, doc, getDocs, setDoc } from "firebase/firestore";
import { db } from "./firebase";
import type { Designer, Project } from "./types";

const designersCol = () => collection(db, "designers");
const projectsCol = () => collection(db, "projects");

export async function loadDesigners(): Promise<Designer[]> {
  const snap = await getDocs(designersCol());
  return snap.docs.map((d) => d.data() as Designer);
}

export async function addProject(project: Project): Promise<void> {
  await setDoc(doc(projectsCol(), project.id), project);
}
