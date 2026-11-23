import { initializeApp } from "firebase/app";
import {
  browserLocalPersistence,
  indexedDBLocalPersistence,
  inMemoryPersistence,
  initializeAuth,
  onAuthStateChanged,
  sendPasswordResetEmail,
  signInWithEmailAndPassword,
  signOut as fbSignOut,
  type User,
} from "firebase/auth";
import { doc, getDoc, initializeFirestore } from "firebase/firestore";

// The web API key comes from VITE_FIREBASE_API_KEY at build time (a GitHub
// Actions variable for deploys, .env.local for local runs; see the README),
// so it isn't kept in this repo.
const apiKey = import.meta.env.VITE_FIREBASE_API_KEY as string | undefined;
if (!apiKey) {
  throw new Error("VITE_FIREBASE_API_KEY is not set. See the README.");
}

// Same wtg-waypoint project as the main Waypoint app. The web `apiKey` is a
// public project identifier (not a secret); access is gated by Firestore
// rules, which only let approved members in: a signed-in email/password user
// with a profile at /designers/{uid}.
const firebaseConfig = {
  apiKey,
  authDomain: "wtg-waypoint.firebaseapp.com",
  projectId: "wtg-waypoint",
  storageBucket: "wtg-waypoint.firebasestorage.app",
  messagingSenderId: "77272476016",
  appId: "1:77272476016:web:7ac5bda22782a636ba87b5",
};

export const firebaseApp = initializeApp(firebaseConfig);

// IndexedDB persistence isn't worth it here: the add-in only does a single
// read (designers list) and a single write (the new project) per session.
export const db = initializeFirestore(firebaseApp, {
  ignoreUndefinedProperties: true,
});

// Keep people signed in between task-pane loads. Firebase tries each
// persistence in order and silently settles on the first one the Office
// iframe allows; if storage is blocked entirely it falls back to memory, and
// the user simply signs in again next time the pane opens.
export const auth = initializeAuth(firebaseApp, {
  persistence: [browserLocalPersistence, indexedDBLocalPersistence, inMemoryPersistence],
});

// Fires immediately with the restored user (or null), then on every change.
export function observeAuth(cb: (user: User | null) => void): () => void {
  return onAuthStateChanged(auth, cb);
}

export function signIn(email: string, password: string): Promise<User> {
  return signInWithEmailAndPassword(auth, email.trim(), password).then((c) => c.user);
}

export function signOut(): Promise<void> {
  return fbSignOut(auth);
}

export function resetPassword(email: string): Promise<void> {
  return sendPasswordResetEmail(auth, email.trim());
}

// An approved member has a profile at /designers/{uid}. The rules always let
// a user read their own profile, so a missing document means "signed in but
// not approved yet" rather than an error.
export async function isApprovedMember(user: User): Promise<boolean> {
  const snap = await getDoc(doc(db, "designers", user.uid));
  return snap.exists();
}
