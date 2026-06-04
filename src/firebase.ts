import { initializeApp } from "firebase/app";
import {
  getAuth,
  onAuthStateChanged,
  signInAnonymously,
  type User,
} from "firebase/auth";
import { initializeFirestore } from "firebase/firestore";

// Same wtg-waypoint project as the main Waypoint app. The web `apiKey` is a
// public project identifier (not a secret); access is gated by Firestore
// rules + Anonymous Auth.
const firebaseConfig = {
  apiKey: "AIzaSyCCeatk1Bv-GwlZOnp_NQgRrMBTLgQNH1o",
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

export const auth = getAuth(firebaseApp);

// Resolves once we have an authenticated user. Office add-in iframes have no
// stored session, so we sign in anonymously on every load.
export function ensureSignedIn(): Promise<User> {
  return new Promise((resolve, reject) => {
    const unsubscribe = onAuthStateChanged(
      auth,
      (user) => {
        if (user) {
          unsubscribe();
          resolve(user);
        }
      },
      (err) => {
        unsubscribe();
        reject(err);
      },
    );
    if (!auth.currentUser) {
      signInAnonymously(auth).catch((err) => {
        unsubscribe();
        reject(err);
      });
    }
  });
}
