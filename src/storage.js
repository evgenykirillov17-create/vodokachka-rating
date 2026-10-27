// Замена внутреннего window.storage (доступного только в артефактах Claude)
// на настоящую базу данных — Firestore. Интерфейс намеренно повторяет исходный
// window.storage.get/set (ключ -> значение), чтобы остальной код почти не менять.

import { initializeApp } from "firebase/app";
import { getFirestore, doc, getDoc, setDoc, deleteDoc } from "firebase/firestore";
import { getAuth, signInWithEmailAndPassword, signOut, onAuthStateChanged } from "firebase/auth";
import { firebaseConfig } from "./firebaseConfig";

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);
const auth = getAuth(app);

// Единственный аккаунт админа. Это не настоящая почта — просто логин в Firebase.
// Пароль хранится только в Firebase (Authentication → Users), в коде его нет.
export const ADMIN_EMAIL = "admin@vodokachka.app";

export async function adminSignIn(password) {
  await signInWithEmailAndPassword(auth, ADMIN_EMAIL, password);
}

export async function adminSignOut() {
  await signOut(auth);
}

// Вызывает cb(true/false) при входе/выходе. Вход запоминается в браузере.
export function onAdminChange(cb) {
  return onAuthStateChanged(auth, (user) => cb(!!user && user.email === ADMIN_EMAIL));
}

const COLLECTION = "ratingApp";

export async function storageGet(key) {
  const ref = doc(db, COLLECTION, key);
  const snap = await getDoc(ref);
  if (!snap.exists()) return null;
  return { key, value: snap.data().value };
}

export async function storageSet(key, value) {
  const ref = doc(db, COLLECTION, key);
  try {
    await setDoc(ref, { value, updatedAt: new Date().toISOString() });
  } catch (e) {
    alert("Не сохранилось в базу: " + (e.code === "permission-denied" ? "нет прав — выйдите и войдите в админку заново." : e.message));
    throw e;
  }
  return { key, value };
}

export async function storageDelete(key) {
  const ref = doc(db, COLLECTION, key);
  await deleteDoc(ref);
}
