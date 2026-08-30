// Замена внутреннего window.storage (доступного только в артефактах Claude)
// на настоящую базу данных — Firestore. Интерфейс намеренно повторяет исходный
// window.storage.get/set (ключ -> значение), чтобы остальной код почти не менять.

import { initializeApp } from "firebase/app";
import { getFirestore, doc, getDoc, setDoc, deleteDoc } from "firebase/firestore";
import { firebaseConfig } from "./firebaseConfig";

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

const COLLECTION = "ratingApp";

export async function storageGet(key) {
  const ref = doc(db, COLLECTION, key);
  const snap = await getDoc(ref);
  if (!snap.exists()) return null;
  return { key, value: snap.data().value };
}

export async function storageSet(key, value) {
  const ref = doc(db, COLLECTION, key);
  await setDoc(ref, { value, updatedAt: new Date().toISOString() });
  return { key, value };
}

export async function storageDelete(key) {
  const ref = doc(db, COLLECTION, key);
  await deleteDoc(ref);
}
