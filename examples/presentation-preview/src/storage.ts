import type { DeckSessionV1 } from "@pptkit/presentation-workflow";

const DATABASE = "pptkit-presentation-preview";
const VERSION = 1;
const SESSIONS = "sessions";
const ASSETS = "assets";

function openDatabase() {
  return new Promise<IDBDatabase>((resolve, reject) => {
    const request = indexedDB.open(DATABASE, VERSION);
    request.onerror = () => reject(request.error ?? new Error("IndexedDB could not be opened."));
    request.onupgradeneeded = () => {
      const database = request.result;
      if (!database.objectStoreNames.contains(SESSIONS)) database.createObjectStore(SESSIONS, { keyPath: "id" });
      if (!database.objectStoreNames.contains(ASSETS)) database.createObjectStore(ASSETS, { keyPath: "key" });
    };
    request.onsuccess = () => resolve(request.result);
  });
}

function complete(transaction: IDBTransaction) {
  return new Promise<void>((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error ?? new Error("IndexedDB transaction failed."));
    transaction.onabort = () => reject(transaction.error ?? new Error("IndexedDB transaction was aborted."));
  });
}

export async function saveSession(session: DeckSessionV1) {
  const database = await openDatabase();
  const transaction = database.transaction(SESSIONS, "readwrite");
  transaction.objectStore(SESSIONS).put(session);
  await complete(transaction);
  database.close();
}

export async function loadSession(id: string) {
  const database = await openDatabase();
  const transaction = database.transaction(SESSIONS, "readonly");
  const request = transaction.objectStore(SESSIONS).get(id);
  const result = await new Promise<DeckSessionV1 | undefined>((resolve, reject) => {
    request.onsuccess = () => resolve(request.result as DeckSessionV1 | undefined);
    request.onerror = () => reject(request.error ?? new Error("Session could not be loaded."));
  });
  await complete(transaction);
  database.close();
  return result;
}

export async function saveAssetBlob(sessionId: string, assetId: string, blob: Blob) {
  const database = await openDatabase();
  const transaction = database.transaction(ASSETS, "readwrite");
  transaction.objectStore(ASSETS).put({ key: `${sessionId}:${assetId}`, blob });
  await complete(transaction);
  database.close();
}

export async function loadAssetBlob(sessionId: string, assetId: string) {
  const database = await openDatabase();
  const transaction = database.transaction(ASSETS, "readonly");
  const request = transaction.objectStore(ASSETS).get(`${sessionId}:${assetId}`);
  const result = await new Promise<{ blob: Blob } | undefined>((resolve, reject) => {
    request.onsuccess = () => resolve(request.result as { blob: Blob } | undefined);
    request.onerror = () => reject(request.error ?? new Error("Asset could not be loaded."));
  });
  await complete(transaction);
  database.close();
  return result?.blob;
}
