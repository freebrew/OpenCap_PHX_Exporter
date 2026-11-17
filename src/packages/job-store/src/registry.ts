import type { JobRecord, MirrorRecord } from "./types.js";

const DB_NAME = "phx-workbook";
const DB_VERSION = 1;

export interface JobRegistry {
  list(): Promise<JobRecord[]>;
  get(jobId: string): Promise<(JobRecord & { handle?: unknown }) | null>;
  put(record: JobRecord & { handle?: unknown }): Promise<void>;
  remove(jobId: string): Promise<void>;
  getMirror(jobId: string): Promise<MirrorRecord | null>;
  putMirror(record: MirrorRecord): Promise<void>;
  removeMirror(jobId: string): Promise<void>;
}

export function requestToPromise<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
    // Chrome aborts the transaction for a value over its size limit and never fails the request,
    // so without this the caller waits forever.
    const transaction = request.transaction;
    transaction?.addEventListener("abort", () => reject(transaction.error ?? new Error("The browser store refused the write.")));
  });
}

export function openRegistry(factory: IDBFactory = indexedDB): Promise<JobRegistry> {
  const open = factory.open(DB_NAME, DB_VERSION);
  open.onupgradeneeded = () => {
    const db = open.result;
    if (!db.objectStoreNames.contains("jobs")) db.createObjectStore("jobs", { keyPath: "jobId" });
    if (!db.objectStoreNames.contains("mirrors")) db.createObjectStore("mirrors", { keyPath: "jobId" });
  };
  return new Promise((resolve, reject) => {
    open.onerror = () => reject(open.error);
    open.onsuccess = () => {
      const db = open.result;
      const tx = (stores: string[], mode: IDBTransactionMode) => db.transaction(stores, mode);
      resolve({
        async list() {
          const store = tx(["jobs"], "readonly").objectStore("jobs");
          const rows = await requestToPromise(store.getAll()) as Array<JobRecord & { handle?: unknown }>;
          return rows.map((row) => ({
            jobId: row.jobId,
            label: row.label,
            fileName: row.fileName,
            lastOpened: row.lastOpened,
          }));
        },
        async get(jobId) {
          const store = tx(["jobs"], "readonly").objectStore("jobs");
          return requestToPromise(store.get(jobId));
        },
        async put(record) {
          const store = tx(["jobs"], "readwrite").objectStore("jobs");
          await requestToPromise(store.put(record));
        },
        async remove(jobId) {
          const transaction = tx(["jobs", "mirrors"], "readwrite");
          await requestToPromise(transaction.objectStore("jobs").delete(jobId));
          await requestToPromise(transaction.objectStore("mirrors").delete(jobId));
        },
        async getMirror(jobId) {
          const store = tx(["mirrors"], "readonly").objectStore("mirrors");
          return requestToPromise(store.get(jobId));
        },
        async putMirror(record) {
          const store = tx(["mirrors"], "readwrite").objectStore("mirrors");
          await requestToPromise(store.put(record));
        },
        async removeMirror(jobId) {
          const store = tx(["mirrors"], "readwrite").objectStore("mirrors");
          await requestToPromise(store.delete(jobId));
        },
      });
    };
  });
}
