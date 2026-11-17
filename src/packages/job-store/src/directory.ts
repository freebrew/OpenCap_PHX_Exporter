import { StoreError } from "./types.js";

export interface ExclusiveWriter {
  write(bytes: Uint8Array): Promise<void>;
  close(): Promise<void>;
  abort(): Promise<void>;
}

export interface StoredFile {
  read(): Promise<Uint8Array | null>;
  writeExclusive(): Promise<ExclusiveWriter>;
  remove(): Promise<void>;
}

/** Narrow folder operations. Callers never receive the browser directory handle. */
export interface StoredDirectory {
  getFile(name: string, create: boolean): Promise<StoredFile | null>;
  list(): Promise<string[]>;
  queryPermission(): Promise<"granted" | "prompt" | "denied">;
  requestPermission(): Promise<"granted" | "denied">;
}

export class MemoryDirectory implements StoredDirectory {
  private readonly files = new Map<string, Uint8Array>();
  private readonly locks = new Set<string>();
  permission: "granted" | "prompt" | "denied" = "granted";

  async getFile(name: string, create: boolean): Promise<StoredFile | null> {
    if (!this.files.has(name) && !create) return null;
    if (!this.files.has(name) && create) this.files.set(name, new Uint8Array());
    return {
      read: async () => {
        const bytes = this.files.get(name);
        return bytes ? bytes.slice() : null;
      },
      writeExclusive: async () => {
        if (this.locks.has(name)) throw new StoreError("concurrent", `${name} is already open for writing.`);
        this.locks.add(name);
        let staged: Uint8Array | null = null;
        let closed = false;
        return {
          write: async (bytes) => {
            staged = bytes.slice();
          },
          close: async () => {
            if (closed) return;
            closed = true;
            if (staged) this.files.set(name, staged);
            this.locks.delete(name);
          },
          abort: async () => {
            closed = true;
            this.locks.delete(name);
          },
        };
      },
      remove: async () => {
        this.files.delete(name);
      },
    };
  }

  async list(): Promise<string[]> {
    return [...this.files.keys()];
  }

  async queryPermission(): Promise<"granted" | "prompt" | "denied"> {
    return this.permission;
  }

  async requestPermission(): Promise<"granted" | "denied"> {
    if (this.permission === "denied") return "denied";
    this.permission = "granted";
    return "granted";
  }
}

interface FileHandleLike {
  getFile(): Promise<Blob>;
  createWritable(options: { mode?: "exclusive" | "siloed" }): Promise<{
    write(data: BufferSource | Blob | string): Promise<void>;
    close(): Promise<void>;
    abort(): Promise<void>;
  }>;
  remove?(): Promise<void>;
}

interface DirectoryHandleLike {
  getFileHandle(name: string, options?: { create?: boolean }): Promise<FileHandleLike>;
  entries(): AsyncIterable<[string, { kind: string }]>;
  queryPermission(descriptor: { mode: "readwrite" }): Promise<"granted" | "prompt" | "denied">;
  requestPermission(descriptor: { mode: "readwrite" }): Promise<"granted" | "prompt" | "denied">;
  removeEntry?(name: string): Promise<void>;
}

export function directoryFromHandle(handle: DirectoryHandleLike): StoredDirectory {
  return {
    async getFile(name, create) {
      try {
        const file = await handle.getFileHandle(name, { create });
        return {
          read: async () => new Uint8Array(await (await file.getFile()).arrayBuffer()),
          writeExclusive: async () => {
            try {
              const writable = await file.createWritable({ mode: "exclusive" });
              return {
                write: async (bytes) => {
                  await writable.write(bytes as Uint8Array<ArrayBuffer>);
                },
                close: async () => {
                  await writable.close();
                },
                abort: async () => {
                  await writable.abort();
                },
              };
            } catch (error) {
              const nameOf = error instanceof DOMException ? error.name : "";
              if (nameOf === "NoModificationAllowedError" || nameOf === "InvalidStateError") {
                throw new StoreError("concurrent", "Another writer has this snapshot open.");
              }
              throw error;
            }
          },
          remove: async () => {
            if (handle.removeEntry) await handle.removeEntry(name);
          },
        };
      } catch (error) {
        if (error instanceof DOMException && error.name === "NotFoundError") return null;
        if (!create && error instanceof Error && /not found/i.test(error.message)) return null;
        throw error;
      }
    },
    async list() {
      const names: string[] = [];
      for await (const [name, entry] of handle.entries()) {
        if (entry.kind === "file") names.push(name);
      }
      return names;
    },
    queryPermission: () => handle.queryPermission({ mode: "readwrite" }),
    requestPermission: async () => {
      const result = await handle.requestPermission({ mode: "readwrite" });
      return result === "granted" ? "granted" : "denied";
    },
  };
}
