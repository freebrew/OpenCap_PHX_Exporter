import { mkdir, open, readdir, readFile, rm, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import type { ExclusiveWriter, StoredDirectory, StoredFile } from "./directory.js";
import { StoreError } from "./types.js";

async function lock(filePath: string): Promise<() => Promise<void>> {
  const lockPath = `${filePath}.lock`;
  try {
    const handle = await open(lockPath, "wx");
    return async () => {
      await handle.close();
      await rm(lockPath, { force: true });
    };
  } catch (error) {
    const code = error && typeof error === "object" && "code" in error ? String(error.code) : "";
    if (code === "EEXIST") throw new StoreError("concurrent", `${path.basename(filePath)} is already open for writing.`);
    throw error;
  }
}

function storedFile(filePath: string): StoredFile {
  return {
    async read() {
      try {
        return new Uint8Array(await readFile(filePath));
      } catch (error) {
        const code = error && typeof error === "object" && "code" in error ? String(error.code) : "";
        if (code === "ENOENT") return null;
        throw error;
      }
    },
    async writeExclusive() {
      const release = await lock(filePath);
      let staged: Uint8Array | null = null;
      const writer: ExclusiveWriter = {
        async write(bytes) {
          staged = bytes;
        },
        async close() {
          try {
            if (staged) await writeFile(filePath, staged);
          } finally {
            await release();
          }
        },
        async abort() {
          await release();
        },
      };
      return writer;
    },
    async remove() {
      await rm(filePath, { force: true });
    },
  };
}

export async function nodeDirectory(root: string): Promise<StoredDirectory> {
  await mkdir(root, { recursive: true });
  return {
    async getFile(name, create) {
      const filePath = path.join(root, name);
      try {
        await readFile(filePath);
      } catch (error) {
        const code = error && typeof error === "object" && "code" in error ? String(error.code) : "";
        if (code === "ENOENT" && !create) return null;
        if (code === "ENOENT" && create) await writeFile(filePath, new Uint8Array());
        else if (code === "ENOENT") return null;
        else throw error;
      }
      return storedFile(filePath);
    },
    async getDirectory(name, create) {
      const dirPath = path.join(root, name);
      if (create) return nodeDirectory(dirPath);
      try {
        const info = await stat(dirPath);
        return info.isDirectory() ? nodeDirectory(dirPath) : null;
      } catch (error) {
        const code = error && typeof error === "object" && "code" in error ? String(error.code) : "";
        if (code === "ENOENT") return null;
        throw error;
      }
    },
    async list() {
      const entries = await readdir(root, { withFileTypes: true });
      return entries.filter((entry) => entry.isFile() && !entry.name.endsWith(".lock")).map((entry) => entry.name);
    },
    async queryPermission() {
      return "granted";
    },
    async requestPermission() {
      return "granted";
    },
  };
}
