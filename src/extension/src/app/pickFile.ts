// Application.FileDialog for the Setup import buttons. The picker has to open in the
// same turn as the click. Awaiting the job folder first spends the user gesture and
// Chrome never shows the dialog. A browser without the file-system picker gets a
// normal file input, which every Chrome window has.

export interface FileEntry {
  readonly kind: "file";
  readonly name: string;
  getFile(): Promise<File>;
  createWritable(): Promise<{ write(data: string): Promise<void>; close(): Promise<void> }>;
}

export interface FolderEntry {
  readonly kind: "directory";
  readonly name: string;
  resolve(possibleDescendant: FileEntry): Promise<string[] | null>;
  getDirectoryHandle(name: string, options?: { create?: boolean }): Promise<FolderEntry>;
  getFileHandle(name: string, options?: { create?: boolean }): Promise<FileEntry>;
  entries(): AsyncIterable<[string, FileEntry | FolderEntry]>;
}

export interface PickerType {
  description: string;
  accept: Record<string, string[]>;
}

export interface PickedFile {
  file: File;
  /** "Job folder\Well Plans\x.pdf" inside the job folder, otherwise the file name. */
  path: string;
  /** A file in the same folder, or null when the pick was outside the job folder. */
  sibling(name: string): Promise<File | null>;
}

type OpenPicker = (options: { id?: string; types: PickerType[]; startIn?: FolderEntry | "documents"; multiple?: false }) => Promise<FileEntry[]>;

async function folderAt(root: FolderEntry, segments: string[]): Promise<FolderEntry> {
  let folder = root;
  for (const segment of segments) folder = await folder.getDirectoryHandle(segment);
  return folder;
}

function acceptList(types: PickerType[]): string {
  return types.flatMap((type) => Object.values(type.accept).flat()).join(",");
}

function nativePicker(picker: OpenPicker): boolean {
  return /\[native code\]/.test(Function.prototype.toString.call(picker));
}

/**
 * Opens on the click, before any await. Cancel resolves to null.
 * The native picker only runs during a real user gesture. A stub (the tests) always
 * runs. Every other Chrome window gets a file input, which does not need that gesture.
 */
function beginPick(id: string, types: PickerType[]): Promise<FileEntry | File | null> {
  const picker = (window as Window & { showOpenFilePicker?: OpenPicker }).showOpenFilePicker;
  const gesture = typeof navigator !== "undefined" && navigator.userActivation?.isActive === true;
  if (picker && (!nativePicker(picker) || gesture)) {
    try {
      return picker({ id, types, startIn: "documents" }).then(
        (entries) => entries[0] ?? null,
        (error: unknown) => {
          if (error instanceof DOMException && (error.name === "AbortError" || error.name === "SecurityError")) return null;
          throw error;
        },
      );
    } catch (error) {
      if (!(error instanceof DOMException && error.name === "SecurityError")) throw error;
    }
  }
  return pickWithInput(types);
}

function pickWithInput(types: PickerType[]): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = acceptList(types);
    input.addEventListener("change", () => resolve(input.files?.[0] ?? null), { once: true });
    input.addEventListener("cancel", () => resolve(null), { once: true });
    input.click();
  });
}

/**
 * The folder may still be resolving. Pass that promise in; do not await it before
 * calling this, or the file dialog will not open.
 */
export function pickFile(id: string, types: PickerType[], jobFolder: Promise<FolderEntry | null> | FolderEntry | null): Promise<PickedFile | null> {
  const opened = beginPick(id, types);
  return Promise.all([opened, Promise.resolve(jobFolder)]).then(([entry, folder]) => finishPick(entry, folder));
}

async function finishPick(entry: FileEntry | File | null, jobFolder: FolderEntry | null): Promise<PickedFile | null> {
  if (!entry) return null;
  if (entry instanceof File) return { file: entry, path: entry.name, sibling: async () => null };
  const file = await entry.getFile();
  const segments = jobFolder ? await jobFolder.resolve(entry).catch(() => null) : null;
  if (!jobFolder || !segments) return { file, path: file.name, sibling: async () => null };
  const parents = segments.slice(0, -1);
  return {
    file,
    path: [jobFolder.name, ...segments].join("\\"),
    sibling: async (name) => {
      try {
        return await (await (await folderAt(jobFolder, parents)).getFileHandle(name)).getFile();
      } catch {
        return null;
      }
    },
  };
}
