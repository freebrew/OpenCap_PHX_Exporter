declare module "virtual:opencap-exporter" {
  /** src/chrome-extension/background.js (after browser-api.js), with `chrome` and `importScripts` supplied. */
  export function startOpenCapExporter(chrome: typeof globalThis.chrome, importScripts: (...urls: string[]) => void): void;
}
