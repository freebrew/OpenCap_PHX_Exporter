// Chromium Edge and Chrome both expose chrome.*.
// A few Edge / Firefox-shaped hosts only expose browser. Bind one chrome
// global so popup, content, and the service worker keep using chrome.*.
(() => {
  const root = typeof globalThis !== "undefined" ? globalThis : self;
  if (root.chrome && root.chrome.runtime) return;
  if (root.browser && root.browser.runtime) {
    root.chrome = root.browser;
  }
})();
