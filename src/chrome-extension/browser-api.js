// One Chrome package, two hosts.
// Chromium Edge exposes chrome.* as well as browser.*. A few hosts only
// expose browser. Bind chrome, then record which browser this copy is
// running in so Edge can take its own fetch and save path.
(() => {
  const root = typeof globalThis !== "undefined" ? globalThis : self;
  if (!(root.chrome && root.chrome.runtime) && root.browser && root.browser.runtime) {
    root.chrome = root.browser;
  }

  const ua = String(root.navigator?.userAgent || "");
  const brands = root.navigator?.userAgentData?.brands;
  const brandEdge = Array.isArray(brands)
    && brands.some((b) => /edge/i.test(String(b?.brand || "")));
  const isEdge = brandEdge || /\bEdg\//.test(ua);

  root.OC = {
    host: isEdge ? "edge" : "chrome",
    isEdge,
    storeName: isEdge ? "Edge Add-ons" : "Chrome Web Store",
  };
})();
