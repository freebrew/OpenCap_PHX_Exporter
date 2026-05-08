(function () {
  "use strict";
  if (window.__fcSpy) return;
  window.__fcSpy = true;

  const post = (url, data) => {
    try {
      window.postMessage({ __FC_ODATA__: 1, url, data }, "*");
    } catch (_) {}
  };

  const _fetch = window.fetch;
  window.fetch = function (input) {
    const url = (typeof input === "string" ? input : input && input.url) || "";
    return _fetch.apply(this, arguments).then(function (resp) {
      if (url.indexOf("/odata/") !== -1) {
        try {
          resp.clone().json().then((data) => post(url, data)).catch(() => {});
        } catch (_) {}
      }
      return resp;
    });
  };

  const _open = XMLHttpRequest.prototype.open;
  const _send = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.open = function (method, url) {
    try {
      this.__fc_url = url || "";
    } catch (_) {}
    return _open.apply(this, arguments);
  };
  XMLHttpRequest.prototype.send = function () {
    this.addEventListener("load", function () {
      try {
        const url = this.__fc_url || "";
        if (url.indexOf("/odata/") === -1) return;
        const ct = (this.getResponseHeader("content-type") || "").toLowerCase();
        if (ct.indexOf("json") === -1) return;
        const txt = this.responseText;
        if (!txt) return;
        post(url, JSON.parse(txt));
      } catch (_) {}
    });
    return _send.apply(this, arguments);
  };
})();
