---
name: ce-frontend-design
description: >-
  Build and improve web UI for PHX_FieldCap, including the Chrome extension
  popup/options pages, any web-based dashboard, or HTML reports generated from
  field data. Use when designing or modifying background.js, popup HTML/CSS,
  or any browser-facing interface. Trigger phrases: 'design this', 'make this
  look better', 'build the UI', 'chrome extension popup', 'web dashboard'.
---

# ce-frontend-design — PHX_FieldCap

## Scope in this project

- Chrome extension UI: popup, options page, injected overlays
- Any HTML reports exported from Excel/PowerShell pipelines
- Web-based field data dashboards (if added)

## Design principles

- **Utility-first** — field crew uses this on-site; clarity beats decoration
- **Readable at a glance** — large tap targets, high contrast, minimal chrome
- **No dependencies** — extension UI must work offline; no CDN imports
- **System fonts only** — `font-family: system-ui, sans-serif`

## Chrome extension conventions

```html
<!-- popup.html pattern -->
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <link rel="stylesheet" href="popup.css">
</head>
<body>
  <!-- Max 800×600px; Chrome clips larger popups -->
  <script src="popup.js"></script>
</body>
</html>
```

- Popup max width: 400px; max height: 600px
- No inline styles — all CSS in external `.css` file
- No `eval()`, no `innerHTML` with untrusted data (CSP violation)
- Use `chrome.storage.local` for state; never `localStorage` in extensions

## Color / spacing tokens (use these, not arbitrary values)

```css
:root {
  --color-bg: #ffffff;
  --color-surface: #f5f5f5;
  --color-border: #e0e0e0;
  --color-text: #1a1a1a;
  --color-text-secondary: #666666;
  --color-accent: #1a56db;   /* PHX blue */
  --color-success: #057a55;
  --color-warning: #b45309;
  --color-error: #c81e1e;
  --radius: 4px;
  --space-sm: 8px;
  --space-md: 16px;
  --space-lg: 24px;
}
```

## Verification step

After any UI change: take a screenshot via the browser tool and confirm layout matches intent before declaring done.
