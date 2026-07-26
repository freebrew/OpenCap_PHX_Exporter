# Chrome Web Store Asset Generator

Remotion source for generating marketing assets from real OpenCap Exporter screenshots.

## Inputs

The renderer uses these source screenshots in `remotion/public/`:

- `FieldCap-Exporter-Extension.png`
- `FieldCap-Exporter-Extension2.png`

## Render

```powershell
.\remotion\render-assets.ps1
```

Generated upload assets are written to `remotion/out/` as JPEG files:

- `screenshot-1280-*.jpg` — 1280x800 screenshots
- `screenshot-640-*.jpg` — 640x400 screenshots
- `small-promo-*.jpg` — 440x280 small promo tiles
- `marquee-*.jpg` — 1400x560 marquee promo tiles

