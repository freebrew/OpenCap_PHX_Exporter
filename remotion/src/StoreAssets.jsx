import React from "react";
import {AbsoluteFill, Img, staticFile} from "remotion";

const accent = "#00a39b";
const amber = "#ff9f2f";
const green = "#52d66b";
const red = "#ff6060";

const variants = {
  export: {
    kicker: "OpenCap Exporter",
    title: "FieldCap data to clean CSV in one click",
    subtitle: "Export job details, crew, BHA equipment, inventory, and daily slide / rotate metres using your active FieldCap session.",
    proof: ["5 structured CSV outputs", "Local file save workflow", "Excel-ready field reporting"],
    highlight: "Fetch & Save CSVs",
    theme: accent,
    screenshot: "FieldCap-Exporter-Extension.png",
  },
  operations: {
    kicker: "Wellsite workflow",
    title: "Built for fast field-to-office handoff",
    subtitle: "A compact operations panel keeps export status, job context, and workbook targeting visible without leaving FieldCap.",
    proof: ["BHA + inventory capture", "Slide / rotate by day", "Live status cards"],
    highlight: "Field reporting pipeline",
    theme: green,
    screenshot: "FieldCap-Exporter-Extension.png",
  },
  diagnostics: {
    kicker: "Diagnostics included",
    title: "Discover tenant-specific OData fields safely",
    subtitle: "Schema, inventory, sniff, and debug tools help validate FieldCap endpoints while keeping request metadata local.",
    proof: ["Schema probe", "Inventory probe", "Local-only metadata"],
    highlight: "No telemetry",
    theme: amber,
    screenshot: "FieldCap-Exporter-Extension2.png",
  },
};

const assets = {
  "screenshot-1280-export": {variant: "export", width: 1280, height: 800, format: "screenshot"},
  "screenshot-1280-operations": {variant: "operations", width: 1280, height: 800, format: "screenshot"},
  "screenshot-1280-diagnostics": {variant: "diagnostics", width: 1280, height: 800, format: "screenshot"},
  "screenshot-640-export": {variant: "export", width: 640, height: 400, format: "screenshot"},
  "screenshot-640-operations": {variant: "operations", width: 640, height: 400, format: "screenshot"},
  "screenshot-640-diagnostics": {variant: "diagnostics", width: 640, height: 400, format: "screenshot"},
  "small-promo-export": {variant: "export", width: 440, height: 280, format: "small"},
  "small-promo-operations": {variant: "operations", width: 440, height: 280, format: "small"},
  "small-promo-diagnostics": {variant: "diagnostics", width: 440, height: 280, format: "small"},
  "marquee-export": {variant: "export", width: 1400, height: 560, format: "marquee"},
  "marquee-operations": {variant: "operations", width: 1400, height: 560, format: "marquee"},
  "marquee-diagnostics": {variant: "diagnostics", width: 1400, height: 560, format: "marquee"},
};

const scale = (width) => width / 1400;

const Background = ({theme}) => (
  <AbsoluteFill
    style={{
      background:
        `radial-gradient(circle at 82% 18%, ${theme}44 0, transparent 28%), ` +
        "radial-gradient(circle at 4% 92%, #ff7a1f33 0, transparent 24%), " +
        "linear-gradient(135deg, #080b0d 0%, #101719 48%, #071010 100%)",
      overflow: "hidden",
    }}
  >
    <div
      style={{
        position: "absolute",
        inset: 20,
        border: `1px solid ${theme}55`,
        borderRadius: 26,
        boxShadow: `inset 0 0 80px ${theme}18`,
      }}
    />
    <div
      style={{
        position: "absolute",
        left: -180,
        bottom: -220,
        width: 620,
        height: 620,
        borderRadius: 999,
        background: "linear-gradient(135deg, #ff7a1f33, transparent 62%)",
        filter: "blur(10px)",
      }}
    />
  </AbsoluteFill>
);

const LogoMark = ({size = 62}) => (
  <div
    style={{
      width: size,
      height: size,
      borderRadius: size * 0.22,
      background: "linear-gradient(135deg, #151718, #050606)",
      border: "1px solid #303739",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      boxShadow: "0 18px 45px #0009",
    }}
  >
    <div
      style={{
        width: size * 0.62,
        height: size * 0.34,
        borderRadius: "60% 18% 60% 18%",
        background: "linear-gradient(135deg, #d36118, #5b2307)",
        transform: "rotate(-24deg)",
      }}
    />
  </div>
);

const Badge = ({children, color = accent}) => (
  <div
    style={{
      color,
      border: `1px solid ${color}88`,
      background: `${color}18`,
      borderRadius: 999,
      padding: "8px 14px",
      fontSize: 18,
      fontWeight: 800,
      letterSpacing: "0.08em",
      textTransform: "uppercase",
      whiteSpace: "nowrap",
    }}
  >
    {children}
  </div>
);

const ScreenshotCard = ({shot, wide = false, compact = false}) => (
  <div
    style={{
      position: "relative",
      borderRadius: compact ? 18 : 28,
      padding: compact ? 6 : 10,
      background: "linear-gradient(135deg, #ffffff24, #ffffff08)",
      border: "1px solid #ffffff2c",
      boxShadow: "0 34px 90px #000c",
      transform: wide ? "rotate(1.2deg)" : "rotate(0.8deg)",
    }}
  >
    <Img
      src={staticFile(shot)}
      style={{
        display: "block",
        width: "100%",
        borderRadius: compact ? 13 : 20,
      }}
    />
  </div>
);

const FeatureList = ({items, theme, compact = false}) => (
  <div style={{display: "flex", flexDirection: "column", gap: compact ? 8 : 12}}>
    {items.map((item) => (
      <div key={item} style={{display: "flex", alignItems: "center", gap: compact ? 8 : 12}}>
        <div
          style={{
            width: compact ? 20 : 26,
            height: compact ? 20 : 26,
            borderRadius: 6,
            background: theme,
            color: "#061112",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            fontSize: compact ? 14 : 18,
            fontWeight: 900,
          }}
        >
          ✓
        </div>
        <div style={{fontSize: compact ? 18 : 24, color: "#edf7f7", fontWeight: 700}}>{item}</div>
      </div>
    ))}
  </div>
);

const ScreenshotLayout = ({asset, data}) => {
  const s = scale(asset.width);
  const compact = asset.width <= 700;
  return (
    <AbsoluteFill style={{fontFamily: "Inter, Segoe UI, Arial, sans-serif"}}>
      <Background theme={data.theme} />
      <div
        style={{
          position: "absolute",
          inset: compact ? 22 : 44,
          display: "grid",
          gridTemplateColumns: "0.92fr 1.08fr",
          gap: compact ? 18 : 42,
          alignItems: "center",
        }}
      >
        <div style={{display: "flex", flexDirection: "column", gap: compact ? 13 : 24}}>
          <div style={{display: "flex", alignItems: "center", gap: compact ? 10 : 16}}>
            <LogoMark size={compact ? 34 : 56} />
            <div>
              <div style={{fontSize: compact ? 13 : 20, color: data.theme, fontWeight: 900, letterSpacing: "0.18em", textTransform: "uppercase"}}>
                {data.kicker}
              </div>
              <div style={{fontSize: compact ? 11 : 15, color: "#9fb4b9", letterSpacing: "0.1em", textTransform: "uppercase"}}>
                FieldCap OData · Excel-ready
              </div>
            </div>
          </div>
          <div
            style={{
              fontSize: compact ? 31 : 62 * s,
              lineHeight: 0.95,
              color: "#ffffff",
              fontWeight: 950,
              letterSpacing: "-0.045em",
            }}
          >
            {data.title}
          </div>
          <div style={{fontSize: compact ? 15 : 24 * s, lineHeight: 1.32, color: "#b8cacf", maxWidth: compact ? 280 : 560}}>
            {data.subtitle}
          </div>
          {!compact && <FeatureList items={data.proof} theme={data.theme} />}
        </div>
        <div style={{position: "relative"}}>
          <ScreenshotCard shot={data.screenshot} compact={compact} />
          <div
            style={{
              position: "absolute",
              right: compact ? 8 : 24,
              bottom: compact ? -14 : -22,
              background: "#071112",
              border: `1px solid ${data.theme}`,
              color: data.theme,
              boxShadow: `0 18px 45px ${data.theme}44`,
              borderRadius: 999,
              padding: compact ? "8px 12px" : "12px 22px",
              fontSize: compact ? 12 : 21,
              fontWeight: 900,
              letterSpacing: "0.08em",
              textTransform: "uppercase",
            }}
          >
            {data.highlight}
          </div>
        </div>
      </div>
    </AbsoluteFill>
  );
};

const SmallPromoLayout = ({data}) => (
  <AbsoluteFill style={{fontFamily: "Inter, Segoe UI, Arial, sans-serif"}}>
    <Background theme={data.theme} />
    <div style={{position: "absolute", inset: 18, display: "grid", gridTemplateColumns: "1fr 0.92fr", gap: 16, alignItems: "center"}}>
      <div style={{display: "flex", flexDirection: "column", gap: 10}}>
        <div style={{display: "flex", alignItems: "center", gap: 10}}>
          <LogoMark size={38} />
          <div style={{fontSize: 13, color: data.theme, fontWeight: 900, letterSpacing: "0.12em", textTransform: "uppercase"}}>OpenCap</div>
        </div>
        <div style={{fontSize: 30, lineHeight: 0.95, color: "#fff", fontWeight: 950, letterSpacing: "-0.04em"}}>
          FieldCap exports, cleaned.
        </div>
        <div style={{fontSize: 14, lineHeight: 1.25, color: "#b8cacf"}}>
          CSVs for job, crew, BHA, inventory, and slide / rotate reporting.
        </div>
        <Badge color={data.theme}>Excel-ready</Badge>
      </div>
      <ScreenshotCard shot={data.screenshot} compact />
    </div>
  </AbsoluteFill>
);

const MarqueeLayout = ({data}) => (
  <AbsoluteFill style={{fontFamily: "Inter, Segoe UI, Arial, sans-serif"}}>
    <Background theme={data.theme} />
    <div style={{position: "absolute", inset: 48, display: "grid", gridTemplateColumns: "0.9fr 1.1fr", gap: 54, alignItems: "center"}}>
      <div style={{display: "flex", flexDirection: "column", gap: 22}}>
        <div style={{display: "flex", alignItems: "center", gap: 18}}>
          <LogoMark size={68} />
          <div>
            <div style={{fontSize: 24, color: data.theme, fontWeight: 950, letterSpacing: "0.18em", textTransform: "uppercase"}}>OpenCap Exporter</div>
            <div style={{fontSize: 16, color: "#9fb4b9", letterSpacing: "0.1em", textTransform: "uppercase"}}>FieldCap OData · Local CSV workflow</div>
          </div>
        </div>
        <div style={{fontSize: 70, lineHeight: 0.92, color: "#fff", fontWeight: 950, letterSpacing: "-0.05em"}}>
          Turn FieldCap sessions into office-ready reports.
        </div>
        <FeatureList items={data.proof} theme={data.theme} compact />
      </div>
      <div style={{position: "relative", height: 430}}>
        <div style={{position: "absolute", width: 470, right: 280, top: 52, opacity: 0.72, transform: "rotate(-4deg) scale(0.92)"}}>
          <ScreenshotCard shot="FieldCap-Exporter-Extension2.png" compact />
        </div>
        <div style={{position: "absolute", width: 520, right: 0, top: 4}}>
          <ScreenshotCard shot={data.screenshot} wide />
        </div>
        <div style={{position: "absolute", right: 28, bottom: 0, display: "flex", gap: 12}}>
          <Badge color={data.theme}>5 CSV outputs</Badge>
          <Badge color={red}>No telemetry</Badge>
        </div>
      </div>
    </div>
  </AbsoluteFill>
);

export const storeAssetIds = Object.keys(assets);

export const StoreAsset = ({assetId}) => {
  const asset = assets[assetId] ?? assets["screenshot-1280-export"];
  const data = variants[asset.variant];

  if (asset.format === "small") return <SmallPromoLayout data={data} />;
  if (asset.format === "marquee") return <MarqueeLayout data={data} />;
  return <ScreenshotLayout asset={asset} data={data} />;
};

export const getAssetConfig = (assetId) => assets[assetId];

