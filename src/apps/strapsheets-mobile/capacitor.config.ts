import type { CapacitorConfig } from "@capacitor/cli";

const config: CapacitorConfig = {
  appId: "com.phx.strapsheets",
  appName: "StrapSheets",
  webDir: "dist",
  android: {
    allowMixedContent: false,
  },
};

export default config;
