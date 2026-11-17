import { LocaleType } from "@univerjs/core";
import { UniverSheetsCoreWorkerPreset } from "@univerjs/preset-sheets-core/worker";
import { createUniver } from "@univerjs/presets";
import { legacyUdfEntries } from "./formulaEntries";

createUniver({
  locale: LocaleType.EN_US,
  presets: [
    UniverSheetsCoreWorkerPreset({
      formula: {
        function: legacyUdfEntries() as never,
      },
    }),
  ],
});
