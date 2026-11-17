export const LEGACY_UDF_NAMES = [
  "CONTDITVDATBIT",
  "PROJACTIVETARGETMDEX",
  "PROJAZMATBIT",
  "PROJBITE",
  "PROJBITN",
  "PROJBURR",
  "PROJCONTDIAIM",
  "PROJCUME",
  "PROJCUMN",
  "PROJDOGLEGDEG",
  "PROJINCATBIT",
  "PROJLANDINGBURR",
  "PROJMETERSTOSLIDE",
  "PROJPARSETF",
  "PROJREQUIREDTF",
  "PROJROLLINGMOTOROUT",
  "PROJSLIDEMETERSBETWEEN",
  "PROJTFTOTARGET",
  "PROJTVDATBIT",
  "TDACTUALBIT",
  "TDESTFINAL",
] as const;

export interface FormulaRegistrar {
  registerFunction(name: string, implementation: (...args: unknown[]) => null, description?: string): void;
}

/** Phase 1 stubs keep the formula parser alive. Phase 2 replaces them with dd-math. */
export function registerLegacyUdfStubs(formula: FormulaRegistrar, extraNames: Iterable<string> = []): void {
  const names = new Set<string>([...LEGACY_UDF_NAMES, ...extraNames]);
  for (const name of names) {
    formula.registerFunction(name, () => null, "Legacy VBA function. This preview returns blank until Phase 2.");
  }
}
