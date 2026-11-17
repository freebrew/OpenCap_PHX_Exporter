import { doglegAngleDeg } from "./angles";
import { closestOffset, type GaugeOffset, type PlanStation } from "./plan";
import {
  activeTargetMd,
  activeTargetMdEx,
  ARG_ERROR,
  azmAtBit,
  bitEast,
  bitNorth,
  burr,
  contDiAim,
  contDiTvdAtBit,
  cumulative,
  incAtBit,
  landingBurr,
  metersToSlide,
  parseToolface,
  requiredTf,
  rollingMotorOut,
  slideMetersBetween,
  tdActualBit,
  tdEstFinal,
  tfToTarget,
  tvdAtBit,
  UdfError,
  type SurveyContext,
} from "./project";

export interface UdfCall {
  value: number | string | null;
  error?: "VALUE" | "NUM" | "NAME";
}

function finish(body: () => number | string | null): UdfCall {
  try {
    return { value: body() };
  } catch (error) {
    if (error instanceof UdfError) return { value: null, error: error.code };
    throw error;
  }
}

export function callUdf(name: string, args: unknown[], context?: SurveyContext): UdfCall {
  const key = name.toUpperCase();
  switch (key) {
    case "PROJPARSETF":
      return finish(() => parseToolface(args[0]));
    case "PROJDOGLEGDEG":
      return finish(() => {
        const values = args.slice(0, 4).map((value) => (typeof value === "number" ? value : Number(value)));
        if (values.some((value) => !Number.isFinite(value))) return null;
        return doglegAngleDeg(values[0] ?? 0, values[1] ?? 0, values[2] ?? 0, values[3] ?? 0);
      });
    case "PROJINCATBIT":
      return finish(() => incAtBit(args[0], args[1], args[2], args[3], args[4], args[5], args[6], context));
    case "PROJAZMATBIT":
      return finish(() => azmAtBit(args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], context));
    case "PROJTVDATBIT":
      return finish(() => tvdAtBit(args[0], args[1], args[2], args[3], args[4], args[5], args[6], context));
    case "PROJCUMN":
      return finish(() => cumulative(args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], "n"));
    case "PROJCUME":
      return finish(() => cumulative(args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], "e"));
    case "PROJBITN":
      return finish(() => bitNorth(args[0], args[1], args[2], args[3], args[4], args[5], args[6]));
    case "PROJBITE":
      return finish(() => bitEast(args[0], args[1], args[2], args[3], args[4], args[5], args[6]));
    case "PROJSLIDEMETERSBETWEEN":
      return finish(() => slideMetersBetween(args[0], args[1], args[2], args[3]));
    case "PROJROLLINGMOTOROUT":
      return finish(() => rollingMotorOut(args[0], args[1]));
    case "PROJACTIVETARGETMD":
      return finish(() => activeTargetMd(args[0], args[1]));
    case "PROJACTIVETARGETMDEX":
      return finish(() => activeTargetMdEx(args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8]));
    case "PROJMETERSTOSLIDE":
      return finish(() => metersToSlide(args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], context));
    case "PROJBURR":
      return finish(() => burr(args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], context));
    case "PROJREQUIREDTF":
      return finish(() => requiredTf(args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], context));
    case "PROJTFTOTARGET":
      return finish(() => tfToTarget(args[0], args[1], args[2], args[3], args[4]));
    case "PROJLANDINGBURR":
      return finish(() => landingBurr(args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], context));
    case "PROJCONTDIAIM":
      return finish(() => contDiAim(args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10], args[11], args[12], context));
    case "CONTDITVDATBIT":
      return finish(() => contDiTvdAtBit(args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10], args[11], args[12]));
    case "TDACTUALBIT":
      return finish(() => tdActualBit(args[0], args[1]));
    case "TDESTFINAL":
      return finish(() => tdEstFinal(args[0], args[1], args[3], context?.plan));
    default:
      return { value: null, error: "NAME" };
  }
}

export function implementedUdf(name: string): boolean {
  return (IMPLEMENTED_UDFS as readonly string[]).includes(name.toUpperCase());
}

export const IMPLEMENTED_UDFS = [
  "CONTDITVDATBIT",
  "PROJACTIVETARGETMD",
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

export function gaugeAt(md: number, n: number, e: number, tvd: number, plan: PlanStation[]): GaugeOffset | null {
  return closestOffset(md, n, e, tvd, plan);
}

export { ARG_ERROR };
export type { SurveyContext };
