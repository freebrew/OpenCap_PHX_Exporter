import type { OracleCell } from "./oracle.js";

const ABS_TOLERANCE = 1e-6;
const REL_TOLERANCE = 1e-9;

export type ActualCell = OracleCell;

function isBlank(cell: OracleCell): boolean {
  return cell.v == null || cell.v === "";
}

export function valuesMatch(expected: OracleCell, actual: ActualCell): boolean {
  if (isBlank(expected) && isBlank(actual)) return true;
  if (expected.t === "e" || actual.t === "e") return expected.t === "e" && actual.t === "e" && String(expected.v) === String(actual.v);
  if (expected.t === "b" || actual.t === "b") return expected.t === "b" && actual.t === "b" && expected.v === actual.v;
  if (expected.t === "n" && actual.t === "n" && typeof expected.v === "number" && typeof actual.v === "number") {
    const delta = Math.abs(expected.v - actual.v);
    const scale = Math.max(Math.abs(expected.v), Math.abs(actual.v), 1);
    return delta <= ABS_TOLERANCE || delta / scale <= REL_TOLERANCE;
  }
  if (expected.v == null && actual.v == null) return true;
  return expected.t === actual.t && String(expected.v) === String(actual.v);
}
