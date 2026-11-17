import type { Ctor, Nullable } from "@univerjs/core";
import {
  BaseFunction,
  ErrorType,
  ErrorValueObject,
  NumberValueObject,
  StringValueObject,
  type BaseReferenceObject,
  type BaseValueObject,
  type IRuntimeUnitDataType,
} from "@univerjs/engine-formula";
import { ARG_ERROR, callUdf, implementedUdf, parseToolface, type PlanStation, type SurveyContext } from "../../../packages/dd-math/src/index";
import { loadPlanGrid } from "../../../packages/dd-math/src/project";
import { LEGACY_UDF_NAMES } from "./legacyUdf";

interface CellMatrix {
  getValue(row: number, column: number): { v?: unknown } | null | undefined;
}

interface CachedGrid {
  unitId: string;
  stamp: string;
  plan: PlanStation[];
  slides: SurveyContext["slides"];
}

let gridCache: CachedGrid | null = null;

class LegacyNameFunction extends BaseFunction {
  constructor(name: string) {
    super(name as never);
  }

  override calculate(): ErrorValueObject {
    return ErrorValueObject.create(ErrorType.NAME);
  }
}

class DirectionalFunction extends BaseFunction {
  constructor(name: string) {
    super(name as never);
    this.needsFormulaDataModel = true;
    this.needsSheetsInfo = true;
    this.needsReferenceObject = true;
    this.minParams = 0;
    this.maxParams = 32;
  }

  override calculate(...args: BaseValueObject[]): BaseValueObject {
    let runtime: IRuntimeUnitDataType | undefined;
    const values = args.map((arg) => {
      if (!arg.isReferenceObject()) return unwrap(arg);
      const reference = arg as unknown as BaseReferenceObject;
      runtime ??= reference.getRuntimeData();
      return unwrap(reference.toArrayValueObject());
    });
    const result = callUdf(String(this.name), values, this.surveyContext(runtime));
    if (result.error === "VALUE") return ErrorValueObject.create(ErrorType.VALUE);
    if (result.error === "NUM") return ErrorValueObject.create(ErrorType.NUM);
    if (result.error === "NAME") return ErrorValueObject.create(ErrorType.NAME);
    if (typeof result.value === "number") return NumberValueObject.create(result.value);
    if (typeof result.value === "string") return StringValueObject.create(result.value);
    return StringValueObject.create("");
  }

  /**
   * The caller row's other columns (AI motor, AL/AM cumulative N/E) are read like VBA's
   * Application.Caller. Values computed earlier in this pass win over stored ones.
   */
  private surveyContext(runtime: IRuntimeUnitDataType | undefined): SurveyContext | undefined {
    try {
      const model = this._formulaDataModel;
      const unitId = this.unitId;
      if (!model || !unitId) return undefined;
      const info = this.getSheetsInfo();
      const sheetId = (name: string) => Object.entries(info.sheetNameMap).find(([, title]) => title === name)?.[0];
      const slidesheetId = sheetId("Slidesheet");
      const planId = sheetId("_OC_Survey");
      const unit = model.getCalculateData().allUnitData[unitId];
      const matrix = (id: string | undefined): CellMatrix | undefined => {
        if (!id) return undefined;
        const stored = unit?.[id]?.cellData as CellMatrix | undefined;
        const live = runtime?.[unitId]?.[id] as CellMatrix | undefined;
        if (!live) return stored;
        return { getValue: (row, column) => live.getValue(row, column) ?? stored?.getValue(row, column) };
      };
      const slidesheet = matrix(slidesheetId);
      const planSheet = matrix(planId);
      const grid = cachedGrid(unitId, slidesheet, planSheet);
      const surveyRow = (md: number) => {
        for (let row = 504; row >= 12; row -= 1) {
          const rowMd = numberAt(slidesheet, row, 4);
          if (rowMd == null || Math.abs(rowMd - md) >= 0.01) continue;
          return { cumN: numberAt(slidesheet, row, 37), cumE: numberAt(slidesheet, row, 38), motor: numberAt(slidesheet, row, 34) };
        }
        return null;
      };
      const shared = { plan: grid.plan, slides: grid.slides, surveyRow };
      if (info.sheetNameMap[this.subUnitId ?? ""] !== "Slidesheet" || !slidesheet) return shared;
      const row = this.row;
      const caller = {
        md: numberAt(slidesheet, row, 4) ?? 0,
        bitMd: numberAt(slidesheet, row, 3) ?? 0,
        inc: numberAt(slidesheet, row, 5) ?? 0,
        azm: numberAt(slidesheet, row, 6) ?? 0,
        tvd: numberAt(slidesheet, row, 7),
        cumN: numberAt(slidesheet, row, 37),
        cumE: numberAt(slidesheet, row, 38),
        motor: numberAt(slidesheet, row, 34),
        course: numberAt(slidesheet, row, 2),
      };
      if (caller.md === 0 && caller.bitMd === 0) return shared;
      return { ...shared, caller };
    } catch {
      return undefined;
    }
  }
}

function cachedGrid(unitId: string, slidesheet: CellMatrix | undefined, planSheet: CellMatrix | undefined): CachedGrid {
  const stamp = `${textAt(planSheet, 2, 0)}|${textAt(slidesheet, 12, 4)}|${textAt(slidesheet, 504, 17)}`;
  if (gridCache && gridCache.unitId === unitId && gridCache.stamp === stamp) return gridCache;
  const header: unknown[] = [];
  for (let column = 0; column < 20; column += 1) header.push(textAt(planSheet, 1, column));
  const rows: unknown[][] = [];
  const mdColumn = header.findIndex((cell) => String(cell).toUpperCase() === "MD");
  let blanks = 0;
  for (let row = 2; row < 5000; row += 1) {
    const md = textAt(planSheet, row, mdColumn);
    if (md == null || md === "") {
      blanks += 1;
      if (blanks > 30 && rows.length > 0) break;
      continue;
    }
    blanks = 0;
    const rowValues: unknown[] = [];
    for (let column = 0; column < 20; column += 1) rowValues.push(textAt(planSheet, row, column));
    rows.push(rowValues);
  }
  const slides: NonNullable<SurveyContext["slides"]> = [];
  for (let row = 12; row <= 504; row += 1) {
    const from = numberAt(slidesheet, row, 17);
    const to = numberAt(slidesheet, row, 18);
    if (from == null || to == null) continue;
    try {
      const tf = parseToolface(textAt(slidesheet, row, 20));
      if (tf == null) continue;
      slides.push({ from, to, tf, row: row + 1 });
    } catch {
      continue;
    }
  }
  gridCache = { unitId, stamp, plan: loadPlanGrid(header, rows), slides };
  return gridCache;
}

function numberAt(matrix: CellMatrix | undefined, row: number, column: number): number | null {
  const value = textAt(matrix, row, column);
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

function textAt(matrix: CellMatrix | undefined, row: number, column: number): unknown {
  if (!matrix || column < 0) return null;
  return matrix.getValue(row, column)?.v ?? null;
}

function unwrap(value: BaseValueObject): unknown {
  if (value.isError()) return ARG_ERROR;
  if (value.isNull()) return null;
  const arrayValue = value as BaseValueObject & { isArray?: () => boolean };
  if (typeof arrayValue.isArray === "function" && arrayValue.isArray()) {
    const rows = (value.getArrayValue() ?? []).map((row) => (row ?? []).map(cellValue));
    // A single-cell reference arrives as a 1x1 array; VBA sees it as the cell value.
    if (rows.length === 1 && rows[0]?.length === 1) return rows[0][0];
    return rows;
  }
  return value.getValue();
}

function cellValue(cell: Nullable<BaseValueObject>): unknown {
  if (cell == null || cell.isNull()) return null;
  if (cell.isError()) return ARG_ERROR;
  return cell.getValue();
}

/** Real directional functions live in the worker. Names still waiting on a port stay #NAME?. */
export function legacyUdfEntries(extraNames: Iterable<string> = []): Array<[Ctor<BaseFunction>, string]> {
  // `_xlfn.` / `_xlws.` are Excel's storage prefixes for built-ins, stripped on import.
  const names = new Set<string>([...LEGACY_UDF_NAMES, ...extraNames].filter((name) => !/^_XL(?:FN|WS)\./i.test(name)));
  return [...names].map((name) => {
    if (implementedUdf(name)) {
      const ctor = class extends DirectionalFunction {
        constructor() {
          super(name);
        }
      };
      return [ctor, name];
    }
    const ctor = class extends LegacyNameFunction {
      constructor() {
        super(name);
      }
    };
    return [ctor, name];
  });
}
