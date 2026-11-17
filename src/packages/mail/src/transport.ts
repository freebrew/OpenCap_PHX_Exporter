export type TransportResult =
  | { status: "created-and-opened"; id?: string; location: string }
  | { status: "created-not-opened"; id: string; location: string; error: string }
  | { status: "no-side-effect-failure"; error: string }
  | { status: "indeterminate"; error: string };

/** A second transport is allowed only when the first one is known not to have created mail. */
export function mayFallback(result: TransportResult): boolean {
  return result.status === "no-side-effect-failure";
}
