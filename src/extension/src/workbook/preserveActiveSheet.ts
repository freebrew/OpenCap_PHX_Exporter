export interface ActiveSheetHost {
  activeSheet(): string;
  activateSheet(name: string): void;
}

/** Keep background workbook updates from navigating the operator away from the current tab. */
export async function preserveActiveSheet<T>(host: ActiveSheetHost, operation: () => Promise<T>): Promise<T> {
  const active = host.activeSheet();
  try {
    return await operation();
  } finally {
    if (active) host.activateSheet(active);
  }
}
