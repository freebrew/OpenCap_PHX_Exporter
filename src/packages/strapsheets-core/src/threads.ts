// Thread types offered on the BOX PIN Thread Type step. The sheet's FormOptions!A2:A50 is the
// live list; this copy (src/python/maintenance/update_threads.py) is what a phone with no sheet
// access shows.

export const DEFAULT_THREADS: string[] = [
  "3 1/2 IF [NC38]",
  "3 1/2 REG",
  "DS 38",
  "DS 40",
  "4 FH [NC40]",
  "4 IF [NC 44]",
  "4 1/2 REG",
  "4 1/2 IF [NC 50]",
  "4 1/2 XH [NC46]",
  "5 1/2 FH",
  "6 5/8 REG",
  "CET43",
  "HLNST-43",
  "DUO400",
];

/**
 * Second names for threads the sheet lists without one. DS connections make up with the API
 * numbered connection of the same size; the premium connections have no API twin, so their maker
 * is shown. REG and 5 1/2 FH have no NC equivalent and get none.
 */
const THREAD_ALIASES: Record<string, string> = {
  "DS 38": "NC38",
  "DS 40": "NC40",
  CET43: "Command",
  "HLNST-43": "Hilong",
  DUO400: "Complete Tubular",
};

/** A thread option split for display: "4 FH [NC40]" is 4 FH over NC40. The stored value is unchanged. */
export function threadParts(option: string): { name: string; alias: string | null } {
  const bracketed = /^(.*?)\s*\[([^\]]*)\]$/.exec(option);
  if (bracketed) return { name: bracketed[1] ?? option, alias: bracketed[2] ?? null };
  return { name: option, alias: THREAD_ALIASES[option] ?? null };
}
