// The job's imported serials under the Serial field. Typing filters the list; tapping one prefills
// the entry. Serials with no strap sheet yet come first; strapped ones are ticked. It is a list,
// not a lock: a serial that is not here can still be typed.

import { filterItems, normSerial, type InventoryItem } from "../../../packages/strapsheets-core/src/index";

const SHOWN = 4;

export function SerialPicker({ items, measured, query, onPick }: { items: InventoryItem[]; measured: Set<string>; query: string; onPick: (item: InventoryItem) => void }) {
  const done = (item: InventoryItem) => measured.has(normSerial(item.serialNumber));
  const left = items.filter((item) => !done(item)).length;
  const matches = filterItems(items, query).sort((a, b) => Number(done(a)) - Number(done(b)));
  const summary = <div className="hint">{left} of {items.length} job tools still need a strap sheet.</div>;
  if (matches.length === 0) return <>{summary}<div className="hint">No imported serial matches. The typed serial will be used.</div></>;
  return (
    <>
      {summary}
      <ul className="serial-list">
        {matches.slice(0, SHOWN).map((item) => (
          <li key={`${item.itemCode}-${item.serialNumber}`}>
            <button type="button" className={done(item) ? "strapped" : ""} onClick={() => onPick(item)}>
              <span className="serial">{item.serialNumber}{done(item) && <span className="tick"> ✓ strapped</span>}</span>
              <span className="meta">
                {item.itemName}
                {item.category && ` · ${item.category}`}
                {item.status && ` · ${item.status}`}
                {item.runHours != null && ` · ${item.runHours} h`}
              </span>
            </button>
          </li>
        ))}
        {matches.length > SHOWN && <li className="hint">{matches.length - SHOWN} more; keep typing to narrow.</li>}
      </ul>
    </>
  );
}
