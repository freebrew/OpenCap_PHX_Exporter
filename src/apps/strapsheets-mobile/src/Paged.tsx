// A list that never scrolls: it shows as many fixed-height rows as fit the space it is given and
// pages through the rest.

import { useLayoutEffect, useRef, useState, type ReactNode } from "react";

interface PagedProps<T> {
  items: T[];
  rowHeight: number;
  keyOf: (item: T) => string;
  render: (item: T) => ReactNode;
  empty?: ReactNode;
}

export function Paged<T>({ items, rowHeight, keyOf, render, empty }: PagedProps<T>) {
  const rows = useRef<HTMLDivElement>(null);
  const [fit, setFit] = useState(1);
  const [page, setPage] = useState(0);

  useLayoutEffect(() => {
    const element = rows.current;
    if (!element) return;
    const observer = new ResizeObserver(([entry]) => {
      if (entry) setFit(Math.max(1, Math.floor(entry.contentRect.height / rowHeight)));
    });
    observer.observe(element);
    return () => observer.disconnect();
  }, [rowHeight]);

  const pages = Math.max(1, Math.ceil(items.length / fit));
  const current = Math.min(page, pages - 1);
  return (
    <div className="paged">
      <div className="paged-rows" ref={rows}>
        {items.length === 0 && empty}
        {items.slice(current * fit, current * fit + fit).map((item) => (
          <div key={keyOf(item)} className="paged-row" style={{ height: rowHeight }}>
            {render(item)}
          </div>
        ))}
      </div>
      {pages > 1 && (
        <div className="pager">
          <button type="button" disabled={current === 0} onClick={() => setPage(current - 1)}>‹ Prev</button>
          <span>{current + 1} / {pages}</span>
          <button type="button" disabled={current >= pages - 1} onClick={() => setPage(current + 1)}>Next ›</button>
        </div>
      )}
    </div>
  );
}
