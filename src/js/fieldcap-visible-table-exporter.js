/*
Paste this into Chrome DevTools Console while the FieldCap Tools table is visible.
It exports the table exactly as rendered in the browser, so values like Job Hours
and HSLS match the UI instead of re-deriving them from OData fields.
*/
(() => {
  const normalize = (value) => String(value ?? "").replace(/\s+/g, " ").trim();

  const csvEscape = (value) => {
    const text = String(value ?? "");
    return /[",\r\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
  };

  const getCellText = (cell) => normalize(cell.innerText || cell.textContent);

  const tables = [...document.querySelectorAll("table")];
  const tableCandidates = tables
    .map((table) => {
      const rows = [...table.querySelectorAll("tr")].map((row) => [...row.children].map(getCellText));
      const flatText = rows.flat().join(" | ");
      const score = ["Serial", "Job Hours", "HSLS", "Shipping Status"].reduce(
        (sum, token) => sum + (flatText.toLowerCase().includes(token.toLowerCase()) ? 1 : 0),
        0
      );

      return { table, rows, score };
    })
    .filter(({ rows, score }) => rows.length > 1 && score >= 2)
    .sort((a, b) => b.score - a.score || b.rows.length - a.rows.length);

  if (tableCandidates.length === 0) {
    throw new Error("Could not find a rendered table with Serial / Job Hours / HSLS headers.");
  }

  const { rows } = tableCandidates[0];
  const headerIndex = rows.findIndex((row) =>
    row.some((cell) => /serial/i.test(cell)) &&
    row.some((cell) => /job\s*hours/i.test(cell))
  );

  if (headerIndex < 0) {
    throw new Error("Found a table, but could not identify its header row.");
  }

  const headers = rows[headerIndex].map((header) => normalize(header));
  const bodyRows = rows
    .slice(headerIndex + 1)
    .map((row) => {
      const output = {};
      headers.forEach((header, index) => {
        if (header) output[header] = row[index] ?? "";
      });
      return output;
    })
    .filter((row) => Object.values(row).some((value) => normalize(value)));

  const preferredColumns = [
    "Description",
    "Size/Description",
    "Serial#",
    "Serial #",
    "Strapped",
    "Dispatched On",
    "Returned On",
    "Shipping Status",
    "Job Hours",
    "HSLS",
  ];

  const columns = [
    ...preferredColumns.filter((column) => headers.includes(column)),
    ...headers.filter((header) => header && !preferredColumns.includes(header)),
  ];

  const csv = [
    columns.join(","),
    ...bodyRows.map((row) => columns.map((column) => csvEscape(row[column])).join(",")),
  ].join("\r\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "fieldcap-visible-tools-table.csv";
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);

  console.table(bodyRows.slice(0, 30));
  console.log(`Downloaded ${bodyRows.length} visible table rows to fieldcap-visible-tools-table.csv`);
})();
