/*
Paste this into Chrome DevTools Console on the logged-in FieldCap page.
It prints every field path containing numbers for one serial so we can find
which raw field maps to the UI's Job Hours / HSLS display.
*/
(async () => {
  const jobId = 20786;
  const targetSerial = "VC550306";
  const endpoint = `/odata/ToolAssemblyItems?$expand=JobTool,JobTool($expand=Item),JobTool($expand=ItemSerial),ToolAssembly&$filter=((null%20eq%20DeletedBy)%20and%20(ToolAssembly/ClientJobId%20eq%20${jobId}))`;

  const response = await fetch(endpoint, {
    credentials: "include",
    headers: { Accept: "application/json" },
  });

  if (!response.ok) {
    throw new Error(`FieldCap request failed: ${response.status} ${response.statusText}`);
  }

  const payload = await response.json();
  const rows = Array.isArray(payload.value) ? payload.value : [];

  const walk = (value, prefix = "") => {
    if (!value || typeof value !== "object" || Array.isArray(value)) return [];

    return Object.entries(value).flatMap(([key, child]) => {
      const path = prefix ? `${prefix}.${key}` : key;
      if (child && typeof child === "object" && !Array.isArray(child)) {
        return walk(child, path);
      }

      return [{ path, value: child }];
    });
  };

  const formatMinutes = (value) => {
    const rawMinutes = Number(value);
    if (!Number.isFinite(rawMinutes)) return "";

    const rounded = Math.trunc(rawMinutes);
    const hours = Math.trunc(rounded / 60);
    const minutes = Math.abs(rounded % 60);
    return `${hours}:${String(minutes).padStart(2, "0")}`;
  };

  const serialMatches = rows.filter((row) => {
    const serial = row.JobTool?.ItemSerial?.SerialNumber ?? row.JobTool?.ItemSerial?.Name;
    return serial === targetSerial;
  });

  if (serialMatches.length === 0) {
    console.warn(`No rows found for serial ${targetSerial}`);
    return;
  }

  serialMatches.forEach((row, rowIndex) => {
    const jobTool = row.JobTool ?? {};
    const item = jobTool.Item ?? {};
    const itemSerial = jobTool.ItemSerial ?? {};
    const itemName = item.ItemName ?? item.Name ?? item.Description ?? jobTool.ItemName ?? "";
    const serial = itemSerial.SerialNumber ?? itemSerial.SerialNo ?? itemSerial.Serial ?? itemSerial.Name ?? "";
    const toolAssemblyId = row.ToolAssemblyId ?? row.ToolAssembly?.ToolAssemblyId ?? "";
    const toolAssemblyNumber = row.ToolAssembly?.ToolAssemblyNumber ?? "";

    const fields = walk(row)
      .filter(({ value }) => Number.isFinite(Number(value)))
      .map(({ path, value }) => ({
        row: rowIndex + 1,
        item: itemName,
        serial,
        toolAssemblyId,
        toolAssemblyNumber,
        path,
        raw: Number(value),
        asHours: formatMinutes(value),
      }))
      .filter(({ raw, path }) => raw !== 0 || /hour|hsls|hsl|total/i.test(path));

    const jobHourBuckets = fields.filter(({ path }) => /JobTool\.JobHours\d+$/i.test(path));
    const jobHourTotal = jobHourBuckets.reduce((sum, field) => sum + field.raw, 0);

    console.group(`Numeric fields for ${serial || targetSerial} row ${rowIndex + 1}`);
    console.log({
      item: itemName,
      serial,
      toolAssemblyId,
      toolAssemblyNumber,
      jobHourBucketTotalRaw: jobHourTotal,
      jobHourBucketTotal: formatMinutes(jobHourTotal),
    });
    console.table(fields);
    console.groupEnd();

    const exactUiMatches = fields.filter(({ raw }) => raw === 3985);
    if (exactUiMatches.length > 0) {
      console.log(`Fields equal to 3985 minutes (66:25) for ${targetSerial}:`);
      console.table(exactUiMatches);
    }
  });
})();
