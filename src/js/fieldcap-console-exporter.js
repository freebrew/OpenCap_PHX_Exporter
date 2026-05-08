/*
Paste this whole file into the Chrome DevTools Console while you are logged in
at https://fieldcap-cdn.phxtech.com/#/Jobs/20786?t=9.

It downloads a CSV using the current Chrome session cookies.
*/
(async () => {
  const jobId = 20786;
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

  const fieldNames = (value, prefix = "") => {
    if (!value || typeof value !== "object" || Array.isArray(value)) return [];

    return Object.entries(value).flatMap(([key, child]) => {
      const path = prefix ? `${prefix}.${key}` : key;
      if (child && typeof child === "object" && !Array.isArray(child)) {
        return [path, ...fieldNames(child, path)];
      }
      return [path];
    });
  };

  const valueAtPath = (value, path) => {
    return path.split(".").reduce((current, part) => {
      if (!current || typeof current !== "object") return undefined;
      return current[part];
    }, value);
  };

  const firstText = (...values) => {
    for (const value of values) {
      if (value !== undefined && value !== null && String(value).trim() !== "") {
        return String(value);
      }
    }
    return "";
  };

  const firstNumber = (...values) => {
    for (const value of values) {
      const number = Number(value);
      if (Number.isFinite(number) && number !== 0) return number;
    }
    return null;
  };

  const formatMinutes = (value) => {
    const rawMinutes = Number(value);
    if (!Number.isFinite(rawMinutes) || rawMinutes === 0) return "0:00";

    const rounded = Math.trunc(rawMinutes);
    const hours = Math.trunc(rounded / 60);
    const minutes = Math.abs(rounded % 60);
    return `${hours}:${String(minutes).padStart(2, "0")}`;
  };

  const firstNumberByExactName = (row, names) => {
    return firstNumber(...names.map((name) => valueAtPath(row, name)));
  };

  const sumNumbersByExactName = (row, names) => {
    const values = names
      .map((name) => Number(valueAtPath(row, name)))
      .filter((value) => Number.isFinite(value));

    if (values.length === 0) return null;
    return values.reduce((sum, value) => sum + value, 0);
  };

  const firstNumberByPathTokens = (row, tokenGroups) => {
    const paths = fieldNames(row);

    for (const tokens of tokenGroups) {
      const match = paths.find((path) => {
        const normalized = path.toLowerCase();
        return tokens.every((token) => normalized.includes(token.toLowerCase()));
      });

      if (match) {
        const value = firstNumber(valueAtPath(row, match));
        if (value !== null) return value;
      }
    }

    return null;
  };

  const shippingStatus = (jobTool) => {
    const explicitStatus = firstText(
      jobTool?.ShippingStatus,
      jobTool?.ShippingStatusName,
      jobTool?.Status
    );

    if (explicitStatus) return explicitStatus;
    if (jobTool?.TransferInDate && !jobTool?.TransferOutDate) return "On Location";
    if (jobTool?.TransferOutDate) return "Transferred Out";
    return "";
  };

  const exportedRows = rows.map((row) => {
    const jobTool = row.JobTool ?? {};
    const item = jobTool.Item ?? {};
    const itemSerial = jobTool.ItemSerial ?? {};

    const jobHourBucketMinutes = sumNumbersByExactName(row, [
        "JobTool.JobHours1",
        "JobTool.JobHours2",
        "JobTool.JobHours3",
        "JobTool.JobHours4",
        "JobTool.JobHours5",
        "JobTool.JobHours6",
        "JobTool.JobHours7",
        "JobTool.JobHours8",
        "JobTool.JobHours9",
        "JobTool.JobHours10",
    ]);

    const jobMinutes =
      jobHourBucketMinutes ??
      firstNumberByExactName(row, [
        "JobHours",
        "TotalHours",
        "ToolHours",
        "JobTool.JobHours",
        "JobTool.TotalHours",
        "JobTool.ToolHours",
      ]) ??
      firstNumberByPathTokens(row, [
        ["job", "hour"],
        ["total", "hour"],
        ["tool", "hour"],
      ]);

    const hslsBucketMinutes = sumNumbersByExactName(row, [
      "JobTool.TransferInHsls1",
      "JobTool.TransferInHsls2",
      "JobTool.TransferInHsls3",
      "JobTool.TransferInHsls4",
      "JobTool.TransferInHsls5",
      "JobTool.TransferInHsls6",
      "JobTool.TransferInHsls7",
      "JobTool.TransferInHsls8",
      "JobTool.TransferInHsls9",
      "JobTool.TransferInHsls10",
      "JobTool.TransferOutHsls1",
      "JobTool.TransferOutHsls2",
      "JobTool.TransferOutHsls3",
      "JobTool.TransferOutHsls4",
      "JobTool.TransferOutHsls5",
      "JobTool.TransferOutHsls6",
      "JobTool.TransferOutHsls7",
      "JobTool.TransferOutHsls8",
      "JobTool.TransferOutHsls9",
      "JobTool.TransferOutHsls10",
    ]);

    const hslsMinutes =
      hslsBucketMinutes ??
      firstNumberByExactName(row, [
        "HSLS",
        "Hsls",
        "HslsHours",
        "TotalHSLS",
        "TotalHsls",
        "JobTool.HSLS",
        "JobTool.Hsls",
        "JobTool.HslsHours",
        "JobTool.TotalHSLS",
        "JobTool.TotalHsls",
      ]) ?? firstNumberByPathTokens(row, [["hsls"], ["hsl"]]);

    return {
      Item: firstText(item.ItemName, item.Name, item.Description, jobTool.ItemName),
      "Serial #": firstText(
        itemSerial.SerialNumber,
        itemSerial.SerialNo,
        itemSerial.Serial,
        itemSerial.ItemSerialNumber,
        itemSerial.Name
      ),
      "Tool Assembly ID": firstText(row.ToolAssemblyId, row.ToolAssembly?.ToolAssemblyId),
      "Tool Assembly #": firstText(row.ToolAssembly?.ToolAssemblyNumber),
      "Shipping Status": shippingStatus(jobTool),
      "Job Hours": formatMinutes(jobMinutes),
      HSLS: formatMinutes(hslsMinutes),
      "Job Hours Raw Minutes": jobMinutes ?? 0,
      "HSLS Raw Minutes": hslsMinutes ?? 0,
    };
  });

  const csvEscape = (value) => {
    const text = String(value ?? "");
    return /[",\r\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
  };

  const columns = [
    "Item",
    "Serial #",
    "Tool Assembly ID",
    "Tool Assembly #",
    "Shipping Status",
    "Job Hours",
    "HSLS",
    "Job Hours Raw Minutes",
    "HSLS Raw Minutes",
  ];

  const csv = [
    columns.join(","),
    ...exportedRows.map((row) => columns.map((column) => csvEscape(row[column])).join(",")),
  ].join("\r\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `fieldcap-job-${jobId}-tools.csv`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);

  console.table(exportedRows.slice(0, 20));
  console.log(`Downloaded ${exportedRows.length} rows to fieldcap-job-${jobId}-tools.csv`);
})();
