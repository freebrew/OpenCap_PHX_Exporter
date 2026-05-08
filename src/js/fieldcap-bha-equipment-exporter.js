/*
Paste this into Chrome DevTools Console while logged in to FieldCap.

It exports one flat CSV that aligns BHA/tool assembly information with each
component serial row. It also scrapes any visible BHAs / Tools table values
currently rendered in the page and keeps those UI values beside raw OData data.
*/
(async () => {
  const jobId = 20786;
  const outputFileName = `fieldcap-job-${jobId}-bha-equipment.csv`;
  const cacheKey = `fieldcap-job-${jobId}-visible-table-cache`;

  const normalize = (value) => String(value ?? "").replace(/\s+/g, " ").trim();

  const csvEscape = (value) => {
    const text = String(value ?? "");
    return /[",\r\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
  };

  const firstText = (...values) => {
    for (const value of values) {
      const text = normalize(value);
      if (text) return text;
    }
    return "";
  };

  const numberOrBlank = (value) => {
    const number = Number(value);
    return Number.isFinite(number) ? number : "";
  };

  const formatMinutes = (value) => {
    const rawMinutes = Number(value);
    if (!Number.isFinite(rawMinutes)) return "";

    const rounded = Math.trunc(rawMinutes);
    const hours = Math.trunc(rounded / 60);
    const minutes = Math.abs(rounded % 60);
    return `${hours}:${String(minutes).padStart(2, "0")}`;
  };

  const byPath = (value, path) => {
    return path.split(".").reduce((current, part) => {
      if (!current || typeof current !== "object") return undefined;
      return current[part];
    }, value);
  };

  const fetchOData = async (path) => {
    const rows = [];
    let nextUrl = path;

    while (nextUrl) {
      const response = await fetch(nextUrl, {
        credentials: "include",
        headers: { Accept: "application/json" },
      });

      if (!response.ok) {
        throw new Error(`FieldCap request failed: ${response.status} ${response.statusText} (${nextUrl})`);
      }

      const payload = await response.json();
      rows.push(...(Array.isArray(payload.value) ? payload.value : [payload]));
      nextUrl = payload["@odata.nextLink"] || "";
    }

    return rows;
  };

  const getCellText = (cell) => normalize(cell.innerText || cell.textContent);

  const getVisibleTables = () => {
    return [...document.querySelectorAll("table")]
      .map((table) => {
        const rows = [...table.querySelectorAll("tr")].map((row) => [...row.children].map(getCellText));
        const headerIndex = rows.findIndex((row) => row.some(Boolean));
        const headers = headerIndex >= 0 ? rows[headerIndex].map(normalize) : [];
        const body = headerIndex >= 0 ? rows.slice(headerIndex + 1) : [];
        const records = body
          .map((row) => {
            const record = {};
            headers.forEach((header, index) => {
              if (header) record[header] = row[index] ?? "";
            });
            return record;
          })
          .filter((record) => Object.values(record).some((value) => normalize(value)));

        return {
          headers,
          records,
          text: rows.flat().join(" | "),
        };
      })
      .filter(({ records }) => records.length > 0);
  };

  const hasHeaders = (table, headerPatterns) => {
    return headerPatterns.every((pattern) => table.headers.some((header) => pattern.test(header)));
  };

  const scrapeVisibleBhaRows = () => {
    const table = getVisibleTables().find((candidate) =>
      hasHeaders(candidate, [/^BHA$/i, /section/i, /total\s*hrs/i, /motor/i])
    );

    return table?.records ?? [];
  };

  const scrapeVisibleToolRows = () => {
    const table = getVisibleTables().find((candidate) =>
      hasHeaders(candidate, [/serial\s*#?/i, /job\s*hours/i, /hsls/i])
    );

    return table?.records ?? [];
  };

  const rowSignature = (row) => JSON.stringify(row);

  const mergeRows = (...rowGroups) => {
    const seen = new Set();
    return rowGroups.flat().filter((row) => {
      const signature = rowSignature(row);
      if (seen.has(signature)) return false;
      seen.add(signature);
      return true;
    });
  };

  const loadVisibleCache = () => {
    try {
      const parsed = JSON.parse(localStorage.getItem(cacheKey) || "{}");
      return {
        bhaRows: Array.isArray(parsed.bhaRows) ? parsed.bhaRows : [],
        toolRows: Array.isArray(parsed.toolRows) ? parsed.toolRows : [],
      };
    } catch {
      return { bhaRows: [], toolRows: [] };
    }
  };

  const saveVisibleCache = (cache) => {
    localStorage.setItem(cacheKey, JSON.stringify(cache));
  };

  const currentVisibleBhaRows = scrapeVisibleBhaRows();
  const currentVisibleToolRows = scrapeVisibleToolRows();
  const visibleCache = loadVisibleCache();
  const visibleBhaRows = mergeRows(visibleCache.bhaRows, currentVisibleBhaRows);
  const visibleToolRows = mergeRows(visibleCache.toolRows, currentVisibleToolRows);
  saveVisibleCache({ bhaRows: visibleBhaRows, toolRows: visibleToolRows });

  const visibleBhaByNumber = new Map(
    visibleBhaRows.map((row) => [firstText(row.BHA, row["BHA #"], row["Tool Assembly #"]), row])
  );

  const visibleToolsBySerial = new Map();
  for (const row of visibleToolRows) {
    const serial = firstText(row["Serial#"], row["Serial #"], row.Serial);
    if (!serial) continue;
    const matches = visibleToolsBySerial.get(serial) ?? [];
    matches.push(row);
    visibleToolsBySerial.set(serial, matches);
  }

  console.log(`[1/4] Fetching BHA list for job ${jobId}...`);
  const assemblies = await fetchOData(
    `/odata/ToolAssemblies?$filter=((null%20eq%20DeletedBy)%20and%20(ClientJobId%20eq%20${jobId}))&$orderby=ToolAssemblyNumber%20desc&$count=true`
  );

  console.log(`[2/4] Fetching BHA details, components, and custom values for ${assemblies.length} BHAs...`);
  const assemblyBundles = [];

  for (let index = 0; index < assemblies.length; index += 1) {
    const assembly = assemblies[index];
    const assemblyId = assembly.ToolAssemblyId;
    const assemblyNumber = firstText(assembly.ToolAssemblyNumber, assembly.BhaNumber, assembly.BHA);
    console.log(`[2/4] [${index + 1}/${assemblies.length}] BHA ${assemblyNumber || assemblyId}`);

    const [details, items, customValues] = await Promise.all([
      fetchOData(
        `/odata/ToolAssemblies?$expand=ClientJob&$filter=((ToolAssemblyId%20eq%20${assemblyId})%20and%20(null%20eq%20DeletedBy))&$top=1`
      ),
      fetchOData(
        `/odata/ToolAssemblyItems?$expand=JobTool,JobTool($expand=Item),JobTool($expand=ItemSerial)&$filter=((null%20eq%20DeletedBy)%20and%20(ToolAssemblyId%20eq%20${assemblyId}))`
      ),
      fetchOData(
        `/odata/ToolAssemblyCustomValues?$filter=((null%20eq%20DeletedBy)%20and%20(ToolAssemblyId%20eq%20${assemblyId}))`
      ),
    ]);

    assemblyBundles.push({
      assembly,
      details: details[0] ?? assembly,
      items,
      customValues,
    });
  }

  const customValueKey = (customValue, index) => {
    return firstText(
      customValue.CustomFieldName,
      customValue.FieldName,
      customValue.Name,
      customValue.Label,
      customValue.ToolAssemblyCustomValueName,
      customValue.CustomValueName,
      customValue.CustomType,
      customValue.CustomTypeId,
      `Custom ${index + 1}`
    );
  };

  const customValueText = (customValue) => {
    return firstText(
      customValue.Value,
      customValue.TextValue,
      customValue.StringValue,
      customValue.NumberValue,
      customValue.BooleanValue,
      customValue.DateValue,
      customValue.CustomValue,
      customValue.CustomText,
      customValue.CustomNumber
    );
  };

  const customMapFromValues = (customValues) => {
    const output = {};
    customValues.forEach((customValue, index) => {
      const key = `Custom:${customValueKey(customValue, index)}`;
      output[key] = customValueText(customValue);
    });
    return output;
  };

  const rawTime = (value) => {
    const text = firstText(value);
    if (!text) return "";

    const number = Number(text);
    if (Number.isFinite(number)) return formatMinutes(number);
    return text;
  };

  const rawBhaFields = (assembly, details) => ({
    "Raw Activated On": firstText(details.ActivatedOn, assembly.ActivatedOn),
    "Raw Completed On": firstText(details.CompletedOn, assembly.CompletedOn),
    "Raw Total Hrs": rawTime(firstText(details.TotalHours, details.TotalHrs, assembly.TotalHours, assembly.TotalHrs)),
    "Raw Below Rot": rawTime(firstText(details.BelowRot, details.BelowRotHours, assembly.BelowRot, assembly.BelowRotHours)),
    "Raw Hrs Slid": rawTime(firstText(details.HrsSlid, details.HoursSlid, assembly.HrsSlid, assembly.HoursSlid)),
    "Raw Hrs Rot": rawTime(firstText(details.HrsRot, details.HoursRot, assembly.HrsRot, assembly.HoursRot)),
    "Raw Hrs Circ": rawTime(firstText(details.HrsCirc, details.HoursCirc, assembly.HrsCirc, assembly.HoursCirc)),
  });

  const componentFields = (itemRow) => {
    const jobTool = itemRow.JobTool ?? {};
    const item = jobTool.Item ?? {};
    const serial = jobTool.ItemSerial ?? {};

    return {
      "Component Row #": firstText(itemRow.LineNumber, itemRow.SortOrder, itemRow.Order, itemRow.Sequence),
      "Serial #": firstText(serial.SerialNumber, serial.SerialNo, serial.Serial, serial.ItemSerialNumber, serial.Name),
      Description: firstText(
        itemRow.Description,
        itemRow.ToolDescription,
        jobTool.Description,
        item.ItemName,
        item.Name,
        item.Description
      ),
      "Sub Description": firstText(itemRow.SubDescription, itemRow.SizeDescription, itemRow.Size, jobTool.TempItemName),
      Item: firstText(item.ItemName, item.Name, item.Description, jobTool.ItemName),
      ItemSerialId: firstText(jobTool.ItemSerialId, serial.ItemSerialId),
      JobToolId: firstText(itemRow.JobToolId, jobTool.JobToolId),
      Top: firstText(itemRow.Top, itemRow.TopConnection, itemRow.TopSub),
      Bottom: firstText(itemRow.Bottom, itemRow.BottomConnection, itemRow.BottomSub),
      "Max OD": firstText(itemRow.MaxOD, itemRow.MaxOd, itemRow.OD, itemRow.OutsideDiameter),
      "Min ID": firstText(itemRow.MinID, itemRow.MinId, itemRow.ID, itemRow.InsideDiameter),
      Length: firstText(itemRow.Length),
      "Accum Length": firstText(itemRow.AccumLength, itemRow.AccumulatedLength),
    };
  };

  const visibleToolFields = (serial) => {
    const matches = visibleToolsBySerial.get(serial) ?? [];
    const row = matches[0] ?? {};

    return {
      "UI Tool Match Count": matches.length || "",
      "UI Tool Job Hours": firstText(row["Job Hours"]),
      "UI Tool HSLS": firstText(row.HSLS),
      "UI Tool Shipping Status": firstText(row["Shipping Status"]),
      Strapped: firstText(row.Strapped),
      "Dispatched On": firstText(row["Dispatched On"]),
      "Returned On": firstText(row["Returned On"]),
      "UI Tool Description": firstText(row.Description),
      "UI Tool Size/Description": firstText(row["Size/Description"]),
    };
  };

  console.log("[3/4] Joining BHA data with component rows...");
  const flatRows = assemblyBundles.flatMap(({ assembly, details, items, customValues }) => {
    const assemblyNumber = firstText(details.ToolAssemblyNumber, assembly.ToolAssemblyNumber, details.BHA, assembly.BHA);
    const visibleBha = visibleBhaByNumber.get(assemblyNumber) ?? {};
    const customFields = customMapFromValues(customValues);
    const base = {
      ClientJobId: firstText(details.ClientJobId, assembly.ClientJobId, details.ClientJob?.ClientJobId),
      "BHA #": assemblyNumber,
      ToolAssemblyId: firstText(details.ToolAssemblyId, assembly.ToolAssemblyId),
      Section: firstText(visibleBha.Section, details.Section, assembly.Section),
      Status: firstText(visibleBha.Status, details.Status, assembly.Status),
      Motor: firstText(visibleBha.Motor, details.Motor, assembly.Motor),
      Guidance: firstText(visibleBha.Guidance, details.Guidance, assembly.Guidance),
      "UI Hrs Slid": firstText(visibleBha["Hrs Slid"]),
      "UI Hrs Rot": firstText(visibleBha["Hrs Rot"]),
      "UI Hrs Circ": firstText(visibleBha["Hrs Circ"]),
      "UI Total Hrs": firstText(visibleBha["Total Hrs"]),
      "UI Below Rot": firstText(visibleBha["Below Rot"]),
      ...rawBhaFields(assembly, details),
      ...customFields,
    };

    const rows = items.length > 0 ? items : [{}];
    return rows.map((itemRow) => {
      const component = componentFields(itemRow);
      return {
        ...base,
        ...component,
        ...visibleToolFields(component["Serial #"]),
      };
    });
  });

  const preferredColumns = [
    "ClientJobId",
    "BHA #",
    "ToolAssemblyId",
    "Section",
    "Status",
    "Motor",
    "Guidance",
    "UI Hrs Slid",
    "UI Hrs Rot",
    "UI Hrs Circ",
    "UI Total Hrs",
    "UI Below Rot",
    "Raw Hrs Slid",
    "Raw Hrs Rot",
    "Raw Hrs Circ",
    "Raw Total Hrs",
    "Raw Below Rot",
    "Raw Activated On",
    "Raw Completed On",
    "Component Row #",
    "Serial #",
    "Description",
    "Sub Description",
    "Item",
    "ItemSerialId",
    "JobToolId",
    "Top",
    "Bottom",
    "Max OD",
    "Min ID",
    "Length",
    "Accum Length",
    "UI Tool Match Count",
    "UI Tool Job Hours",
    "UI Tool HSLS",
    "UI Tool Shipping Status",
    "Strapped",
    "Dispatched On",
    "Returned On",
    "UI Tool Description",
    "UI Tool Size/Description",
  ];

  const dynamicColumns = [
    ...new Set(flatRows.flatMap((row) => Object.keys(row)).filter((column) => !preferredColumns.includes(column))),
  ].sort();
  const columns = [...preferredColumns, ...dynamicColumns];

  console.log("[4/4] Building CSV download...");
  const csv = [
    columns.join(","),
    ...flatRows.map((row) => columns.map((column) => csvEscape(row[column])).join(",")),
  ].join("\r\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = outputFileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);

  const bhaFiveRows = flatRows.filter((row) => String(row["BHA #"]) === "5");
  const knownSerialRows = flatRows.filter((row) =>
    ["VC550213", "VG550005PTS", "VL550095", "VX550172", "VC550306"].includes(row["Serial #"])
  );

  console.table(flatRows.slice(0, 25));
  console.log(`Downloaded ${flatRows.length} component rows to ${outputFileName}`);
  console.log(`Visible BHAs scraped: ${visibleBhaRows.length}; visible Tools rows scraped: ${visibleToolRows.length}`);
  console.log("Validation preview: BHA #5 rows", bhaFiveRows.slice(0, 20));
  console.log("Validation preview: known serial rows", knownSerialRows);
})();
