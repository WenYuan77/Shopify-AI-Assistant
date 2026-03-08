import ExcelJS from "exceljs";
import { ChartJSNodeCanvas } from "chartjs-node-canvas";
import type { ChartConfiguration } from "chart.js";
import type { Attachment } from "./ai.server";

type GqlFn = (
  q: string,
  o?: { variables?: Record<string, unknown> },
) => Promise<Response>;

async function fetchAllNodes(
  admin: { graphql: GqlFn },
  query: string,
  variables: Record<string, unknown>,
  dataPath: string,
): Promise<Array<Record<string, unknown>>> {
  const allNodes: Array<Record<string, unknown>> = [];
  let cursor: string | null = null;

  do {
    if (cursor) variables.after = cursor;
    let resp: Response;
    try {
      resp = await admin.graphql(query, { variables });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      console.error(`[Fetch] GraphQL threw:`, msg);
      break;
    }
    const json = (await resp.json()) as Record<string, unknown>;
    if (json.errors) {
      console.error(`[Fetch] GraphQL errors:`, JSON.stringify(json.errors).slice(0, 300));
    }
    const nodes = extractRows(json, dataPath);
    allNodes.push(...nodes);

    const data = json.data as Record<string, unknown> | undefined;
    const resource = data ? getNestedValue(data, dataPath) : undefined;
    const pageInfo =
      resource && typeof resource === "object"
        ? ((resource as Record<string, unknown>).pageInfo as {
            hasNextPage?: boolean;
            endCursor?: string;
          })
        : undefined;
    cursor = pageInfo?.hasNextPage ? (pageInfo.endCursor ?? null) : null;
  } while (cursor);

  return allNodes;
}

function safeFilename(t: string): string {
  return t.replace(/[^a-zA-Z0-9\u4e00-\u9fff_-]/g, "_").slice(0, 60);
}

function getNestedValue(obj: unknown, path: string): unknown {
  return path.split(".").reduce((cur, key) => {
    if (cur == null) return undefined;
    if (typeof cur !== "object") return undefined;
    return (cur as Record<string, unknown>)[key];
  }, obj);
}

function extractRows(
  gqlResponse: Record<string, unknown>,
  dataPath: string,
): Array<Record<string, unknown>> {
  const data = gqlResponse.data as Record<string, unknown> | undefined;
  if (!data) return [];

  const resource = getNestedValue(data, dataPath);
  if (!resource || typeof resource !== "object") return [];

  const edges = (resource as Record<string, unknown>).edges;
  if (!Array.isArray(edges)) return [];

  return edges.map((e: { node?: Record<string, unknown> }) => e.node ?? {});
}

function resolveField(node: Record<string, unknown>, key: string): unknown {
  const parts = key.split(".");
  let val: unknown = node;
  for (const p of parts) {
    if (val == null || typeof val !== "object") return "";
    val = (val as Record<string, unknown>)[p];
  }
  if (val && typeof val === "object" && "amount" in (val as Record<string, unknown>)) {
    return (val as Record<string, unknown>).amount;
  }
  return val ?? "";
}

/* ── export_to_excel ── */

export interface RowFilter {
  field: string;
  operator: "eq" | "ne" | "lt" | "lte" | "gt" | "gte" | "contains";
  value: unknown;
}

function applyFilters(
  nodes: Array<Record<string, unknown>>,
  filters?: RowFilter[],
): Array<Record<string, unknown>> {
  if (!filters?.length) return nodes;
  return nodes.filter((node) =>
    filters.every((f) => {
      const raw = resolveField(node, f.field);
      const val = typeof raw === "string" && !isNaN(Number(raw)) ? Number(raw) : raw;
      const cmp = f.value;
      switch (f.operator) {
        case "eq":  return val == cmp;
        case "ne":  return val != cmp;
        case "lt":  return Number(val) <  Number(cmp);
        case "lte": return Number(val) <= Number(cmp);
        case "gt":  return Number(val) >  Number(cmp);
        case "gte": return Number(val) >= Number(cmp);
        case "contains":
          return String(val).toLowerCase().includes(String(cmp).toLowerCase());
        default: return true;
      }
    }),
  );
}

export async function exportToExcel(
  admin: { graphql: GqlFn },
  args: {
    query: string;
    variables?: Record<string, unknown>;
    title: string;
    columns: Array<{ header: string; key: string; width?: number }>;
    dataPath: string;
    rowFilters?: RowFilter[];
  },
): Promise<{ result: unknown; attachment: Attachment }> {
  const { query, title, columns, dataPath } = args;
  const variables = args.variables ?? {};
  if (/\$first/i.test(query) && !variables.first) variables.first = 250;

  console.log(`[Export Excel] query:`, query.slice(0, 300));
  console.log(`[Export Excel] vars:`, JSON.stringify(variables));
  console.log(`[Export Excel] filters:`, JSON.stringify(args.rowFilters ?? []));

  const allNodes = await fetchAllNodes(admin, query, variables, dataPath);

  const filtered = applyFilters(allNodes, args.rowFilters);
  console.log(`[Export Excel] ${allNodes.length} total, ${filtered.length} after filters`);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "AI Report Assistant";
  const sheet = workbook.addWorksheet(title);

  sheet.columns = columns.map((c) => ({
    header: c.header,
    key: c.key,
    width: c.width ?? 18,
  }));
  const hdr = sheet.getRow(1);
  hdr.font = { bold: true, color: { argb: "FFFFFFFF" } };
  hdr.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4472C4" } };

  for (const node of filtered) {
    const row: Record<string, unknown> = {};
    for (const col of columns) {
      row[col.key] = resolveField(node, col.key);
    }
    sheet.addRow(row);
  }

  const buf = await workbook.xlsx.writeBuffer();
  const filename = `report_${safeFilename(title)}_${new Date().toISOString().slice(0, 10)}.xlsx`;

  return {
    result: { success: true, filename, rowCount: filtered.length },
    attachment: {
      base64: Buffer.from(buf as ArrayBuffer).toString("base64"),
      filename,
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    },
  };
}

/* ── export_chart ── */

export async function exportChart(
  admin: { graphql: GqlFn },
  args: {
    query: string;
    variables?: Record<string, unknown>;
    title: string;
    chartType: string;
    dataPath: string;
    labelField: string;
    valueField: string;
    valueLabel: string;
    groupBy?: string;
  },
): Promise<{ result: unknown; attachment: Attachment }> {
  const { query, title, dataPath, labelField, valueField, valueLabel } = args;
  const variables = args.variables ?? {};
  if (/\$first/i.test(query) && !variables.first) variables.first = 250;
  const chartType = (
    args.chartType === "bar" ? "bar" : args.chartType === "pie" ? "pie" : "line"
  ) as "bar" | "pie" | "line";

  const allNodes = await fetchAllNodes(admin, query, variables, dataPath);
  console.log(`[Chart] ${allNodes.length} nodes fetched`);

  let labels: string[];
  let values: number[];

  if (args.groupBy === "month") {
    const monthNames = ["1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"];
    const byMonth: Record<number, number> = {};
    for (let i = 1; i <= 12; i++) byMonth[i] = 0;
    for (const node of allNodes) {
      const dateVal = resolveField(node, labelField);
      const val = Number(resolveField(node, valueField)) || 0;
      if (dateVal) {
        const month = new Date(String(dateVal)).getMonth() + 1;
        if (month >= 1 && month <= 12) byMonth[month] += val || 1;
      }
    }
    labels = monthNames;
    values = monthNames.map((_, i) => byMonth[i + 1]);
  } else {
    labels = allNodes.map((n) => String(resolveField(n, labelField) ?? ""));
    values = allNodes.map((n) => Number(resolveField(n, valueField)) || 0);
  }

  console.log(`[Chart] ${allNodes.length} nodes, ${labels.length} data points`);

  const chartConfig: ChartConfiguration = {
    type: chartType,
    data: {
      labels,
      datasets: [{
        label: valueLabel,
        data: values,
        backgroundColor: chartType === "pie"
          ? ["#4472C4","#5B9BD5","#70AD47","#FFC000","#ED7D31","#A5A5A5","#FF5050","#7030A0","#00B0F0","#92D050","#FFD966","#C55A11"]
          : "rgba(68,114,196,0.8)",
        borderColor: chartType === "pie" ? "white" : "rgba(68,114,196,1)",
        borderWidth: 1,
        fill: chartType === "line",
      }],
    },
    options: {
      animation: false,
      responsive: false,
      plugins: { legend: { position: "bottom" } },
      scales: chartType !== "pie" ? { y: { beginAtZero: true } } : undefined,
    },
    plugins: [{
      id: "white-bg",
      beforeDraw: (chart) => {
        const ctx = chart.ctx;
        ctx.save();
        ctx.fillStyle = "white";
        ctx.fillRect(0, 0, 640, 360);
        ctx.restore();
      },
    }],
  };

  const canvas = new ChartJSNodeCanvas({ width: 640, height: 360, backgroundColour: "white" });
  const chartBuffer = await canvas.renderToBuffer(chartConfig);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "AI Report Assistant";
  const sheet = workbook.addWorksheet(title);

  sheet.columns = [
    { header: "类别", key: "label", width: 22 },
    { header: valueLabel, key: "value", width: 18 },
  ];
  const hdr = sheet.getRow(1);
  hdr.font = { bold: true, color: { argb: "FFFFFFFF" } };
  hdr.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4472C4" } };
  labels.forEach((l, i) => sheet.addRow({ label: l, value: values[i] ?? 0 }));

  const imageId = workbook.addImage({
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    buffer: Buffer.from(chartBuffer) as any,
    extension: "png",
  });
  sheet.addImage(imageId, {
    tl: { col: 0, row: labels.length + 2 },
    ext: { width: 640, height: 360 },
    editAs: "oneCell",
  });

  const buf = await workbook.xlsx.writeBuffer();
  const filename = `chart_${safeFilename(title)}_${new Date().toISOString().slice(0, 10)}.xlsx`;

  return {
    result: { success: true, filename, dataPoints: labels.length },
    attachment: {
      base64: Buffer.from(buf as ArrayBuffer).toString("base64"),
      filename,
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    },
  };
}
