import ExcelJS from "exceljs";
import { ChartJSNodeCanvas } from "chartjs-node-canvas";
import type { ChartConfiguration } from "chart.js";
import type { Attachment } from "./ai.server";

function flattenNode(obj: Record<string, unknown>): Record<string, unknown> {
  const flat: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(obj)) {
    if (v && typeof v === "object" && !Array.isArray(v)) {
      const inner = v as Record<string, unknown>;
      if ("edges" in inner && Array.isArray(inner.edges)) {
        const firstNode = (inner.edges as Array<{ node?: unknown }>)[0]?.node;
        if (firstNode && typeof firstNode === "object") {
          for (const [nk, nv] of Object.entries(
            firstNode as Record<string, unknown>,
          )) {
            if (nv && typeof nv === "object" && "amount" in (nv as Record<string, unknown>)) {
              flat[nk] = (nv as Record<string, unknown>).amount;
            } else if (nv === null || typeof nv !== "object") {
              flat[nk] = nv;
            }
          }
        }
      } else if ("amount" in inner) {
        flat[k] = inner.amount;
      } else {
        for (const [nk, nv] of Object.entries(inner)) {
          if (nv === null || typeof nv !== "object") flat[nk] = nv;
        }
      }
    } else {
      flat[k] = v;
    }
  }
  return flat;
}

function normalizeRows(
  raw: unknown,
  columnKeys: string[],
): Array<Record<string, unknown>> {
  if (Array.isArray(raw)) {
    return raw.map((item) => {
      if (item && typeof item === "object") {
        const obj = item as Record<string, unknown>;
        if ("node" in obj && typeof obj.node === "object" && obj.node) {
          return flattenNode(obj.node as Record<string, unknown>);
        }
        return flattenNode(obj);
      }
      return { value: item };
    });
  }

  if (raw && typeof raw === "object") {
    const obj = raw as Record<string, unknown>;
    if ("edges" in obj && Array.isArray(obj.edges)) {
      return normalizeRows(obj.edges, columnKeys);
    }
    if ("data" in obj && typeof obj.data === "object" && obj.data) {
      const data = obj.data as Record<string, unknown>;
      const firstKey = Object.keys(data)[0];
      if (firstKey) return normalizeRows(data[firstKey], columnKeys);
    }
    const vals = Object.values(obj);
    if (vals.length > 0 && vals.every((v) => typeof v === "object" && v))
      return vals.map((v) => flattenNode(v as Record<string, unknown>));
  }

  return [];
}

function safeFilename(title: string): string {
  return title.replace(/[^a-zA-Z0-9\u4e00-\u9fff_-]/g, "_").slice(0, 60);
}

export async function generateChart(args: {
  title: string;
  chartType: string;
  labels: string[];
  values: number[];
  valueLabel: string;
}): Promise<{ result: unknown; attachment: Attachment }> {
  const { title, labels, values, valueLabel } = args;
  const chartType = (
    args.chartType === "bar" ? "bar" : args.chartType === "pie" ? "pie" : "line"
  ) as "bar" | "pie" | "line";

  const chartConfig: ChartConfiguration = {
    type: chartType,
    data: {
      labels,
      datasets: [
        {
          label: valueLabel,
          data: values,
          backgroundColor:
            chartType === "pie"
              ? [
                  "#4472C4", "#5B9BD5", "#70AD47", "#FFC000", "#ED7D31",
                  "#A5A5A5", "#FF5050", "#7030A0", "#00B0F0", "#92D050",
                  "#FFD966", "#C55A11",
                ]
              : "rgba(68, 114, 196, 0.8)",
          borderColor:
            chartType === "pie" ? "white" : "rgba(68, 114, 196, 1)",
          borderWidth: 1,
          fill: chartType === "line",
        },
      ],
    },
    options: {
      animation: false,
      responsive: false,
      plugins: { legend: { position: "bottom" } },
      scales: chartType !== "pie" ? { y: { beginAtZero: true } } : undefined,
    },
    plugins: [
      {
        id: "white-bg",
        beforeDraw: (chart) => {
          const ctx = chart.ctx;
          ctx.save();
          ctx.fillStyle = "white";
          ctx.fillRect(0, 0, 640, 360);
          ctx.restore();
        },
      },
    ],
  };

  const canvas = new ChartJSNodeCanvas({
    width: 640,
    height: 360,
    backgroundColour: "white",
  });
  const chartBuffer = await canvas.renderToBuffer(chartConfig);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "AI Report Assistant";
  const sheet = workbook.addWorksheet(title);

  sheet.columns = [
    { header: "类别", key: "label", width: 22 },
    { header: valueLabel, key: "value", width: 18 },
  ];
  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
  headerRow.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF4472C4" },
  };

  labels.forEach((label, i) =>
    sheet.addRow({ label, value: values[i] ?? 0 }),
  );

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
    result: { success: true, filename },
    attachment: {
      base64: Buffer.from(buf as ArrayBuffer).toString("base64"),
      filename,
      mimeType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    },
  };
}

export async function generateExcel(args: {
  title: string;
  columns: Array<{ header: string; key: string; width?: number }>;
  rows: unknown;
}): Promise<{ result: unknown; attachment: Attachment }> {
  const { title, columns } = args;

  const rows = normalizeRows(args.rows, columns.map((c) => c.key));

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "AI Report Assistant";
  const sheet = workbook.addWorksheet(title);

  sheet.columns = columns.map((c) => ({
    header: c.header,
    key: c.key,
    width: c.width ?? 18,
  }));
  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
  headerRow.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF4472C4" },
  };

  for (const row of rows) sheet.addRow(row);

  const buf = await workbook.xlsx.writeBuffer();
  const filename = `report_${safeFilename(title)}_${new Date().toISOString().slice(0, 10)}.xlsx`;

  return {
    result: { success: true, filename, rowCount: rows.length },
    attachment: {
      base64: Buffer.from(buf as ArrayBuffer).toString("base64"),
      filename,
      mimeType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    },
  };
}
